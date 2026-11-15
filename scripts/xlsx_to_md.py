#!/usr/bin/env python3
"""xlsx_to_md.py - Excel workbook -> AI-readable Markdown, with images put back where they live.

What it extracts
  * Every sheet as a Markdown table (and a CSV), full size, no truncation.
  * In-cell images, the kind that sit *inside* a cell:
      - WPS Office "embed in cell" images: the cell holds =DISPIMG("ID_...",1),
        the picture lives in xl/cellimages.xml.
      - Microsoft 365 "Place in Cell" images (IMAGE / picture in cell): the cell has a
        vm="N" value-metadata index that points into xl/richData.
    Each one is saved under images/ and written back into its own cell of the table as
    ![...](images/...), so the picture appears exactly where it sits in the sheet.
  * Floating images (pictures placed over the grid), saved with their anchor cell and,
    by default, also linked inside that anchor cell.
  * Text in floating shapes / text boxes (form numbers often live there), with anchor cell.
  * Structure facts: sheet state (hidden / veryHidden), merged ranges, hidden rows and
    columns, formula cells (formula text + cached value), comments, charts present.
  * Optional OCR of every image with Tesseract (if installed), plus a review page that
    lists every image with its location and nearby cell text, for an AI agent or a person
    to look at and describe.

It reads the .xlsx package directly (zip + XML) for images, so formats that openpyxl
drops (EMF/WMF) are still saved. Cell values come from openpyxl.

Usage
  python xlsx_to_md.py INPUT [INPUT ...] -o OUTDIR [--ocr auto|off] [--ocr-lang eng+chi_sim]
                       [--no-floating-inline]
  INPUT may be a .xlsx / .xlsm / .xls file or a folder (searched recursively).
  .xls is first converted to .xlsx with Microsoft Excel (Windows, pywin32) or LibreOffice;
  if neither is available, only cell values are read (no images) and a warning is recorded.

Output per workbook: OUTDIR/<name>/
  <name>.md               workbook index + one Markdown table per sheet, images inline
  <name>.images.md        one entry per image: picture, location, nearby text, OCR, description slot
  images/                 every extracted image
  sheets_csv/<sheet>.csv  sheet values; image cells hold [image: images/...]
  image_manifest.csv      sheet, cell, kind, file, size, OCR text, ...
  formula_manifest.csv    sheet, cell, formula, cached value
  shapes_manifest.csv     floating text boxes / shapes with text, anchor cell
  conversion_report.json  counts and warnings (unresolved images, EMF/WMF, charts, OCR state)

Requires: Python 3.9+, openpyxl. Optional: Pillow (image size, EMF/WMF to PNG for OCR),
Tesseract OCR on PATH or in TESSERACT_CMD, xlrd (values-only fallback for .xls).
License: MIT
"""
from __future__ import annotations

import argparse
import csv
import datetime as _dt
import hashlib
import json
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

try:
    import openpyxl
    from openpyxl.utils import get_column_letter, column_index_from_string
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is required: pip install openpyxl")

try:
    from PIL import Image  # optional
except Exception:  # pragma: no cover
    Image = None

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
R_EMBED = "{%s}embed" % NS["r"]
R_ID = "{%s}id" % NS["r"]
DISPIMG_RE = re.compile(r'DISPIMG\(\s*"([^"]+)"', re.I)
CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")


# ----------------------------------------------------------------------------- helpers
def local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def read_xml(z: zipfile.ZipFile, name: str):
    try:
        return ET.fromstring(z.read(name))
    except KeyError:
        return None


def rels_of(z: zipfile.ZipFile, part: str) -> dict:
    """Relationships of a package part: {rId: (type, target_part_path)}."""
    d, b = posixpath.split(part)
    rel_name = posixpath.join(d, "_rels", b + ".rels")
    root = read_xml(z, rel_name)
    out = {}
    if root is None:
        return out
    for rel in root:
        rid, typ, tgt = rel.get("Id"), rel.get("Type", ""), rel.get("Target", "")
        if rel.get("TargetMode") == "External":
            out[rid] = (typ, None)
            continue
        if tgt.startswith("/"):
            path = tgt.lstrip("/")
        else:
            path = posixpath.normpath(posixpath.join(d, tgt))
        out[rid] = (typ, path)
    return out


def safe_name(s: str, maxlen: int = 40) -> str:
    s = re.sub(r'[\\/:*?"<>|\s]+', "_", s).strip("._")
    return (s or "sheet")[:maxlen]


def fmt_val(v) -> str:
    """Render a cell value the way a person reads it."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        if v.is_integer() and abs(v) < 1e15:
            return str(int(v))
        return repr(v)
    if isinstance(v, _dt.datetime):
        if v.hour == v.minute == v.second == 0 and v.microsecond == 0:
            return v.date().isoformat()
        return v.isoformat(sep=" ")
    if isinstance(v, (_dt.date, _dt.time)):
        return v.isoformat()
    return str(v)


def md_cell(s: str) -> str:
    return s.replace("\\", "\\\\").replace("|", "\\|").replace("\r\n", "<br>").replace("\n", "<br>").replace("\r", "<br>")


def cell_ref(col0: int, row0: int) -> str:
    return f"{get_column_letter(col0 + 1)}{row0 + 1}"


def split_ref(ref: str):
    m = CELL_RE.match(ref)
    if not m:
        return None
    return column_index_from_string(m.group(1)), int(m.group(2))


# ----------------------------------------------------------------------------- package model
class Package:
    """Image / shape facts read straight from the .xlsx zip."""

    def __init__(self, path: Path):
        self.z = zipfile.ZipFile(path)
        self.names = set(self.z.namelist())
        self.warnings: list[str] = []
        self.sheets = self._sheets()           # [(name, state, part_path)]
        self.wps_images = self._wps_cellimages()   # {ID: media_path}
        self.rich_images = self._richdata_images()  # [media_path or None] by vm index (0-based)

    # workbook / sheets
    def _sheets(self):
        wb = read_xml(self.z, "xl/workbook.xml")
        rels = rels_of(self.z, "xl/workbook.xml")
        out = []
        if wb is None:
            return out
        for el in wb.iter():
            if local(el.tag) == "sheet":
                rid = el.get(R_ID)
                part = rels.get(rid, ("", None))[1]
                out.append((el.get("name"), el.get("state", "visible"), part))
        return out

    # WPS: =DISPIMG("ID_xxx",1) + xl/cellimages.xml
    def _wps_cellimages(self) -> dict:
        part = None
        for _rid, (typ, tgt) in rels_of(self.z, "xl/workbook.xml").items():
            if tgt and "cellimage" in (typ or "").lower():
                part = tgt
        if part is None and "xl/cellimages.xml" in self.names:
            part = "xl/cellimages.xml"
        if part is None:
            return {}
        root = read_xml(self.z, part)
        rels = rels_of(self.z, part)
        out = {}
        if root is None:
            return out
        for pic in root.iter():
            if local(pic.tag) != "pic":
                continue
            name, embed, descr = None, None, ""
            for el in pic.iter():
                t = local(el.tag)
                if t == "cNvPr":
                    name = el.get("name")
                    descr = el.get("descr", "") or ""
                elif t == "blip":
                    embed = el.get(R_EMBED)
            media = rels.get(embed, ("", None))[1] if embed else None
            if name:
                out[name] = {"media": media, "descr": descr}
        return out

    # Microsoft 365 "Place in Cell": cell vm=N -> metadata.xml -> richData -> media
    def _richdata_images(self) -> list:
        meta = read_xml(self.z, "xl/metadata.xml")
        if meta is None:
            return []
        type_names = []
        future = {}   # name -> [rvb index or None per bk]
        value_bk = []  # [(t, v)]
        for el in meta:
            t = local(el.tag)
            if t == "metadataTypes":
                type_names = [mt.get("name") for mt in el if local(mt.tag) == "metadataType"]
            elif t == "futureMetadata":
                lst = []
                for bk in el:
                    if local(bk.tag) != "bk":
                        continue
                    idx = None
                    for x in bk.iter():
                        if local(x.tag) == "rvb":
                            idx = int(x.get("i"))
                    lst.append(idx)
                future[el.get("name")] = lst
            elif t == "valueMetadata":
                for bk in el:
                    if local(bk.tag) != "bk":
                        continue
                    rc = next((x for x in bk if local(x.tag) == "rc"), None)
                    value_bk.append((int(rc.get("t")), int(rc.get("v"))) if rc is not None else (None, None))
        # rich values
        rv_root = read_xml(self.z, "xl/richData/rdrichvalue.xml")
        st_root = read_xml(self.z, "xl/richData/rdrichvaluestructure.xml")
        rel_root = read_xml(self.z, "xl/richData/richValueRel.xml")
        if rv_root is None or st_root is None or rel_root is None:
            if value_bk:
                self.warnings.append("workbook has value metadata but no complete xl/richData parts; in-cell images (Microsoft 365) may be missing")
            return []
        structures = []
        for s in st_root:
            if local(s.tag) == "s":
                structures.append([k.get("n") for k in s if local(k.tag) == "k"])
        rich = []
        for rv in rv_root:
            if local(rv.tag) != "rv":
                continue
            keys = structures[int(rv.get("s", 0))] if structures else []
            vals = [v.text for v in rv if local(v.tag) == "v"]
            rich.append(dict(zip(keys, vals)))
        rel_ids = [el.get(R_ID) for el in rel_root if local(el.tag) == "rel"]
        rel_map = rels_of(self.z, "xl/richData/richValueRel.xml")
        out = []
        for t, v in value_bk:
            media, alt = None, ""
            if t is not None and 0 < t <= len(type_names) and type_names[t - 1] == "XLRICHVALUE":
                lst = future.get("XLRICHVALUE", [])
                rv_i = lst[v] if v is not None and v < len(lst) else None
                if rv_i is not None and rv_i < len(rich):
                    rv = rich[rv_i]
                    li = rv.get("_rvRel:LocalImageIdentifier")
                    alt = rv.get("Text", "") or ""
                    if li is not None and int(li) < len(rel_ids):
                        media = rel_map.get(rel_ids[int(li)], ("", None))[1]
            out.append({"media": media, "descr": alt})
        return out

    # cells carrying in-cell images, read from the sheet XML itself
    def incell_placements(self, sheet_part: str) -> list:
        out = []
        if not sheet_part or sheet_part not in self.names:
            return out
        with self.z.open(sheet_part) as fh:
            for _ev, el in ET.iterparse(fh):
                if local(el.tag) != "c":
                    continue
                ref = el.get("r")
                vm = el.get("vm")
                texts = [x.text or "" for x in el if local(x.tag) in ("f", "v", "is")]
                texts += [t.text or "" for t in el.iter() if local(t.tag) == "t"]
                joined = " ".join(texts)
                m = DISPIMG_RE.search(joined)
                if m:
                    info = self.wps_images.get(m.group(1))
                    out.append({"cell": ref, "kind": "in-cell (WPS DISPIMG)", "key": m.group(1),
                                "media": info["media"] if info else None,
                                "descr": info["descr"] if info else ""})
                elif vm is not None and self.rich_images:
                    i = int(vm) - 1
                    info = self.rich_images[i] if 0 <= i < len(self.rich_images) else None
                    if info and info["media"]:
                        out.append({"cell": ref, "kind": "in-cell (Microsoft 365)", "key": f"vm={vm}",
                                    "media": info["media"], "descr": info["descr"]})
                el.clear()
        return out

    # floating pictures, shapes with text, charts
    def drawing_objects(self, sheet_part: str):
        pics, shapes, charts = [], [], 0
        if not sheet_part:
            return pics, shapes, charts
        for _rid, (typ, tgt) in rels_of(self.z, sheet_part).items():
            if not tgt or not typ.endswith("/drawing"):
                continue
            root = read_xml(self.z, tgt)
            drels = rels_of(self.z, tgt)
            if root is None:
                continue
            for anchor in root:
                kind = local(anchor.tag)
                if kind not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
                    continue
                cell = ""
                frm = next((x for x in anchor if local(x.tag) == "from"), None)
                if frm is not None:
                    col = int(next(x.text for x in frm if local(x.tag) == "col"))
                    row = int(next(x.text for x in frm if local(x.tag) == "row"))
                    cell = cell_ref(col, row)
                to_cell = ""
                to = next((x for x in anchor if local(x.tag) == "to"), None)
                if to is not None:
                    col = int(next(x.text for x in to if local(x.tag) == "col"))
                    row = int(next(x.text for x in to if local(x.tag) == "row"))
                    to_cell = cell_ref(col, row)
                for el in anchor.iter():
                    t = local(el.tag)
                    if t == "pic":
                        embed, name, descr = None, "", ""
                        for x in el.iter():
                            lt = local(x.tag)
                            if lt == "blip":
                                embed = x.get(R_EMBED)
                            elif lt == "cNvPr":
                                name, descr = x.get("name", ""), x.get("descr", "") or ""
                        media = drels.get(embed, ("", None))[1] if embed else None
                        pics.append({"cell": cell, "to_cell": to_cell, "kind": "floating", "key": name,
                                     "media": media, "descr": descr})
                    elif t == "sp":
                        txt = "".join(x.text or "" for x in el.iter() if local(x.tag) == "t").strip()
                        if txt:
                            name = next((x.get("name", "") for x in el.iter() if local(x.tag) == "cNvPr"), "")
                            shapes.append({"cell": cell, "to_cell": to_cell, "name": name, "text": txt})
                    elif t == "graphicFrame":
                        charts += 1
        return pics, shapes, charts

    def media_bytes(self, media: str) -> bytes | None:
        try:
            return self.z.read(media)
        except KeyError:
            return None


# ----------------------------------------------------------------------------- OCR
class OCR:
    def __init__(self, mode: str, langs: str, warnings: list):
        self.cmd, self.langs = None, None
        if mode == "off":
            return
        cand = [os.environ.get("TESSERACT_CMD"), shutil.which("tesseract"),
                r"C:\Program Files\Tesseract-OCR\tesseract.exe", "/usr/bin/tesseract", "/opt/homebrew/bin/tesseract"]
        self.cmd = next((c for c in cand if c and Path(c).exists()), None)
        if not self.cmd:
            warnings.append("OCR skipped: tesseract not found (set TESSERACT_CMD or install Tesseract)")
            return
        try:
            out = subprocess.run([self.cmd, "--list-langs"], capture_output=True, text=True, timeout=30).stdout
            have = {l.strip() for l in out.splitlines()[1:] if l.strip()}
        except Exception:
            have = set()
        want = [l for l in langs.split("+") if l]
        use = [l for l in want if l in have] or (["eng"] if "eng" in have else [])
        missing = [l for l in want if l not in have]
        if missing:
            warnings.append(f"OCR languages not installed: {'+'.join(missing)}; using {'+'.join(use) or 'none'}")
        self.langs = "+".join(use) if use else None
        if not self.langs:
            self.cmd = None

    def run(self, img_path: Path) -> str:
        if not self.cmd:
            return ""
        src = img_path
        tmp = None
        if img_path.suffix.lower() in (".emf", ".wmf") and Image is not None:
            try:
                tmp = Path(tempfile.mkstemp(suffix=".png")[1])
                Image.open(img_path).save(tmp)
                src = tmp
            except Exception:
                return ""
        try:
            r = subprocess.run([self.cmd, str(src), "stdout", "-l", self.langs, "--psm", "6"],
                               capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120)
            return re.sub(r"\s+\n", "\n", r.stdout).strip()
        except Exception:
            return ""
        finally:
            if tmp and tmp.exists():
                tmp.unlink()


# ----------------------------------------------------------------------------- .xls
def xls_to_xlsx(path: Path, warnings: list) -> Path | None:
    out_dir = Path(tempfile.mkdtemp(prefix="xls2xlsx_"))
    target = out_dir / (path.stem + ".xlsx")
    if sys.platform == "win32":
        try:
            import win32com.client  # type: ignore
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            try:
                wb = excel.Workbooks.Open(str(path.resolve()), ReadOnly=True)
                wb.SaveAs(str(target.resolve()), FileFormat=51)
                wb.Close(False)
            finally:
                excel.Quit()
            if target.exists():
                warnings.append(".xls converted to .xlsx with Microsoft Excel")
                return target
        except Exception as e:
            warnings.append(f".xls conversion with Excel failed: {e.__class__.__name__}")
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    for c in (r"C:\Program Files\LibreOffice\program\soffice.exe", "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if not soffice and Path(c).exists():
            soffice = c
    if soffice:
        try:
            subprocess.run([soffice, "--headless", "--convert-to", "xlsx", "--outdir", str(out_dir), str(path)],
                           capture_output=True, timeout=300)
            if target.exists():
                warnings.append(".xls converted to .xlsx with LibreOffice")
                return target
        except Exception:
            pass
    return None


def xls_values_only(path: Path):
    import xlrd  # type: ignore
    book = xlrd.open_workbook(str(path))
    sheets = []
    for sh in book.sheets():
        rows = [[fmt_val(sh.cell_value(r, c)) for c in range(sh.ncols)] for r in range(sh.nrows)]
        sheets.append((sh.name, rows))
    return sheets


# ----------------------------------------------------------------------------- convert
def nearby_text(grid: dict, col: int, row: int) -> str:
    """Left neighbour, column header (first non-empty cell above) and same-row text."""
    parts = []
    left = next((grid.get((row, c)) for c in range(col - 1, 0, -1) if grid.get((row, c))), "")
    if left:
        parts.append(f"left: {left}")
    above = next((grid.get((r, col)) for r in range(row - 1, 0, -1) if grid.get((r, col))), "")
    if above:
        parts.append(f"above: {above}")
    row_txt = " / ".join(v for (r, c), v in sorted(grid.items()) if r == row and c != col and v and "![" not in v)
    if row_txt:
        parts.append(f"row: {row_txt[:200]}")
    return "; ".join(parts)


def convert(path: Path, outroot: Path, ocr_mode: str, ocr_lang: str, floating_inline: bool) -> dict:
    warnings: list[str] = []
    src_display = str(path)
    work = path
    if path.suffix.lower() == ".xls":
        conv = xls_to_xlsx(path, warnings)
        if conv is None:
            return convert_xls_values_only(path, outroot, warnings)
        work = conv

    out = outroot / safe_name(path.stem, 80)
    img_dir = out / "images"
    csv_dir = out / "sheets_csv"
    img_dir.mkdir(parents=True, exist_ok=True)
    csv_dir.mkdir(parents=True, exist_ok=True)

    pkg = Package(work)
    warnings += pkg.warnings
    ocr = OCR(ocr_mode, ocr_lang, warnings)
    wb_v = openpyxl.load_workbook(work, data_only=True)
    wb_f = openpyxl.load_workbook(work, data_only=False)

    image_rows, formula_rows, shape_rows = [], [], []
    md = [f"# {path.name}", "",
          f"- source: `{src_display}`",
          f"- converted: {_dt.datetime.now().isoformat(timespec='seconds')}",
          f"- sheets: {len(pkg.sheets)}", ""]
    totals = {"in_cell_wps": 0, "in_cell_m365": 0, "floating": 0, "unresolved": 0, "shapes_with_text": 0, "charts": 0}

    for si, (sname, state, part) in enumerate(pkg.sheets, 1):
        if sname not in wb_v.sheetnames:
            warnings.append(f"sheet '{sname}' not readable by openpyxl")
            continue
        ws, wsf = wb_v[sname], wb_f[sname]
        max_r, max_c = ws.max_row or 0, ws.max_column or 0
        grid: dict[tuple[int, int], str] = {}
        for row in ws.iter_rows(min_row=1, max_row=max_r, max_col=max_c):
            for cell in row:
                if cell.value is not None:
                    grid[(cell.row, cell.column)] = fmt_val(cell.value)
        for row in wsf.iter_rows(min_row=1, max_row=max_r, max_col=max_c):
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("="):
                    formula_rows.append({"sheet": sname, "cell": cell.coordinate, "formula": v,
                                         "cached_value": grid.get((cell.row, cell.column), "")})
                elif type(v).__name__ == "ArrayFormula":
                    formula_rows.append({"sheet": sname, "cell": cell.coordinate, "formula": getattr(v, "text", str(v)),
                                         "cached_value": grid.get((cell.row, cell.column), "")})

        # images
        placements = pkg.incell_placements(part)
        pics, shapes, charts = pkg.drawing_objects(part)
        totals["charts"] += charts
        n = 0
        img_links: dict[tuple[int, int], list[str]] = {}
        incell_keys: set[tuple[int, int]] = set()
        for pl in placements + pics:
            rc = split_ref(pl["cell"]) if pl["cell"] else None
            if pl["kind"].startswith("in-cell (WPS"):
                totals["in_cell_wps"] += 1
            elif pl["kind"].startswith("in-cell (Microsoft"):
                totals["in_cell_m365"] += 1
            else:
                totals["floating"] += 1
            data = pkg.media_bytes(pl["media"]) if pl["media"] else None
            if data is None:
                totals["unresolved"] += 1
                warnings.append(f"{sname}!{pl['cell']}: image not found in package ({pl['kind']}, key {pl['key']})")
                if rc and pl["kind"].startswith("in-cell"):
                    grid[(rc[1], rc[0])] = f"[image missing: {pl['key']}]"
                continue
            n += 1
            ext = posixpath.splitext(pl["media"])[1].lower() or ".bin"
            tag = "cell" if pl["kind"].startswith("in-cell") else "float"
            fname = f"{si:02d}_{safe_name(sname, 24)}_{pl['cell'] or 'abs'}_{tag}{n}{ext}"
            fpath = img_dir / fname
            fpath.write_bytes(data)
            w = h = ""
            if Image is not None:
                try:
                    with Image.open(fpath) as im:
                        w, h = im.size
                except Exception:
                    pass
            if ext in (".emf", ".wmf"):
                warnings.append(f"{sname}!{pl['cell']}: {ext} image saved as-is; convert to PNG to view it")
            rel = f"images/{fname}"
            alt = f"{sname}!{pl['cell']}" + (f" {pl['descr']}" if pl["descr"] else "")
            link = f"![{md_cell(alt)}]({rel})"
            if rc:
                key = (rc[1], rc[0])
                if pl["kind"].startswith("in-cell"):
                    grid.pop(key, None)         # drop =DISPIMG(...) / #VALUE!: the picture *is* the cell value
                    incell_keys.add(key)
                    img_links.setdefault(key, []).append(link)
                elif floating_inline:
                    img_links.setdefault(key, []).append(link + " (floating)")
                max_r, max_c = max(max_r, rc[1]), max(max_c, rc[0])
            image_rows.append({"sheet": sname, "cell": pl["cell"], "to_cell": pl.get("to_cell", ""),
                               "kind": pl["kind"], "image_file": rel, "source_part": pl["media"],
                               "key": pl["key"], "alt_text": pl["descr"], "width": w, "height": h,
                               "bytes": len(data), "sha1": hashlib.sha1(data).hexdigest()[:12],
                               "ocr_text": "", "nearby_text": ""})
        for sh in shapes:
            totals["shapes_with_text"] += 1
            shape_rows.append({"sheet": sname, "cell": sh["cell"], "to_cell": sh["to_cell"], "name": sh["name"], "text": sh["text"]})

        # nearby text + OCR (after all values are in place)
        for r in image_rows:
            if r["sheet"] != sname:
                continue
            rc = split_ref(r["cell"]) if r["cell"] else None
            if rc:
                r["nearby_text"] = nearby_text(grid, rc[0], rc[1])
            r["ocr_text"] = ocr.run(out / r["image_file"])

        # merged cells: keep the value in the top-left cell only
        merged = [str(m) for m in ws.merged_cells.ranges]
        hidden_rows = [r for r, d in ws.row_dimensions.items() if d.hidden]
        hidden_cols = [c for c, d in ws.column_dimensions.items() if d.hidden]
        comments = [(c.coordinate, c.comment.text.strip()) for row in ws.iter_rows() for c in row if c.comment]

        # sheet section
        md += [f"## Sheet {si}: {sname}", ""]
        facts = [f"state: {state}", f"used range: A1:{get_column_letter(max(max_c,1))}{max(max_r,1)}"]
        if merged:
            facts.append(f"merged: {', '.join(merged[:60])}" + (" ..." if len(merged) > 60 else ""))
        if hidden_rows:
            facts.append(f"hidden rows: {', '.join(map(str, hidden_rows[:60]))}")
        if hidden_cols:
            facts.append(f"hidden columns: {', '.join(hidden_cols[:60])}")
        n_img = sum(1 for r in image_rows if r["sheet"] == sname)
        if n_img:
            facts.append(f"images: {n_img}")
        if charts:
            facts.append(f"charts: {charts} (not rendered; open the workbook to see them)")
        md += [f"- {f}" for f in facts] + [""]
        for coord, text in comments:
            md.append(f"- comment {coord}: {md_cell(text)}")
        for sh in shapes:
            md.append(f"- text box at {sh['cell'] or '?'}: {md_cell(sh['text'])}")
        if comments or shapes:
            md.append("")

        if max_r and max_c:
            header = ["#"] + [get_column_letter(c) for c in range(1, max_c + 1)]
            md.append("| " + " | ".join(header) + " |")
            md.append("|" + "---|" * len(header))
            rows_csv = []
            for r in range(1, max_r + 1):
                md_vals, csv_vals = [], []
                for c in range(1, max_c + 1):
                    text = grid.get((r, c), "")
                    links = img_links.get((r, c), [])
                    md_vals.append("<br>".join(([md_cell(text)] if text else []) + links))
                    csv_vals.append("\n".join(([text] if text else []) +
                                              [re.sub(r"!\[[^\]]*\]\((images/[^)]+)\)", r"[image: \1]", l) for l in links]))
                md.append("| " + " | ".join([str(r)] + md_vals) + " |")
                rows_csv.append(csv_vals)
            with open(csv_dir / f"{si:02d}_{safe_name(sname)}.csv", "w", newline="", encoding="utf-8-sig") as fh:
                csv.writer(fh).writerows(rows_csv)
        else:
            md.append("_(empty sheet)_")
        md.append("")

    # files
    (out / f"{safe_name(path.stem, 80)}.md").write_text("\n".join(md), encoding="utf-8")
    write_csv(out / "image_manifest.csv", image_rows,
              ["sheet", "cell", "to_cell", "kind", "image_file", "source_part", "key", "alt_text",
               "width", "height", "bytes", "sha1", "ocr_text", "nearby_text"])
    write_csv(out / "formula_manifest.csv", formula_rows, ["sheet", "cell", "formula", "cached_value"])
    write_csv(out / "shapes_manifest.csv", shape_rows, ["sheet", "cell", "to_cell", "name", "text"])
    write_images_page(out / f"{safe_name(path.stem, 80)}.images.md", path.name, image_rows)
    report = {"source": src_display, "output_dir": str(out), "sheets": len(pkg.sheets),
              "images": totals, "formulas": len(formula_rows), "ocr": {"engine": ocr.cmd, "langs": ocr.langs},
              "warnings": warnings,
              "review_required": bool(totals["unresolved"] or totals["charts"] or any("emf" in w or "wmf" in w for w in warnings))}
    (out / "conversion_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    pkg.z.close()
    return report


def convert_xls_values_only(path: Path, outroot: Path, warnings: list) -> dict:
    out = outroot / safe_name(path.stem, 80)
    (out / "sheets_csv").mkdir(parents=True, exist_ok=True)
    warnings.append(".xls read values-only (no Excel or LibreOffice found): images, shapes and formulas NOT extracted")
    md = [f"# {path.name}", "", f"- source: `{path}`", "- WARNING: values only, images not extracted", ""]
    sheets = xls_values_only(path)
    for si, (name, rows) in enumerate(sheets, 1):
        md += [f"## Sheet {si}: {name}", ""]
        if rows and rows[0]:
            ncol = max(len(r) for r in rows)
            md.append("| # | " + " | ".join(get_column_letter(c) for c in range(1, ncol + 1)) + " |")
            md.append("|" + "---|" * (ncol + 1))
            for i, r in enumerate(rows, 1):
                md.append(f"| {i} | " + " | ".join(md_cell(v) for v in r + [""] * (ncol - len(r))) + " |")
        md.append("")
        with open(out / "sheets_csv" / f"{si:02d}_{safe_name(name)}.csv", "w", newline="", encoding="utf-8-sig") as fh:
            csv.writer(fh).writerows(rows)
    (out / f"{safe_name(path.stem, 80)}.md").write_text("\n".join(md), encoding="utf-8")
    report = {"source": str(path), "output_dir": str(out), "sheets": len(sheets), "images": None,
              "warnings": warnings, "review_required": True}
    (out / "conversion_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def write_csv(path: Path, rows: list, fields: list) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def write_images_page(path: Path, title: str, rows: list) -> None:
    lines = [f"# Images in {title}", "",
             f"{len(rows)} image(s). Each entry: where it sits, what is around it, OCR text, and a description slot.",
             "An agent should open every image file, look at it, and replace the description placeholder.", ""]
    for i, r in enumerate(rows, 1):
        loc = f"{r['sheet']}!{r['cell']}" + (f":{r['to_cell']}" if r.get("to_cell") else "")
        lines += [f"## {i}. {loc} - {r['kind']}", "",
                  f"![{loc}]({r['image_file']})", "",
                  f"- file: `{r['image_file']}` ({r['width']}x{r['height']} px, {r['bytes']} bytes)",
                  f"- package part: `{r['source_part']}`" + (f", key `{r['key']}`" if r["key"] else "")]
        if r["alt_text"]:
            lines.append(f"- alt text: {r['alt_text']}")
        if r["nearby_text"]:
            lines.append(f"- nearby cells: {r['nearby_text']}")
        lines.append("- OCR: " + (("\n\n```text\n" + r["ocr_text"] + "\n```") if r["ocr_text"] else "_(none)_"))
        lines += ["- description: _(to be filled after looking at the image)_", ""]
    path.write_text("\n".join(lines), encoding="utf-8")


def iter_inputs(items):
    for it in items:
        p = Path(it)
        if p.is_dir():
            for q in sorted(p.rglob("*")):
                if q.suffix.lower() in (".xlsx", ".xlsm", ".xls") and not q.name.startswith("~$"):
                    yield q
        elif p.exists():
            yield p
        else:
            print(f"not found: {p}", file=sys.stderr)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Excel -> Markdown with in-cell and floating images put back in place.")
    ap.add_argument("inputs", nargs="+", help=".xlsx/.xlsm/.xls files or folders")
    ap.add_argument("-o", "--out", required=True, help="output folder")
    ap.add_argument("--ocr", choices=["auto", "off"], default="auto", help="OCR images with Tesseract if available")
    ap.add_argument("--ocr-lang", default="eng+chi_sim", help="Tesseract languages, e.g. eng+chi_sim")
    ap.add_argument("--no-floating-inline", action="store_true", help="do not link floating images inside their anchor cell")
    a = ap.parse_args(argv)
    outroot = Path(a.out)
    outroot.mkdir(parents=True, exist_ok=True)
    summary = []
    for p in iter_inputs(a.inputs):
        try:
            rep = convert(p, outroot, a.ocr, a.ocr_lang, not a.no_floating_inline)
        except Exception as e:  # keep the batch going
            rep = {"source": str(p), "error": f"{e.__class__.__name__}: {e}", "review_required": True}
        summary.append(rep)
        img = rep.get("images") or {}
        print(f"{p.name}: in-cell WPS {img.get('in_cell_wps', 0)}, in-cell M365 {img.get('in_cell_m365', 0)}, "
              f"floating {img.get('floating', 0)}, unresolved {img.get('unresolved', 0)}"
              + (f"  ERROR {rep['error']}" if "error" in rep else ""))
    (outroot / "batch_summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if all("error" not in r for r in summary) else 1


if __name__ == "__main__":
    sys.exit(main())
