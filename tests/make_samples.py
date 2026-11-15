#!/usr/bin/env python3
"""Build small synthetic workbooks that exercise every image path of xlsx_to_md.py.

  m365_in_cell.xlsx  Microsoft 365 "Place in Cell" images (xlsxwriter embed_image),
                     a floating picture, a text box, a merged range, a formula, a hidden sheet.
  wps_dispimg.xlsx   WPS-style in-cell images: =DISPIMG("ID_...",1) + xl/cellimages.xml,
                     plus one DISPIMG that points to a missing picture.

No real company data: pictures are generated text labels.
Usage: python make_samples.py OUTDIR   (needs Pillow, xlsxwriter, openpyxl)
"""
import io
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def label_png(text: str, color: str, size=(360, 160)) -> bytes:
    im = Image.new("RGB", size, color)
    d = ImageDraw.Draw(im)
    try:
        font = ImageFont.truetype("arial.ttf", 56)
    except Exception:
        try:
            font = ImageFont.truetype("DejaVuSans.ttf", 56)
        except Exception:
            font = ImageFont.load_default()
    box = d.textbbox((0, 0), text, font=font)
    d.text(((size[0] - (box[2] - box[0])) / 2, (size[1] - (box[3] - box[1])) / 2 - box[1]), text, fill="black", font=font)
    buf = io.BytesIO()
    im.save(buf, "PNG")
    return buf.getvalue()


def make_m365(out: Path) -> Path:
    import xlsxwriter
    p = out / "m365_in_cell.xlsx"
    wb = xlsxwriter.Workbook(str(p))
    ws = wb.add_worksheet("Parts")
    bold = wb.add_format({"bold": True})
    ws.write_row("A1", ["Part", "Photo", "Note", "Qty x2"], bold)
    ws.set_column("B:B", 30)
    for r, (name, label, color) in enumerate([("Bracket", "PART-A", "#9fd3ff"), ("Spring", "PART-B", "#ffd59f")], start=2):
        ws.write(f"A{r}", name)
        ws.set_row(r - 1, 80)
        ws.embed_image(f"B{r}", "p.png", {"image_data": io.BytesIO(label_png(label, color))})
        ws.write(f"C{r}", f"{name} | in-cell picture")
        ws.write_formula(f"D{r}", f"={r}*2", None, r * 2)
    ws.merge_range("A5:C5", "Merged title across A5:C5")
    ws.insert_image("F2", "f.png", {"image_data": io.BytesIO(label_png("FLOAT-1", "#c8f7c5"))})
    ws.insert_textbox("F10", "Form No. QR-TEST-01", {"width": 220, "height": 40})
    hid = wb.add_worksheet("HiddenSheet")
    hid.write("A1", "secret-but-recorded")
    hid.hide()
    wb.close()
    return p


CELLIMAGES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<etc:cellImages xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData">{items}</etc:cellImages>"""
ITEM = """<etc:cellImage><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{n}" name="{id}" descr="{descr}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId{n}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic></etc:cellImage>"""


def make_wps(out: Path) -> Path:
    import openpyxl
    p = out / "wps_dispimg.xlsx"
    tmp = Path(tempfile.mkdtemp()) / "base.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BOM"
    ws.append(["Item", "Picture", "Spec"])
    rows = [("Bracket", "ID_TESTA1", "SUS304"), ("Spring", "ID_TESTB2", "0.5 mm wire"), ("Ghost", "ID_NOPE", "picture missing on purpose")]
    for i, (item, pid, spec) in enumerate(rows, start=2):
        ws.cell(i, 1, item)
        ws.cell(i, 2, f'=_xlfn.DISPIMG("{pid}",1)')
        ws.cell(i, 3, spec)
    wb.save(tmp)
    pics = [("ID_TESTA1", "bracket photo", label_png("WPS-A", "#e0c3fc")), ("ID_TESTB2", "spring photo", label_png("WPS-B", "#fcd5c3"))]
    with zipfile.ZipFile(tmp) as zin, zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                s = data.decode("utf-8")
                if 'Extension="png"' not in s:
                    s = s.replace("<Default ", '<Default Extension="png" ContentType="image/png"/><Default ', 1)
                s = s.replace("</Types>", '<Override PartName="/xl/cellimages.xml" ContentType="application/vnd.wps-officedocument.cellimage+xml"/></Types>')
                data = s.encode("utf-8")
            elif item.filename == "xl/_rels/workbook.xml.rels":
                s = data.decode("utf-8")
                s = s.replace("</Relationships>", '<Relationship Id="rIdCellImg" Type="http://www.wps.cn/officeDocument/2020/cellImage" Target="cellimages.xml"/></Relationships>')
                data = s.encode("utf-8")
            zout.writestr(item, data)
        items = "".join(ITEM.format(n=n, id=pid, descr=descr) for n, (pid, descr, _png) in enumerate(pics, start=1))
        zout.writestr("xl/cellimages.xml", CELLIMAGES.format(items=items))
        rels = "".join(f'<Relationship Id="rId{n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image{n}.png"/>'
                       for n in range(1, len(pics) + 1))
        zout.writestr("xl/_rels/cellimages.xml.rels",
                      f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>')
        for n, (_pid, _d, png) in enumerate(pics, start=1):
            zout.writestr(f"xl/media/image{n}.png", png)
    shutil.rmtree(tmp.parent, ignore_errors=True)
    return p


if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "samples")
    out.mkdir(parents=True, exist_ok=True)
    for f in (make_m365(out), make_wps(out)):
        print("wrote", f)
