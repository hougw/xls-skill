#!/usr/bin/env python3
"""Self-test: build the synthetic samples, convert them, check every image landed in its cell.
Usage: python tests/run_tests.py
"""
import csv
import json
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
WORK = HERE / "_out"
SAMPLES = WORK / "samples"
OUT = WORK / "converted"
fails = []


def check(cond, msg):
    print(("PASS " if cond else "FAIL ") + msg)
    if not cond:
        fails.append(msg)


def manifest(name):
    with open(OUT / name / "image_manifest.csv", encoding="utf-8-sig") as fh:
        return list(csv.DictReader(fh))


shutil.rmtree(WORK, ignore_errors=True)
subprocess.run([sys.executable, str(HERE / "make_samples.py"), str(SAMPLES)], check=True)
subprocess.run([sys.executable, str(ROOT / "scripts" / "xlsx_to_md.py"), str(SAMPLES), "-o", str(OUT), "--ocr-lang", "eng"], check=True)

# Microsoft 365 "Place in Cell"
rows = manifest("m365_in_cell")
cells = {(r["kind"], r["cell"]) for r in rows}
check(("in-cell (Microsoft 365)", "B2") in cells and ("in-cell (Microsoft 365)", "B3") in cells, "M365 in-cell images found at B2 and B3")
check(("floating", "F2") in cells, "floating picture found, anchored at F2")
for r in rows:
    check((OUT / "m365_in_cell" / r["image_file"]).exists(), f"image file written: {r['image_file']}")
md = (OUT / "m365_in_cell" / "m365_in_cell.md").read_text(encoding="utf-8")
row2 = next(l for l in md.splitlines() if l.startswith("| 2 |"))
check("![Parts!B2" in row2.split("|")[3], "Markdown: the picture sits in column B of row 2")
check("#VALUE!" not in row2, "Markdown: the #VALUE! placeholder of the in-cell image is gone")
check("state: hidden" in md, "hidden sheet recorded")
check("merged: A5:C5" in md, "merged range recorded")
check("Bracket \\| in-cell picture" in md, "pipe in cell text escaped")
shapes = (OUT / "m365_in_cell" / "shapes_manifest.csv").read_text(encoding="utf-8-sig")
check("QR-TEST-01" in shapes, "text box text extracted (Form No. QR-TEST-01)")
formulas = (OUT / "m365_in_cell" / "formula_manifest.csv").read_text(encoding="utf-8-sig")
check("D2" in formulas and "=2*2" in formulas, "formula recorded with cell")
rep = json.loads((OUT / "m365_in_cell" / "conversion_report.json").read_text(encoding="utf-8"))
if rep["ocr"]["engine"]:
    ocr = {r["cell"]: r["ocr_text"] for r in rows}
    check("PART-A" in ocr.get("B2", "").replace(" ", ""), "OCR read the label inside the B2 picture")
else:
    print("SKIP OCR (tesseract not installed)")

# WPS DISPIMG
rows = manifest("wps_dispimg")
cells = {(r["kind"], r["cell"]) for r in rows}
check(("in-cell (WPS DISPIMG)", "B2") in cells and ("in-cell (WPS DISPIMG)", "B3") in cells, "WPS in-cell images found at B2 and B3")
check(any(r["alt_text"] == "bracket photo" for r in rows), "WPS picture alt text kept")
md = (OUT / "wps_dispimg" / "wps_dispimg.md").read_text(encoding="utf-8")
row3 = next(l for l in md.splitlines() if l.startswith("| 3 |"))
check("![BOM!B3" in row3.split("|")[3] and "DISPIMG" not in row3, "Markdown: =DISPIMG(...) replaced by the picture in B3")
check("[image missing: ID_NOPE]" in md, "a DISPIMG with no picture is flagged, not dropped silently")
rep = json.loads((OUT / "wps_dispimg" / "conversion_report.json").read_text(encoding="utf-8"))
check(rep["images"]["unresolved"] == 1 and rep["review_required"], "report counts 1 unresolved image and asks for review")
csv_txt = (OUT / "wps_dispimg" / "sheets_csv" / "01_BOM.csv").read_text(encoding="utf-8-sig")
check("[image: images/" in csv_txt, "CSV keeps an [image: ...] marker in the picture cell")

print()
print("ALL PASS" if not fails else f"{len(fails)} FAILED")
sys.exit(1 if fails else 0)
