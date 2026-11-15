---
name: xls-skill
description: Convert Excel workbooks (.xlsx/.xlsm/.xls) into AI-readable Markdown with every picture put back where it sits - including images embedded INSIDE cells (WPS Office =DISPIMG and Microsoft 365 "Place in Cell"), which normal Excel readers silently miss - plus floating pictures, text boxes, formulas, merged/hidden structure and OCR. Use when reading, converting, auditing or answering questions about spreadsheets that contain photos, labels, logos, screenshots, drawings or signatures in cells or over the grid.
---

# xls-skill: Excel -> Markdown, pictures kept in their cells

## Why this skill exists

Most Excel readers (openpyxl, pandas, CSV export) only see cell values. Two kinds of pictures get lost:

1. **In-cell images.** The picture is the content of a cell.
   - WPS Office "embed in cell": the cell only shows `=DISPIMG("ID_...",1)`; the picture is in `xl/cellimages.xml`.
   - Microsoft 365 "Place in Cell" / `IMAGE`: the cell shows `#VALUE!`; the picture is reached through `xl/metadata.xml` and `xl/richData/`.
2. **Floating pictures and text boxes** drawn over the grid (process photos, labels, form numbers).

This skill extracts all of them, saves each picture as a file, and writes it back into the Markdown table **at the cell where it lives**, so a reader sees the table and its pictures together:

```
| # | A | B | C | D |
|---|---|---|---|---|
| 2 | Bracket | ![Parts!B2](images/01_Parts_B2_cell1.png) | Bracket \| in-cell picture | 4 |
```

## How to run

```bash
python <skill-dir>/scripts/xlsx_to_md.py <file-or-folder> [...] -o <out-dir> [--ocr auto|off] [--ocr-lang eng+chi_sim]
```

- Accepts `.xlsx`, `.xlsm`, `.xls`, or folders (recursive). Prints one line per workbook with picture counts.
- Needs Python 3.9+ and `openpyxl`. Optional: `Pillow` (image sizes), Tesseract OCR (text inside pictures), Microsoft Excel or LibreOffice (to open `.xls`).
- Self-test: `python <skill-dir>/tests/run_tests.py` builds synthetic WPS and Microsoft 365 workbooks and checks every picture lands in its cell.

## What you get per workbook (`<out-dir>/<name>/`)

| File | Content |
|---|---|
| `<name>.md` | Workbook index, then each sheet as a full Markdown table (no truncation) with pictures inline in their cells; sheet facts (hidden state, merged ranges, hidden rows/columns, comments, text boxes, chart count) above each table |
| `<name>.images.md` | One entry per picture: the picture, sheet!cell, kind, size, nearby cell text, OCR text, and a description slot |
| `images/` | Every picture as a file, named `<sheet#>_<sheet>_<cell>_<cell|float><n>.<ext>` |
| `sheets_csv/` | Sheet values; picture cells hold `[image: images/...]` |
| `image_manifest.csv` | sheet, cell, kind (`in-cell (WPS DISPIMG)` / `in-cell (Microsoft 365)` / `floating`), file, package part, alt text, size, OCR, nearby text |
| `formula_manifest.csv` | Formula text and cached value per cell |
| `shapes_manifest.csv` | Text from floating text boxes / shapes, with anchor cell |
| `conversion_report.json` | Counts, warnings, `review_required` |

## Workflow for the agent

1. **Convert.** Run the script on the workbook(s).
2. **Read the report first.** Open `conversion_report.json`. Act on every warning:
   - `image not found in package` -> the cell is marked `[image missing: ID]`; say so, do not guess the picture.
   - `.emf/.wmf saved as-is` -> convert to PNG (e.g. Pillow on Windows, LibreOffice elsewhere) before looking at it.
   - `charts` > 0 -> charts are counted, not rendered; open or export the sheet if the chart matters.
   - OCR skipped or a language missing -> rely on looking at the pictures yourself.
3. **Look at the pictures.** Open every file listed in `<name>.images.md` and replace each `description` placeholder with what the picture shows: object, readable text (labels, model numbers, marks), symbols, and how it relates to the nearby cells. OCR is a hint, not the answer; small label text often needs a crop or zoom.
4. **Answer from the table and the pictures together.** Use the Markdown table for exact cell values, formulas and structure; use the pictures for "what does it look like / what is written on the part". Cite `sheet!cell` for both.
5. **Report gaps honestly.** If anything in step 2 stayed unresolved, mark the workbook as needing manual review.

## Rules

- A CSV or value-only table is not complete evidence when the workbook has pictures, text boxes, colours or hidden content. Check the sheet facts and manifests.
- Pair every picture with its cell and neighbours; never rely on picture order alone.
- Text that matters may live outside cells: text boxes (form numbers, revision marks), comments, headers/footers. Text boxes and comments are extracted; headers/footers are not - check them in Excel if a document number is missing.
- Colour can be data (pass/fail, risk level, status). This script does not decode fills or conditional formatting; if colours carry meaning, inspect them with openpyxl or a page export and state the mapping.
- Merged ranges keep their value in the top-left cell only; read the `merged:` line before interpreting blank cells.
- Hidden sheets, rows and columns are converted and listed, not dropped. Say when an answer comes from hidden content.
- Do not trust mojibake paths in a terminal; check the real path exists before reporting a file as missing.

## Limitations

- Floating pictures are linked in their top-left anchor cell; a picture spanning many cells is listed with its `to_cell` in the manifest.
- Grouped shapes are read, but arrows and callouts drawn as separate shapes keep their relation to a photo only in a page render. When layout matters ("which arrow points where"), export the sheet to PDF/PNG with Excel or LibreOffice and look at it alongside the Markdown.
- Pictures in headers/footers (VML) and in comments are not extracted.
- `.xls` needs Microsoft Excel (Windows, pywin32) or LibreOffice to convert first; without them only cell values are read and the report says so.
