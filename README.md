# xls-skill

**Turn Excel workbooks into AI-readable Markdown — and keep every picture in the cell where it lives.**

Most spreadsheet readers only see cell values. Pictures embedded *inside* cells come out as
`=DISPIMG("ID_…",1)` or `#VALUE!`, and floating photos, labels and text boxes disappear.
`xls-skill` reads the workbook package directly, pulls every picture out, and writes it back
into the Markdown table at its own cell, so an AI agent (or a person) sees the table and its
pictures together.

It is an Agent Skill (a folder with a `SKILL.md`) for Claude Code and other agents that read
skills, plus a standalone Python script you can run on its own.

## Highlight: images inside cells

| Where the picture lives | What other readers show | What xls-skill does |
|---|---|---|
| **WPS Office "embed in cell"** (`xl/cellimages.xml`) | `=DISPIMG("ID_9115…",1)` | Resolves the ID → saves the picture → puts it in that cell |
| **Microsoft 365 "Place in Cell" / `IMAGE`** (`xl/richData/`) | `#VALUE!` | Follows `vm` → `metadata.xml` → rich value → picture → puts it in that cell |
| **Floating pictures** (drawings over the grid) | nothing | Saves them with anchor cell and links them in that cell |
| **Text boxes / shapes** (form numbers, notes) | nothing | Extracts the text with its anchor cell |

Before (what a normal reader gives you):

```
| Bracket | =DISPIMG("ID_TESTA1",1) | SUS304 |
```

After (`xls-skill`):

```
| 2 | Bracket | ![BOM!B2 bracket photo](images/01_BOM_B2_cell1.png) | SUS304 |
```

See [`examples/output/m365_in_cell/m365_in_cell.md`](examples/output/m365_in_cell/m365_in_cell.md)
and [`examples/output/wps_dispimg/wps_dispimg.md`](examples/output/wps_dispimg/wps_dispimg.md) —
GitHub renders the pictures right inside the tables.

Every picture also gets an entry in `<name>.images.md` with its location, the text of the
neighbouring cells, OCR text (if Tesseract is installed) and a slot for the agent to write
what the picture shows after looking at it.

## Also extracted

- Every sheet as a full Markdown table and a CSV (no row/column/text truncation)
- Hidden sheets (`hidden` / `veryHidden`), hidden rows and columns, merged ranges
- Formulas with their cached values (`formula_manifest.csv`)
- Cell comments, chart count
- EMF/WMF pictures are saved as-is (openpyxl drops them)
- A `conversion_report.json` with counts, warnings and `review_required`

## Install

As a Claude Code skill:

```bash
git clone https://github.com/<you>/xls-skill ~/.claude/skills/xls-skill
pip install openpyxl pillow        # pillow is optional
```

Then just ask Claude to read or convert an Excel file; the skill tells it how to run the
converter, look at the pictures and answer with `sheet!cell` references.

Optional:
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) for text inside pictures
  (add `chi_sim`, etc. for other languages). Found on `PATH` or via `TESSERACT_CMD`.
- Microsoft Excel (Windows, `pywin32`) or LibreOffice to open legacy `.xls`.

## Use the script directly

```bash
python scripts/xlsx_to_md.py report.xlsx other.xlsx some_folder/ -o out/
python scripts/xlsx_to_md.py report.xlsx -o out/ --ocr-lang eng+chi_sim
python scripts/xlsx_to_md.py report.xlsx -o out/ --ocr off --no-floating-inline
```

Output per workbook (`out/<name>/`):

```
<name>.md               sheets as Markdown tables, pictures inline in their cells
<name>.images.md        one entry per picture: location, nearby text, OCR, description slot
images/                 every picture
sheets_csv/             sheet values; picture cells hold [image: images/...]
image_manifest.csv      sheet, cell, kind, file, package part, alt text, size, OCR, nearby text
formula_manifest.csv    formula and cached value per cell
shapes_manifest.csv     text boxes / shapes with text and anchor cell
conversion_report.json  counts and warnings
```

## How it works

An `.xlsx` file is a zip of XML parts. Cell values come from `openpyxl`; pictures come from
the parts themselves:

- **WPS:** find `=DISPIMG("ID",…)` in the cell's formula/value → look the ID up in
  `xl/cellimages.xml` → follow its relationship to `xl/media/…`.
- **Microsoft 365:** a picture cell carries `vm="N"` → `xl/metadata.xml` value-metadata block N
  → `XLRICHVALUE` future-metadata → rich value in `xl/richData/rdrichvalue.xml` →
  `_rvRel:LocalImageIdentifier` → `richValueRel.xml` → `xl/media/…`.
- **Floating:** sheet → drawing part → anchors (`from` cell) → `blip` → `xl/media/…`.

## Test

```bash
python tests/run_tests.py
```

Builds two synthetic workbooks (no real data): one with Microsoft 365 in-cell pictures, a
floating picture, a text box, a merged range, a formula and a hidden sheet; one with WPS
`DISPIMG` pictures including one deliberately missing. Then converts them and checks each
picture landed in the right cell (19 checks; the OCR check is skipped without Tesseract).

It has also been run on real WPS work instructions and BOM workbooks: 12 of 12 workbooks with
every in-cell picture resolved and placed; one work instruction alone had 115 floating photos
and 24 text boxes, all extracted with their anchor cells.

## Limitations

- Charts are counted, not rendered. Pictures in headers/footers and comments are not extracted.
- Arrows/callouts drawn as separate shapes keep their relation to a photo only in a page
  render; export the sheet to PDF/PNG when layout matters.
- Cell colours / conditional formatting are not decoded.
- `.xls` without Excel or LibreOffice: values only.

## License

MIT

---

## 中文简介

把 Excel 转成 AI 能读的 Markdown，**图片留在它所在的单元格里**。

- 重点：能读**单元格里的图片**：WPS 的"嵌入单元格图片"（单元格里只显示 `=DISPIMG(...)`），以及 Microsoft 365 的"放在单元格中"（单元格显示 `#VALUE!`）。每张图都取出来存成文件，再放回表格里原来那一格。
- 浮在表格上的图片、文本框里的文字（比如表单编号）也一起取出，记下锚定的单元格。
- 每张图都有一条记录：位置、旁边单元格的文字、OCR 文字，以及留给 AI 看图后写描述的位置。
- 同时输出完整表格（不截断）、公式、合并单元格、隐藏的工作表/行/列，以及一份带警告的转换报告。

用法：`python scripts/xlsx_to_md.py 文件或文件夹 -o 输出目录`；自测：`python tests/run_tests.py`。
