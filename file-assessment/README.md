# docstats — table & figure statistics for Word and PDF

![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![License](https://img.shields.io/badge/license-MIT-green)

Scan a whole folder of **Word (`.docx`, `.doc`)** and **PDF** files and answer:
how many **tables**, how many **figures**, how many **complex tables**, how many
tables **break across a page**, how many are **nested inside another table**, how
many contain **merged cells** — and click any number to see *which file, which
page*.

Built for document QA: the checks it runs are the ones a reviewer would
otherwise do by hand, page by page.

---

## Install

```bash
pip install -r requirements.txt
```

Requires Python 3.9+. MS Word is optional — it is only used to get real page
numbers for Word files (see [How it works](#how-it-works)). PDFs never need it.

## Run

Double-click `run_gui.bat` (or `docstats_gui.pyw`), or:

```bash
python docstats_gui.pyw
```

Open a folder and scan it immediately:

```bash
python docstats_gui.pyw "D:\Documents"
```

Command line only, writing an Excel report:

```bash
python cli.py "D:\Documents" --excel report.xlsx --list
```

Change the area a picture must cover to count as a figure (default 25% of the
page, `0` counts every image):

```bash
python cli.py "D:\Documents" --fig-min-area 10
```

## Using the GUI

1. **Browse…** to a folder, then press **Scan**. Word and PDF files are scanned
   in the same pass.
2. The tiles along the top are folder-wide totals. **Clicking a tile filters the
   list below to exactly that category** (complex / split / nested / merged /
   split without header) — always **across every file**: clicking a tile clears
   any file filter and the search box, so the number on the tile matches the
   number of rows below it. The **Files** and **Figures** tiles are clickable
   too and jump to the *By file* / *Figures* tab.
3. **By file** tab: one row per file, with a *Type* column (Word/PDF).
   **Double-click a file** to jump to the Tables tab already filtered to it.
   The *Small figs* column counts figures dropped for being under the area
   threshold. The **Figures** tab shows *Width×Height (cm)* and *% of page*;
   tick **Show the small figures that were dropped** to check that the tool
   really did drop only logos and stamps (grey rows are not counted).
4. **Tables** tab: one row per table, with a **Page** column. **Double-click a
   row** (or press *Open document*) and the Word file opens with **that exact
   table selected**; a PDF opens **at the page** the table is on. The panel at
   the bottom spells out the page, the size, the merged-cell count, which row
   the page break falls after, and so on.
5. **Export Excel** writes an `.xlsx` with four sheets: Summary, By file,
   Tables, Figures — auto-filter on, header row frozen.

Row colours: red = split across pages, purple = nested, yellow = complex,
blue = has merged cells.

## What each number means

| Metric | How it is computed |
|---|---|
| **Tables** | Includes tables nested inside other tables. Children are numbered `Table 3.1`, `Table 3.1.1`. |
| **Figures** | Images, charts, SmartArt and vector drawings. Text boxes, table ruling lines and full-page scans are **not** counted. By default only figures covering **≥ 25% of the page** count, which drops logos, stamps and icons — adjustable in the toolbar, or turn the limit off to count everything. Figures whose size cannot be measured are still counted, so nothing is dropped silently. |
| **Complexity score** | The number of **visible cells** (a merged block counts as one). The default "complex" threshold is **≥ 50 cells**, adjustable in the toolbar — changing it recomputes instantly, without rescanning. |
| **Split across pages** | The table occupies two or more pages. The *Break after row* column says exactly which row the break falls after. |
| **Nested tables** | A table inside a cell of another table. Both the parent and the child are counted. |
| **Merged cells** | A table with nothing merged has exactly `rows × columns` cells. However many cells are missing from that number is how many were **merged away**: `merged = rows × columns − actual cells`. The *Horiz./Vert. merges* columns count merge **regions** (one cell spanning 4 columns is 1 region but 3 merged-away cells). |
| **Split without header** | The table breaks across a page but the header row is not repeated on the next one. This is the single most common layout defect. |

A `.docx` file and a PDF printed from that same document produce **identical**
numbers — verified table by table: row count, cell count, merges, page, break
position, and whether the header repeats.

## How it works

**Word** — reads the OOXML inside the file directly (`gridSpan`, `vMerge`,
`tblHeader`, nested tables), so the table structure is exact. But a `.docx` file
does **not** store pagination (Word computes it when opening the file), so to
know which page a table is on, or whether it breaks, the tool borrows MS Word:
it opens the file hidden and read-only. Nothing is written back to the file.

| | **Word: use MS Word** on (default) | off |
|---|---|---|
| Speed | ~1–3 s/file | very fast |
| Page numbers, split detection | ✅ exact | ❌ unavailable |
| Legacy `.doc` files | ✅ readable | ❌ skipped |

*Figure size* is the size **rendered on the page** (Word's `wp:extent`, the image
frame in a PDF), not the pixel dimensions of the image file — so a 4000px image
scaled down to a 2cm stamp is still dropped correctly. Page area comes from the
paper size declared in the file (`w:pgSz`), defaulting to A4 when absent.
Cross-checked: the same document as `.docx` and as printed PDF yields the same
measurements (e.g. a 16×12cm figure = 31.8% of a Letter page).

**PDF** — a PDF stores no table structure, only strokes and glyphs, so tables
are recovered from the page geometry (PyMuPDF). Page numbers are always
available, no Word needed. PDF scanning is fast (~0.05 s/file).

- *merged cells*: a grid is rebuilt from the ruling lines; a cell wider than one
  column or taller than one row is merged. The final number still follows
  `rows × columns − actual cells`.
- *nested tables*: a table entirely inside another table's bounding box.
- *split tables*: the bottom fragment of page N and the top fragment of page N+1
  are joined when they share the same column count and the same vertical rule
  positions.
- *repeated header*: whether the first row of the continuation matches the first
  row of the previous fragment.

PDFs have two limits, and the tool reports both on the **Log / Errors** tab
instead of failing quietly:

- **Borderless tables** are not detected by default → tick *PDF: detect
  borderless tables* (which infers columns from text alignment, and is more
  prone to false positives, hence off by default).
- **Scanned PDFs** (each page is one image) have no text to work with → the tool
  reports "N/M pages are scanned images (OCR needed)" rather than silently
  returning zero tables.
- A table **merged down to a single cell** is just a rectangle in a PDF,
  indistinguishable from a text box, so it is not counted as a table. The Word
  version of the same document still counts it.

Password-protected and corrupt PDFs are reported clearly and do not abort the
rest of the scan.

## Command line

```
python cli.py FOLDER [options]

  --no-recursive       do not scan subfolders
  --no-word            Word files: skip MS Word (no page numbers / split detection)
  --pdf-borderless     PDF: also detect tables without ruling lines
  --threshold N        cells needed to count as complex (default 50)
  --fig-min-area PCT   min % of the page for a figure to count (default 25, 0 = all)
  --excel FILE.xlsx    write an Excel report
  --csv FOLDER         write files.csv, tables.csv and figures.csv
  --list               print every table that has an issue
```

## Use as a library

```python
from docstats import analyze_folder

result = analyze_folder(r"D:\Documents", threshold=50, use_word=True)
print(result.totals())

for t in result.all_tables():
    if t.is_split:
        print(t.file_name, t.label, "pages", t.page_start, "->", t.page_end)
```

## Project layout

```
run_gui.bat          launch the GUI
docstats_gui.pyw     GUI entry point (no console window)
cli.py               command line / report generation
docstats/
  analyzer.py        Word analysis + shared orchestration
  pdf_analyzer.py    PDF analysis
  gui.py             tkinter interface
  export.py          Excel / CSV export
```

## Requirements

- Python 3.9+
- `python-docx`, `pymupdf`, `openpyxl`, `pywin32` (see `requirements.txt`)
- MS Word — optional, only for page numbers in Word files

## License

MIT — see [LICENSE](LICENSE).
