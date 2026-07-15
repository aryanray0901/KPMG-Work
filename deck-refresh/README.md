# Deck Refresh

**A local web application that updates the numbers in a PowerPoint or Excel file using new source data, while preserving all existing formatting, layout, and charts.**

Refreshing a quarterly report by hand typically involves retyping figures into table cells and charts one at a time, with a meaningful risk of errors or inconsistent formatting. This application automates that process: the user uploads the file to be updated and the new data, reviews the matched values, and receives the same file back with figures updated and formatting intact.

![Executive summary slide, rendered by the app](screenshots/executive-summary.png)

## What it does

- **Identifies every number in the file.** Text boxes, table cells, and charts (pie, bar, and line) in PowerPoint; cells and charted ranges in Excel. Each value is associated with a label inferred from its surrounding text or row and column headers.
- **Matches values by meaning, not position.** Labels are matched against the new data using fuzzy matching, so a Q3 deck can be updated from Q4 data without manual remapping. A keyword-conflict check prevents mismatches such as a "Costs" figure being filled in with "Revenue" data due to similar row names.
- **Updates headings, not just values.** The tool detects the reporting period referenced in the file and in the new data, and if they differ, updates titles, table headers, and chart labels accordingly.
- **Leaves unmatched values unchanged.** Any figure without a confident match is flagged for manual review rather than replaced automatically.
- **Displays the result before it is finalized.** A synchronized viewer renders the original and updated PowerPoint files side by side, using LibreOffice, with shared navigation and zoom so changes can be reviewed slide by slide.
- **Preserves the file structure.** PowerPoint charts are updated using `python-pptx`'s native data-replacement method, which retains colors, legends, and chart type. Excel charts are left untouched; because they reference live cell ranges, they update automatically when the underlying cell values change.

![Data table and chart slide, before and after](screenshots/data-slide.png)

## Tech stack

Python, Flask, python-pptx, openpyxl, pandas, RapidFuzz for fuzzy matching, LibreOffice and PyMuPDF for slide rendering, and a vanilla HTML, CSS, and JavaScript frontend.

## Setup

- **Mac:** double-click `Start on Mac.command`
- **Windows:** double-click `Start on Windows.bat`

The first run installs dependencies automatically and opens the application in a browser. A fictional KPMG-styled sample deck is included in `sample_files/`. Upload `kpmg_advisory_q3_original.pptx` together with `kpmg_advisory_q4_data.xlsx` to test the workflow.

Manual setup:
```bash
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py
```
Then open `http://127.0.0.1:5050`.

> Rendered slide previews require LibreOffice (free, available at libreoffice.org) or, on Windows, Microsoft PowerPoint automation. Without either, the application still functions fully for matching, editing, and downloading files, using a simplified data and chart comparison view in place of rendered slide images.

## Included sample

`sample_files/` contains a fictional nine-slide KPMG-styled advisory practice review, including KPI cards, six data tables, six native PowerPoint charts, and 120 numeric targets, along with the corresponding Q4 source data, a verified expected output file, and an automated verification script.

```bash
python tools/verify_sample_update.py
```

Most recent verification result: 120 of 120 targets matched, all table and chart values correct, all headings updated from Q3 to Q4, and slide, shape, and chart structure fully preserved.

## How it works

1. The user uploads the file to update, in `.pptx` or `.xlsx` format, along with the new data.
2. The application extracts every numeric target and its inferred label.
3. Labels are matched against the new data using fuzzy matching with a keyword-conflict check, so terms such as Revenue, Costs, Actual, and Budget are not cross-matched incorrectly.
4. Reporting-period headings are detected and updated to reflect the new data.
5. The user reviews and confirms each match, with the option to edit or exclude individual values.
6. The original and updated files are compared side by side before the updated file is downloaded.

All processing takes place locally. No file is transmitted to an external service.
