# Variance Insights

**Converts a budget-versus-actual data table into a formatted variance analysis slide and workbook, with supporting commentary drafted automatically.**

Preparing a variance analysis typically involves calculating the difference between actual and budgeted figures, determining which variances are significant enough to explain, and writing commentary describing what occurred, formatted for client presentation. This application performs the calculation and classification steps automatically and drafts an initial commentary for each material variance, which the reviewer can then refine with the underlying explanation.

![Generated variance slide](screenshots/variance-slide.png)

## What it does

- **Detects the relevant columns automatically.** Given a spreadsheet with a label column and two numeric columns, such as "Q4 Actual" and "Q4 Budget" or "Actual" and "Prior Year," the application identifies which column represents which value.
- **Classifies each line item by type and direction.** Values are categorized as currency, percentage, or count, and each is evaluated based on whether an increase represents a favorable or unfavorable result, using the wording of the line item itself. A "Tax Revenue" line and a "Tax Expense" line are treated correctly as opposites, rather than matched only on the shared word "tax."
- **Flags material variances.** A configurable threshold, set to 10 percent by default, determines which variances are significant enough to warrant commentary.
- **Drafts commentary for each material variance.** For example: "Days Sales Outstanding came in up 6 days (15.0%) versus Budget, an unfavorable variance that warrants explanation." Each draft includes a placeholder marking where the underlying explanation should be added by the reviewer.
- **Exports two formats.** A KPMG-styled PowerPoint slide with color-coded favorable and unfavorable values, and a formatted Excel workbook containing the variance table and a separate commentary sheet.
- **Displays both outputs before download.** The result page shows the PowerPoint slide and every Excel sheet side by side. PowerPoint is used for slide rendering on Windows when installed, with LibreOffice as the fallback. A browser text and data preview remains available when neither renderer is installed.

## Setup

- **Mac:** double-click `Start on Mac.command`
- **Windows:** double-click `Start on Windows.bat`

A sample dataset is included at `sample_files/sample_variance_input.xlsx`, a fictional Q4 advisory practice scorecard that combines currency, percentage, and count metrics to demonstrate all three classification types.

Manual setup:
```bash
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py
```
Then open `http://127.0.0.1:5060`.

## Design notes

- **Commentary is drafted, not fully generated, by design.** Determining why a figure differed from budget requires context beyond what a spreadsheet provides. The application drafts the factual portion of each variance statement, covering magnitude, direction, and classification, and clearly marks where explanatory context is needed rather than generating an unsupported explanation.
- **A single materiality threshold is used rather than commentary on every line.** A transparent, adjustable percentage threshold determines which variances are flagged, consistent with how variance analyses are typically prepared.
- All processing takes place locally. No file is transmitted to an external service.
