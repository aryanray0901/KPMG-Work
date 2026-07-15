# Variance Insights

**Turns a budget-vs-actual data table into a client-ready variance analysis slide and workbook — with the commentary already drafted.**

This is the task that eats a rising senior's first few weeks on an engagement: take a set of Actual vs. Budget (or Prior Period, or Forecast) figures, compute the variances, figure out which ones are big enough to matter, and write a few sentences per line explaining what happened — formatted the way the client expects to see it. This app does the mechanical 80% of that automatically and hands you a clean starting point for the last 20%: the actual "why."

![Generated variance slide](screenshots/variance-slide.png)

## What it does

- **Auto-detects your columns** — point it at a spreadsheet with a label column and two numeric columns (any reasonable names: "Q4 Actual"/"Q4 Budget", "Actual"/"Prior Year", etc.) and it figures out which is which.
- **Classifies each line item** by type — currency, percentage, or count — and by whether an increase is good news or bad news, using the metric's own name (revenue/profit/margin lines: higher is better; cost/expense/tax lines: lower is better). A "Tax Revenue" line and a "Tax Expense" line get opposite treatment correctly, not just pattern-matched on the word "tax."
- **Flags what's material** against a threshold you set (default 10%), so the commentary focuses on what actually deserves a sentence, not every rounding-level wiggle.
- **Drafts the commentary** — a factual, professionally-worded starting sentence per material variance ("Days Sales Outstanding came in up 6 days (15.0%) versus Budget, an unfavorable variance that warrants explanation."), with a clear `[Add driver commentary]` marker where the real analysis — the "why" — still needs a human who actually knows the client.
- **Exports both formats** — a KPMG-styled PowerPoint slide (color-coded favorable/unfavorable, ready to paste into a deck) and a formatted Excel workbook (variance table + a separate commentary sheet).
- **Shows you the real rendered slide** before you download, using the same LibreOffice-based rendering approach as Deck Refresh.

## Try it

- **Mac:** double-click `Start on Mac.command`
- **Windows:** double-click `Start on Windows.bat`

A sample dataset is included — `sample_files/sample_variance_input.xlsx`, a fictional Q4 advisory-practice scorecard mixing currency, percentage, and count metrics, to exercise all three classification paths at once.

Manual setup:
```bash
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py
```
Then open `http://127.0.0.1:5060`.

## Design notes

- **Why draft commentary instead of trying to fully explain variances?** No tool can know *why* revenue beat budget without more context than a spreadsheet provides. Pretending otherwise would mean either hallucinating a plausible-sounding reason (actively misleading) or staying silent (not useful). Drafting the factual half — magnitude, direction, favorable/unfavorable — and clearly marking the part that needs a human is the honest version of "helpful."
- **Why a materiality threshold instead of commentary on everything?** Real variance reports don't explain every line; they draw attention to what matters. A single global % threshold is a simple, transparent rule you can see and adjust, rather than a black-box judgment call.
- All processing is local. Nothing is uploaded to an external service.
