# KPMG Work

A collection of small internal tools built to automate the tedious parts of quarterly reporting and client deliverable prep — the kind of work that eats a rising senior's time during busy season.

All data in both tools is fictional, built for demonstration only.

## [Deck Refresh](deck-refresh/)

Updates the numbers in a PowerPoint or Excel file from new data — while keeping every color, font, layout, and chart exactly as it was. Upload a deck, upload the new quarter's numbers, review the matches, and get back the same file with updated figures. Includes a synchronized side-by-side viewer that renders the actual before/after presentation so you can see precisely what changed before committing to it.

![Deck Refresh](deck-refresh/screenshots/executive-summary.png)

## [Variance Insights](variance-insights/)

Turns a budget-vs-actual data table into a client-ready variance analysis slide and workbook, with commentary already drafted. Computes variances, classifies each line as favorable or unfavorable based on what kind of metric it is, flags what's material against a threshold you set, and drafts the factual half of the commentary — leaving a clear marker for the human judgment part ("why") that no tool should pretend to know.

![Variance Insights](variance-insights/screenshots/variance-slide.png)

## Running either tool

Each folder is a fully self-contained local Flask app — see the README inside each for setup. Both:
- Run entirely locally; no data leaves your machine
- Ship with a double-click launcher for Mac and Windows
- Include a fictional sample dataset so you can try it immediately
