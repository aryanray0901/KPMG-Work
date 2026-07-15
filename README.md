# KPMG Work

A collection of internal tools for automating recurring tasks in quarterly reporting and client deliverable preparation.

All data included is fictional and intended for demonstration purposes only.

## [Deck Refresh](deck-refresh/)

Updates the numbers in a PowerPoint or Excel file using new source data, while preserving every color, font, layout, and chart. The user uploads a file and the updated figures, reviews the proposed changes, and receives the same file back with the values updated. Includes a synchronized side-by-side viewer that renders the original and updated presentation so changes can be verified before the file is finalized.

![Deck Refresh](deck-refresh/screenshots/executive-summary.png)

## [Variance Insights](variance-insights/)

Converts a budget-versus-actual data table into a formatted variance analysis slide and workbook, with supporting commentary drafted automatically. The tool calculates variances, classifies each line item as favorable or unfavorable based on the type of metric, flags variances that exceed a defined materiality threshold, and drafts commentary describing the size and direction of each material variance. Explanatory context is left for the reviewer to complete.

![Variance Insights](variance-insights/screenshots/variance-slide.png)

## [Engagement Hub](engagement-hub/)

A workbench combining three tools for recurring engagement work. The Action Tracker converts meeting notes into a reviewed action item list, an Excel tracker, a follow-up email draft, and a summary slide. The Executive Summary module converts a period-over-period data table into a KPI-and-narrative summary slide. The Benchmarking module compares a client's metrics against a peer set and produces a quartile-ranked benchmarking slide and workbook.

![Engagement Hub](engagement-hub/screenshots/benchmarking.png)

## Running any tool

Each folder is a self-contained local Flask application. See the README in each folder for setup instructions. All three:
- Run entirely on the local machine; no data is transmitted externally
- Include a double-click launcher for Mac and Windows
- Include a fictional sample dataset for immediate testing
