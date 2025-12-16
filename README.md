# Excel Monthly Structure Auditor

A read-only auditing script that scans a year/month folder layout and validates that each monthly Excel workbook contains the expected sheets and minimum required columns. It also compares which months exist on disk vs which "Source" values exist inside the combined workbook.

This is useful when dashboards show missing periods ("blanks") and you need to quickly identify whether the cause is missing files, missing sheets, or shifted headers/columns.

## What it checks

### 1) Monthly files (per month folder)
For each month folder found, it checks:

- The monthly file exists (e.g. `CanalesSCMG.xlsx`)
- P&L sheet exists (tries multiple names)
- BS sheet exists
- DB sheet exists
- Minimum required columns exist (structure sanity check)

It also attempts to infer the header row when headers are shifted.

### 2) Combined workbook coverage
If the combined workbook exists (e.g. `CanalesSCMG combined.xlsx`), it reads the `Source` column from each combined sheet and reports:

- Months present in folders but missing in combined
- Sources present in combined but not present on disk

## Expected folder layout

The script discovers months using this pattern:

