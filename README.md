# Flexible Master Monthly Report Builder

This website package does not require `orignal(2).xlsx` and does not require a fixed number of contractor files.

## What it does
- Accepts one or many Excel workbooks
- Automatically detects the best template/base workbook
- Merges compatible monthly-report sheets into one master workbook
- Preserves formulas from the template workbook
- Adds a `Merge_Summary` sheet showing:
  - which files were loaded
  - which workbook became the template
  - which sheets were merged
  - which sheets were skipped because the layout was too different

## Important
This version uses a **compatible-sheet merge** approach:
- same sheet type
- similar sheet size/layout
- same cell position for numeric values

That makes it flexible and much safer than forcing exact filenames.

## How to use
1. Upload all files in this ZIP to your GitHub Pages repo.
2. Open the website.
3. Upload any mix of monthly Excel files.
4. Click **Review uploaded files**.
5. Click **Generate master workbook**.
6. Open the downloaded workbook in Excel and save once to refresh calculations.

## Recommended
If you have a clean original/master template workbook, upload that too. The app will usually choose it automatically as the base workbook.
