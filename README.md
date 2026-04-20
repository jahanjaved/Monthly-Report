# Flexible Master Monthly Report Builder

## What changed
- No fixed required file list
- No `orignal(2).xlsx` dependency
- Generate button becomes active when at least one `.xlsx` file is loaded
- The website auto-selects the best workbook as template
- Matching sheets from uploaded workbooks are merged into one master workbook
- A `Merge_Summary` worksheet is added to the output

## How to use
1. Delete the old GitHub website files first.
2. Upload all files from this zip to GitHub.
3. Hard refresh the website.
4. Upload one or more monthly Excel files.
5. Click **Generate master workbook**.

## Important note
This version performs a flexible browser-side merge by summing matching numeric cells into the detected template workbook.
If some contractor files have very different layouts, those parts may not line up perfectly and you should review the `Merge_Summary` sheet after download.
