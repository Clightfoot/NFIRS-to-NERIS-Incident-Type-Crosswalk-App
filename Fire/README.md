# NFIRS to NERIS Incident Type Crosswalk App

This local app crosswalks an uploaded NFIRS incident type file to NERIS incident type category fields.

## Start the app

Double-click `launch_app.bat`, then open:

```text
http://127.0.0.1:8765
```

## Uploads

The app accepts:

- `.csv`
- `.xlsx`
- `.xlsm`
- `.xls`

For older binary `.xls` files, the app first tries the local reader. If that fails, it tries to use Microsoft Excel on this computer to convert the file quietly to `.xlsx` before analysis. If Excel is not installed or cannot open the file, save a copy as `.xlsx` or `.csv` and upload that copy.

For Excel workbooks with multiple sheets, the app uses the best non-empty sheet it can find and shows the sheet name in the browser summary.

## Outputs

After analysis, the page provides:

- A downloadable Excel summary workbook
- A downloadable CSV with the original rows plus NERIS crosswalk columns

The summary workbook includes:

- Overall match counts
- NERIS category totals
- NFIRS code totals
- Crosswalked source rows
