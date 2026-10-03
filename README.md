# Business Time & Mileage Tracker

`Business-Time-Mileage-Tracker.xlsx` is a ready-to-use Excel template for a small rental-property business.

## Use the template

1. Open the workbook and read the `Instructions` sheet.
2. Set the reporting year on `Summary`.
3. Replace the sample entries in `Time Log` and `Mileage Log`.
4. Add new records in the first blank table row. The tables, totals, and charts update automatically in Excel.
5. Save a separate copy for each tax year.

The workbook stores decimal hours, miles driven, properties, business purposes, locations, and notes. Property and activity columns use customizable dropdown lists. It contains no macros.

## Rebuild the workbook

Install `openpyxl`, then run:

```powershell
python -m pip install openpyxl
python .\create_template.py
```

The generator also validates the workbook structure, formulas, dropdowns, charts, and absence of VBA content.
