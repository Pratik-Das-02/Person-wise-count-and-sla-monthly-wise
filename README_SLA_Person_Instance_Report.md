# Final SLA Person + Instance Report

A Google Colab Python script that creates an Excel report from a Submission History worksheet.

## Features

- Finds `Submission History` or `Submissions History`.
- Detects Assigned To, Instance, Actual SLA and Overall SLA columns.
- Cleans employee and Instance text.
- Groups recognized Morning, Afternoon, Evening, Night and Shift A-G entries under the base employee name.
- Keeps exact Assigned To names without a recognized shift separate with the `(W)` suffix.
- Removes blank or zero Instance records and invalid Assigned To records.
- Converts SLA values to seconds for calculation.
- Uses Overall SLA minus one minute when Actual SLA is blank and Overall SLA is available.
- Calculates counts and average Actual SLA by employee + Instance.
- Exports `SLA Summary`, `Processed Data` and `Validation` worksheets.
- Formats the Excel workbook and downloads it from Google Colab.

## Requirements

Run in Google Colab. Install dependencies in a separate cell:

```python
!pip -q install pandas openpyxl xlrd
```

## How to run

1. Open [Google Colab](https://colab.research.google.com/).
2. Create a new notebook.
3. Copy the contents of `Final_SLA_Person_Instance_Report.py` into a code cell, or upload the `.py` file and run it from Colab.
4. Upload an Excel workbook when prompted.
5. Ensure the workbook contains a worksheet named `Submission History` or `Submissions History`.
6. The script downloads `SLA_Person_Instance_Summary.xlsx`.

## Output worksheets

| Worksheet | Description |
|---|---|
| `SLA Summary` | Person heading rows and Instance-wise counts / average Actual SLA |
| `Processed Data` | Filtered source rows with output person, assignment type and calculated Actual SLA |
| `Validation` | Row removal, inclusion and SLA availability / fallback counts |

## Business rules

- Exact name without a recognized shift: `Abhinav` → `Abhinav (W)`.
- Recognized shift suffix: `Abhinav Morning` or `Abhinav Shift A` → `Abhinav`.
- All recognized shifts are combined under the base person name.
- Person heading rows show Count only; Average Actual SLA is blank.
- Average Actual SLA is calculated at Person + Instance level.
- Missing Actual SLA is replaced with Overall SLA minus 60 seconds, when Overall SLA exists.

## Important scope notes

- Only the first uploaded workbook is processed.
- Only the Submission History / Submissions History worksheet is processed.
- Counts are row counts, not distinct Submission ID counts.
- The script does not perform fuzzy employee-name matching.
- The script does not classify Within SLA / Out of SLA.
