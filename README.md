# EasyXLS GUI

`EasyXLS-GUI.ps1` is a PowerShell Windows Forms tool that opens an Excel workbook and builds one consolidated transaction table.

## Requirements

- Windows PowerShell.
- Microsoft Excel installed locally for legacy `.xls` workbooks.
- `.xlsx` and `.xlsm` workbooks are read directly by the script without Excel COM.

## How To Run

Double-click:

```text
EasyXLS.exe
```

Keep `EasyXLS-GUI.ps1` in the same folder as `EasyXLS.exe`. The executable launches that script.

From PowerShell:

```powershell
Set-Location "C:\Users\Forensic Workstation\Documents\Cursor Workspaces\EasyXLS"
.\EasyXLS-GUI.ps1
```

If script execution is blocked for the current session:

```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Users\Forensic Workstation\Documents\Cursor Workspaces\EasyXLS\EasyXLS-GUI.ps1"
```

## Detection Rules

- The script scans each worksheet and picks the column with the most date-like cells as that sheet's date column.
- For each row with a valid date, it scans left-to-right across the same row and uses the first currency-formatted value it finds, excluding the date cell. For `.xlsx` files with plain numeric transaction cells, it falls back to the first non-date numeric value on that row.
- The final grid has `Date` in the first column and one column per worksheet.
- If the same sheet has multiple transactions for the same date, the first matching row for that date is kept.
