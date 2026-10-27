# Weekly Reports Automation

Automatically creates the weekly report workbook, names it with the correct week-ending date, and updates the date inside the file, so nobody has to rebuild it by hand every week.

![PowerShell](https://img.shields.io/badge/PowerShell-5391FE?logo=powershell&logoColor=white)
![Python](https://img.shields.io/badge/Python-3776AB?logo=python&logoColor=white)
![Excel](https://img.shields.io/badge/Excel-217346?logo=microsoftexcel&logoColor=white)

## The problem

Every week, a set of Excel reports has to be recreated before the daily data can be added. For each one, someone had to:

1. Copy last week's workbook.
2. Rename it with the new week-ending date.
3. Open it and change the date inside.

Across 3 reports and 7 managers, this was repetitive, easy to get wrong (wrong date, wrong file name), and took time away from actual work.

## The solution

One script does the setup in a few seconds:

- Finds the most recent report workbook in the reports folder.
- Works out the upcoming **week-ending Sunday**.
- Saves a copy with the new name, e.g. `KEV EON Reports WE 01.11.26.xlsx`.
- Updates the date cell inside the workbook (`B2` → `WE 01.11.26`).

Managers no longer have to do the initial creation. The files are ready to receive the week's data.

## How it works

```
updateReports.ps1                          edit_excel.py
─────────────────                          ─────────────
1. find newest .xlsx in Reports folder
2. calculate next Sunday (week end)
3. copy it to "<name> WE dd.MM.yy.xlsx"
4. call Python ───────────────────────────▶ 5. open the new copy (openpyxl)
                                            6. write "WE <date>" into B2
                                            7. save
```

| File | Purpose |
|------|---------|
| `updateReports.ps1` | Main script. Finds the latest report, calculates the week-ending date, copies and renames the file, then calls the Python script. |
| `edit_excel.py` | Opens the new workbook with `openpyxl`, updates the date cell, and saves it. |
| `MASTER/` | Master/template reports. |
| `KEV EON Report W.E 00.00.00.xlsx` | Example/template workbook with a placeholder date. |

### Date logic

The week always ends on a **Sunday**. If you run the script on a Sunday, that day is used; on any other day it uses the coming Sunday.

| Run on | Week ending used |
|--------|------------------|
| Mon 26 Oct 2026 | 01.11.26 |
| Wed 28 Oct 2026 | 01.11.26 |
| Sun 01 Nov 2026 | 01.11.26 |

## Requirements

- Windows with PowerShell 5.1 or later
- [Python 3](https://www.python.org/downloads/) available on `PATH`
- [openpyxl](https://openpyxl.readthedocs.io/)

```powershell
pip install openpyxl
```

## Setup

1. Clone the repo:
   ```powershell
   git clone https://github.com/garciahelena/weekly-reports-automation.git
   cd weekly-reports-automation
   ```
2. Open `updateReports.ps1` and set the two folders at the top:
   ```powershell
   $SourceDir = "C:\Path\To\Reports"   # folder holding last week's report(s)
   $DestDir   = "C:\Path\To\Output"    # where the new file should be saved
   ```
3. (Optional) Change the file name pattern in `$newName` to match your report.
4. If your date lives in a different cell, change `ws["B2"]` in `edit_excel.py`.

## Usage

Run from the repo folder (the script calls `edit_excel.py` by relative path):

```powershell
powershell -ExecutionPolicy Bypass -File .\updateReports.ps1
```

Expected output:

```
Excel updated successfully
 Running Python
Report created and updated: KEV EON Reports WE 01.11.26.xlsx
```

`edit_excel.py` can also be run on its own:

```powershell
python edit_excel.py "path\to\report.xlsx" "01.11.26"
```

### Run it automatically every week (optional)

Use Windows Task Scheduler to run the script once a week, for example every Monday morning:

1. Open **Task Scheduler** → **Create Basic Task**.
2. Trigger: **Weekly**, pick the day and time.
3. Action: **Start a program**
   - Program: `powershell.exe`
   - Arguments: `-ExecutionPolicy Bypass -File "C:\path\to\updateReports.ps1"`
   - Start in: `C:\path\to\weekly-reports-automation`

## Impact

- Removes a manual weekly task that 7 managers previously shared.
- Consistent file names and dates, so no more typos in the week-ending date.
- Setup that used to take several minutes per report now takes seconds.

## Known limitations and next steps

- [ ] Folder paths are hard-coded in the script; move them to a config file or parameters.
- [ ] Handles one report at a time; extend it to loop over every report in the folder.
- [ ] Only updates one cell (`B2`); make the cell(s) configurable.
- [ ] Add a check for when no `.xlsx` file is found in the source folder.
- [ ] Add a dry-run mode that prints what would be created without writing files.

## Notes

The workbooks in this repo contain placeholder data only. No real business or customer data should be committed.
