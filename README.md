# Expense-RG (Expense Report Generator)

**Combines per-location QuickBooks expense reports into one Excel workbook with month-by-month, year-over-year and year-to-date comparisons, built in seconds instead of hours.**

## Background

Built in 2020 to automate bookkeeping at a multi-location metals recycling company. Every location and set of books had its own monthly expense report, and management wanted one view comparing every location this year against last year. That used to mean copying a dozen reports into a master spreadsheet by hand. Expense-RG builds the whole workbook automatically from the reports that [Expense Extract](https://github.com/Weston-Matthews/Expense-Extract) produces.

It was one piece of a broader automation suite Weston built for the company's back office, which together saved roughly $100K–$120K a year in labor.

## Features

- Pick any subset of locations from a list defined in **`config.json`** (no names in the code)
- Copies each location's current-year and prior-year report into its own tab, **keeping formatting**: fonts, fills, borders, number formats, column widths, row heights and freeze panes
- Detects the latest month in the reports automatically, so it handles partial years
- Builds a **Summary** sheet with:
  - a current-year table and a prior-year table, with one row per location and one column per month, plus **Average/Location** and **Total** columns
  - monthly totals across all locations and a running average
  - **Summary YTD**: each location's average monthly expense this year (through the latest month) vs. its average month last year, with **Dollar Change** and **% Change**
  - **Latest month** (for example "September"): that month this year vs. the same month last year, with Dollar Change and % Change
- Uses Excel formulas throughout (`=SUM`, `=AVERAGE`, `=IFERROR(...)`), so the summary recalculates when someone edits a value
- Conditional number formats (negative changes show red, in parentheses) and banded rows
- Ships as a stand-alone Windows `.exe` via PyInstaller

## How it works

1. **Ask for the year** being reported. The prior year is derived from it.
2. **Load the locations** from `config.json` and let the user pick which to include (`easygui.multchoicebox`).
3. **Select files.** For each chosen location, the user picks this year's and last year's expense report (`.xlsx`).
4. **Copy reports in.** Each report is copied cell by cell, with styles, into a tab named `<code> <yy>` (for example `NORTH 20`, `NORTH 19`).
5. **Find the latest month.** `detect_latest_month()` scans each sheet's header row, and the earliest "latest month" across locations sets the comparison window.
6. **Fill the summary.** `find_month_total()` reads each month's total from the bottom of its column, then the summary tables, totals, averages and change formulas are written. A location's row is fixed by its position in the config, so the layout stays the same from month to month.
7. **Save** `Expense_Report.xlsx` to the working directory.

## Architecture

A single script, `Expense-RG.py`:

| Function | Role |
|---|---|
| `find_config()` / `load_locations()` | Locate and validate `config.json` (falls back to `config.example.json`) |
| `detect_latest_month(sheet)` | Latest month present in a report's header row |
| `copy_cell(write_cell, trans_cell)` | Copies a cell's value and full style between workbooks |
| `find_month_total(sheet, month)` | Bottom-line total for a given month column |
| `main()` | Prompts → copy reports → build current-year, prior-year, YTD and latest-month tables → save |

## Tech stack

Python 3.8 · openpyxl · easygui · JSON config · Poetry · PyInstaller · pytest

## Requirements

- Windows for the packaged `.exe` and the GUI as designed. The script itself is pure Python and also runs anywhere `tkinter` is available.
- Expense reports exported from QuickBooks Desktop (QBFC13 SDK) by **Expense Extract**, one per location per year. This tool doesn't connect to QuickBooks itself.
- Python 3.8
- Poetry

## Install

```bat
git clone https://github.com/Weston-Matthews/Expense-RG.git
cd Expense-RG
poetry install
copy config.example.json config.json
```

Without Poetry: `pip install -r requirements.txt`

## Run

```bat
poetry run python Expense-RG.py
```

To try it without any real data, generate fictional reports first:

```bat
poetry run python samples\make_sample_reports.py sample_reports 2020 9
```

This writes `sample_reports\NORTH 2020.xlsx`, `sample_reports\NORTH 2019.xlsx` and so on, for every location in `config.example.json`, with January–September for 2020 and a full year for 2019.

Run the tests:

```bat
poetry run pytest
```

## Build (stand-alone .exe)

```bat
poetry run pyinstaller -y Expense-RG.spec
```

Output: `dist\Expense-RG\Expense-RG.exe`, with `config.example.json` bundled next to it. Put a `config.json` beside the exe (or in the working directory) to use your own locations. The spec file uses only relative paths.

## Configuration

`config.json` (copy from `config.example.json`):

```json
{
  "locations": [
    {"name": "Site North", "code": "NORTH"},
    {"name": "Site South", "code": "SOUTH"},
    {"name": "Location A", "code": "LOC-A"},
    {"name": "Book A",     "code": "BOOK-A"}
  ]
}
```

| Key | Meaning |
|---|---|
| `name` | Shown in the location picker |
| `code` | Short label used for the tab names (`<code> <yy>`) and the summary rows. Keep it short, with no digits. |
| order | Sets the row order in the summary tables (up to 12 locations) |

`config.json` is git-ignored, so real location names stay out of the repository.

## Sample output

This is a real run against the generated sample reports (`samples/make_sample_reports.py`, fictional data) for **Site North**, **Site South** and **Location A**, for 2020 through September. It produces `Expense_Report.xlsx` with tabs `Summary`, `NORTH 20`, `NORTH 19`, `SOUTH 20`, `SOUTH 19`, `LOC-A 20`, `LOC-A 19`.

Summary, current-year table (rows 2–18):

```
        Jan         Feb         Mar      ...   Sep         Oct Nov Dec   Average/Location   Total for the 9 months
NORTH   $51,141.23  $38,681.27  $54,260.87 ... $50,325.86   -   -   -    $50,551.36         $454,962.28
SOUTH   $50,203.95  $52,690.63  $36,219.03 ... $77,246.53   -   -   -    $53,107.35         $477,966.16
LOC-A   ...                                                              $49,910.50         $449,194.52
(totals row: =SUM per month; "Running Average" below it)
```

The prior-year table (rows 24–40) has the same layout for 2019. On the right are the comparison blocks:

```
                 Summary YTD                                     September
        2020        2019        Dollar Change  % Change    2020        2019        Dollar Change  % Change
NORTH   $50,551.36  $54,592.49  ($4,041)       (7%)        $50,325.86  $57,734.56  ($7,409)       (13%)
SOUTH   $53,107.35  $55,437.70  ($2,330)       (4%)        $77,246.53  $41,833.90  $35,413        85%
LOC-A   $49,910.50  $55,494.55  ($5,584)       (10%)       $47,265.12  $49,618.28  ($2,353)       (5%)
```

Every summary value is a live formula. For example, NORTH's row is `=AVERAGE(D4:L4)`, `=SUM(D4:L4)`, `=P4`, `=P26`, `=T4-U4` and `=IFERROR((T4-U4)/U4,0)`. Negative changes show in red, in parentheses.

## Limitations

- Expects the column layout produced by Expense Extract: month names in row 1, followed by a `TOTAL` column, with totals in the last row.
- The summary layout has room for 12 locations.
- Tabs are matched to locations by the two-digit year in their names, so location codes shouldn't contain digits.
- If a report file is skipped in the file picker, the remaining tabs can be matched to the wrong summary rows.
- Summary YTD compares this year's average month (through the latest month) with last year's average month across the whole year. It doesn't compare the same months year over year.
- Every location in the config gets a summary row, even ones that weren't selected (they stay blank).
- Months after the latest reported month are filled with `=A1` placeholders, which display as blank or zero.
- Written in 2020 for Python 3.8 and PyInstaller 3.6. It hasn't been updated for newer versions.

## Author

**Weston Matthews** · [github.com/Weston-Matthews](https://github.com/Weston-Matthews)

## License

TBD
