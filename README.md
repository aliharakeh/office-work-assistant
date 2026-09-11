# Work Assistant

A WPF desktop app for Windows 7 SP1 and later, built on .NET Framework 4.8. It bundles five tools for everyday file and Excel chores. Everything runs locally, and each tool previews what it will do before it changes anything.

## Tools

### Excel processing
Compare two `.xlsx` workbooks and pull out the rows you need.

1. Load file A and file B, and pick a sheet in each.
2. Set match rules. Each rule has a formula for side A and a formula for side B, joined by `==`, `!=`, `>`, `>=`, `<` or `<=`. `$A` is the value from A's chosen column, `$B` from B's. You can also split a column on a separator (space, underscore, dash or a custom character) and compare parts. `$1` is the first part, `$2` the second, and so on.
3. Pick the output: rows found in A but not B, rows found in B but not A, or rows common to both.
4. Tick the columns to keep and save. When both files have a column with the same name, the saved sheet keeps both, with `_A` and `_B` suffixes.

Workbooks are read with a shared file handle, so a file that is open in Excel still loads. The grid previews 200, 1000 or 5000 rows; saving always runs on the full sheet.

### Copy files
Copy files into a destination folder under new names, leaving the originals alone.

- File filter and folder filter are formulas. Leave them blank to take everything, or narrow the list with something like `CONTAINS($Name, "report") && $Ext == ".pdf"`.
- File pattern renames each file. Split the name on a separator and rebuild it from the parts. `$2-$1` turns `report_2024.pdf` into `2024-report.pdf`. `$Stem`, `$Name`, `$Ext`, `$Size` and `$Modified` are available too.
- Create wrapper folder puts each file in a subfolder named by the folder pattern, for example `$1` to group by the first part of the source folder name.
- Preview lists every file with its new path and a status: Ready, Exists, Same path or Bad name. Only Ready rows are copied, and an existing destination is never overwritten.

### Merge duplicates
Combine folders that are near-copies of each other.

- Choose a parent folder and a separator, then write a match formula that compares two folder names. `$A1 == $B1` groups folders that share their first split part, and `CONTAINS($A, "2024")` can narrow it further.
- Preview lists each group of matching folders.
- Per group, pick the folder to keep, then tick the folders to merge into it. The ticked folders are emptied into the keeper.
- Merge moves files. A name that already exists in the target is renamed to `name (2).ext`, files that are locked or fail to move are skipped and counted, empty subfolders are cleaned up, and a merged-away folder is deleted only once it is empty.
- The result grid reports per folder: moved, renamed, skipped, and any error.

### Templates
Build a reusable Excel transform and run it on any sheet.

- Keep all source columns, or tick only the ones you want.
- Add extra columns of five kinds: fixed value, copy of a column, combined text (`$A $B`), math (`$A * $B`, `ADDDAYS($Today, 7)`), or conditional (`IF($A > 10, "Big", "Small")`).
- Save template writes the definition to an XML file. Load template brings it back later.
- Preview shows the generated sheet and Save Excel writes it out. The preview caps at 200, 1000 or 5000 rows; saving runs the whole sheet.

### Filter & sort
Filter rows and reorder columns into a new workbook.

- Conditions work on a column (pick a column, an operator and a value) or as a free formula like `$A > 10 && $B == "OK"`. Column operators: `==`, `!=`, `>`, `>=`, `<`, `<=`, contains, starts with, ends with, is empty, is not empty.
- Match all conditions or any condition.
- Pick the columns to keep. Unchecked columns are dropped from the result.
- Add any number of sort keys, each ascending or descending, and move them up or down to set priority. Ties keep the original row order.
- Preview the result and save it as a new `.xlsx` file.

## Formulas
All five tools share one formula language.

Columns are `$A`, `$B`, `$C`, the first, second and third column of the sheet. Split values add `$1`, `$2` for the parts.

Operators: `+ - * / ^`, `&` to glue text together, `== != > >= < <=`, `&& || !`, and `IF(condition, then, else)`.

Functions: `TODAY()`, `NOW()`, `DATE()`, `YEAR()`, `MONTH()`, `DAY()`, `WEEKDAY()`, `ADDDAYS()`, `ADDWEEKS()`, `ADDMONTHS()`, `ADDYEARS()`, `STARTOFWEEK()`, `ENDOFWEEK()`, `STARTOFMONTH()`, `ENDOFMONTH()`, `FORMAT(date, fmt)`, `TRIM()`, `REMOVEDIGITS()`, `CLEARSYMBOLS()`, `FIRSTWORD()`, `LASTWORD()`, `CONTAINS()`, `STARTSWITH()`, `ENDSWITH()`.

Date variables: `$Today`, `$Yesterday`, `$Tomorrow`, `$WeekAgo`, `$WeekLater`, `$MonthAgo`, `$MonthLater`, `$StartOfWeek`, `$EndOfWeek`, `$StartOfMonth`, `$EndOfMonth`, `$StartOfYear`, `$EndOfYear`, `$CurrentYear`, `$CurrentMonth`, `$CurrentDay`, `$Now`, plus the start and end of next and previous week and month.

Each page has a Variables button that lists every variable with the value it has right now.

## Requirements

- Windows 7 SP1 or later
- .NET Framework 4.8 runtime. No SDK is needed to run the app.
- Excel support covers `.xlsx` only, for both reading and writing. A `.xls` file renamed to `.xlsx` will not open.

## Build and run

Use Visual Studio or Framework MSBuild. No modern .NET SDK is required.

```
%WINDIR%\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe WorkAssistant.csproj /p:Configuration=Release /v:minimal
bin\Release\WorkAssistant.exe
```

Close a running instance before rebuilding, otherwise MSBuild cannot overwrite the exe.

## Project layout

```
Views/             Shell and navigation. MainWindow hosts a Frame, HomePage links to the tools.
Features/<Name>/   One tool per folder: <Name>Page.xaml (UI) plus <Name>Work.cs (logic, no WPF).
Expressions/       Shared formula engine and the Variables help window.
```

Each feature keeps its logic in the `Work` class so the page stays thin, and features do not reference each other. See `AGENTS.md` for the rules on adding a new tool.
