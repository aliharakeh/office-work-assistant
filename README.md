# Office Work Assistant

A WPF desktop app for Windows 7 SP1 and later, built on .NET Framework 4.8. It bundles seven tools for everyday file and Excel chores, plus a pipeline that chains the Excel tools together. Everything runs locally, and each tool previews what it will do before it changes anything.

## Tools

### Excel processing
Compare two `.xlsx` workbooks and pull out the rows you need.

1. Load file A and file B, and pick a sheet in each.
2. Set match rules. Each rule has a formula for side A and a formula for side B, joined by `==`, `!=`, `>`, `>=`, `<` or `<=`. `$A` is the value from A's chosen column, `$B` from B's. You can also split a column on a separator (space, underscore, dash or a custom character) and compare parts. `$1` is the first part, `$2` the second, and so on.
3. Pick the output: rows found in A but not B, rows found in B but not A, or rows common to both.
4. Tick the columns to keep and save. When both files have a column with the same name, the saved sheet keeps both, with `_A` and `_B` suffixes.

Workbooks are read with a shared file handle, so a file that is open in Excel still loads. The grid previews 200, 1000 or 5000 rows; saving always runs on the full sheet. Saving (Excel processing, Fill columns, Filter & sort, Templates) asks whether to write a new `.xlsx` file or add a new sheet to a loaded `.xlsx` file; `.xlsm` files are not offered because saving would drop their macros.

### Fill columns
Fill columns of workbook B with values from workbook A, row by row, using formula rules.

1. Load file A (the values come from here) and file B (the rows get filled), and pick a sheet in each.
2. Set match rules. Each rule is a formula on the A row `==` a formula on the B row, for example `TRIM($ID)` == `$Code`. Each side only sees its own file. All rules must match, text compares ignore case, and when several A rows share a key the first one is used (the status line counts these).
3. Add fill rules. Each rule has a target column (one of B's columns, or a new name to add a column), an optional When condition, and a Value formula. Rules run top to bottom, and the first rule that fires for a column fills that cell. A cell that no rule fills keeps B's value. Use Up to reorder.
4. Inside fill formulas, `$A_Price` or `$A_C` reads the matched A row (blank when nothing matched). `$B_Price`, `$Price` or `$C` reads the B row, including values that earlier rules wrote. A name that B does not have falls back to A. `$Value` is the target cell's current value, and `$Matched` is true when an A row matched.
5. The preview updates as you type and lists any `$names` that match no column. Save writes B with the filled columns, either as a new file or as a new sheet.

Example: target `Price`, When `$Matched && $Value == ""`, Value `$A_Price` fills only the empty prices. A second `Price` rule below it, with When `$Matched` and Value `$Value * 1.1`, raises the prices that B already had.

### Copy files
Copy files into a destination folder under new names, leaving the originals alone.

- File filter and folder filter are formulas. Leave them blank to take everything, or narrow the list with something like `CONTAINS($Name, "report") && $Ext == ".pdf"`.
- File pattern renames each file. Split the name on a separator and rebuild it from the parts. `$2-$1` turns `report_2024.pdf` into `2024-report.pdf`. `$Stem`, `$Name`, `$Ext`, `$Size` and `$Modified` are available too.
- All filters and patterns have explicit source variables: `$FileName` includes the extension, `$FileStem` omits it, `$RootFolderName` is the selected Source folder's name, and `$ParentFolderName` is the file's immediate containing folder name. With Source `C:\Reports`, the file `C:\Reports\Invoices\bill.pdf` gives `bill.pdf`, `bill`, `Reports`, and `Invoices`. For files directly in Source, both folder variables give `Reports`; for deeper files, the parent is always the immediate folder, not the first subfolder. Existing `$Name`, `$Stem`, and `$FolderName` remain aliases.
- File pattern automatically appends the original extension, so use `$FileStem`, not `$FileName`, to keep the original name. `$RootFolderName & "-" & $ParentFolderName & "-" & $FileStem` produces `Reports-Invoices-bill.pdf` in the example above.
- Create wrapper folder puts each file in a subfolder named by the folder pattern, for example `$ParentFolderName` to group by the immediate source folder or `$RootFolderName & "-" & $ParentFolderName` to combine the two names. `$1` still means the first split part of the file stem.
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

### Formula guide
Learn the formula language without touching a file.

- Eight lessons walk from plain numbers through text, dates, decisions and finished multi-part formulas. Each step shows the formula and the result the engine gives on the sample values.
- The tester at the bottom runs anything you type. Edit the sample values, or add rows, to stand in for your own columns and fields.
- Try on any example loads it into the tester, ready to tweak, and the lesson results refresh whenever you edit a sample value.

### Pipeline
Chain the Excel tools on a canvas, so one step's result feeds the next, and rerun the whole chain later on new files.

- Add steps from the toolbar: Load file, Filter & Sort, Template, Compare A/B, Fill columns and Save file. When a step is selected, the new step is placed next to it and linked to it.
- Link steps by dragging from a step's right dot to another step's left dot. Compare A/B and Fill columns have two inputs, A and B, so two branches can join. A step's result can feed several later steps. Links that would make a loop are refused. Select a link or step and press Delete to remove it.
- Double-click a step (or click Edit step) to set it up. This opens the normal tool page with the earlier step's result already loaded in place of a file. Set it up as usual and click Use in pipeline. Back cancels the edit. Load and Save steps are set up in the panel on the right.
- Run preview runs the selected step and the steps it needs, and shows its result below the canvas. Nothing is written. A step that worked turns green. A step that failed turns red and shows why.
- Run and save files reads every file again and runs every step. Only when every step has worked does it write the Save steps' files: each one either replaces a file or adds a new sheet to a workbook.
- Save... writes the pipeline to an XML file, and Open... loads it back. To rerun the chain on next month's files, open the pipeline, point the Load steps at the new files, and click Run and save files.
- Steps remember columns by header name, not by position. A step therefore keeps working when an earlier step adds, removes or reorders columns. If a column it needs disappears, the step fails and names the missing column. Formulas such as `$A` still mean "first column", so prefer header names (`$Qty`) in pipelines.

## Formulas
All seven tools share one formula language.

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
%WINDIR%\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe OfficeWorkAssistant.csproj /p:Configuration=Release /v:minimal
bin\Release\OfficeWorkAssistant.exe
```

Close a running instance before rebuilding, otherwise MSBuild cannot overwrite the exe.

## Project layout

```
Views/             Shell and navigation. MainWindow hosts a Frame, HomePage links to the tools.
Features/<Name>/   One tool per folder: <Name>Page.xaml (UI) plus <Name>Work.cs (logic, no WPF).
Expressions/       Shared formula engine and the Variables help window.
```

Each feature keeps its logic in the `Work` class so the page stays thin, and features do not reference each other. See `AGENTS.md` for the rules on adding a new tool.
