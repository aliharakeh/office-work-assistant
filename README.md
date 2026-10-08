# Office Work Assistant

A WPF desktop app for Windows 7 SP1 and later, built on .NET Framework 4.8. The whole app is one pipeline canvas. You chain steps (load Excel files, filter, compare, fill, find files, copy, move, delete) into a pipeline, save it, and rerun it next month on new files. Everything runs locally. Run preview never changes anything on disk, and file changes happen only when you click Run and save files and confirm.

## The pipeline window

- **Saved pipelines** (top left) lists the pipelines in `Documents\Office Work Assistant\Pipelines`, newest first. Click one to open it. The last saved pipeline opens when the app starts. New starts an empty pipeline. Duplicate copies the selected one. Delete moves it to the Recycle Bin. Import... copies a pipeline file from elsewhere into the list, and Folder opens the folder in Explorer.
- **The name box** at the top is the pipeline's name. Save (Ctrl+S) stores the pipeline under that name, with no file dialog. Changing the name and saving renames the saved pipeline. Saving under a name another pipeline already uses asks first. Switching pipelines or closing the app with unsaved changes offers to save them.
- **Add a step** (bottom left) lists every step in three groups, each with its own colour: Excel data, Files and folders, and Change files. Click a step to add it next to the selected step, linked to it. You can also drag a step onto the canvas to place it.
- **The canvas:**
  - Link steps by dragging from a step's right dot to another step's left dot. Compare A/B and Fill columns have two inputs, A and B.
  - One step's result can feed several later steps.
  - Links that would make a loop are refused.
  - To remove a link or step, select it and press Delete.
- **Editing a step:** double-click it, or select it and click Edit step... This opens the step's editor with the earlier step's result already loaded. Click Use in pipeline to keep the changes, or Back to cancel. Load file and Save file are edited in the Step panel on the right.
- **Run preview** (F5) runs the selected step and the steps it needs, or every step when none is selected, and shows the result below the canvas. A step that worked turns green. A step that failed turns red and shows why.
- **Run and save files** reads every file again and runs every step without changing anything. It then lists the files it will replace and how many items each Change files step will copy, move, delete or merge, and asks you to confirm. Only then does it carry out those steps in order and write the Save files.
- **Formula guide** opens the formula lessons and tester. Back returns to the canvas.
- **Column references:** steps remember columns by header name, not by position, so a step keeps working when an earlier step adds, removes or reorders columns. If a column it needs disappears, the step fails and names the missing column. Formulas such as `$A` still mean "first column", so prefer header names such as `$Qty` in pipelines.

## Excel steps

### Load file / Save file
Load file reads one sheet of an `.xlsx` or `.xlsm` file, even when it is open in Excel. Save file either writes a new `.xlsx` file or adds a new sheet to a workbook. It writes only on Run and save files.

### Filter & Sort
- Conditions work on a column (pick a column, an operator and a value) or as a free formula such as `$A > 10 && $B == "OK"`. The column operators are `==`, `!=`, `>`, `>=`, `<`, `<=`, contains, starts with, ends with, is empty and is not empty.
- Rows can match all conditions or any one of them.
- Unchecked columns are dropped from the result.
- Add any number of sort keys, each ascending or descending, and move them up or down to set priority. Ties keep the original row order.

### Template
Keep all source columns, or tick only the ones you want, and add extra columns. An extra column can be a fixed value, a copy of a column, combined text (`$A $B`), math (`$A * $B`, `ADDDAYS($Today, 7)`) or a condition (`IF($A > 10, "Big", "Small")`). Save template... and Load template... share a template between pipelines as an XML file.

### Compare A/B
Two inputs, A and B.
- Match rules compare a formula on A with a formula on B, using `==`, `!=`, `>`, `>=`, `<` or `<=`. A column can be split on a separator, and its parts compared as `$1`, `$2` and so on.
- The result is the rows only in A, the rows only in B, or the rows common to both. When both sides have a column with the same name, the result keeps both, with `_A` and `_B` suffixes.

### Fill columns
Fills columns of B with values from the matching row of A.
- **Match rules** compare a formula on the A row with a formula on the B row, for example `TRIM($ID)` == `$Code`. All rules must match, text compares ignore case, and the first matching A row is used.
- **Fill rules** each have a target column (one of B's columns, or a new name to add a column), an optional When condition, and a Value formula. For each column, the first rule that fires fills the cell.
- **Variables:**
  - `$A_Price` reads the matched A row.
  - `$Price` reads the B row.
  - `$Value` is the cell's current value.
  - `$Matched` is true when an A row matched.

## File steps

These steps link Excel data to files and folders on disk. Each step's result is a normal table, so you can filter, compare and save it like any other step.

- **List / set / map** builds keys from the step before it with a key formula, such as `$Code` or `TRIM($A)`:
  - A list keeps every row.
  - A set keeps unique keys.
  - A map keeps unique keys, each with the value from a value formula.

  You can also choose Stored values and type the keys in, or click Capture from input to freeze what the input gives today. Stored values are kept in the pipeline, and the step then needs no input.
- **List folder** lists the files, the folders, or both in a folder, with or without its subfolders. A file filter (for example `$Ext == ".pdf"`) and a separate folder filter (for example `STARTSWITH($Name, "20")`) narrow the list. The folder filter only decides which folders are listed: files inside a folder that is not listed are still searched.
- **Find files** searches a folder for the items that match each key of the step before it. An item can match when its name equals the key, when its name without extension equals the key, or when its name contains, starts with or ends with the key. Formula mode uses one formula for files and another for folders, for example `$1 == $Key && $Ext == ".pdf"` for files and `$Name == $Key` for folders. Keep keys with no match adds a row with `Found = FALSE` for each key that matched nothing. When a folder matches, the items inside it are skipped, so the folder is handled as a whole.

## Change files steps

These steps change the disk. Run preview and their editors only show the plan, with a Status for each row: Ready, Overwrite, Exists - skip, Missing, Duplicate, and so on. The changes happen only on Run and save files. A Save file step after one of these steps saves the real outcome: Copied, Moved, Recycled, Merged or Failed: ...

- **File action** copies, moves or deletes the paths in one column of the step before it (`Path` by default).
  - Copy and Move take a destination folder and an option to keep the folder structure. Files and folders each have their own optional subfolder formula (for example `$Key`) and new-name formula. A file keeps its extension.
  - When the target exists, the step can skip it, overwrite it, or rename to `name (2).ext`.
  - Delete sends items to the Recycle Bin unless Delete permanently is ticked.
- **Merge folders** finds subfolders of a folder that belong together and moves everything from the others into one of them.
  - A match formula compares two folder names. `$A` and `$B` are the names, and `$A1` and `$B2` are their parts after a split. `$A1 == $B1` groups folders that share their first part.
  - A keep rule picks the folder that stays: first or last by name, shortest or longest name, or newest or oldest change.
  - Name clashes are renamed to `name (2).ext`.
  - A merged-away folder is deleted only once it is empty.

All file formulas can use `$Name $Stem $Ext $Path $Folder $FolderName $Relative $IsFolder $Size $Modified $1 $2`, plus `$Key` and `$Value`. File action can also use any column of its input row.

Example: Load file -> List / set / map (a set of `$Code`) -> Find files (name without extension = key) -> File action (copy to `D:\Out`, subfolder `$Key`) -> Save file (report).

To find files that are not in the Excel list, link List folder to Compare A/B as A, with the Excel list as B, and keep the rows only in A. Then send the result to a File action.

## Formulas
Every step shares one formula language.

Columns are `$A`, `$B`, `$C`, the first, second and third column of the sheet. Split values add `$1`, `$2` for the parts.

Operators: `+ - * / ^`, `&` to glue text together, `== != > >= < <=`, `&& || !`, and `IF(condition, then, else)`.

Functions: `TODAY()`, `NOW()`, `DATE()`, `YEAR()`, `MONTH()`, `DAY()`, `WEEKDAY()`, `ADDDAYS()`, `ADDWEEKS()`, `ADDMONTHS()`, `ADDYEARS()`, `STARTOFWEEK()`, `ENDOFWEEK()`, `STARTOFMONTH()`, `ENDOFMONTH()`, `FORMAT(date, fmt)`, `TRIM()`, `REMOVEDIGITS()`, `CLEARSYMBOLS()`, `FIRSTWORD()`, `LASTWORD()`, `CONTAINS()`, `STARTSWITH()`, `ENDSWITH()`.

Date variables: `$Today`, `$Yesterday`, `$Tomorrow`, `$WeekAgo`, `$WeekLater`, `$MonthAgo`, `$MonthLater`, `$StartOfWeek`, `$EndOfWeek`, `$StartOfMonth`, `$EndOfMonth`, `$StartOfYear`, `$EndOfYear`, `$CurrentYear`, `$CurrentMonth`, `$CurrentDay`, `$Now`, plus the start and end of next and previous week and month.

Each step editor has a Variables button that lists every variable with the value it has right now.

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
Views/             MainWindow: a Frame that shows the pipeline page, with step editors opening on top of it.
Features/Pipeline/ The main page: canvas, saved pipelines, running.
Features/<Name>/   One step per folder: <Name>Page.xaml (the step editor) plus <Name>Work.cs (logic, no WPF).
Excel/             Shared .xlsx reading and writing.
Expressions/       Shared formula engine and the Variables help window.
```

Each step keeps its logic in its `Work` class so the editor stays thin. Steps do not reference each other; only the pipeline references them. See `AGENTS.md` for the rules on adding a new step.
