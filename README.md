# Office Work Assistant

A WPF desktop app for Windows 7 SP1 and later, built on .NET Framework 4.8. The whole app is one pipeline canvas. You chain steps (load Excel files, filter, compare, fill, find files, copy, move, delete) into a pipeline, save it, and rerun it next month on new files. Everything runs locally. Run preview never changes anything on disk, and file changes happen only when you click Run and save files and confirm.

## The pipeline window

- **Saved pipelines** (top left) lists the pipelines in `Documents\Office Work Assistant\Pipelines`, newest first. Click one to open it. The last saved pipeline opens when the app starts. New starts an empty pipeline. Duplicate copies the selected one. Delete moves it to the Recycle Bin. Import... copies a pipeline file from elsewhere into the list, and Folder opens the folder in Explorer.
- **The name box** at the top is the pipeline's name. Save (Ctrl+S) stores the pipeline under that name, with no file dialog. Changing the name and saving renames the saved pipeline. Saving under a name another pipeline already uses asks first. Switching pipelines or closing the app with unsaved changes offers to save them.
- **Add a step** (bottom left) lists every step in two groups, Excel and Files and folders. Each step has its own icon and a colour for its kind (data in/out, Excel work, style, files, steps that change files). Click a step to add it next to the selected step, linked to it. You can also drag a step onto the canvas to place it.
- **The canvas:**
  - Link steps by dragging from a step's right dot to another step's left dot. Compare A/B and Fill columns have two inputs, A and B. Append tables always has one free input dot more than it has links, so you can keep adding tables.
  - One step's result can feed several later steps.
  - Links that would make a loop are refused.
  - To remove a link or step, select it and press Delete.
- **Editing a step:** select it. Its settings open in the Step panel on the right, with the earlier step's result already loaded. Click Apply to step to keep the changes (switching to another step discards unapplied ones). Drag the divider to make the panel wider. Double-click a Load file or Save file step to browse for its file.
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

### Add columns
Keep all source columns, or tick only the ones you want, and add extra columns. An extra column can be a fixed value, a copy of a column, combined text (`$A $B`), math (`$A * $B`, `ADDDAYS($Today, 7)`) or a condition (`IF($A > 10, "Big", "Small")`).
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

## Tidy and reshape steps

- **Clean text** applies every option you tick to the text cells of the ticked columns, or of every column when none is ticked. The options are:
  - trim spaces and turn runs of spaces into one
  - turn line breaks into spaces
  - remove hidden characters (non-breaking spaces, zero-width marks, control codes)
  - change case: upper, lower, Proper or Sentence
  - turn text such as `1,234.50` into a number
  - leave a cell empty when nothing is left

  Numbers and dates are not touched.
- **Split column** cuts one column into several, placed right after it. It can split:
  - on a delimiter (`\t` is a tab)
  - by fixed widths, such as `3,5,2`
  - by a regular expression: without `( )` groups it cuts wherever the pattern matches; with groups, each group is one part, for example `(\w+)@(.+)`

  Max parts makes the last part keep the rest. You can name the new columns, and keep or drop the original.
- **Combine columns** joins the picked columns, in the order you list them, with a separator. It can skip empty values so no separator is doubled. With Remove the joined columns ticked, the result takes the place of the leftmost one; otherwise it goes after the rightmost.
- **Remove duplicates** compares rows on the ticked columns, or on whole rows when none is ticked. Or choose a formula: rows are duplicates when the formula gives the same result, for example `LEFT($Code, 5)` or `LOWER(TRIM($Name)) & "|" & $City`. It can ignore case and spaces at the start and end. It keeps one of:
  - the first of each
  - the last of each
  - only rows that appear once
  - only rows that repeat, for review

  An optional count column says how many rows shared the key. Kept rows stay in their original order.
- **Arrange columns** lists every column in output order. Click a column's name to select it and use Up and Down to reorder; tick its box to keep it, and type a new name to rename it. Columns an earlier step adds later can be kept at the end or dropped.
- **Append tables** stacks the rows of every linked step, in link order: input 1 first. Columns match by header (the 2nd `Qty` of one table with the 2nd `Qty` of another) or by position. Every column is kept, or only those all tables have. An optional source column holds each row's input step title, so rename the Load steps to label the rows, for example `January`. Deleting a link renumbers the remaining inputs.

## Style steps

These steps decide how the Save file step writes the sheet. The data itself does not change. The result grid shows the colours. **Put them right before Save:** any other step after them drops the styling, because it can move or remove rows. Highlight and Format sheet can follow each other in either order; Highlight colours stay on top.

- **Highlight** colours cells or rows that match a condition. Each rule has a formula, what to style, and a look. What to style is one of:
  - the whole row
  - the picked columns in matching rows
  - each picked cell tested on its own, with `$Value` as that cell

  Examples:
  - `$Status == "Late"`: whole row light red
  - `$Value < 0` on Amount and Balance: red bold text
  - `$Due < $Today`: Due column yellow

  The look can set a fill colour, text colour, bold, italic, underline, strikethrough, border, alignment, wrap and number format. Colours are a name from the list or `#RRGGBB`. Later rules paint over earlier ones, unless a rule is set to stop. Build a condition fills in the formula from a column, an operator and a value.
- **Format sheet** sets the sheet's look:
  - header row style (bold on light blue with thin borders by default)
  - freeze the header row and the first N columns
  - filter buttons on the header
  - column widths fitted to their contents
  - thin borders on every cell
  - every second row shaded
  - per-column width, number format (for example `#,##0.00` or `dd/MM/yyyy`), alignment and wrap

Example: Load file -> Clean text -> Remove duplicates -> Highlight (`$Qty == 0` row light red) -> Format sheet -> Save file.

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

Functions: `TODAY()`, `NOW()`, `DATE()`, `YEAR()`, `MONTH()`, `DAY()`, `WEEKDAY()`, `ADDDAYS()`, `ADDWEEKS()`, `ADDMONTHS()`, `ADDYEARS()`, `STARTOFWEEK()`, `ENDOFWEEK()`, `STARTOFMONTH()`, `ENDOFMONTH()`, `FORMAT(date, fmt)`, `TRIM()`, `REMOVEDIGITS()`, `CLEARSYMBOLS()`, `FIRSTWORD()`, `LASTWORD()`, `CONTAINS()`, `STARTSWITH()`, `ENDSWITH()`, `UPPER()`, `LOWER()`, `PROPER()`, `LEN()`, `LEFT(text, n)`, `RIGHT(text, n)`, `MID(text, start, n)`, `SUBSTITUTE(text, old, new)`, `ISBLANK()`.

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
Views/             MainWindow: a Frame that shows the pipeline page; step editors live in its right-hand Step panel.
Features/Pipeline/ The main page: canvas, saved pipelines, running.
Features/<Name>/   One step per folder: <Name>Page.xaml (the step editor, a UserControl shown in the Step panel) plus <Name>Work.cs (logic, no WPF).
Excel/             Shared .xlsx reading and writing, cell styles (SheetStyle) and the style picker, column letters.
Expressions/       Shared formula engine and the Variables help window.
```

Each step keeps its logic in its `Work` class so the editor stays thin. Steps do not reference each other; only the pipeline references them. See `AGENTS.md` for the rules on adding a new step.
