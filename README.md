# Office Work Assistant

A WPF desktop app for Windows 7 SP1 and later, built on .NET Framework 4.8. It has two modes: **Builder**, where you chain steps (load Excel files, filter, compare, fill, find files, copy, move, delete) into a pipeline and save it, and **Run**, where someone else picks that saved pipeline, points it at this month's files and runs it. Everything runs locally. Run preview never changes anything on disk, and file changes happen only when you click Run and save files and confirm.

## Two modes

A switch at the top of the window moves between them: **Builder** to make pipelines, **Run** to run one. The app opens in Run mode, ready for the person who just needs this month's result.

### Builder mode

This is the pipeline canvas and everything around it: steps, the Steps sidebar, saved pipelines, Run preview, Run and save files and the Formula guide. It works exactly as it always did; "## The pipeline window" below describes it.

### Run mode

Run mode is for the person who just runs a saved pipeline, not the person who builds it. The saved pipelines are cards on the left, each with its name, when it was last saved, and a plain-language headline such as "5 steps - 2 inputs - 1 output". Pick one and the form on the right shows only what you have to decide: the Excel file or folder each step reads, the output file each step writes and the destination folder of a Copy or Move, and one plain sentence per step saying what it does.

Every step has a **Show in Run panel** checkbox in Builder mode. Checked steps show a section in Run mode with any file choices and an **Edit step settings...** button, which opens the same editor used in Builder mode. All settings can be changed, including columns, formulas, matching rules, sorting, styles and headers. Apply to step uses the changes for this run; closing the editor without applying cancels them. Unchecked steps run with their saved settings and have no editable section. The checkbox defaults to unchecked for every step. A missing setting also defaults to unchecked; explicitly saved choices are preserved. Nothing in Run mode saves changes back to the pipeline. **Check** runs every step and changes nothing on disk; **Run** asks you to confirm replacements and file actions before doing the work, and **Open output folder** opens where the results went.

Run mode uses one tab per branch, named after its final step, just like the Builder sidebar. Each tab shows the steps leading to that result. Shared file choices stay synchronized across tabs. One Check button and one Run button below the tabs validate and run every branch together, regardless of which tab is selected.

## The pipeline window

- **Saved pipelines** (top left) lists the pipelines in `Documents\Office Work Assistant\Pipelines`, newest first. Click one to open it. The last saved pipeline opens when the app starts. New starts an empty pipeline. Duplicate copies the selected one. Delete moves it to the Recycle Bin. Import... copies a pipeline file from elsewhere into the list, and Folder opens the folder in Explorer.
- **The name box** at the top is the pipeline's name. Save (Ctrl+S) stores the pipeline under that name, with no file dialog. Changing the name and saving renames the saved pipeline. Saving under a name another pipeline already uses asks first. Switching pipelines or closing the app with unsaved changes offers to save them.
- **Add a step** (bottom left) lists every step in two groups, Excel and Files and folders. Each step has its own icon and a colour for its kind (data in/out, Excel work, style, files, steps that change files). Click a step to add it next to the selected step, linked to it. You can also drag a step onto the canvas to place it.
- **The canvas:**
  - Link steps by dragging from a step's right dot to another step's left dot. After you connect an input, a new socket appears for the next one. You can connect several Load file steps to one Filter & Sort or Format sheet step and set it up once for every workbook.
  - Steps that process workbooks separately show one numbered output socket per result. Connect each socket to the destination it needs. A socket can feed several branches. Adding the next step from the palette connects all those sockets to it.
  - Compare A/B, Append tables, and Merge PDF show one output socket. Existing links that carried several results continue to carry them and appear as several lines to the same input.
  - Compare A/B, Fill columns, and Arrange PDF keep their A and B inputs. Connect one result on one side to reuse it with each result on the other side. When both sides have several results, they must have the same count and are paired in input order.
  - Links that would make a loop are refused.
  - To remove a link or step, select it and press Delete.
- **Editing steps:** the Steps sidebar on the right shows every step's settings at once, one section per step in running order, with the earlier step's result already loaded. A pipeline that splits (one file feeding two branches) gets one tab per branch, so each branch's settings stay separate; a shared step appears in each tab. Selecting a step on the canvas jumps to its section; clicking a section selects its step. Changes apply by themselves a moment after you edit (there is no Apply button); a step that cannot apply yet shows why in its section. Steps are named automatically (Filter / Sort 1, 2, ...). Double-click a Load file or Save file step to browse for its file.
- **Run preview** (F5) runs the selected step and the steps it needs, or every step when none is selected, and shows the result below the canvas. Use the result picker to inspect each workbook in a batch. Step editors show the first input and share the settings across the batch. A step that worked turns green. A step that failed turns red and shows why.
- **Run and save files** reads every file again and runs every step without changing anything. It then lists the Excel and text / JSON files it will replace and how many items each Change files step will copy, move, delete or merge, and asks you to confirm. Only then does it carry out those steps in order and write the Save files. This is the flow Run mode gives the person who is not building the pipeline, without the canvas.
- **Formula guide** opens the formula lessons and tester. Back returns to the canvas.
- **Column references:** steps remember columns by header name, not by position, so a step keeps working when an earlier step adds, removes or reorders columns. If a column it needs disappears, the step fails and names the missing column. Formulas such as `$A` still mean "first column", so prefer header names such as `$Qty` in pipelines.

## Excel steps

### Load file / Save file
Load file reads one sheet of an `.xlsx` or `.xlsm` file, even when it is open in Excel. Save file either writes a new `.xlsx` file or adds a new sheet to a workbook. It writes only on Run and save files.

For multiple workbooks, connect each Load file to the same processing step, then connect its output to Save file. You can also connect the Load file steps directly to Save file. With one result, the Save step uses the chosen filename. With several results, `result.xlsx` becomes `result-January.xlsx`, `result-February.xlsx`, and so on. Repeated source names get a numeric suffix. New sheet mode uses the same naming rule and adds the configured sheet in each destination workbook. Save text / JSON and Save PDF also add the source name when their input wire carries several results. Run mode's **Check** lists the actual output paths.

For a sheet without headers, tick **No header row (first row is data)** in Load file. This preserves the first data row and starts with blank headers. Existing blank header cells also stay blank.

### Filter & Sort
- Conditions work on a column (pick a column, an operator and a value) or as a free formula such as `$A > 10 && $B == "OK"`. The column operators are `==`, `!=`, `>`, `>=`, `<`, `<=`, contains, starts with, ends with, is empty and is not empty.
- Rows can match all conditions or any one of them.
- Unchecked columns are dropped from the result.
- Add any number of sort keys, each ascending or descending, and move them up or down to set priority. Ties keep the original row order.

### Add columns
Keep all source columns, or tick only the ones you want, and add extra columns. An extra column can be a fixed value, a copy of a column, combined text (`$A $B`), math (`$A * $B`, `ADDDAYS($Today, 7)`) or a condition (`IF($A > 10, "Big", "Small")`). **Add column** opens a window for one column; the step lists each added column with **Edit** and **Remove**.
### Compare A/B
Two inputs, A and B.
- Match rules compare a formula on A with a formula on B, using `==`, `!=`, `>`, `>=`, `<` or `<=`. A column can be split on a separator, and its parts compared as `$1`, `$2` and so on. **Add** opens a window for one rule; the step lists each rule with **Edit** and **Remove**.
- The result is the rows only in A, the rows only in B, or the rows common to both. When both sides have a column with the same name, the result keeps both, with `_A` and `_B` suffixes.

### Fill columns
Fills columns of B with values from the matching row of A.
- **Match rules** compare a formula on the A row with a formula on the B row, for example `TRIM($ID)` == `$Code`. Choose AND to require all rules or OR to accept any rule. Text compares ignore case, and the first matching A row is used. **Add** opens a window for one key; the step lists each key with **Edit** and **Remove**.
- **Fill rules** each have a target column (one of B's columns, or a new name to add a column), an optional When condition, and a Value formula. For each column, the first rule that fires fills the cell. **Add** opens a window for one rule; the step lists each rule with **Edit**, **Up** and **Remove**.
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
- **Edit headers** adds, edits or clears individual header names without removing their data. Choose Text, Formula or Clear for each column; Keep leaves the all-headers formula's result unchanged. An optional formula changes all existing headers first, then individual edits override it. Add blank column appends a new column with empty data cells. Formulas use the shared formula engine: `$Value` / `$Header` is the original header, `$Column` is its Excel letter, `$Index` is its number (starting at 1), and `$A` / `$Name` reads another original header. For example, `upper(trim($Value))` cleans all names, `"Field_" + $Column` generates names, and `if(isblank($Value), "Field_" + $Column, $Value)` fills only missing headers. Clear removes the name, keeping the column. Header settings stay in Builder mode.
- **Append tables** stacks the rows of every linked step, in link order: input 1 first. Columns match by header (the 2nd `Qty` of one table with the 2nd `Qty` of another) or by position. Every column is kept, or only those all tables have. An optional source column holds each row's input step title, so rename the Load steps to label the rows, for example `January`. Deleting a link renumbers the remaining inputs.

## Style steps

These steps decide how the Save file step writes the sheet. The data itself does not change. The result grid shows the colours. **Put them right before Save:** any other step after them drops the styling, because it can move or remove rows. Highlight and Format sheet can follow each other in either order; Highlight colours stay on top.

    - **Highlight** colours cells or rows that match a condition. Each rule has a formula, what to style, and a look. **Add rule** opens a window for one rule; the step lists each rule with **Edit**, **Up**, **Down** and **Remove**. What to style is one of:
  - the whole row
  - the picked columns in matching rows
  - each picked cell tested on its own, with `$Value` as that cell

  Examples:
  - `$Status == "Late"`: whole row light red
  - `$Value < 0` on Amount and Balance: red bold text
  - `$Due < $Today`: Due column yellow

  The look can set a fill colour, text colour, bold, italic, underline, strikethrough, border, alignment, wrap and number format. Colours are a name from the list or `#RRGGBB`. Later rules paint over earlier ones, unless a rule is set to stop. Build a condition fills in the formula from a column, an operator and a value.
- **Format sheet** has five sections. Each offers fill and text colours, bold, text wrap, borders, horizontal and vertical alignment, and Excel number formats.
  - **Sheet** applies to the header and the data range. Sheet borders can cover the entire range or only cells with values. This section also controls frozen panes and automatic column widths.
  - **Header row** controls filter buttons and header styling. The default header is bold on light blue with thin borders.
  - **Data rows** styles the rows below the header and can shade every second row.
  - **Column defaults** sets a common width and style for all data columns.
  - **Individual columns** lets you choose any input column and give it its own width and cell style.

  Settings you leave blank or unchanged inherit the broader style. For data cells, the sheet style comes first, followed by data rows, column defaults, row banding, and individual columns. Highlight rules apply on top.

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
- **Save text / JSON** writes one column, or a list / set / map, to a UTF-8 text or JSON file. It never changes the table it reads.
  - **Text column**: one row per line, in row order, with no header. Choose any column, such as `Key`. Line breaks inside a value are kept.
  - **Text map**: one `key` + tab + `value` line per row, in row order. Choose the key column (usually `Key`) and the value column.
  - **JSON array (list / set)**: a JSON array, `["one", "two"]`. Row order is kept and duplicate values stay, so a set is just the unique list the List / set / map step already made. Numbers, booleans and blank cells stay typed (`null`).
  - **JSON object (map)**: a JSON object, `{"one": "First item"}`. Keys are case-sensitive; duplicate or blank keys stop the step, so use List / set / map (a map) first.

  The output file defaults to `.txt` for the two text formats and `.json` for the two JSON formats. Any explicit extension is kept, and changing the extension does not change the format. Type a full path, or use **Choose output file...**; double-clicking the step also opens the picker. The folder must already exist. Preview shows the plan with `Path`, `Format`, `Rows`, `Content` and `Status` (Ready / Overwrite) and never writes. Existing files are replaced only on Run and save files, after the usual confirmation, and a failed write never truncates the old file. Two save steps (of any kind) may not write the same file in one pipeline.

  Examples:
  - List / set / map (a set) -> Save text / JSON (JSON array) -> `["A100", "B200"]`.
  - List / set / map (a map: key `$Code`, value `$Name`) -> Save text / JSON (JSON object) -> `{"A100": "First item"}`.

## PDF steps

- **Load PDF** opens a PDF chosen in the sidebar. To process several files, connect a table with full PDF paths, such as **List folder** filtered to PDF files. Choose the path column if its header is not Path. Preview shows each page's source, number, size and rotation.
- **Arrange PDF pages** uses an ordered list of pages from input A. Use a list like `1-2, 5, 3-4, 5@90` to reorder, duplicate and rotate pages. Omit a page to remove it. Use `all`, `last`, `odd`, `even`, descending ranges such as `last-1`, or `blank` for an A4 page. To insert from a second PDF, connect it to B and use `1-2, B:1, 3-last`.
- **Merge PDFs** takes PDFs in numbered port order. Connect separate Load PDF steps or a Load PDF step that reads a list of files. Interleave alternates pages from exactly two documents and can reverse the second document first.
- **Split PDF** divides each input PDF into one page per file by default. Enter a different page count or groups separated by semicolons, such as `1-3; 4-last`. Connect Save PDF to write the parts.
- **Save PDF** takes a full output file path. If the step receives multiple documents, `result.pdf` becomes `result-001.pdf`, `result-002.pdf`, and so on. Preview lists every file and page count. The pipeline confirms replacements before writing and requires a different path from any source PDF.

These steps copy PDF pages and their visible content. Encrypted PDFs are unsupported. Bookmarks, forms and other document-level features may be lost, and signatures do not carry over to the new PDF. PDF steps use the managed PDFsharp-WPF 6.2.4 Framework build. The app can run them offline without a separate PDF viewer, native renderer or OCR models.

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

Restore packages with the standalone NuGet CLI (`nuget restore OfficeWorkAssistant.sln`) or Visual Studio first.

```
%WINDIR%\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe OfficeWorkAssistant.csproj /p:Configuration=Release /v:minimal
bin\Release\OfficeWorkAssistant.exe
```

Close a running instance before rebuilding, otherwise MSBuild cannot overwrite the exe.

The Save text / JSON smoke checks can be run after building Release.

```
powershell -NoProfile -File Tests\RunSaveTextTests.ps1
```

They cover text and JSON content (including Arabic, quotes, line breaks and typed values), list/set/map results, `.txt` / `.json` defaults and custom extensions, column remapping, pipeline XML and dispatch, preview writing nothing, collision checks between save steps, and that a failed save leaves the old file untouched.

The Run mode smoke checks can be run after building Release.

```
powershell -NoProfile -File Tests\RunRunnerTests.ps1
```

They cover the choices Run mode asks for (which file or folder each step needs, the sheet list of the chosen workbook, the output extension, what is still missing), the plain-language summary and headline, a dry run and a real run with the user's own choices, that the pipeline file is never written by a run, what Check reports for each output and file action, the wording of a failure, and the folders offered after a run.

Run the PDF checks after building Release. On 64-bit Windows, run both process architectures.

```
powershell -NoProfile -File Tests\RunPdfTests.ps1 -Platform x86
powershell -NoProfile -File Tests\RunPdfTests.ps1 -Platform x64
```

## Project layout

```
Views/             MainWindow: a Frame that shows the pipeline page; step editors live in its right-hand Step panel.
Features/Pipeline/ The main page (Builder mode): canvas, saved pipelines, running. Also Run mode: RunnerPage.xaml (the page the user sees) and RunnerWork.cs (its logic).
Features/<Name>/   One step per folder: <Name>Page.xaml (the step editor, a UserControl shown in the Step panel) plus <Name>Work.cs (logic, no WPF).
Excel/             Shared .xlsx reading and writing, cell styles (SheetStyle) and the style picker, column letters.
Expressions/       Shared formula engine and the Variables help window.
```

Each step keeps its logic in its `Work` class so the editor stays thin. Steps do not reference each other; only the pipeline references them. See `AGENTS.md` for the rules on adding a new step.

Match rules in Compare, Fill columns and Filter & Sort support **AND (all rules)** and **OR (any rule)**. Existing pipelines keep AND unless changed. Fill columns uses the first matching A row in either mode; fill and highlight actions retain their order.

In Builder, enable **Allow editing match rules in Run mode** separately on each step that has matching conditions. It starts off. Run mode then shows that step's match rules, filters or conditions, including its AND/OR choice where applicable. Each rule editor has a **Variables...** button with help for that rule's variables, examples and the shared formula functions. Other step settings remain fixed, and rule edits apply only to the loaded run; the saved pipeline is never overwritten.
