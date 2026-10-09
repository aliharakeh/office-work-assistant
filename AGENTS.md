# Office Work Assistant

This is a **.NET Framework 4.8** WPF app for **Windows 7 SP1**.

- Target `v4.8` only. Do not retarget to `net6`, `net8`, `net10`, or any SDK-style modern .NET. Those do not run on Windows 7.
- Keep the classic `.csproj` (`ToolsVersion`, `TargetFrameworkVersion`). Do not convert to an SDK-style project.
- Do not install or require a modern .NET SDK to work on this repo. Build with Visual Studio or Framework MSBuild.
- End users need the **.NET Framework 4.8 runtime**, not an SDK. Do not add APIs, NuGet packages, or OS features that need Windows 8+.
- Prefer Framework BCL + WPF. If a package needs `netstandard2.1` or a TFM newer than `net48`, do not use it.
- Framework MSBuild compiles with the C# 5 compiler: no `$""` strings, `?.`, `nameof`, `out var` or expression-bodied members.

## Compile

Close a running instance first if MSBuild cannot overwrite the exe:

```
%WINDIR%\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe OfficeWorkAssistant.csproj /p:Configuration=Release /v:minimal
```

## Run

Compile Release first, then start that build. Visual Studio/`devenv` is not required.

```
bin\Release\OfficeWorkAssistant.exe
```

## Architecture

The app is one pipeline canvas. Every tool is a pipeline step; there is no Home page and no stand-alone tool page.

Feature-folder layout (classic `.csproj`, no SDK globbing — every file needs an explicit entry):

```
App.xaml / App.xaml.cs
Views/
  MainWindow.xaml(.cs)   -> namespace OfficeWorkAssistant.Views   (Frame showing PipelinePage)
Features/
  Pipeline/              -> the main page: canvas, saved pipelines (PipelineStore), running (PipelineWork)
  <Name>/                -> namespace OfficeWorkAssistant.Features.<Name>
    <Name>Page.xaml(.cs)    step editor (UserControl)
    <Name>Work.cs           pure logic, no WPF
  (ExcelProcessing, FillColumns, FilterSort, Templates, CleanText, SplitCombine, RemoveDuplicates,
   ArrangeColumns, Append, Highlight, FormatSheet, FileOps, MergeDuplicates, FormulaGuide)
Excel/                   -> namespace OfficeWorkAssistant.Excel        (shared .xlsx read/write, grid headers, column
                            letters and RowLookup, SheetStyle/CellStyle, CellStylePicker control)
Expressions/             -> namespace OfficeWorkAssistant.Expressions  (shared formula engine)
```

- `MainWindow` navigates its `Frame` to one `PipelinePage` and asks it (`CanClose`) before closing. Step editors are `UserControl`s hosted in the pipeline page's right-hand Step panel (`EditorHost`); only the Formula guide is navigated to, and goes back with `NavigationService.GoBack()`.
- `Features/<Name>/` holds one self-contained step: the editor control plus `<Name>Work.cs`. `FileOps` holds several file steps (List / set / map, List folder, Find files, File action), and `SplitCombine` holds Split column and Combine columns. `FormulaGuide` is a help page, not a step.
- Data passes as `DataTable` with `object` columns; the real header is `DataColumn.Caption` (use `ExcelFile.Header` / `ExcelFile.FindColumn` / `ExcelFile.AddColumn`).
- Namespaces must match folders: `OfficeWorkAssistant.Views`, `OfficeWorkAssistant.Features.<Name>`. XAML `x:Class` must match the code-behind namespace.
- `App.xaml` `StartupUri` is `Views\MainWindow.xaml`.
- Saved pipelines are `<name>.xml` in `Documents\Office Work Assistant\Pipelines` (`PipelineStore`). The file name is the pipeline name; there is no save dialog.

## Pipeline and steps

`Features/Pipeline` is the only feature allowed to reference other features (their Work, settings and Page classes). No feature may reference `OfficeWorkAssistant.Features.Pipeline`, and features do not reference each other.

A step provides, inside its own folder:

1. An XML-serializable settings class in `<Name>Work.cs`: public get/set properties, a parameterless constructor that fills lists and sets defaults, no `char` or interfaces.
2. A static method in `<Name>Work.cs` that takes the input `DataTable`(s) and the settings and returns a new `DataTable`, never changing its inputs.
3. An editor `UserControl` (named `<Name>Page`) whose only constructor is `(DataTable input(s), string label(s), <Name>Settings settings, Action<<Name>Settings> use)` (steps without inputs drop the tables and labels). It shows the inputs, has an "Apply to step" button that validates and calls `use(...)`. It never navigates and has no Back button or title: the pipeline page shows the step name above it. No Browse, sheet picker or Save: Load file and Save file are pipeline steps. The N-input Append step takes `(IList<DataTable> inputs, IList<string> labels, IList<string> names, AppendSettings settings, Action<AppendSettings> use)`.

Adding a step kind touches:
- `PipelineStepKind` (always append: pipelines store the kind by name). Add file kinds to `IsFileKind` and style kinds to `IsStyleKind`.
- An `[XmlElement]` on `PipelineNode.Settings`.
- In `PipelineWork`: `Ports`, `KindLabel`, `KindDescription`, `SettingsFit`, `Summary` and `RunNode`. Also `IsActionKind` / `DescribeAction` / `RunActions` if it changes the disk.
- In `PipelinePage`: `PaletteGroups` and `EditStep`.
- In `OfficeWorkAssistant.csproj`: explicit `<Page>` + `<Compile>` (with `<DependentUpon>`) entries for the editor, and a `<Compile>` entry for the Work file.
- In `README.md`: a short description.

Rules:
- Pages and Work store column references as Excel letters. A pipeline stores header names so steps survive upstream column changes; `PipelineWork.*ToNames` / `*ToLetters` convert at the edges. Formulas (`$Qty`, `$A`) are passed through unchanged.
- `PipelineWork.Run` computes every step first. Run preview stops there and never touches the disk.
- Steps that change the disk (File action, Merge folders: `IsActionKind`) only **plan** in `RunNode`, giving a table with a `Status` column. On Run and save files, and only after every step has worked, `RunActions` carries them out in order, planning each one again against the current disk first. It then runs their downstream steps again, so a Save gets the real outcome. Only after that does `WriteSaves` write the Excel files (new files before new sheets).
- `PipelinePage.RunAll_Click` confirms the replaced files and each action's item count before running.
- Ports are fixed per kind (`Ports`), except Append: `InputPorts` gives it `In1..In(n+1)`, one per link plus a free one. `CompactPorts` renumbers its links after one is removed, and the page rebuilds its box (`RebuildView`).
- Styling travels on the table: `ExcelFile.WithStyle` returns a copy carrying a `SheetStyle` in `ExtendedProperties`, and `ExcelFile.Save` writes it. Never change a cached table or a `SheetStyle` already on a table; copy first. `PipelineWork.RunStep` clears the style from the result of every step except `IsStyleKind` steps and Save, because other steps may move rows. Formula conditions resolve names with `RowLookup` (letters, then headers, then built-ins, plus `$Value`).
- File step results are plain tables: lists have `Key`/`Value` columns, and file lists have `Type, Name, Path, Folder, Relative, Size, Modified`. File and folder items each have their own formulas (filters, match, subfolder, name).
