# Office Work Assistant

This is a **.NET Framework 4.8** WPF app for **Windows 7 SP1**.

- Target `v4.8` only. Do not retarget to `net6`, `net8`, `net10`, or any SDK-style modern .NET. Those do not run on Windows 7.
- Keep the classic `.csproj` (`ToolsVersion`, `TargetFrameworkVersion`). Do not convert to an SDK-style project.
- Do not install or require a modern .NET SDK to work on this repo. Build with Visual Studio or Framework MSBuild.
- End users need the **.NET Framework 4.8 runtime**, not an SDK. Do not add APIs, NuGet packages, or OS features that need Windows 8+.
- Prefer Framework BCL + WPF. If a package needs `netstandard2.1` or a TFM newer than `net48`, do not use it.

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

Feature-folder layout (classic `.csproj`, no SDK globbing — every file needs an explicit entry):

```
App.xaml / App.xaml.cs
Views/
  MainWindow.xaml(.cs)   -> namespace OfficeWorkAssistant.Views
  HomePage.xaml(.cs)     -> namespace OfficeWorkAssistant.Views
Features/
  <Name>/                -> namespace OfficeWorkAssistant.Features.<Name>
    <Name>Page.xaml(.cs)
    <Name>Work.cs
  (ExcelProcessing, MergeColumns, FilterSort, Templates, CopyFiles,
   MergeDuplicates, FormulaGuide, Pipeline)
Excel/                   -> namespace OfficeWorkAssistant.Excel        (shared .xlsx read/write, save dialog)
Expressions/             -> namespace OfficeWorkAssistant.Expressions  (shared formula engine)
```

- `Views/` holds shell/navigation only. `MainWindow` hosts a `Frame`; `HomePage` links to features.
- `Features/<Name>/` holds one self-contained tool: `<Name>Page.xaml(.cs)` (UI) + `<Name>Work.cs` (pure logic, no WPF).
- Excel features pass data as `DataTable` with `object` columns; the real header is `DataColumn.Caption` (use `ExcelFile.Header` / `ExcelFile.FindColumn`).
- Namespaces must match folders: `OfficeWorkAssistant.Views`, `OfficeWorkAssistant.Features.<Name>`. XAML `x:Class` must match the code-behind namespace.
- `App.xaml` `StartupUri` is `Views\MainWindow.xaml`.

## Add a new feature

1. Create `Features\<Name>\` with `<Name>Page.xaml(.cs)` and `<Name>Work.cs`.
2. In `OfficeWorkAssistant.csproj`, add explicit entries (keep classic style):
   `<Page Include="Features\<Name>\<Name>Page.xaml">` + `<Compile Include="Features\<Name>\<Name>Page.xaml.cs">` with `<DependentUpon>`, plus `<Compile Include="Features\<Name>\<Name>Work.cs" />`.
3. Set page namespace to `OfficeWorkAssistant.Features.<Name>` in both `.xaml` (`x:Class`) and `.xaml.cs`.
4. Wire navigation from `Views\HomePage.xaml(.cs)`: add `using OfficeWorkAssistant.Features.<Name>;`, add a card/button, call `NavigationService.Navigate(new <Name>Page());`.
5. Keep logic UI-free in `<Name>Work.cs` so pages stay thin. Do not add cross-feature references (the one exception is `Features/Pipeline`, below).
6. Build Release per above to verify.

## Pipeline (orchestrator exception)

`Features/Pipeline` chains Excel features on a canvas. It is the only feature allowed to reference other features (their Work, settings and Page classes). No feature may reference `OfficeWorkAssistant.Features.Pipeline`.

A pipeline-capable feature provides, inside its own folder:

1. An XML-serializable `<Name>Settings` class in `<Name>Work.cs`: public get/set properties, a parameterless constructor that fills lists, no `char` or interfaces.
2. A static `Run(...)` in `<Name>Work.cs` that takes the input `DataTable`(s) and the settings and returns a new `DataTable`, never changing its inputs.
3. A page constructor `(DataTable input(s), string label(s), <Name>Settings, Action<<Name>Settings> use)`. It hides Browse, the sheet row and Save, renames Home to Back, shows `UseBtn` ("Use in pipeline"), uses the given tables instead of reading files (`FullSource`/`FullTable` and reload handlers check the pipeline input first), then calls `use(...)` and `NavigationService.GoBack()`. The default constructor must behave as before.

Rules:
- Pages and Work store column references as Excel letters. A pipeline stores header names so steps survive upstream column changes; `PipelineWork.*ToNames` / `*ToLetters` convert at the edges. Formulas (`$Qty`, `$A`) are passed through unchanged.
- `PipelineWork.Run` computes every step first and writes Save files only after all steps worked (new files before new sheets).
- Adding a step kind touches: `PipelineStepKind`, an `[XmlElement]` on `PipelineNode.Settings`, `PipelineWork.Ports` / `KindLabel` / `SettingsFit` / `Summary` / `RunNode`, and the palette button plus `EditStep` in `PipelinePage`.
