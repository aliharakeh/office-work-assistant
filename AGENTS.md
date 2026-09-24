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
  Excel/
    ExcelPage.xaml(.cs)  -> namespace OfficeWorkAssistant.Features.Excel
    ExcelWork.cs         -> namespace OfficeWorkAssistant.Features.Excel
  Files/
    FilesPage.xaml(.cs)  -> namespace OfficeWorkAssistant.Features.Files
    FilesWork.cs         -> namespace OfficeWorkAssistant.Features.Files
```

- `Views/` holds shell/navigation only. `MainWindow` hosts a `Frame`; `HomePage` links to features.
- `Features/<Name>/` holds one self-contained tool: `<Name>Page.xaml(.cs)` (UI) + `<Name>Work.cs` (pure logic, no WPF).
- Namespaces must match folders: `OfficeWorkAssistant.Views`, `OfficeWorkAssistant.Features.<Name>`. XAML `x:Class` must match the code-behind namespace.
- `App.xaml` `StartupUri` is `Views\MainWindow.xaml`.

## Add a new feature

1. Create `Features\<Name>\` with `<Name>Page.xaml(.cs)` and `<Name>Work.cs`.
2. In `OfficeWorkAssistant.csproj`, add explicit entries (keep classic style):
   `<Page Include="Features\<Name>\<Name>Page.xaml">` + `<Compile Include="Features\<Name>\<Name>Page.xaml.cs">` with `<DependentUpon>`, plus `<Compile Include="Features\<Name>\<Name>Work.cs" />`.
3. Set page namespace to `OfficeWorkAssistant.Features.<Name>` in both `.xaml` (`x:Class`) and `.xaml.cs`.
4. Wire navigation from `Views\HomePage.xaml(.cs)`: add `using OfficeWorkAssistant.Features.<Name>;`, add a card/button, call `NavigationService.Navigate(new <Name>Page());`.
5. Keep logic UI-free in `<Name>Work.cs` so pages stay thin. Do not add cross-feature references.
6. Build Release per above to verify.
