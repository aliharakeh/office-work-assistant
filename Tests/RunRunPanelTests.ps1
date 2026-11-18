param([string]$BuildDirectory = 'bin\Release')
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$release = Join-Path $root $BuildDirectory
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
$framework = Join-Path ${env:ProgramFiles(x86)} 'Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8'
$test = Join-Path $release 'RunPanelSmokeTests.exe'
& $compiler /nologo /target:exe "/out:$test" `
    "/reference:$release\OfficeWorkAssistant.exe" /reference:System.Data.dll /reference:System.Xml.dll `
    "/reference:$framework\WindowsBase.dll" "/reference:$framework\PresentationCore.dll" `
    "/reference:$framework\PresentationFramework.dll" /reference:System.Xaml.dll `
    (Join-Path $PSScriptRoot 'RunPanelSmokeTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Run panel test compilation failed. Build first.' }
& $test
if ($LASTEXITCODE -ne 0) { throw 'Run panel smoke tests failed.' }
