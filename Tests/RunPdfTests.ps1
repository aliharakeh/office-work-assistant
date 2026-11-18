param([ValidateSet('x86', 'x64')][string]$Platform = 'x86')

$ErrorActionPreference = 'Stop'
$release = Join-Path (Split-Path $PSScriptRoot -Parent) 'bin\Release'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
$programFiles = ${env:ProgramFiles(x86)}
if (-not $programFiles) { $programFiles = $env:ProgramFiles }
$framework = Join-Path $programFiles 'Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8'
$test = Join-Path $release "PdfSmokeTests-$Platform.exe"
& $compiler /nologo /target:exe "/platform:$Platform" "/out:$test" `
    "/reference:$release\OfficeWorkAssistant.exe" "/reference:$release\PdfSharp-wpf.dll" `
    /reference:System.Data.dll /reference:System.Xaml.dll `
    "/reference:$framework\WindowsBase.dll" "/reference:$framework\PresentationCore.dll" `
    "/reference:$framework\PresentationFramework.dll" (Join-Path $PSScriptRoot 'PdfSmokeTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'PDF test compilation failed. Build Release first.' }
Copy-Item -LiteralPath "$release\OfficeWorkAssistant.exe.config" -Destination "$test.config"
& $test
if ($LASTEXITCODE -ne 0) { throw 'PDF smoke tests failed.' }

