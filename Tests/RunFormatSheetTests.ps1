$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$release = Join-Path $root 'bin\Release'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
$test = Join-Path $release 'FormatSheetSmokeTests.exe'
& $compiler /nologo /target:exe "/out:$test" `
    "/reference:$release\OfficeWorkAssistant.exe" "/reference:$release\ClosedXML.dll" `
    /reference:System.Data.dll /reference:System.Xml.dll `
    (Join-Path $PSScriptRoot 'FormatSheetSmokeTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Format sheet test compilation failed. Build Release first.' }
& $test
if ($LASTEXITCODE -ne 0) { throw 'Format sheet smoke tests failed.' }
