$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$release = Join-Path $root 'bin\Release'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
$test = Join-Path $release 'EditHeadersSmokeTests.exe'
& $compiler /nologo /target:exe "/out:$test" `
    "/reference:$release\OfficeWorkAssistant.exe" "/reference:$release\ClosedXML.dll" `
    /reference:System.Data.dll /reference:System.Xml.dll `
    (Join-Path $PSScriptRoot 'EditHeadersSmokeTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Header test compilation failed. Build Release first.' }
& $test
if ($LASTEXITCODE -ne 0) { throw 'Header smoke tests failed.' }
