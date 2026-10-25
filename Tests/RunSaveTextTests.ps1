$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$release = Join-Path $root 'bin\Release'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
$test = Join-Path $release 'SaveTextSmokeTests.exe'

& $compiler /nologo /target:exe "/out:$test" `
    "/reference:$release\OfficeWorkAssistant.exe" /reference:System.Data.dll /reference:System.Web.Extensions.dll `
    (Join-Path $PSScriptRoot 'SaveTextSmokeTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Text / JSON test compilation failed. Build Release first.' }
& $test
if ($LASTEXITCODE -ne 0) { throw 'Text / JSON smoke tests failed.' }
