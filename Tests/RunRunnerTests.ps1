$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$release = Join-Path $root 'bin\Release'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
$test = Join-Path $release 'RunnerSmokeTests.exe'

& $compiler /nologo /target:exe "/out:$test" `
    "/reference:$release\OfficeWorkAssistant.exe" /reference:System.Data.dll `
    (Join-Path $PSScriptRoot 'RunnerSmokeTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Run mode test compilation failed. Build Release first.' }
& $test
if ($LASTEXITCODE -ne 0) { throw 'Run mode smoke tests failed.' }
