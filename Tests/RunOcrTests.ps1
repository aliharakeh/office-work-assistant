param(
    [ValidateSet('x86', 'x64')]
    [string]$Platform = 'x86',
    [switch]$SkipNative
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$release = Join-Path $root 'bin\Release'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
$test = Join-Path $release ('OcrSmokeTests-' + $Platform + '.exe')

& $compiler /nologo /target:exe "/platform:$Platform" "/out:$test" `
    "/reference:$release\OfficeWorkAssistant.exe" /reference:System.Data.dll /reference:System.Drawing.dll `
    (Join-Path $PSScriptRoot 'OcrSmokeTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'OCR test compilation failed. Build Release first.' }

if ($SkipNative) { & $test --skip-native }
else { & $test }
if ($LASTEXITCODE -ne 0) { throw 'OCR smoke tests failed.' }
