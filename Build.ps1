param([switch]$Test)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (!(Test-Path -LiteralPath $compiler)) { $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
if (!(Test-Path -LiteralPath $compiler)) { throw '.NET Framework 4.x compiler is required.' }
$bin = Join-Path $root 'bin'
New-Item -ItemType Directory -Path $bin -Force | Out-Null
$sources = @(Get-ChildItem -LiteralPath (Join-Path $root 'src') -Filter '*.cs' | ForEach-Object FullName)
& $compiler /nologo /target:library /platform:anycpu /optimize+ '/r:System.Windows.Forms.dll' '/r:System.Drawing.dll' '/r:System.Web.Extensions.dll' '/r:Microsoft.CSharp.dll' ("/out:" + (Join-Path $bin 'OpenSlideCharts.dll')) $sources
if ($LASTEXITCODE -ne 0) { throw 'Add-in build failed.' }
if ($Test) {
    & $compiler /nologo /target:exe ("/out:" + (Join-Path $bin 'CoreTests.exe')) (Join-Path $root 'src\ChartCore.cs') (Join-Path $root 'src\SeriesChart.cs') (Join-Path $root 'tests\CoreTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Test build failed.' }
    & (Join-Path $bin 'CoreTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Core tests failed.' }
    & $compiler /nologo /target:exe '/r:System.Windows.Forms.dll' ('/r:' + (Join-Path $bin 'OpenSlideCharts.dll')) ('/out:' + (Join-Path $bin 'EditorTests.exe')) (Join-Path $root 'tests\EditorTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Editor test build failed.' }
    & (Join-Path $bin 'EditorTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Editor tests failed.' }
}
Write-Output ('Built: ' + (Join-Path $bin 'OpenSlideCharts.dll'))
