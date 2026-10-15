param([switch]$Test)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (!(Test-Path -LiteralPath $compiler)) { $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
if (!(Test-Path -LiteralPath $compiler)) { throw '.NET Framework 4.x compiler is required.' }
$bin = Join-Path $root 'bin'
New-Item -ItemType Directory -Path $bin -Force | Out-Null
$sources = @(Get-ChildItem -LiteralPath (Join-Path $root 'src') -Filter '*.cs' | ForEach-Object FullName)
& $compiler /nologo /target:library /platform:anycpu /optimize+ '/r:System.Windows.Forms.dll' '/r:System.Drawing.dll' '/r:System.Web.Extensions.dll' '/r:Microsoft.CSharp.dll' '/r:System.Numerics.dll' ("/out:" + (Join-Path $bin 'OpenSlideCharts.dll')) $sources
if ($LASTEXITCODE -ne 0) { throw 'Add-in build failed.' }
if ($Test) {
    $coreSources = @('src\ChartCore.cs', 'src\ChartAppearance.cs', 'src\SeriesChart.cs', 'src\WaterfallModel.cs', 'src\WaterfallSolver.cs', 'src\WaterfallGeometry.cs', 'src\WaterfallChart.cs', 'tests\CoreTests.cs') | ForEach-Object { Join-Path $root $_ }
    & $compiler /nologo /target:exe '/r:System.Numerics.dll' ("/out:" + (Join-Path $bin 'CoreTests.exe')) $coreSources
    if ($LASTEXITCODE -ne 0) { throw 'Test build failed.' }
    & (Join-Path $bin 'CoreTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Core tests failed.' }
    & $compiler /nologo /target:exe '/r:System.Numerics.dll' ('/r:' + (Join-Path $bin 'OpenSlideCharts.dll')) ('/out:' + (Join-Path $bin 'WaterfallTests.exe')) (Join-Path $root 'tests\WaterfallTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Waterfall test build failed.' }
    & (Join-Path $bin 'WaterfallTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Waterfall tests failed.' }
    & $compiler /nologo /target:exe '/r:System.Numerics.dll' ('/r:' + (Join-Path $bin 'OpenSlideCharts.dll')) ('/out:' + (Join-Path $bin 'WaterfallInteractionTests.exe')) (Join-Path $root 'tests\WaterfallInteractionTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Waterfall interaction test build failed.' }
    & (Join-Path $bin 'WaterfallInteractionTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Waterfall interaction tests failed.' }
    & $compiler /nologo /target:exe '/r:System.Numerics.dll' ('/r:' + (Join-Path $bin 'OpenSlideCharts.dll')) ('/out:' + (Join-Path $bin 'WaterfallCommandTests.exe')) (Join-Path $root 'tests\WaterfallCommandTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Waterfall command test build failed.' }
    & (Join-Path $bin 'WaterfallCommandTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Waterfall command tests failed.' }
    & $compiler /nologo /target:exe ('/out:' + (Join-Path $bin 'PowerPointEventsTests.exe')) (Join-Path $root 'src\PowerPointEvents.cs') (Join-Path $root 'tests\PowerPointEventsTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'PowerPoint event test build failed.' }
    & (Join-Path $bin 'PowerPointEventsTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'PowerPoint event tests failed.' }
    & $compiler /nologo /target:exe '/r:System.Windows.Forms.dll' '/r:System.Drawing.dll' ('/r:' + (Join-Path $bin 'OpenSlideCharts.dll')) ('/out:' + (Join-Path $bin 'PowerPointTimerTests.exe')) (Join-Path $root 'tests\PowerPointTimerTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'PowerPoint timer test build failed.' }
    & (Join-Path $bin 'PowerPointTimerTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'PowerPoint timer tests failed.' }
    & $compiler /nologo /target:exe '/r:System.Windows.Forms.dll' ('/r:' + (Join-Path $bin 'OpenSlideCharts.dll')) ('/out:' + (Join-Path $bin 'EditorTests.exe')) (Join-Path $root 'tests\EditorTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Editor test build failed.' }
    & (Join-Path $bin 'EditorTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Editor tests failed.' }
}
Write-Output ('Built: ' + (Join-Path $bin 'OpenSlideCharts.dll'))
