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
    $coreSources = @('src\ChartCore.cs', 'src\ChartAppearance.cs', 'src\SeriesChart.cs', 'src\PieChart.cs', 'src\ScatterChart.cs', 'src\ScatterRegression.cs', 'src\GanttModel.cs', 'src\GanttCalendar.cs', 'src\GanttChart.cs', 'src\WaterfallModel.cs', 'src\WaterfallSolver.cs', 'src\WaterfallGeometry.cs', 'src\WaterfallChart.cs', 'tests\CoreTests.cs') | ForEach-Object { Join-Path $root $_ }
    & $compiler /nologo /target:exe '/r:System.Numerics.dll' ("/out:" + (Join-Path $bin 'CoreTests.exe')) $coreSources
    if ($LASTEXITCODE -ne 0) { throw 'Test build failed.' }
    & (Join-Path $bin 'CoreTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Core tests failed.' }
    & $compiler /nologo /target:exe ('/r:' + (Join-Path $bin 'OpenSlideCharts.dll')) ('/out:' + (Join-Path $bin 'PieTests.exe')) (Join-Path $root 'tests\PieTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Pie test build failed.' }
    & (Join-Path $bin 'PieTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Pie tests failed.' }
    & $compiler /nologo /target:exe '/r:System.Numerics.dll' ('/r:' + (Join-Path $bin 'OpenSlideCharts.dll')) ('/out:' + (Join-Path $bin 'ScatterTests.exe')) (Join-Path $root 'tests\ScatterTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Scatter test build failed.' }
    & (Join-Path $bin 'ScatterTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Scatter tests failed.' }
    & $compiler /nologo /target:exe '/r:System.Numerics.dll' ('/r:' + (Join-Path $bin 'OpenSlideCharts.dll')) ('/out:' + (Join-Path $bin 'ScatterRegressionTests.exe')) (Join-Path $root 'tests\ScatterRegressionTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Scatter regression test build failed.' }
    & (Join-Path $bin 'ScatterRegressionTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Scatter regression tests failed.' }
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
    & $compiler /nologo /target:exe '/r:System.Windows.Forms.dll' '/r:System.Drawing.dll' ('/r:' + (Join-Path $bin 'OpenSlideCharts.dll')) ('/out:' + (Join-Path $bin 'ConnectorHandleOverlayTests.exe')) (Join-Path $root 'tests\ConnectorHandleOverlayTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Connector overlay test build failed.' }
    & (Join-Path $bin 'ConnectorHandleOverlayTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Connector overlay tests failed.' }
    & $compiler /nologo /target:exe '/r:System.Drawing.dll' ('/r:' + (Join-Path $bin 'OpenSlideCharts.dll')) ('/out:' + (Join-Path $bin 'ConnectorScreenProjectionTests.exe')) (Join-Path $root 'tests\ConnectorScreenProjectionTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Connector projection test build failed.' }
    & (Join-Path $bin 'ConnectorScreenProjectionTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Connector projection tests failed.' }
    & $compiler /nologo /target:exe '/r:System.Windows.Forms.dll' ('/out:' + (Join-Path $bin 'OfficeEditKeysTests.exe')) (Join-Path $root 'src\OfficeEditKeys.cs') (Join-Path $root 'tests\OfficeEditKeysTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Office edit key test build failed.' }
    & (Join-Path $bin 'OfficeEditKeysTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Office edit key tests failed.' }
    & $compiler /nologo /target:exe '/r:System.Windows.Forms.dll' ('/r:' + (Join-Path $bin 'OpenSlideCharts.dll')) ('/out:' + (Join-Path $bin 'EditorTests.exe')) (Join-Path $root 'tests\EditorTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Editor test build failed.' }
    & (Join-Path $bin 'EditorTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Editor tests failed.' }
    & $compiler /nologo /target:exe '/r:System.Windows.Forms.dll' ('/r:' + (Join-Path $bin 'OpenSlideCharts.dll')) ('/out:' + (Join-Path $bin 'PieEditorTests.exe')) (Join-Path $root 'tests\PieEditorTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Pie editor test build failed.' }
    & (Join-Path $bin 'PieEditorTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Pie editor tests failed.' }
    & $compiler /nologo /target:exe '/r:System.Windows.Forms.dll' ('/r:' + (Join-Path $bin 'OpenSlideCharts.dll')) ('/out:' + (Join-Path $bin 'ScatterEditorTests.exe')) (Join-Path $root 'tests\ScatterEditorTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Scatter editor test build failed.' }
    & (Join-Path $bin 'ScatterEditorTests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Scatter editor tests failed.' }
    foreach ($suite in @('GanttCalendarTests', 'GanttTests', 'GanttEditorTests')) {
        & $compiler /nologo /target:exe '/r:System.Windows.Forms.dll' ('/r:' + (Join-Path $bin 'OpenSlideCharts.dll')) ('/out:' + (Join-Path $bin ($suite + '.exe'))) (Join-Path $root ('tests\' + $suite + '.cs'))
        if ($LASTEXITCODE -ne 0) { throw ($suite + ' build failed.') }
        & (Join-Path $bin ($suite + '.exe'))
        if ($LASTEXITCODE -ne 0) { throw ($suite + ' failed.') }
    }
}
Write-Output ('Built: ' + (Join-Path $bin 'OpenSlideCharts.dll'))
