param([switch]$RendererOnly, [string]$InputJson)
$ErrorActionPreference = 'Stop'
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { throw "Fermez PowerPoint avant les tests d'intégration pour isoler les documents de test." }
# Excel automation creates a separate instance; the test never uses GetActiveObject.
& (Join-Path $PSScriptRoot 'Build.ps1') -Test
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (!(Test-Path -LiteralPath $compiler)) { $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
$bin = Join-Path $PSScriptRoot 'bin'
& $compiler /nologo /target:exe /r:Microsoft.CSharp.dll /r:System.Web.Extensions.dll ('/r:' + (Join-Path $bin 'OpenSlideCharts.dll')) ('/out:' + (Join-Path $bin 'OfficeSmoke.exe')) (Join-Path $PSScriptRoot 'tests\OfficeSmoke.cs')
if ($LASTEXITCODE -ne 0) { throw 'Échec de compilation des tests Office.' }
if (!$RendererOnly) {
    $registration = Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CLSID'
    $clsid = $registration.'(default)'
    $server = (Get-ItemProperty -LiteralPath ('Registry::HKEY_CLASSES_ROOT\CLSID\' + $clsid + '\LocalServer32')).'(default)'
    if ($server -match '^"([^\"]+\.exe)"') { $executable = $Matches[1] }
    elseif ($server -match '^(.*?\.exe)(?:\s|$)') { $executable = $Matches[1] }
    else { throw 'Exécutable PowerPoint introuvable.' }
    # A normal startup is required: Office automation-only startup may omit user add-ins.
    Start-Process -FilePath $executable -WindowStyle Hidden
    Start-Sleep -Seconds 3
}
$output = Join-Path $PSScriptRoot 'test-results'
$arguments = @($output)
if ($RendererOnly) { $arguments += '--renderer-only' }
if ($InputJson) { $arguments += @('--input-json', (Resolve-Path -LiteralPath $InputJson).Path) }
& (Join-Path $bin 'OfficeSmoke.exe') @arguments
if ($LASTEXITCODE -ne 0) { throw 'Échec des tests Office. Consulter le diagnostic précédent.' }
