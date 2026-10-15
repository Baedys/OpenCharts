param([string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    throw 'Fermez PowerPoint avant les tests dispersion/bulles pour isoler les documents de test.'
}
# Renderer-only. No Excel, add-in registration or COMAddIns state changes.
& (Join-Path $PSScriptRoot 'Build.ps1') -Test
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (!(Test-Path -LiteralPath $compiler)) { $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
if (!(Test-Path -LiteralPath $compiler)) { throw 'Compilateur .NET Framework 4.x introuvable.' }
$bin = Join-Path $PSScriptRoot 'bin'
$windowsBase = Join-Path (Split-Path -Parent $compiler) 'WPF\WindowsBase.dll'
if (!(Test-Path -LiteralPath $windowsBase)) { throw 'WindowsBase.dll (.NET Framework WPF) introuvable.' }
& $compiler /nologo /target:exe /r:Microsoft.CSharp.dll ('/r:' + $windowsBase) ('/r:' + (Join-Path $bin 'OpenSlideCharts.dll')) ('/out:' + (Join-Path $bin 'ScatterOffice.exe')) (Join-Path $PSScriptRoot 'tests\ScatterOffice.cs')
if ($LASTEXITCODE -ne 0) { throw 'Échec de compilation des tests dispersion/bulles Office.' }
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { throw 'PowerPoint a été ouvert pendant la compilation ; fermez-le avant les tests.' }
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { $OutputDirectory = Join-Path $PSScriptRoot 'test-results\scatter' }
$outputPath = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
& (Join-Path $bin 'ScatterOffice.exe') $outputPath
if ($LASTEXITCODE -ne 0) { throw 'Échec des tests dispersion/bulles Office. Consulter le diagnostic précédent.' }
Write-Output ('Résultats dispersion/bulles : ' + $outputPath)
