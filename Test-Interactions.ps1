param()
$ErrorActionPreference = 'Stop'
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { throw 'Fermez PowerPoint avant les tests pour isoler les documents de test.' }
& (Join-Path $PSScriptRoot 'Build.ps1') -Test
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (!(Test-Path -LiteralPath $compiler)) { $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
$bin = Join-Path $PSScriptRoot 'bin'
& $compiler /nologo /target:exe /r:Microsoft.CSharp.dll /r:System.Web.Extensions.dll /r:System.Windows.Forms.dll ('/r:' + (Join-Path $bin 'OpenSlideCharts.dll')) ('/out:' + (Join-Path $bin 'InteractionOffice.exe')) (Join-Path $PSScriptRoot 'tests\InteractionOffice.cs')
if ($LASTEXITCODE -ne 0) { throw 'Échec compilation interaction Office.' }
$clsid = (Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CLSID').'(default)'
$server = (Get-ItemProperty -LiteralPath ('Registry::HKEY_CLASSES_ROOT\CLSID\' + $clsid + '\LocalServer32')).'(default)'
if ($server -match '^"([^\"]+\.exe)"') { $executable = $Matches[1] }
elseif ($server -match '^(.*?\.exe)(?:\s|$)') { $executable = $Matches[1] }
else { throw 'Exécutable PowerPoint introuvable.' }
Start-Process -FilePath $executable -WindowStyle Hidden
Start-Sleep -Seconds 3
& (Join-Path $bin 'InteractionOffice.exe') (Join-Path $PSScriptRoot 'test-results\interactions')
if ($LASTEXITCODE -ne 0) { throw 'Échec tests interactions Office.' }
