param([switch]$Remove, [ValidateSet('auto','x64','x86')][string]$OfficeArchitecture = 'auto')
$ErrorActionPreference = 'Stop'
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { throw 'Fermez PowerPoint avant installation ou désinstallation.' }
if ($OfficeArchitecture -eq 'auto') {
    $config = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
    if ($config.Platform -in @('x64','x86')) { $OfficeArchitecture = $config.Platform }
    else { throw 'Architecture Office indéterminée. Précisez -OfficeArchitecture x64 ou x86.' }
}
$view = if ($OfficeArchitecture -eq 'x64') { [Microsoft.Win32.RegistryView]::Registry64 } else { [Microsoft.Win32.RegistryView]::Registry32 }
$registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
$clsid = '{F42D7316-0B29-4DB1-A11D-6D889F02DDBA}'
$progId = 'OpenSlideCharts.Connect'
$classPath = 'SOFTWARE\Classes\CLSID\' + $clsid
$progPath = 'SOFTWARE\Classes\' + $progId
$addinPath = 'SOFTWARE\Microsoft\Office\PowerPoint\Addins\' + $progId
try {
    if ($Remove) {
        foreach ($path in @($addinPath,$progPath,$classPath)) { $registry.DeleteSubKeyTree($path, $false) }
        Write-Output 'OpenSlideCharts désinscrit pour cet utilisateur. Sources et présentations conservées.'
        return
    }
    $dll = Join-Path $PSScriptRoot 'bin\OpenSlideCharts.dll'
    if (!(Test-Path -LiteralPath $dll)) { throw 'Exécutez Build.ps1 avant installation.' }
    $assemblyName = [Reflection.AssemblyName]::GetAssemblyName($dll).FullName
    $codeBase = ([Uri]$dll).AbsoluteUri
    function Set-RegistryValues([string]$path, [hashtable]$values) {
        $key = $registry.CreateSubKey($path)
        try { foreach ($name in $values.Keys) { $key.SetValue($name,$values[$name]) } } finally { $key.Dispose() }
    }
    Set-RegistryValues $classPath @{''='OpenSlideCharts'}
    Set-RegistryValues ($classPath + '\InprocServer32') @{''='mscoree.dll';ThreadingModel='Both';Class='OpenSlideCharts.Connect';Assembly=$assemblyName;RuntimeVersion='v4.0.30319';CodeBase=$codeBase}
    Set-RegistryValues ($classPath + '\InprocServer32\0.1.0.0') @{Class='OpenSlideCharts.Connect';Assembly=$assemblyName;RuntimeVersion='v4.0.30319';CodeBase=$codeBase}
    Set-RegistryValues ($classPath + '\ProgId') @{''=$progId}
    Set-RegistryValues ($classPath + '\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}') @{}
    Set-RegistryValues $progPath @{''='OpenSlideCharts'}
    Set-RegistryValues ($progPath + '\CLSID') @{''=$clsid}
    Set-RegistryValues $addinPath @{FriendlyName='OpenSlideCharts';Description='Graphiques éditables open source (prototype)';LoadBehavior=[int]3;CommandLineSafe=[int]1}
    Write-Output ('Complément inscrit pour cet utilisateur, Office ' + $OfficeArchitecture + '. Ouvrez PowerPoint : onglet OpenSlideCharts.')
} finally { $registry.Dispose() }
