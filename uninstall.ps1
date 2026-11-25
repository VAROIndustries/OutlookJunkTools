<# Removes OutlookJunkTools for the current user. Keeps your block list and log in %AppData%\OutlookJunkTools. #>
$ErrorActionPreference = 'Stop'
$ProgId = 'OutlookJunkTools.Connect'
$Clsid  = '{6F1C2B7A-3D54-4E0B-9A7E-2C1D8E5F4A10}'

if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) {
    throw 'Outlook is running. Close it, then run this script again.'
}

foreach ($view in [Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32) {
    $hkcu = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
    foreach ($path in "Software\Classes\$ProgId", "Software\Classes\CLSID\$Clsid", "Software\Microsoft\Office\Outlook\Addins\$ProgId") {
        $hkcu.DeleteSubKeyTree($path, $false)
    }
    $hkcu.Close()
}

$dest = Join-Path $env:LOCALAPPDATA 'OutlookJunkTools\app'
if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
Write-Host 'Uninstalled.'
