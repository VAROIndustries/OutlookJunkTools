<#
  Registers OutlookJunkTools for the current user. No admin rights needed.
  Close Outlook first when updating an existing install. Run uninstall.ps1 to remove.
#>
param([string]$Configuration = 'Release')
$ErrorActionPreference = 'Stop'

$ProgId = 'OutlookJunkTools.Connect'
$Clsid  = '{6F1C2B7A-3D54-4E0B-9A7E-2C1D8E5F4A10}'
$Class  = 'OutlookJunkTools.AddIn.Connect'

$src = Join-Path $PSScriptRoot "src\OutlookJunkTools.AddIn\bin\$Configuration\net48"
$dll = Join-Path $src 'OutlookJunkTools.AddIn.dll'
if (-not (Test-Path $dll)) { throw "Build output not found at $dll. Run: dotnet build -c $Configuration" }

$dest = Join-Path $env:LOCALAPPDATA 'OutlookJunkTools\app'
# An installed copy is locked while Outlook has it loaded. A first install is not.
if ((Test-Path (Join-Path $dest 'OutlookJunkTools.AddIn.dll')) -and (Get-Process OUTLOOK -ErrorAction SilentlyContinue)) {
    throw 'Outlook is running with a previous install loaded. Close it, then run this script again.'
}
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item -Path (Join-Path $src '*') -Destination $dest -Recurse -Force
$destDll = Join-Path $dest 'OutlookJunkTools.AddIn.dll'
$asmName = [System.Reflection.AssemblyName]::GetAssemblyName($destDll).FullName
$codeBase = ([System.Uri]$destDll).AbsoluteUri

# COM class keys must go in the registry view that matches Outlook's bitness.
$platform = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue).Platform
$view = if ($platform -eq 'x86') { [Microsoft.Win32.RegistryView]::Registry32 } else { [Microsoft.Win32.RegistryView]::Registry64 }
$hkcu = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)

function Set-Key($path, [hashtable]$values) {
    $k = $hkcu.CreateSubKey($path)
    foreach ($name in $values.Keys) {
        $v = $values[$name]
        if ($v -is [int]) { $k.SetValue($name, $v, [Microsoft.Win32.RegistryValueKind]::DWord) }
        else { $k.SetValue($name, [string]$v, [Microsoft.Win32.RegistryValueKind]::String) }
    }
    $k.Close()
}

Set-Key "Software\Classes\$ProgId"       @{ '' = $ProgId }
Set-Key "Software\Classes\$ProgId\CLSID" @{ '' = $Clsid }
Set-Key "Software\Classes\CLSID\$Clsid"  @{ '' = $ProgId }
Set-Key "Software\Classes\CLSID\$Clsid\ProgId" @{ '' = $ProgId }
Set-Key "Software\Classes\CLSID\$Clsid\InprocServer32" @{
    ''             = 'mscoree.dll'
    ThreadingModel = 'Both'
    Class          = $Class
    Assembly       = $asmName
    RuntimeVersion = 'v4.0.30319'
    CodeBase       = $codeBase
}
Set-Key "Software\Microsoft\Office\Outlook\Addins\$ProgId" @{
    FriendlyName = 'Outlook Junk Tools'
    Description  = 'Block sender, block domain, junk and unsubscribe commands.'
    LoadBehavior = 3
}
$hkcu.Close()

Write-Host "Installed to $dest ($platform Outlook). Start Outlook."
