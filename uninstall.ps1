$ErrorActionPreference = 'Stop'

if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    throw '请先关闭 PowerPoint，再卸载。'
}

$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw '请以管理员身份运行 PowerShell。'
}

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$dll = Join-Path $root 'src\LeanDesignTools\bin\Release\net48\LeanDesignTools.dll'
$regAsm = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
if (Test-Path $dll) { & $regAsm $dll /unregister | Out-Host }

Remove-Item 'HKLM:\Software\Classes\CLSID\{AA7C4797-FD89-49BC-8E1A-A0C63A3788F0}\Control' -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\LeanDesignTools.AddIn' -Recurse -Force -ErrorAction SilentlyContinue
Write-Host '卸载完成。'
