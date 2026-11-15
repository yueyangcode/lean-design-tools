$ErrorActionPreference = 'Stop'

if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    throw '请先保存工作并关闭 PowerPoint，再运行安装脚本。'
}

$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw '标准 CLR COM 注册需要管理员权限，请以管理员身份运行 PowerShell。'
}

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$dll = Join-Path $root 'src\LeanDesignTools\bin\Release\net48\LeanDesignTools.dll'
if (-not (Test-Path $dll)) { & (Join-Path $root 'build.ps1') }

$regAsm = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
& $regAsm $dll /codebase
if ($LASTEXITCODE -ne 0) { throw "RegAsm 注册失败：$LASTEXITCODE" }

New-Item 'HKLM:\Software\Classes\CLSID\{AA7C4797-FD89-49BC-8E1A-A0C63A3788F0}\Control' -Force | Out-Null

$addin = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\LeanDesignTools.AddIn'
New-Item $addin -Force | Out-Null
Set-ItemProperty $addin -Name FriendlyName -Value '设计工具（精简版）'
Set-ItemProperty $addin -Name Description -Value '独立的 PowerPoint 设计工具任务窗格'
Set-ItemProperty $addin -Name LoadBehavior -Type DWord -Value 3

Write-Host '安装完成。重新打开 PowerPoint，在“设计工具”选项卡中打开任务窗格。'
