$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
& (Join-Path $root 'build.ps1')

$assemblyInfo = Get-Content (Join-Path $root 'src\LeanDesignTools\AssemblyInfo.cs') -Raw
$match = [regex]::Match($assemblyInfo, 'AssemblyVersion\("(\d+\.\d+\.\d+)\.\d+"\)')
if (-not $match.Success) { throw '无法从 AssemblyInfo.cs 读取版本号。' }
$version = $match.Groups[1].Value

$compilers = @(
    (Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe'),
    'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
    'C:\Program Files\Inno Setup 6\ISCC.exe'
)
$compiler = $compilers | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $compiler) { throw '未找到 Inno Setup 6。请先安装 JRSoftware.InnoSetup。' }

& $compiler "/DAppVersion=$version" (Join-Path $root 'installer\LeanDesignTools.iss')
if ($LASTEXITCODE -ne 0) { throw "安装包构建失败：$LASTEXITCODE" }

$setup = Join-Path $root "dist\LeanDesignTools-Setup-x64-$version.exe"
if (-not (Test-Path $setup)) { throw "安装包未生成：$setup" }
Write-Host "安装包已生成：$setup"

