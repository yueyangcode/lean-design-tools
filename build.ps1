$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
dotnet build "$root\src\LeanDesignTools\LeanDesignTools.csproj" -c Release
if ($LASTEXITCODE -ne 0) { throw '插件构建失败。' }
dotnet run --project "$root\tests\LeanDesignTools.Tests\LeanDesignTools.Tests.csproj" -c Release
if ($LASTEXITCODE -ne 0) { throw '测试失败。' }
dotnet build "$root\tests\LeanDesignTools.Preview\LeanDesignTools.Preview.csproj" -c Release
if ($LASTEXITCODE -ne 0) { throw '预览程序构建失败。' }
