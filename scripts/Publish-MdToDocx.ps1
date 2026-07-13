param(
    [string]$Runtime = "win-x64",
    [string]$Configuration = "Release"
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptRoot
$projectPath = Join-Path $projectRoot "src\MdToDocx\MdToDocx.csproj"
$outputPath = Join-Path $projectRoot "bin"

dotnet publish $projectPath -c $Configuration -r $Runtime --self-contained true -o $outputPath /p:PublishSingleFile=true /p:PublishTrimmed=false

Write-Host ("mdtodocx.exe 已发布到: " + (Join-Path $outputPath "mdtodocx.exe"))
