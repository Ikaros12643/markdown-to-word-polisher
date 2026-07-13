param(
    [Parameter(Mandatory = $true)]
    [string]$Version,

    [string]$Title,

    [string]$Notes,

    [string]$Runtime = "win-x64",

    [switch]$Draft,

    [switch]$Prerelease,

    [switch]$SkipTag,

    [switch]$SkipPush,

    [switch]$Clobber
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

function Test-CommandExists {
    param([string]$Name)

    $command = Get-Command $Name -ErrorAction SilentlyContinue
    return $null -ne $command
}

function Copy-DirectoryContent {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Source,

        [Parameter(Mandatory = $true)]
        [string]$Destination
    )

    if ((Test-Path -LiteralPath $Destination) -eq $false) {
        New-Item -ItemType Directory -Path $Destination | Out-Null
    }

    Get-ChildItem -LiteralPath $Source -Force | ForEach-Object {
        $target = Join-Path $Destination $_.Name
        if ($_.PSIsContainer) {
            if ($_.Name -eq "bin") {
                return
            }

            Copy-DirectoryContent -Source $_.FullName -Destination $target
        } else {
            Copy-Item -LiteralPath $_.FullName -Destination $target -Force
        }
    }
}

if ($Version -notmatch '^v\d+\.\d+(?:\.\d+)?([-.][0-9A-Za-z.-]+)?$') {
    throw "Version 必须使用版本 tag 格式，例如 v1.0、v0.1.0 或 v0.2.0-beta.1"
}

foreach ($commandName in @("git", "gh", "dotnet")) {
    if ((Test-CommandExists -Name $commandName) -eq $false) {
        throw "未找到 $commandName，请先安装并确保可在 PATH 中调用。"
    }
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptRoot
$skillSourcePath = Join-Path $projectRoot "skills\markdown-to-word-polisher"
$projectPath = Join-Path $projectRoot "src\MdToDocx\MdToDocx.csproj"
$outPath = Join-Path $projectRoot "out"
$stagingRoot = Join-Path $outPath ("release-staging-" + $Version)
$stagedSkillPath = Join-Path $stagingRoot "markdown-to-word-polisher"
$stagedBinPath = Join-Path $stagedSkillPath "bin"
$publishPath = Join-Path $stagingRoot "publish"
$assetName = "markdown-to-word-polisher-skill-$Version.zip"
$assetPath = Join-Path $outPath $assetName

if ((Test-Path -LiteralPath $skillSourcePath) -eq $false) {
    throw "skill 目录不存在: $skillSourcePath"
}

if ((Test-Path -LiteralPath $projectPath) -eq $false) {
    throw "mdtodocx 项目不存在: $projectPath"
}

$status = git -C $projectRoot status --porcelain
if ([string]::IsNullOrWhiteSpace($status) -eq $false) {
    throw "工作区存在未提交更改，请先提交或清理后再发布。"
}

if ((Test-Path -LiteralPath $outPath) -eq $false) {
    New-Item -ItemType Directory -Path $outPath | Out-Null
}

if ((Test-Path -LiteralPath $assetPath) -and (-not $Clobber)) {
    throw "发布包已存在: $assetPath。使用 -Clobber 覆盖。"
}

if ((Test-Path -LiteralPath $stagingRoot)) {
    Remove-Item -LiteralPath $stagingRoot -Recurse -Force
}

New-Item -ItemType Directory -Path $stagingRoot | Out-Null

dotnet publish $projectPath -c Release -r $Runtime --self-contained true -o $publishPath /p:PublishSingleFile=true /p:PublishTrimmed=false

Copy-DirectoryContent -Source $skillSourcePath -Destination $stagedSkillPath
New-Item -ItemType Directory -Path $stagedBinPath | Out-Null
Copy-Item -LiteralPath (Join-Path $publishPath "mdtodocx.exe") -Destination (Join-Path $stagedBinPath "mdtodocx.exe") -Force

if ((Test-Path -LiteralPath $assetPath) -and $Clobber) {
    Remove-Item -LiteralPath $assetPath -Force
}

Compress-Archive -Path $stagedSkillPath -DestinationPath $assetPath -Force
Write-Host ("已生成发布包: " + $assetPath)

$existingTag = git -C $projectRoot tag --list $Version
if ([string]::IsNullOrWhiteSpace($existingTag)) {
    if (-not $SkipTag) {
        git -C $projectRoot tag $Version
        Write-Host ("已创建 tag: " + $Version)
    }
} else {
    Write-Host ("tag 已存在: " + $Version)
}

if (-not $SkipPush) {
    git -C $projectRoot push origin $Version
    Write-Host ("已推送 tag: " + $Version)
}

if ([string]::IsNullOrWhiteSpace($Title)) {
    $Title = $Version
}

if ([string]::IsNullOrWhiteSpace($Notes)) {
    $Notes = "Markdown To Word Polisher skill release $Version."
}

$releaseArgs = @(
    "release", "create", $Version, $assetPath,
    "--title", $Title,
    "--notes", $Notes
)

if ($Draft) {
    $releaseArgs += "--draft"
}

if ($Prerelease) {
    $releaseArgs += "--prerelease"
}

gh @releaseArgs
Write-Host ("GitHub Release 已创建: " + $Version)
