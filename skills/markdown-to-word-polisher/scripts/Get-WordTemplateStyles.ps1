param(
    [Parameter(Mandatory = $true)]
    [string]$TemplatePath,

    [string]$OutputPath,

    [switch]$Visible
)

$ErrorActionPreference = "Stop"

function Resolve-ExistingFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathValue,
        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $resolved = Resolve-Path -LiteralPath $PathValue -ErrorAction SilentlyContinue
    if ($null -eq $resolved) {
        throw "$Label 不存在: $PathValue"
    }

    return $resolved.ProviderPath
}

function ConvertTo-JsonCompat {
    param([object]$Value)

    return $Value | ConvertTo-Json -Depth 6
}

$resolvedTemplate = Resolve-ExistingFile -PathValue $TemplatePath -Label "Word 模板"
$word = $null
$document = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$Visible
    $document = $word.Documents.Open($resolvedTemplate, $false, $true)

    $paragraphStyles = New-Object System.Collections.ArrayList
    $characterStyles = New-Object System.Collections.ArrayList
    $tableStyles = New-Object System.Collections.ArrayList

    foreach ($style in $document.Styles) {
        $name = [string]$style.NameLocal
        if ([string]::IsNullOrWhiteSpace($name)) {
            $name = [string]$style.Name
        }

        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        if ($style.Type -eq 1) {
            [void]$paragraphStyles.Add($name)
        } elseif ($style.Type -eq 2) {
            [void]$characterStyles.Add($name)
        } elseif ($style.Type -eq 3) {
            [void]$tableStyles.Add($name)
        }
    }

    $result = [ordered]@{
        template = $resolvedTemplate
        paragraphStyles = @($paragraphStyles | Sort-Object -Unique)
        characterStyles = @($characterStyles | Sort-Object -Unique)
        tableStyles = @($tableStyles | Sort-Object -Unique)
    }

    $json = ConvertTo-JsonCompat -Value $result
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        Write-Output $json
    } else {
        $fullOutput = [System.IO.Path]::GetFullPath($OutputPath)
        $directory = [System.IO.Path]::GetDirectoryName($fullOutput)
        if ([string]::IsNullOrEmpty($directory) -eq $false -and (Test-Path -LiteralPath $directory) -eq $false) {
            New-Item -ItemType Directory -Path $directory | Out-Null
        }
        Set-Content -LiteralPath $fullOutput -Value $json -Encoding UTF8
        Write-Host ("已写入模板样式清单: " + $fullOutput)
    }
} finally {
    if ($null -ne $document) {
        $document.Close($false)
    }
    if ($null -ne $word) {
        $word.Quit()
    }
}
