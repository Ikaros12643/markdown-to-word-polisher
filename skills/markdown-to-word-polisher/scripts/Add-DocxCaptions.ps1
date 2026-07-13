param(
    [Parameter(Mandatory = $true)]
    [string]$DocumentPath,

    [string]$TableCaptionsPath,

    [string]$CaptionStyleName = "题注",

    [switch]$AddImageCaptions,

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

function Read-TableCaptionTitles {
    param([string]$PathValue)

    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        return @()
    }

    $resolved = Resolve-ExistingFile -PathValue $PathValue -Label "表题注配置"
    $json = Get-Content -LiteralPath $resolved -Raw -Encoding UTF8
    $data = $json | ConvertFrom-Json

    $items = @()
    if ($null -ne $data.tables) {
        $items = @($data.tables)
    } else {
        $items = @($data)
    }

    $titles = New-Object System.Collections.ArrayList
    foreach ($item in $items) {
        if ($null -ne $item.PSObject.Properties["title"]) {
            [void]$titles.Add([string]$item.title)
        } else {
            [void]$titles.Add([string]$item)
        }
    }

    return @($titles)
}

function Ensure-CaptionLabel {
    param(
        [object]$Word,
        [string]$LabelName,
        [bool]$IncludeChapterNumber
    )

    $label = $null
    try {
        $label = $Word.CaptionLabels.Item($LabelName)
    } catch {
        $label = $Word.CaptionLabels.Add($LabelName)
    }

    $label.NumberStyle = 0
    $label.IncludeChapterNumber = $IncludeChapterNumber
    if ($IncludeChapterNumber -eq $true) {
        $label.ChapterStyleLevel = 1
        $label.Separator = 0
    }
}

function Set-CaptionStyleForNearbyParagraphs {
    param(
        [object]$Document,
        [string]$StyleName
    )

    if ([string]::IsNullOrWhiteSpace($StyleName)) {
        return
    }

    foreach ($paragraph in $Document.Paragraphs) {
        $text = ([string]$paragraph.Range.Text).Trim()
        if ($text.StartsWith("表 ") -or $text.StartsWith("图 ")) {
            try {
                $paragraph.Range.Style = $StyleName
            } catch {
                Write-Warning ("题注样式不存在或无法应用: " + $StyleName)
                return
            }
        }
    }
}

function Append-TextToPreviousParagraph {
    param(
        [object]$Paragraph,
        [string]$Text
    )

    if ($null -eq $Paragraph -or [string]::IsNullOrWhiteSpace($Text)) {
        return
    }

    $appendRange = $Paragraph.Range.Duplicate()
    if ($appendRange.End -gt $appendRange.Start) {
        $appendRange.End = $appendRange.End - 1
    }

    $wdCollapseEnd = 0
    $appendRange.Collapse($wdCollapseEnd)
    $appendRange.InsertAfter($Text)
}

function Insert-TableCaption {
    param(
        [object]$Table,
        [string]$Title
    )

    $range = $Table.Range.Duplicate()
    $range.InsertCaption("表", "", $null, 0, $false)

    $captionParagraph = $null
    try {
        $captionParagraph = $Table.Range.Paragraphs.First.Previous()
    } catch {
        $captionParagraph = $null
    }

    Append-TextToPreviousParagraph -Paragraph $captionParagraph -Text (" " + $Title)
}

function Set-TablesAutoFitToWindow {
    param([object]$Document)

    $wdAutoFitWindow = 2
    $count = $Document.Tables.Count
    for ($i = 1; $i -le $count; $i++) {
        $table = $Document.Tables.Item($i)
        try {
            $table.AllowAutoFit = $true
            $table.AutoFitBehavior($wdAutoFitWindow)
        } catch {
            Write-Warning ("表格按窗口自动调整失败，序号: " + $i)
        }
    }

    return $count
}

$resolvedDocument = Resolve-ExistingFile -PathValue $DocumentPath -Label "Word 文档"
$tableTitles = @(Read-TableCaptionTitles -PathValue $TableCaptionsPath)

$word = $null
$document = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$Visible
    Ensure-CaptionLabel -Word $word -LabelName "表" -IncludeChapterNumber $true
    Ensure-CaptionLabel -Word $word -LabelName "图" -IncludeChapterNumber $false

    $document = $word.Documents.Open($resolvedDocument)

    $tableCount = $document.Tables.Count
    $autoFitCount = Set-TablesAutoFitToWindow -Document $document
    if ($tableTitles.Count -gt 0 -and $tableTitles.Count -ne $tableCount) {
        Write-Warning ("表题注数量(" + $tableTitles.Count + ")与文档表格数量(" + $tableCount + ")不一致，将按较小数量处理。")
    }

    $limit = $tableTitles.Count
    if ($tableCount -lt $limit) {
        $limit = $tableCount
    }

    for ($i = $limit; $i -ge 1; $i--) {
        $title = [string]$tableTitles[$i - 1]
        $table = $document.Tables.Item($i)
        Insert-TableCaption -Table $table -Title $title
    }

    if ($AddImageCaptions) {
        $imageCount = $document.InlineShapes.Count
        for ($i = $imageCount; $i -ge 1; $i--) {
            $shape = $document.InlineShapes.Item($i)
            $range = $shape.Range.Duplicate()
            $range.InsertCaption("图", "", $null, 1, $false)
        }
    }

    $document.Fields.Update() | Out-Null
    Set-CaptionStyleForNearbyParagraphs -Document $document -StyleName $CaptionStyleName
    $document.Save()

    Write-Host ("已完成题注后处理: " + $resolvedDocument)
    Write-Host ("表格数量: " + $tableCount + "，插入表题注: " + $limit)
    Write-Host ("按窗口自动调整表格: " + $autoFitCount)
    if ($AddImageCaptions) {
        Write-Host ("插入图题注: " + $imageCount)
    }
} finally {
    if ($null -ne $document) {
        $document.Close($true)
    }
    if ($null -ne $word) {
        $word.Quit()
    }
}
