param(
    [string]$ProjectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..\..\..")).Path,
    [string]$TemplatePath = (Join-Path $ProjectRoot "template-ppt\Orange.template.pptx"),
    [string]$InputJson = (Join-Path $PSScriptRoot "slides.example.json"),
    [string]$LayoutsJson = (Join-Path $PSScriptRoot "layouts.json"),
    [string]$OutputPath = (Join-Path $ProjectRoot "Presentation.pptx"),
    [switch]$Visible
)

$ErrorActionPreference = "Stop"

function ConvertTo-Hashtable {
    param([object]$Object)

    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IEnumerable] -and $Object -isnot [string] -and $Object -isnot [System.Management.Automation.PSCustomObject]) {
        $array = @()
        foreach ($item in $Object) { $array += ,(ConvertTo-Hashtable $item) }
        return $array
    }
    if ($Object -is [System.Management.Automation.PSCustomObject]) {
        $hash = @{}
        foreach ($property in $Object.PSObject.Properties) {
            $hash[$property.Name] = ConvertTo-Hashtable $property.Value
        }
        return $hash
    }
    return $Object
}

function Get-TextShapes {
    param($Slide)

    $shapes = @()
    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        $shape = $Slide.Shapes.Item($i)
        try {
            if ($shape.Type -eq 6) {
                for ($j = 1; $j -le $shape.GroupItems.Count; $j++) {
                    $groupItem = $shape.GroupItems.Item($j)
                    try {
                        if ($groupItem.HasTextFrame -and $groupItem.TextFrame.HasText) {
                            $text = $groupItem.TextFrame.TextRange.Text
                            if (-not [string]::IsNullOrWhiteSpace($text)) {
                                $shapes += [pscustomobject]@{
                                    Shape = $groupItem
                                    Text = $text
                                    Top = [double]$groupItem.Top
                                    Left = [double]$groupItem.Left
                                }
                            }
                        }
                    } catch {
                    }
                }
            }
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $text = $shape.TextFrame.TextRange.Text
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    $shapes += [pscustomobject]@{
                        Shape = $shape
                        Text = $text
                        Top = [double]$shape.Top
                        Left = [double]$shape.Left
                    }
                }
            }
        } catch {
            # Some grouped or special shapes do not expose a regular text frame.
        }
    }

    return $shapes | Sort-Object Top, Left
}

function Set-ShapeTextPreservingStyle {
    param($Shape, [string]$Value)

    if ($null -eq $Value) { $Value = "" }
    $Shape.TextFrame.TextRange.Text = $Value
}

function ConvertTo-SlideText {
    param([object]$Value)

    if ($null -eq $Value) { return "" }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string] -and $Value -isnot [hashtable]) {
        $lines = @()
        foreach ($item in $Value) {
            if (-not [string]::IsNullOrWhiteSpace([string]$item)) {
                $lines += [string]$item
            }
        }
        return ($lines -join "`r")
    }
    return [string]$Value
}

function Get-DefaultAuthor {
    $registryPaths = @(
        "HKCU:\Software\Microsoft\Office\Common\UserInfo",
        "HKCU:\Software\Microsoft\Office\16.0\Common\UserInfo"
    )

    foreach ($path in $registryPaths) {
        try {
            if (Test-Path -LiteralPath $path) {
                $userName = (Get-ItemProperty -LiteralPath $path).UserName
                if (-not [string]::IsNullOrWhiteSpace($userName)) { return [string]$userName }
            }
        } catch {
        }
    }

    try {
        if (-not [string]::IsNullOrWhiteSpace($env:USERNAME)) { return [string]$env:USERNAME }
    } catch {
    }

    return ""
}

function Test-InvalidAuthor {
    param([object]$Author)

    if ([string]::IsNullOrWhiteSpace([string]$Author)) { return $true }
    $value = ([string]$Author).Trim()
    return @("OpenCode", "opencode", "Assistant", "AI", "IA") -contains $value
}

function Replace-Placeholders {
    param($Slide, [hashtable]$Values)

    $replacementCount = 0
    foreach ($entry in Get-TextShapes $Slide) {
        $text = $entry.Shape.TextFrame.TextRange.Text
        $newText = $text
        foreach ($key in $Values.Keys) {
            if ($newText.Contains($key)) {
                $newText = $newText.Replace($key, (ConvertTo-SlideText $Values[$key]))
            }
        }
        if ($newText -ne $text) {
            Set-ShapeTextPreservingStyle $entry.Shape $newText
            $replacementCount++
        }
    }
    return $replacementCount
}

function Apply-FallbackSlots {
    param($Slide, [string]$LayoutName, [hashtable]$Values)

    $slotMap = @{
        "cover" = @("{{DECK_TITLE}}", "{{DECK_SUBTITLE}}", "{{DATE}}", "{{AUTHORS}}")
        "classic_bullets" = @("{{BODY}}", "{{TITLE}}", "{{SUBTITLE}}")
        "two_columns_bullets" = @("{{TITLE}}", "{{SUBTITLE}}", "{{LEFT_TITLE}}", "{{LEFT_BODY}}", "{{RIGHT_TITLE}}", "{{RIGHT_BODY}}")
    }

    if (-not $slotMap.ContainsKey($LayoutName)) { return 0 }

    $textShapes = @(Get-TextShapes $Slide)
    $slots = @($slotMap[$LayoutName])
    $limit = [Math]::Min($textShapes.Count, $slots.Count)
    $count = 0

    for ($i = 0; $i -lt $limit; $i++) {
        $slot = $slots[$i]
        if ($Values.ContainsKey($slot)) {
            Set-ShapeTextPreservingStyle $textShapes[$i].Shape (ConvertTo-SlideText $Values[$slot])
            $count++
        }
    }

    return $count
}

function Get-FirstTableShape {
    param($Slide)

    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        $shape = $Slide.Shapes.Item($i)
        try {
            if ($shape.HasTable) { return $shape }
        } catch {
        }
    }
    return $null
}

function Get-TitleShape {
    param($Slide, [string]$Title)

    $textShapes = @(Get-TextShapes $Slide)
    if (-not [string]::IsNullOrWhiteSpace($Title)) {
        $matchingShapes = @($textShapes | Where-Object { $_.Text.Trim() -eq $Title.Trim() } | Sort-Object Top, Left)
        if ($matchingShapes.Count -gt 0) { return $matchingShapes[0].Shape }
    }

    $candidateShapes = @($textShapes | Where-Object {
        $_.Top -lt 140 -and
        $_.Text -notmatch '^\s*\d+\s*/\s*\d+\s*$' -and
        $_.Text -notmatch '^\s*\{\{PAGE_NUMBER\}\}\s*$' -and
        $_.Text -notmatch '^\s*\{\{FOOTER_TITLE\}\}\s*$'
    } | Sort-Object Top, Left)

    if ($candidateShapes.Count -gt 0) { return $candidateShapes[0].Shape }
    return $null
}

function Align-TableToTitle {
    param($TableShape, $TitleShape)

    if (-not $TableShape -or -not $TitleShape) { return }
    try {
        $TableShape.Left = [double]$TitleShape.Left
        $TableShape.Width = [double]$TitleShape.Width
    } catch {
    }
}

function Set-TableText {
    param($Cell, [object]$Value)

    $text = ConvertTo-SlideText $Value
    $Cell.Shape.TextFrame.TextRange.Text = $text
}

function Resize-Table {
    param($Table, [int]$ColumnCount, [int]$RowCount)

    while ($Table.Columns.Count -lt $ColumnCount) { [void]$Table.Columns.Add() }
    while ($Table.Columns.Count -gt $ColumnCount) { $Table.Columns.Item($Table.Columns.Count).Delete() }
    while ($Table.Rows.Count -lt $RowCount) { [void]$Table.Rows.Add() }
    while ($Table.Rows.Count -gt $RowCount) { $Table.Rows.Item($Table.Rows.Count).Delete() }
}

function Normalize-TableColumns {
    param($Table, [double]$TotalWidth)

    if ($Table.Columns.Count -le 0) { return }
    $columnWidth = $TotalWidth / $Table.Columns.Count
    for ($i = 1; $i -le $Table.Columns.Count; $i++) {
        try { $Table.Columns.Item($i).Width = $columnWidth } catch { }
    }
}

function Set-DynamicTable {
    param($Slide, [hashtable]$SlideSpec, [hashtable]$Layout)

    if (-not $SlideSpec.ContainsKey("table") -or -not $SlideSpec.table) { return }

    $tableShape = Get-FirstTableShape $Slide
    if (-not $tableShape) { throw "Aucun tableau PowerPoint trouve pour le layout table_dynamic" }

    $columns = @($SlideSpec.table.columns)
    $rows = @($SlideSpec.table.rows)

    if ($columns.Count -lt 2) { throw "table_dynamic exige au moins 2 colonnes" }

    foreach ($row in $rows) {
        if (@($row).Count -ne $columns.Count) {
            throw "Chaque ligne table_dynamic doit contenir le meme nombre de cellules que les colonnes"
        }
    }

    $table = $tableShape.Table
    $titleShape = Get-TitleShape $Slide ([string]$SlideSpec.placeholders["{{TITLE}}"])
    Align-TableToTitle $tableShape $titleShape
    $targetWidth = [double]$tableShape.Width
    if ($titleShape) { $targetWidth = [double]$titleShape.Width }
    $targetRows = [Math]::Max(2, $rows.Count + 1)
    Resize-Table $table $columns.Count $targetRows
    Align-TableToTitle $tableShape $titleShape
    Normalize-TableColumns $table $targetWidth

    for ($c = 1; $c -le $columns.Count; $c++) {
        Set-TableText $table.Cell(1, $c) $columns[$c - 1]
    }

    for ($r = 1; $r -le $rows.Count; $r++) {
        $row = @($rows[$r - 1])
        for ($c = 1; $c -le $columns.Count; $c++) {
            Set-TableText $table.Cell($r + 1, $c) $row[$c - 1]
        }
    }

    for ($r = $rows.Count + 2; $r -le $table.Rows.Count; $r++) {
        for ($c = 1; $c -le $table.Columns.Count; $c++) {
            Set-TableText $table.Cell($r, $c) ""
        }
    }
}

function Get-RemainingPlaceholders {
    param($Presentation)

    $remaining = @()
    for ($s = 1; $s -le $Presentation.Slides.Count; $s++) {
        foreach ($entry in Get-TextShapes $Presentation.Slides.Item($s)) {
            if ($entry.Text -match "\{\{[^}]+\}\}") {
                $remaining += "Slide ${s}: $($entry.Text)"
            }
        }
    }
    return $remaining
}

function Test-TextOverflow {
    param($Presentation)

    $warnings = @()
    for ($s = 1; $s -le $Presentation.Slides.Count; $s++) {
        foreach ($entry in Get-TextShapes $Presentation.Slides.Item($s)) {
            try {
                $boundHeight = [double]$entry.Shape.TextFrame2.TextRange.BoundHeight
                if ($boundHeight -gt ([double]$entry.Shape.Height + 4)) {
                    $preview = $entry.Shape.TextFrame.TextRange.Text.Replace("`r", " ").Replace("`n", " ")
                    if ($preview.Length -gt 80) { $preview = $preview.Substring(0, 80) + "..." }
                    $warnings += "Slide ${s}: texte potentiellement trop haut ($preview)"
                }
            } catch {
            }
        }
    }
    return $warnings
}

if (-not (Test-Path -LiteralPath $TemplatePath)) { throw "Template introuvable: $TemplatePath" }
if (-not (Test-Path -LiteralPath $InputJson)) { throw "JSON d'entree introuvable: $InputJson" }
if (-not (Test-Path -LiteralPath $LayoutsJson)) { throw "Catalogue de layouts introuvable: $LayoutsJson" }

$TemplatePath = (Resolve-Path -LiteralPath $TemplatePath).Path
$InputJson = (Resolve-Path -LiteralPath $InputJson).Path
$LayoutsJson = (Resolve-Path -LiteralPath $LayoutsJson).Path
$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)

$inputData = ConvertTo-Hashtable ((Get-Content -LiteralPath $InputJson -Raw -Encoding UTF8) | ConvertFrom-Json)
$layoutData = ConvertTo-Hashtable ((Get-Content -LiteralPath $LayoutsJson -Raw -Encoding UTF8) | ConvertFrom-Json)

$layoutsByName = @{}
foreach ($layout in $layoutData.layouts) { $layoutsByName[$layout.name] = $layout }

Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath -Force

$powerPoint = $null
$presentation = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    try {
        $powerPoint.Visible = if ($Visible) { -1 } else { 0 }
    } catch {
        $powerPoint.Visible = -1
    }
    try { $powerPoint.DisplayAlerts = 1 } catch { }
    $presentation = $powerPoint.Presentations.Open($OutputPath, $false, $false, $Visible.IsPresent)

    $generatedSlides = @()

    $coverLayout = $layoutsByName["cover"]
    $cover = $presentation.Slides.Item([int]$coverLayout.sourceSlide)
    $coverAuthors = $inputData.cover.authors
    if (Test-InvalidAuthor $coverAuthors) { $coverAuthors = Get-DefaultAuthor }
    $coverValues = @{
        "{{DECK_TITLE}}" = $inputData.cover.deckTitle
        "{{DECK_SUBTITLE}}" = $inputData.cover.deckSubtitle
        "{{DATE}}" = $inputData.cover.date
        "{{AUTHORS}}" = $coverAuthors
        "{{FOOTER_TITLE}}" = $inputData.cover.deckTitle
    }

    try {
        $presentation.BuiltInDocumentProperties.Item("Author").Value = $coverAuthors
        $presentation.BuiltInDocumentProperties.Item("Last Author").Value = $coverAuthors
    } catch {
    }
    $replaced = Replace-Placeholders $cover $coverValues
    if ($replaced -eq 0) { [void](Apply-FallbackSlots $cover "cover" $coverValues) }
    $generatedSlides += 1

    foreach ($slideSpec in $inputData.slides) {
        $layoutName = [string]$slideSpec.layout
        if (-not $layoutsByName.ContainsKey($layoutName)) { throw "Layout inconnu: $layoutName" }

        $layout = $layoutsByName[$layoutName]
        $sourceSlide = $presentation.Slides.Item([int]$layout.sourceSlide)
        $sourceSlide.Copy()
        $pasted = $presentation.Slides.Paste($presentation.Slides.Count + 1)
        $newSlide = $pasted.Item(1)
        $generatedSlides += $newSlide.SlideIndex

        $values = @{}
        foreach ($property in $slideSpec.placeholders.Keys) {
            $values[$property] = $slideSpec.placeholders[$property]
        }
        if ($slideSpec.body -and -not $values.ContainsKey("{{BODY}}")) {
            $values["{{BODY}}"] = $slideSpec.body
        }
        if (-not $values.ContainsKey("{{FOOTER_TITLE}}")) {
            $values["{{FOOTER_TITLE}}"] = $inputData.cover.deckTitle
        }
        if ($inputData.ContainsKey("defaults") -and $inputData.defaults) {
            $values["{{FOOTER_LEFT}}"] = $inputData.defaults.footerLeft
            $values["{{FOOTER_RIGHT}}"] = $inputData.defaults.footerRight
        }

        $replaced = Replace-Placeholders $newSlide $values
        if ($replaced -eq 0) { [void](Apply-FallbackSlots $newSlide $layoutName $values) }
        if ($layoutName -eq "table_dynamic") { Set-DynamicTable $newSlide $slideSpec $layout }
    }

    for ($i = $presentation.Slides.Count; $i -ge 1; $i--) {
        if ($generatedSlides -notcontains $i) {
            $presentation.Slides.Item($i).Delete()
        }
    }

    $contentCount = $presentation.Slides.Count - 1
    for ($i = 2; $i -le $presentation.Slides.Count; $i++) {
        $pageValue = [string]::Format($layoutData.rules.contentPageNumberFormat, ($i - 1), $contentCount)
        [void](Replace-Placeholders $presentation.Slides.Item($i) @{ "{{PAGE_NUMBER}}" = $pageValue })
    }

    $remaining = @(Get-RemainingPlaceholders $presentation)
    $overflow = @(Test-TextOverflow $presentation)

    $presentation.Save()

    "Generation terminee: $OutputPath"
    if ($remaining.Count -gt 0) {
        "Placeholders restants:"
        $remaining | ForEach-Object { "- $_" }
    }
    if ($overflow.Count -gt 0) {
        "Alertes rendu visuel:"
        $overflow | ForEach-Object { "- $_" }
    }
} finally {
    if ($presentation) { $presentation.Close() }
    if ($powerPoint) { $powerPoint.Quit() }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
