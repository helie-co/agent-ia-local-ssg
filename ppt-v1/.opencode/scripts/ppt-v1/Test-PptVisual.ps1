param(
    [Parameter(Mandatory = $true)]
    [string]$PresentationPath,
    [string]$ExportDirectory = (Join-Path (Get-Location) "template-ppt\visual-check"),
    [string]$ReportPath = (Join-Path (Get-Location) "template-ppt\visual-check-report.json"),
    [double]$MinimumTextGap = 0,
    [switch]$Visible
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $PresentationPath)) { throw "Presentation introuvable: $PresentationPath" }

$PresentationPath = (Resolve-Path -LiteralPath $PresentationPath).Path
$ExportDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportDirectory)
$ReportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportPath)

if (-not (Test-Path -LiteralPath $ExportDirectory)) {
    New-Item -ItemType Directory -Path $ExportDirectory | Out-Null
}

function Get-TextShapes {
    param($Slide)

    $items = @()
    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        $shape = $Slide.Shapes.Item($i)
        try {
            if ($shape.Type -eq 6) {
                for ($j = 1; $j -le $shape.GroupItems.Count; $j++) {
                    $groupItem = $shape.GroupItems.Item($j)
                    try {
                        if ($groupItem.HasTextFrame -and $groupItem.TextFrame.HasText) {
                            $items += [pscustomobject]@{
                                Shape = $groupItem
                                Text = $groupItem.TextFrame.TextRange.Text
                                Top = [double]$groupItem.Top
                                Left = [double]$groupItem.Left
                                Width = [double]$groupItem.Width
                                Height = [double]$groupItem.Height
                            }
                        }
                    } catch {
                    }
                }
            }
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $items += [pscustomobject]@{
                    Shape = $shape
                    Text = $shape.TextFrame.TextRange.Text
                    Top = [double]$shape.Top
                    Left = [double]$shape.Left
                    Width = [double]$shape.Width
                    Height = [double]$shape.Height
                }
            }
        } catch {
        }
    }
    return $items
}

$powerPoint = $null
$presentation = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    try { $powerPoint.Visible = if ($Visible) { -1 } else { 0 } } catch { $powerPoint.Visible = -1 }
    try { $powerPoint.DisplayAlerts = 1 } catch { }

    $presentation = $powerPoint.Presentations.Open($PresentationPath, $false, $false, $Visible.IsPresent)

    $issues = @()
    $slides = @()

    for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
        $slide = $presentation.Slides.Item($slideIndex)
        $imagePath = Join-Path $ExportDirectory ("slide-{0:D2}.png" -f $slideIndex)
        $slide.Export($imagePath, "PNG", 1600, 900)

        $slideIssues = @()
        $textEntries = @(Get-TextShapes $slide)

        if ($slideIndex -eq 1) {
            $coverTitleEntry = @($textEntries | Sort-Object Top, Left | Select-Object -First 1)
            if ($coverTitleEntry.Count -gt 0) {
                $coverTitle = ([string]$coverTitleEntry[0].Text).Replace("`r", " ").Replace("`n", " ").Trim()
                $lineCount = (([string]$coverTitleEntry[0].Text) -split "(`r`n|`r|`n)").Count
                if ($coverTitle.Length -gt 45 -or $lineCount -gt 1) {
                    $slideIssues += [pscustomobject]@{
                        type = "cover_title_too_long"
                        message = "Titre de cover trop long ou sur plusieurs lignes"
                        text = $coverTitle
                        length = $coverTitle.Length
                        maxLength = 45
                    }
                }
            }
        }

        foreach ($entry in $textEntries) {
            $text = [string]$entry.Text
            $shape = $entry.Shape

            if ($text -match "\{\{[^}]+\}\}") {
                $slideIssues += [pscustomobject]@{
                    type = "placeholder"
                    message = "Placeholder non remplace"
                    text = $text
                }
            }

            try {
                $boundHeight = [double]$shape.TextFrame2.TextRange.BoundHeight
                $boundWidth = [double]$shape.TextFrame2.TextRange.BoundWidth
                $heightLimit = [double]$shape.Height + 4
                $widthLimit = [double]$shape.Width + 4

                $isDecorativeNumber = $text.Trim() -match "^\d{1,2}$"
                if (-not $isDecorativeNumber -and ($boundHeight -gt $heightLimit -or $boundWidth -gt $widthLimit)) {
                    $preview = $text.Replace("`r", " ").Replace("`n", " ").Trim()
                    if ($preview.Length -gt 120) { $preview = $preview.Substring(0, 120) + "..." }
                    $slideIssues += [pscustomobject]@{
                        type = "overflow"
                        message = "Texte potentiellement hors zone"
                        text = $preview
                        boundWidth = [math]::Round($boundWidth, 2)
                        shapeWidth = [math]::Round([double]$shape.Width, 2)
                        boundHeight = [math]::Round($boundHeight, 2)
                        shapeHeight = [math]::Round([double]$shape.Height, 2)
                    }
                }
            } catch {
            }
        }

        $orderedTextEntries = @($textEntries | Sort-Object Top, Left)
        for ($i = 0; $i -lt $orderedTextEntries.Count; $i++) {
            if ($slideIndex -eq 1) { continue }
            for ($j = $i + 1; $j -lt $orderedTextEntries.Count; $j++) {
                $upper = $orderedTextEntries[$i]
                $lower = $orderedTextEntries[$j]

                if (([string]$upper.Text).Trim() -match "^\d{1,2}$" -or ([string]$lower.Text).Trim() -match "^\d{1,2}$") { continue }

                if ($lower.Top -lt $upper.Top) { continue }
                if ($upper.Top -lt 45) { continue }

                $left = [math]::Max([double]$upper.Left, [double]$lower.Left)
                $right = [math]::Min(([double]$upper.Left + [double]$upper.Width), ([double]$lower.Left + [double]$lower.Width))
                $horizontalOverlap = $right - $left
                if ($horizontalOverlap -le 20) { continue }

                $upperBottom = [double]$upper.Top + [double]$upper.Height
                try {
                    $boundHeight = [double]$upper.Shape.TextFrame2.TextRange.BoundHeight
                    if ($boundHeight -gt 0) { $upperBottom = [double]$upper.Top + $boundHeight }
                } catch {
                }

                $minimumGap = $MinimumTextGap
                $gap = [double]$lower.Top - $upperBottom
                if ($gap -lt $minimumGap) {
                    $upperPreview = ([string]$upper.Text).Replace("`r", " ").Replace("`n", " ").Trim()
                    $lowerPreview = ([string]$lower.Text).Replace("`r", " ").Replace("`n", " ").Trim()
                    if ($upperPreview.Length -gt 70) { $upperPreview = $upperPreview.Substring(0, 70) + "..." }
                    if ($lowerPreview.Length -gt 70) { $lowerPreview = $lowerPreview.Substring(0, 70) + "..." }
                    $slideIssues += [pscustomobject]@{
                        type = "overlap"
                        message = "Espacement insuffisant entre deux zones de texte"
                        text = "$upperPreview -> $lowerPreview"
                        gap = [math]::Round($gap, 2)
                        minimumGap = $minimumGap
                    }
                }
            }
        }

        $slides += [pscustomobject]@{
            index = $slideIndex
            image = $imagePath
            issues = $slideIssues
        }

        foreach ($issue in $slideIssues) {
            $issues += [pscustomobject]@{
                slide = $slideIndex
                type = $issue.type
                message = $issue.message
                text = $issue.text
            }
        }
    }

    $report = [pscustomobject]@{
        presentation = $PresentationPath
        exportDirectory = $ExportDirectory
        slideCount = $presentation.Slides.Count
        issueCount = $issues.Count
        issues = $issues
        slides = $slides
    }

    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ReportPath -Encoding UTF8

    "Verification visuelle terminee: $ReportPath"
    "Slides exportees: $ExportDirectory"
    "Issues=$($issues.Count)"
    foreach ($issue in $issues) {
        "- Slide $($issue.slide): $($issue.type) - $($issue.text)"
    }
} finally {
    if ($presentation) { $presentation.Close() }
    if ($powerPoint) { $powerPoint.Quit() }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
