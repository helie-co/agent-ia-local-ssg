param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePptx,
    [string]$OutputMarkdown = (Join-Path (Join-Path (Get-Location) "ppt") "deck.extracted.md"),
    [switch]$NoOcr
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $SourcePptx)) { throw "PowerPoint source introuvable: $SourcePptx" }

$SourcePptx = (Resolve-Path -LiteralPath $SourcePptx).Path
$OutputMarkdown = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputMarkdown)
$outputParent = Split-Path -Parent $OutputMarkdown
if ($outputParent -and -not (Test-Path -LiteralPath $outputParent)) {
    New-Item -ItemType Directory -Path $outputParent | Out-Null
}

function Get-TesseractPath {
    $tesseractCmd = Get-Command tesseract -ErrorAction SilentlyContinue
    if ($tesseractCmd) { return $tesseractCmd.Source }

    $candidatePaths = @(
        (Join-Path $env:LOCALAPPDATA "Programs\Tesseract-OCR\tesseract.exe"),
        "C:\Program Files\Tesseract-OCR\tesseract.exe",
        "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
    )

    foreach ($candidatePath in $candidatePaths) {
        if (Test-Path -LiteralPath $candidatePath) { return $candidatePath }
    }

    return $null
}

function Get-TesseractLanguage {
    param([string]$TesseractPath)

    if (-not $TesseractPath) { return $null }
    $tessData = Join-Path (Split-Path -Parent $TesseractPath) "tessdata"
    if (Test-Path -LiteralPath (Join-Path $tessData "fra.traineddata")) { return "fra+eng" }
    return "eng"
}

function Normalize-ExtractedText {
    param([string]$Text)

    if ($null -eq $Text) { return "" }
    return ([string]$Text).
        Replace([char]0x000B, "`n").
        Replace([char]0x00A0, " ").
        Trim()
}

function Split-TextLines {
    param([string]$Text)

    $cleanText = Normalize-ExtractedText $Text
    if ([string]::IsNullOrWhiteSpace($cleanText)) { return @() }

    $result = @()
    foreach ($line in ($cleanText -split "(`r`n|`r|`n)")) {
        $clean = $line.Trim()
        if (-not [string]::IsNullOrWhiteSpace($clean)) { $result += $clean }
    }
    return $result
}

function Escape-MarkdownCell {
    param([string]$Text)

    $clean = Normalize-ExtractedText $Text
    $clean = (($clean -split "(`r`n|`r|`n)") | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join " / "
    return $clean.Replace("|", "\|")
}

function Test-OcrTextUseful {
    param([string[]]$Lines)

    if (-not $Lines -or $Lines.Count -eq 0) { return $false }
    $joined = ($Lines -join " ")
    $signal = ([regex]::Matches($joined, "[\p{L}\p{Nd}]")).Count
    return ($signal -ge 12)
}

function Add-TextShape {
    param($Shape, [System.Collections.ArrayList]$Items)

    try {
        if ($Shape.Type -eq 6) {
            for ($i = 1; $i -le $Shape.GroupItems.Count; $i++) {
                Add-TextShape $Shape.GroupItems.Item($i) $Items
            }
            return
        }

        if ($Shape.HasTextFrame -and $Shape.TextFrame.HasText) {
            $text = $Shape.TextFrame.TextRange.Text.Trim()
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                [void]$Items.Add([pscustomobject]@{
                    Kind = "text"
                    Text = $text
                    Top = [double]$Shape.Top
                    Left = [double]$Shape.Left
                })
            }
        }
    } catch {
    }
}

function Add-TableShape {
    param($Shape, [System.Collections.ArrayList]$Items)

    try {
        if ($Shape.Type -eq 6) {
            for ($i = 1; $i -le $Shape.GroupItems.Count; $i++) {
                Add-TableShape $Shape.GroupItems.Item($i) $Items
            }
            return
        }

        if (-not $Shape.HasTable) { return }

        $table = $Shape.Table
        $rows = @()
        for ($r = 1; $r -le $table.Rows.Count; $r++) {
            $row = @()
            for ($c = 1; $c -le $table.Columns.Count; $c++) {
                $cellText = ""
                try { $cellText = $table.Cell($r, $c).Shape.TextFrame.TextRange.Text } catch { }
                $row += (Escape-MarkdownCell $cellText)
            }
            $rows += ,$row
        }

        [void]$Items.Add([pscustomobject]@{
            Kind = "table"
            Rows = $rows
            RowCount = [int]$table.Rows.Count
            ColumnCount = [int]$table.Columns.Count
            Top = [double]$Shape.Top
            Left = [double]$Shape.Left
        })
    } catch {
    }
}

function Get-OcrShapeItems {
    param(
        $Shape,
        [string]$TesseractPath,
        [string]$TesseractLanguage,
        [string]$OcrDirectory,
        [int]$SlideIndex,
        [ref]$ImageIndex
    )

    if (-not $TesseractPath) { return @() }

    $results = @()

    try {
        if ($Shape.Type -eq 6) {
            for ($i = 1; $i -le $Shape.GroupItems.Count; $i++) {
                $results += @(Get-OcrShapeItems $Shape.GroupItems.Item($i) $TesseractPath $TesseractLanguage $OcrDirectory $SlideIndex $ImageIndex)
            }
            return $results
        }

        $isPicture = ($Shape.Type -eq 13 -or $Shape.Type -eq 11)
        if (-not $isPicture) { return @() }

        # Ignore small logos and decorative icons; OCR is useful mostly on screenshots/photos.
        if ([double]$Shape.Width -lt 120 -or [double]$Shape.Height -lt 60) { return @() }

        $ImageIndex.Value++
        $baseName = "slide_{0:D3}_image_{1:D2}" -f $SlideIndex, $ImageIndex.Value
        $imagePath = Join-Path $OcrDirectory ($baseName + ".png")
        $outputBase = Join-Path $OcrDirectory $baseName
        $outputTextPath = $outputBase + ".txt"

        try { $Shape.Export($imagePath, 2) } catch { return @() }
        if (-not (Test-Path -LiteralPath $imagePath)) { return @() }

        $null = & $TesseractPath $imagePath $outputBase -l $TesseractLanguage --psm 6 2>$null
        if (-not (Test-Path -LiteralPath $outputTextPath)) { return @() }

        $ocrText = [System.IO.File]::ReadAllText($outputTextPath, [System.Text.Encoding]::UTF8)
        $ocrLines = @(Split-TextLines $ocrText)
        if (-not (Test-OcrTextUseful $ocrLines)) { return @() }

        $results += [pscustomobject]@{
            Kind = "ocr"
            Text = ($ocrLines -join "`n")
            ImageNumber = [int]$ImageIndex.Value
            Top = [double]$Shape.Top
            Left = [double]$Shape.Left
        }
    } catch {
    }

    return $results
}

$powerPoint = $null
$presentation = $null
$ocrDirectory = Join-Path $outputParent "deck.extracted.ocr"
$tesseractPath = if ($NoOcr) { $null } else { Get-TesseractPath }
$tesseractLanguage = if ($tesseractPath) { Get-TesseractLanguage $tesseractPath } else { $null }

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    try { $powerPoint.Visible = 0 } catch { $powerPoint.Visible = -1 }
    try { $powerPoint.DisplayAlerts = 1 } catch { }
    $presentation = $powerPoint.Presentations.Open($SourcePptx, $false, $false, $false)

    $lines = @()
    $lines += "# Reprise de presentation"
    $lines += ""
    $lines += "Source: $SourcePptx"
    if ($NoOcr) {
        $lines += "OCR images: desactive"
    } elseif ($tesseractPath) {
        $lines += "OCR images: $tesseractPath"
        $lines += "OCR langue: $tesseractLanguage"
        if (Test-Path -LiteralPath $ocrDirectory) { Remove-Item -LiteralPath $ocrDirectory -Recurse -Force }
        New-Item -ItemType Directory -Path $ocrDirectory | Out-Null
    } else {
        $lines += "OCR images: tesseract introuvable, lancer /ppt-v1 --install pour extraire le texte des photos"
    }
    $lines += ""

    for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
        $slide = $presentation.Slides.Item($slideIndex)
        $items = New-Object System.Collections.ArrayList
        $imageIndex = 0
        for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
            $shape = $slide.Shapes.Item($shapeIndex)
            Add-TableShape $shape $items
            Add-TextShape $shape $items
            foreach ($ocrItem in @(Get-OcrShapeItems $shape $tesseractPath $tesseractLanguage $ocrDirectory $slideIndex ([ref]$imageIndex))) {
                [void]$items.Add($ocrItem)
            }
        }

        if ($tesseractPath -and -not (@($items | Where-Object { $_.Kind -eq "ocr" }).Count)) {
            $ocrPattern = "slide_{0:D3}_image_*.txt" -f $slideIndex
            $ocrTextFiles = @(Get-ChildItem -LiteralPath $ocrDirectory -Filter $ocrPattern -ErrorAction SilentlyContinue | Sort-Object Name)
            foreach ($ocrTextFile in $ocrTextFiles) {
                $ocrText = [System.IO.File]::ReadAllText($ocrTextFile.FullName, [System.Text.Encoding]::UTF8)
                $ocrLines = @(Split-TextLines $ocrText)
                if (Test-OcrTextUseful $ocrLines) {
                    $imageNumber = 1
                    if ($ocrTextFile.BaseName -match "image_(\d+)$") { $imageNumber = [int]$Matches[1] }
                    [void]$items.Add([pscustomobject]@{
                        Kind = "ocr"
                        Text = ($ocrLines -join "`n")
                        ImageNumber = $imageNumber
                        Top = 99999
                        Left = $imageNumber
                    })
                }
            }
        }

        $orderedItems = @($items | Sort-Object Top, Left)
        $texts = @($orderedItems | Where-Object { $_.Kind -eq "text" } | ForEach-Object { $_.Text })
        $expanded = @()
        foreach ($text in $texts) {
            $expanded += @(Split-TextLines $text)
        }

        $title = if ($expanded.Count -gt 0) { $expanded[0] } else { "Slide $slideIndex" }
        $subtitle = if ($expanded.Count -gt 1) { $expanded[1] } else { "" }
        $body = @()
        if ($expanded.Count -gt 2) { $body = @($expanded | Select-Object -Skip 2) }

        $tables = @($orderedItems | Where-Object { $_.Kind -eq "table" })
        $ocrItems = @($orderedItems | Where-Object { $_.Kind -eq "ocr" })
        $compatibleTables = @($tables | Where-Object { $_.ColumnCount -ge 2 -and $_.RowCount -ge 2 })
        $layout = if ($slideIndex -eq 1) { "cover" } elseif ($tables.Count -eq 1 -and $compatibleTables.Count -eq 1 -and $body.Count -le 2) { "table_dynamic" } else { "classic_bullets" }
        $lines += "## Slide $slideIndex - $title"
        $lines += "Layout: $layout"
        $lines += ""

        if ($slideIndex -eq 1) {
            $lines += "Titre: $title"
            $lines += "Sous-titre: $subtitle"
            $lines += "Date: " + ($(if ($body.Count -gt 0) { $body[0] } else { "" }))
            $lines += "Auteurs: " + ($(if ($body.Count -gt 1) { $body[1] } else { "" }))
        } else {
            $lines += "Titre: $title"
            $lines += "Sous-titre: $subtitle"
            $lines += ""
            if ($body.Count -gt 0) {
                $lines += "Points:"
                foreach ($item in $body) { $lines += "- $item" }
                $lines += ""
            }

            foreach ($tableItem in $tables) {
                if ($tableItem.Rows.Count -gt 0) {
                    $header = @($tableItem.Rows[0])
                    $lines += "Colonnes:"
                    foreach ($column in $header) { $lines += "- $column" }
                    $lines += ""

                    if ($tableItem.Rows.Count -gt 1) {
                        $lines += "Lignes:"
                        for ($rowIndex = 1; $rowIndex -lt $tableItem.Rows.Count; $rowIndex++) {
                            $lines += "- " + (@($tableItem.Rows[$rowIndex]) -join " | ")
                        }
                        $lines += ""
                    }
                }
            }

            foreach ($ocrItem in $ocrItems) {
                $label = if ($ocrItems.Count -gt 1) { "Texte OCR image $($ocrItem.ImageNumber):" } else { "Texte OCR image:" }
                $lines += $label
                foreach ($ocrLine in (Split-TextLines $ocrItem.Text)) { $lines += "- $ocrLine" }
                $lines += ""
            }
        }
        $lines += ""
    }

    [System.IO.File]::WriteAllLines($OutputMarkdown, $lines, [System.Text.UTF8Encoding]::new($false))
    "Extraction terminee: $OutputMarkdown"
} finally {
    if ($presentation) { $presentation.Close() }
    if ($powerPoint) { $powerPoint.Quit() }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
