param(
    [string]$OutputPath = (Join-Path $PSScriptRoot "CardTemplate.template.pptx"),
    [switch]$Visible
)

$ErrorActionPreference = "Stop"

$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$outputParent = Split-Path -Parent $OutputPath
if ($outputParent -and -not (Test-Path -LiteralPath $outputParent)) {
    New-Item -ItemType Directory -Path $outputParent | Out-Null
}

function Add-TextBox {
    param($Slide, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [string]$Text, [int]$FontSize = 18, [bool]$Bold = $false)

    $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height) | Out-Null
    $shape = $Slide.Shapes($Slide.Shapes.Count)
    $shape.TextFrame.TextRange.Text = $Text
    $shape.TextFrame.TextRange.Font.Size = $FontSize
    $shape.TextFrame.TextRange.Font.Bold = $Bold
}

function Add-RoundedRect {
    param($Slide, [double]$Left, [double]$Top, [double]$Width, [double]$Height)

    $Slide.Shapes.AddShape(5, $Left, $Top, $Width, $Height) | Out-Null
    $shape = $Slide.Shapes($Slide.Shapes.Count)
    $shape.Fill.ForeColor.RGB = 0xF0F0F0
    $shape.Fill.Visible = -1
    $shape.Line.Visible = $false
}

$powerPoint = $null
$presentation = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    try { $powerPoint.Visible = if ($Visible) { -1 } else { 0 } } catch { $powerPoint.Visible = -1 }
    try { $powerPoint.DisplayAlerts = 1 } catch { }

    $presentation = $powerPoint.Presentations.Add()

    $slideWidth = 720
    $slideHeight = 540
    $margin = 36
    $cardSpacing = 12

    # Add cover slide (presentation starts empty)
    $cover = $presentation.Slides.Add(1, 1)
    $coverWidth = $slideWidth
    Add-TextBox $cover 36 120 ($coverWidth - 72) 60 "{{DECK_TITLE}}" 36 $true
    Add-TextBox $cover 36 190 ($coverWidth - 72) 40 "{{DECK_SUBTITLE}}" 20 $false
    Add-TextBox $cover 36 ($slideHeight - 100) 200 30 "{{DATE}}" 16 $false
    Add-TextBox $cover 36 ($slideHeight - 70) 400 30 "{{AUTHORS}}" 14 $false

    function New-CardSlide {
        param($Presentation, [int]$CardCount, [int]$Columns)

        $slide = $Presentation.Slides.Add(1, 1)
        Add-TextBox $slide 36 20 ($slideWidth - 72) 36 "{{TITLE}}" 28 $true
        Add-TextBox $slide 36 58 ($slideWidth - 72) 20 "{{SUBTITLE}}" 16 $false

        $rows = [Math]::Ceiling([double]$CardCount / $Columns)
        $cardWidth = ($slideWidth - 2 * $margin - ($Columns - 1) * $cardSpacing) / $Columns
        $cardAreaTop = 90
        $cardAreaHeight = $slideHeight - $cardAreaTop - $margin
        $cardHeight = ($cardAreaHeight - ($rows - 1) * $cardSpacing) / $rows

        for ($i = 0; $i -lt $CardCount; $i++) {
            $col = $i % $Columns
            $row = [Math]::Floor([double]$i / $Columns)
            $left = $margin + $col * ($cardWidth + $cardSpacing)
            $top = $cardAreaTop + $row * ($cardHeight + $cardSpacing)
            $num = $i + 1

            Add-RoundedRect $slide $left $top $cardWidth $cardHeight
            Add-TextBox $slide ($left + 8) ($top + 6) ($cardWidth - 16) 22 "{{CARD${num}_TITLE}}" 14 $true
            Add-TextBox $slide ($left + 8) ($top + 30) ($cardWidth - 16) ($cardHeight - 38) "{{CARD${num}_BODY}}" 12 $false
        }
        return $slide
    }

    [void](New-CardSlide $presentation 2 2)
    [void](New-CardSlide $presentation 3 3)
    [void](New-CardSlide $presentation 4 2)
    [void](New-CardSlide $presentation 6 3)

    $classic = $presentation.Slides.Add(1, 1)
    Add-TextBox $classic 36 20 ($slideWidth - 72) 36 "{{TITLE}}" 28 $true
    Add-TextBox $classic 36 58 ($slideWidth - 72) 20 "{{SUBTITLE}}" 16 $false
    Add-TextBox $classic 36 100 ($slideWidth - 72) 380 "{{BODY}}" 16 $false

    $presentation.SaveAs($OutputPath, 24)
    $presentation.Saved = $true
    "Template cree: $OutputPath"
    "Layouts: cover, card_2, card_3, card_4, card_6, classic_bullets"
} finally {
    if ($presentation) { try { $presentation.Close() } catch { } }
    if ($powerPoint) { try { $powerPoint.Quit() } catch { } }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
