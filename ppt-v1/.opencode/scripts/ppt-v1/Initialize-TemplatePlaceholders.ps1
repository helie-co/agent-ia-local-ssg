param(
    [string]$ProjectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..\..\..")).Path,
    [string]$TemplatePath = (Join-Path $ProjectRoot "ppt-generator\Orange.template.pptx"),
    [switch]$Visible
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $TemplatePath)) { throw "Template introuvable: $TemplatePath" }

$TemplatePath = (Resolve-Path -LiteralPath $TemplatePath).Path

function Set-FirstMatchingText {
    param($Slide, [string]$Contains, [string]$Replacement)

    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        $shape = $Slide.Shapes.Item($i)
        try {
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $text = $shape.TextFrame.TextRange.Text
                if ($text -like "*$Contains*") {
                    $shape.TextFrame.TextRange.Text = $Replacement
                    return $true
                }
            }
        } catch {
        }
    }
    return $false
}

function Set-ClassicSlidePlaceholders {
    param($Slide)

    $textShapes = @()
    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        $shape = $Slide.Shapes.Item($i)
        try {
            if ($shape.Type -eq 6) {
                for ($j = 1; $j -le $shape.GroupItems.Count; $j++) {
                    $groupItem = $shape.GroupItems.Item($j)
                    try {
                        if ($groupItem.HasTextFrame -and $groupItem.TextFrame.HasText) {
                            $groupItem.TextFrame.TextRange.Text = ""
                        }
                    } catch {
                    }
                }
            }
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $textShapes += [pscustomobject]@{
                    Shape = $shape
                    Top = [double]$shape.Top
                    Left = [double]$shape.Left
                }
            }
        } catch {
        }
    }

    $ordered = @($textShapes | Sort-Object Top, Left)
    for ($i = 0; $i -lt $ordered.Count; $i++) {
        if ($i -eq 0) {
            $ordered[$i].Shape.TextFrame.TextRange.Text = "{{TITLE}}`r{{SUBTITLE}}"
        } else {
            $ordered[$i].Shape.TextFrame.TextRange.Text = "{{BODY}}"
        }
    }
}

function Set-TwoColumnsSlidePlaceholders {
    param($Slide)

    $textShapes = @()
    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        $shape = $Slide.Shapes.Item($i)
        try {
            if ($shape.Type -eq 6) {
                for ($j = 1; $j -le $shape.GroupItems.Count; $j++) {
                    $groupItem = $shape.GroupItems.Item($j)
                    try {
                        if ($groupItem.HasTextFrame -and $groupItem.TextFrame.HasText) {
                            $groupItem.TextFrame.TextRange.Text = ""
                        }
                    } catch {
                    }
                }
            }
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $textShapes += [pscustomobject]@{
                    Shape = $shape
                    Top = [double]$shape.Top
                    Left = [double]$shape.Left
                    Text = $shape.TextFrame.TextRange.Text
                }
            }
        } catch {
        }
    }

    $titleShapes = @($textShapes | Where-Object { $_.Top -lt 80 } | Sort-Object Top, Left)
    if ($titleShapes.Count -ge 1) { $titleShapes[0].Shape.TextFrame.TextRange.Text = "{{TITLE}}" }
    if ($titleShapes.Count -ge 2) { $titleShapes[1].Shape.TextFrame.TextRange.Text = "{{SUBTITLE}}" }

    $bodyShapes = @($textShapes | Where-Object { $_.Top -ge 80 } | Sort-Object Left)
    if ($bodyShapes.Count -ge 1) { $bodyShapes[0].Shape.TextFrame.TextRange.Text = "{{LEFT_BODY}}" }
    if ($bodyShapes.Count -ge 2) { $bodyShapes[1].Shape.TextFrame.TextRange.Text = "{{RIGHT_BODY}}" }
}

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
    $presentation = $powerPoint.Presentations.Open($TemplatePath, $false, $false, $Visible.IsPresent)

    $cover = $presentation.Slides.Item(1)
    [void](Set-FirstMatchingText $cover "Agile Day" "{{DECK_TITLE}}")
    [void](Set-FirstMatchingText $cover "REX Programme" "{{DECK_TITLE}}")
    [void](Set-FirstMatchingText $cover "juin" "{{DATE}}")
    [void](Set-FirstMatchingText $cover "Virginie" "{{AUTHORS}}")

    if ($presentation.Slides.Count -ge 2) {
        $classic = $presentation.Slides.Item(2)
        Set-ClassicSlidePlaceholders $classic
    }

    if ($presentation.Slides.Count -ge 3) {
        $twoColumns = $presentation.Slides.Item(3)
        Set-TwoColumnsSlidePlaceholders $twoColumns
    }

    $presentation.Save()
    "Template initialise: $TemplatePath"
} finally {
    if ($presentation) { $presentation.Close() }
    if ($powerPoint) { $powerPoint.Quit() }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
