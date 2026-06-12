param(
    [string]$TemplatePath = (Join-Path $PSScriptRoot "CardTemplate.template.pptx"),
    [string]$InputJson = (Join-Path $PSScriptRoot "slides.example.json"),
    [string]$LayoutsJson = (Join-Path $PSScriptRoot "layouts.json"),
    [string]$OutputPath = (Join-Path (Get-Location) "Presentation.pptx"),
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
                                $shapes += [pscustomobject]@{ Shape = $groupItem; Text = $text; Top = [double]$groupItem.Top; Left = [double]$groupItem.Left }
                            }
                        }
                    } catch { }
                }
            }
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $text = $shape.TextFrame.TextRange.Text
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    $shapes += [pscustomobject]@{ Shape = $shape; Text = $text; Top = [double]$shape.Top; Left = [double]$shape.Left }
                }
            }
        } catch { }
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
            if (-not [string]::IsNullOrWhiteSpace([string]$item)) { $lines += [string]$item }
        }
        return ($lines -join "`r")
    }
    return [string]$Value
}

function Get-DefaultAuthor {
    $registryPaths = @("HKCU:\Software\Microsoft\Office\Common\UserInfo", "HKCU:\Software\Microsoft\Office\16.0\Common\UserInfo")
    foreach ($path in $registryPaths) {
        try {
            if (Test-Path -LiteralPath $path) {
                $userName = (Get-ItemProperty -LiteralPath $path).UserName
                if (-not [string]::IsNullOrWhiteSpace($userName)) { return [string]$userName }
            }
        } catch { }
    }
    try { if (-not [string]::IsNullOrWhiteSpace($env:USERNAME)) { return [string]$env:USERNAME } } catch { }
    return ""
}

function Test-InvalidAuthor {
    param([object]$Author)

    if ([string]::IsNullOrWhiteSpace([string]$Author)) { return $true }
    $value = ([string]$Author).Trim()
    return @("OpenCode", "opencode", "Assistant", "AI", "IA") -contains $value
}

function Get-OpenPowerPointPresentationPaths {
    $paths = @{}
    try {
        $activePowerPoint = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
        for ($i = 1; $i -le $activePowerPoint.Presentations.Count; $i++) {
            try {
                $fullName = [string]$activePowerPoint.Presentations.Item($i).FullName
                if (-not [string]::IsNullOrWhiteSpace($fullName)) { $paths[$fullName.ToLowerInvariant()] = $true }
            } catch { }
        }
    } catch { }
    return $paths
}

function Close-PresentationIfOpenedByScript {
    param($Presentation, [hashtable]$ExistingPresentationPaths)

    if (-not $Presentation) { return }
    try {
        $fullName = [string]$Presentation.FullName
        if ([string]::IsNullOrWhiteSpace($fullName) -or -not $ExistingPresentationPaths.ContainsKey($fullName.ToLowerInvariant())) {
            $Presentation.Close()
        }
    } catch { }
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
            } catch { }
        }
    }
    return $warnings
}

function Build-SlideValues {
    param([hashtable]$SlideSpec, [hashtable]$CoverData, [string]$DeckTitle)

    $values = @{}

    foreach ($key in $SlideSpec.placeholders.Keys) {
        $values[$key] = $SlideSpec.placeholders[$key]
    }

    if (-not $values.ContainsKey("{{FOOTER_TITLE}}")) {
        $values["{{FOOTER_TITLE}}"] = $DeckTitle
    }

    if ($SlideSpec.ContainsKey("cards") -and $SlideSpec.cards) {
        $cardIndex = 1
        foreach ($card in $SlideSpec.cards) {
            $titleKey = "{{CARD${cardIndex}_TITLE}}"
            $bodyKey = "{{CARD${cardIndex}_BODY}}"
            if ($card.ContainsKey("title")) { $values[$titleKey] = $card.title }
            if ($card.ContainsKey("body")) { $values[$bodyKey] = $card.body }
            $cardIndex++
        }
    }

    return $values
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
$powerPointProcessCountBefore = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue).Count
$openPresentationPathsBefore = Get-OpenPowerPointPresentationPaths

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    try { $powerPoint.Visible = if ($Visible) { -1 } else { 0 } } catch { $powerPoint.Visible = -1 }
    try { $powerPoint.DisplayAlerts = 1 } catch { }
    $useTempPath = $null
    if ([System.IO.Path]::GetPathRoot($OutputPath) -like "*OneDrive*" -or $OutputPath -match "OneDrive") {
        $useTempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "ppt-v2-gen-" + [System.IO.Path]::GetRandomFileName() + ".pptx")
        Copy-Item -LiteralPath $OutputPath -Destination $useTempPath -Force
        $presentation = $powerPoint.Presentations.Open($useTempPath, $false, $false, $Visible.IsPresent)
    } else {
        $presentation = $powerPoint.Presentations.Open($OutputPath, $false, $false, $Visible.IsPresent)
    }

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
    } catch { }

    $replaced = Replace-Placeholders $cover $coverValues
    $generatedSlides += $cover.SlideIndex

    foreach ($slideSpec in $inputData.slides) {
        $layoutName = [string]$slideSpec.layout
        if (-not $layoutsByName.ContainsKey($layoutName)) { throw "Layout inconnu: $layoutName" }

        $layout = $layoutsByName[$layoutName]
        $sourceSlide = $presentation.Slides.Item([int]$layout.sourceSlide)
        $duplicated = $sourceSlide.Duplicate()
        $newSlide = $duplicated.Item(1)
        $newSlide.MoveTo($presentation.Slides.Count)
        $generatedSlides += $newSlide.SlideIndex

        $values = Build-SlideValues $slideSpec $inputData.cover $inputData.cover.deckTitle

        $replaced = Replace-Placeholders $newSlide $values
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

    if ($useTempPath) {
        Copy-Item -LiteralPath $useTempPath -Destination $OutputPath -Force
        Remove-Item -LiteralPath $useTempPath -Force -ErrorAction SilentlyContinue
    }

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
    Close-PresentationIfOpenedByScript $presentation $openPresentationPathsBefore
    if ($powerPoint -and $powerPointProcessCountBefore -eq 0) { $powerPoint.Quit() }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
