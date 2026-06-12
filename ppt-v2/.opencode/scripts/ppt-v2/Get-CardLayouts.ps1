param(
    [Parameter(Mandatory = $true)]
    [string]$TemplatePath,
    [string]$OutputJson = (Join-Path $PSScriptRoot "layouts.json")
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $TemplatePath)) { throw "Template introuvable: $TemplatePath" }

$TemplatePath = (Resolve-Path -LiteralPath $TemplatePath).Path
$OutputJson = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputJson)
$outputParent = Split-Path -Parent $OutputJson
if ($outputParent -and -not (Test-Path -LiteralPath $outputParent)) {
    New-Item -ItemType Directory -Path $outputParent | Out-Null
}

function Get-TextShapes {
    param($ShapeCollection)

    $items = @()
    for ($i = 1; $i -le $ShapeCollection.Count; $i++) {
        $shape = $ShapeCollection.Item($i)
        try {
            if ($shape.Type -eq 6) {
                $items += @(Get-TextShapes $shape.GroupItems)
                continue
            }
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $items += [pscustomobject]@{
                    Text = [string]$shape.TextFrame.TextRange.Text
                    Top = [double]$shape.Top
                    Left = [double]$shape.Left
                }
            }
        } catch { }
    }
    return $items
}

function Get-Placeholders {
    param([string[]]$Texts)

    $set = [ordered]@{}
    foreach ($text in $Texts) {
        foreach ($match in [regex]::Matches([string]$text, "\{\{[^}]+\}\}")) {
            $set[$match.Value] = $true
        }
    }
    return @($set.Keys)
}

function Test-ContainsAll {
    param([string[]]$Values, [string[]]$Required)

    foreach ($item in $Required) {
        if ($Values -notcontains $item) { return $false }
    }
    return $true
}

function New-Layout {
    param(
        [string]$Name,
        [int]$SourceSlide,
        [string]$Description,
        [string[]]$RequiredPlaceholders,
        [string[]]$OptionalPlaceholders = @()
    )

    $layout = [ordered]@{
        name = $Name
        sourceSlide = $SourceSlide
        description = $Description
        requiredPlaceholders = @($RequiredPlaceholders)
    }

    if ($OptionalPlaceholders.Count -gt 0) { $layout.optionalPlaceholders = @($OptionalPlaceholders) }

    return $layout
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

function Get-CardCount {
    param([string[]]$Placeholders)

    $max = 0
    foreach ($ph in $Placeholders) {
        if ($ph -match '\{\{CARD(\d+)_') {
            $num = [int]$matches[1]
            if ($num -gt $max) { $max = $num }
        }
    }
    return $max
}

$powerPoint = $null
$presentation = $null
$powerPointProcessCountBefore = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue).Count
$openPresentationPathsBefore = Get-OpenPowerPointPresentationPaths

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    try { $powerPoint.Visible = 0 } catch { $powerPoint.Visible = -1 }
    try { $powerPoint.DisplayAlerts = 1 } catch { }
    $presentation = $powerPoint.Presentations.Open($TemplatePath, $true, $false, $false)

    $layouts = @()
    $seenLayouts = @{}

    for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
        $slide = $presentation.Slides.Item($slideIndex)
        $textShapes = @(Get-TextShapes $slide.Shapes)
        $texts = @($textShapes | Sort-Object Top, Left | ForEach-Object { $_.Text })
        $placeholders = @(Get-Placeholders $texts)
        $cardCount = Get-CardCount $placeholders

        $layout = $null

        if (Test-ContainsAll $placeholders @("{{DECK_TITLE}}", "{{DATE}}")) {
            $required = @("{{DECK_TITLE}}", "{{DATE}}")
            if ($placeholders -contains "{{AUTHORS}}") { $required += "{{AUTHORS}}" }
            $optional = @($placeholders | Where-Object { $required -notcontains $_ })
            $layout = New-Layout "cover" $slideIndex "Page de garde" $required $optional
        } elseif ($cardCount -ge 2 -and $placeholders -contains "{{CARD1_TITLE}}") {
            $required = @("{{TITLE}}")
            if ($placeholders -contains "{{SUBTITLE}}") { $required += "{{SUBTITLE}}" }
            for ($i = 1; $i -le $cardCount; $i++) {
                $required += "{{CARD${i}_TITLE}}"
                if ($placeholders -contains "{{CARD${i}_BODY}}") { $required += "{{CARD${i}_BODY}}" }
            }
            $optional = @($placeholders | Where-Object { $required -notcontains $_ })
            $layout = New-Layout "card_${cardCount}" $slideIndex "Slide avec $cardCount cartes" $required $optional
        } elseif (($placeholders -contains "{{BODY}}") -and ($placeholders -contains "{{TITLE}}")) {
            $required = @("{{TITLE}}", "{{BODY}}")
            if ($placeholders -contains "{{SUBTITLE}}") { $required += "{{SUBTITLE}}" }
            $optional = @($placeholders | Where-Object { $required -notcontains $_ })
            $layout = New-Layout "classic_bullets" $slideIndex "Slide bullets" $required $optional
        }

        if ($layout -and -not $seenLayouts.ContainsKey($layout.name)) {
            $seenLayouts[$layout.name] = $true
            $layouts += $layout
        }
    }

    if (-not ($layouts | Where-Object { $_.name -eq "cover" })) { throw "Aucun layout cover detecte dans le template." }
    if (@($layouts | Where-Object { $_.name -ne "cover" }).Count -eq 0) { throw "Aucun layout de contenu detecte dans le template." }

    $result = [ordered]@{
        version = 1
        template = $TemplatePath
        rules = [ordered]@{
            preserveTemplateFormatting = $true
            doNotChangeFonts = $true
            doNotChangeColors = $true
            coverHasPageNumber = $false
            contentPageNumberFormat = "{0} / {1}"
            maxTitleCharacters = 90
            maxSubtitleCharacters = 130
            maxBulletCharacters = 120
            recommendedMaxBulletsPerSlide = 5
            allowBodyBullets = $true
        }
        layouts = @($layouts)
    }

    $json = $result | ConvertTo-Json -Depth 20
    [System.IO.File]::WriteAllText($OutputJson, $json, [System.Text.UTF8Encoding]::new($false))
    "Catalogue detecte: $OutputJson"
    "Layouts: " + (@($layouts | ForEach-Object { "$($_.name):slide$($_.sourceSlide)" }) -join ", ")
} finally {
    Close-PresentationIfOpenedByScript $presentation $openPresentationPathsBefore
    if ($powerPoint -and $powerPointProcessCountBefore -eq 0) { $powerPoint.Quit() }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
