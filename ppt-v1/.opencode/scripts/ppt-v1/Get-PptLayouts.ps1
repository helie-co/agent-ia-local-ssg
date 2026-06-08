param(
    [Parameter(Mandatory = $true)]
    [string]$TemplatePath,
    [string]$OutputJson = (Join-Path (Join-Path (Get-Location) "ppt-v1") "layouts.detected.json")
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
        } catch {
        }
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

function Test-HasTable {
    param($Slide)

    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        $shape = $Slide.Shapes.Item($i)
        try {
            if ($shape.Type -eq 6) {
                for ($j = 1; $j -le $shape.GroupItems.Count; $j++) {
                    try { if ($shape.GroupItems.Item($j).HasTable) { return $true } } catch { }
                }
            }
            if ($shape.HasTable) { return $true }
        } catch {
        }
    }
    return $false
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
        [string[]]$OptionalPlaceholders = @(),
        [object]$TableRules = $null,
        [string[]]$SelectionCriteria = @()
    )

    $layout = [ordered]@{
        name = $Name
        sourceSlide = $SourceSlide
        description = $Description
        requiredPlaceholders = @($RequiredPlaceholders)
    }

    if ($OptionalPlaceholders.Count -gt 0) { $layout.optionalPlaceholders = @($OptionalPlaceholders) }
    if ($null -ne $TableRules) { $layout.tableRules = $TableRules }
    if ($SelectionCriteria.Count -gt 0) { $layout.selectionCriteria = @($SelectionCriteria) }

    return $layout
}

$powerPoint = $null
$presentation = $null

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
        $hasTable = Test-HasTable $slide

        $layout = $null

        if (Test-ContainsAll $placeholders @("{{DECK_TITLE}}", "{{DATE}}")) {
            $required = @("{{DECK_TITLE}}", "{{DATE}}")
            if ($placeholders -contains "{{AUTHORS}}") { $required += "{{AUTHORS}}" }
            $optional = @($placeholders | Where-Object { $required -notcontains $_ })
            $layout = New-Layout "cover" $slideIndex "Page de garde detectee automatiquement." $required $optional
        } elseif (Test-ContainsAll $placeholders @("{{SECTION_TITLE}}")) {
            $required = @("{{SECTION_TITLE}}")
            if ($placeholders -contains "{{SECTION_SUBTITLE}}") { $required += "{{SECTION_SUBTITLE}}" }
            if ($placeholders -contains "{{SECTION_NUMBER}}") { $required += "{{SECTION_NUMBER}}" }
            $optional = @($placeholders | Where-Object { $required -notcontains $_ })
            $layout = New-Layout "section_header" $slideIndex "Slide de chapitre detectee automatiquement." $required $optional
        } elseif (Test-ContainsAll $placeholders @("{{LEFT_BODY}}", "{{RIGHT_BODY}}")) {
            $required = @("{{TITLE}}", "{{SUBTITLE}}", "{{LEFT_BODY}}", "{{RIGHT_BODY}}") | Where-Object { $placeholders -contains $_ }
            foreach ($candidate in @("{{LEFT_TITLE}}", "{{RIGHT_TITLE}}")) {
                if ($placeholders -contains $candidate) { $required += $candidate }
            }
            $optional = @($placeholders | Where-Object { $required -notcontains $_ })
            $layout = New-Layout "two_columns_bullets" $slideIndex "Slide deux colonnes detectee automatiquement." $required $optional $null @(
                "Comparer deux categories equilibrees",
                "Avantages / inconvenients",
                "Pour / contre",
                "Forces / limites",
                "Opportunites / risques"
            )
        } elseif ($hasTable -and ($placeholders -contains "{{TITLE}}")) {
            $required = @("{{TITLE}}")
            if ($placeholders -contains "{{SUBTITLE}}") { $required += "{{SUBTITLE}}" }
            $optional = @($placeholders | Where-Object { $required -notcontains $_ })
            $layout = New-Layout "table_dynamic" $slideIndex "Slide tableau detectee automatiquement." $required $optional ([ordered]@{
                minColumns = 2
            }) @(
                "Tableau, matrice ou grille",
                "Comparaison avec colonnes homogenes"
            )
        } elseif (($placeholders -contains "{{BODY}}") -and ($placeholders -contains "{{TITLE}}")) {
            $required = @("{{TITLE}}", "{{BODY}}")
            if ($placeholders -contains "{{SUBTITLE}}") { $required += "{{SUBTITLE}}" }
            $optional = @($placeholders | Where-Object { $required -notcontains $_ })
            $layout = New-Layout "classic_bullets" $slideIndex "Slide bullets detectee automatiquement." $required $optional
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
    if ($presentation) { $presentation.Close() }
    if ($powerPoint) { $powerPoint.Quit() }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
