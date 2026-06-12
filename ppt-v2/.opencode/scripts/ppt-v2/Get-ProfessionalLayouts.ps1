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

function ConvertTo-Utf8NoBom($path, $content) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($content)
    [System.IO.File]::WriteAllBytes($path, $bytes)
}

function Get-TextShapes($Slide) {
    $shapes = @()
    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        $shape = $Slide.Shapes.Item($i)
        try {
            if ($shape.Type -eq 6) {
                for ($j = 1; $j -le $shape.GroupItems.Count; $j++) {
                    $gi = $shape.GroupItems.Item($j)
                    try { if ($gi.HasTextFrame -and $gi.TextFrame.HasText) { $shapes += [pscustomobject]@{Shape=$gi;Text=$gi.TextFrame.TextRange.Text} } } catch {}
                }
            }
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $shapes += [pscustomobject]@{Shape=$shape;Text=$shape.TextFrame.TextRange.Text}
            }
        } catch {}
    }
    return $shapes
}

function Get-Placeholders($Slide) {
    $phs = @{}
    foreach ($entry in Get-TextShapes $Slide) {
        $matches = [regex]::Matches($entry.Text, '\{\{[^}]+\}\}')
        foreach ($m in $matches) { $phs[$m.Value] = $true }
    }
    return $phs.Keys
}

function Get-MaxCardNumber($placeholders) {
    $max = 0
    foreach ($ph in $placeholders) {
        if ($ph -match 'CARD(\d+)_') { $n = [int]$Matches[1]; if ($n -gt $max) { $max = $n } }
    }
    return $max
}

function Get-MaxKpiNumber($placeholders) {
    $max = 0
    foreach ($ph in $placeholders) {
        if ($ph -match 'KPI(\d+)_') { $n = [int]$Matches[1]; if ($n -gt $max) { $max = $n } }
    }
    return $max
}

function Get-MaxStepNumber($placeholders) {
    $max = 0
    foreach ($ph in $placeholders) {
        if ($ph -match 'STEP(\d+)(?:_|$)') { $n = [int]$Matches[1]; if ($n -gt $max) { $max = $n } }
    }
    return $max
}

function Get-MaxLessonNumber($placeholders) {
    $max = 0
    foreach ($ph in $placeholders) {
        if ($ph -match 'LESSON_(\d+)') { $n = [int]$Matches[1]; if ($n -gt $max) { $max = $n } }
    }
    return $max
}

function Detect-Layout($Slide) {
    $placeholders = Get-Placeholders $Slide
    $phSet = @{}
    foreach ($p in $placeholders) { $phSet[$p] = $true }
    $required = @{
        "title_slide" = @("{{DECK_TITLE}}")
        "context_kpi_slide" = @("{{KPI1_VALUE}}")
        "problem_slide" = @("{{PROBLEM}}", "{{IMPACTS}}")
        "question_answer_slide" = @("{{QUESTION}}", "{{ANSWER}}")
        "three_cards_slide" = @("{{CARD1_TITLE}}", "{{CARD2_TITLE}}", "{{CARD3_TITLE}}")
        "four_cards_slide" = @("{{CARD1_TITLE}}", "{{CARD2_TITLE}}", "{{CARD3_TITLE}}", "{{CARD4_TITLE}}")
        "process_slide" = @("{{STEP1}}")
        "role_focus_slide" = @("{{ROLE_NAME}}")
        "lessons_slide" = @("{{LESSON_1}}")
        "adoption_loop_slide" = @("{{PHASE1}}")
        "conclusion_slide" = @("{{MAIN_MESSAGE}}")
    }
    $bestMatch = $null; $bestScore = 0
    foreach ($layout in $required.Keys) {
        $score = 0; $total = $required[$layout].Count
        foreach ($req in $required[$layout]) { if ($phSet.ContainsKey($req)) { $score++ } }
        if ($score -eq $total -and $score -gt $bestScore) { $bestMatch = $layout; $bestScore = $score }
    }
    if (-not $bestMatch) { $bestMatch = "unknown" }
    $layoutInfo = @{ name = $bestMatch; placeholders = @($placeholders) }
    if ($bestMatch -eq "context_kpi_slide") { $layoutInfo.kpiCount = Get-MaxKpiNumber $placeholders }
    if ($bestMatch -eq "three_cards_slide" -or $bestMatch -eq "four_cards_slide") { $layoutInfo.cardCount = Get-MaxCardNumber $placeholders }
    if ($bestMatch -eq "process_slide") { $layoutInfo.stepCount = Get-MaxStepNumber $placeholders }
    if ($bestMatch -eq "lessons_slide") { $layoutInfo.lessonCount = Get-MaxLessonNumber $placeholders }
    return $layoutInfo
}

$powerPoint = $null
$presentation = $null
$powerPointProcessCountBefore = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue).Count

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    try { $powerPoint.Visible = 0 } catch { }
    try { $powerPoint.DisplayAlerts = 1 } catch { }
    $presentation = $powerPoint.Presentations.Open($TemplatePath, $false, $false, $false)
    $layouts = @()
    for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
        $slide = $presentation.Slides.Item($i)
        $detected = Detect-Layout $slide
        $detected.sourceSlide = $i
        $layouts += $detected
    }
    $catalog = @{
        version = 2
        template = $TemplatePath
        slideSize = "widescreen"
        colors = @{ primary = "#FF7900"; dark = "#333333"; medium = "#666666"; light = "#F5F2F0"; conclusion = "#FF7900" }
        rules = @{
            preserveTemplateFormatting = $true
            doNotChangeFonts = $true
            doNotChangeColors = $true
            coverHasPageNumber = $false
            contentPageNumberFormat = "{0} / {1}"
            maxTitleCharacters = 100
            maxSubtitleCharacters = 150
            maxCardCharacters = 120
            recommendedMaxCardsPerSlide = 5
            requireConclusion = $true
            maxWordsPerCard = 12
        }
        layouts = $layouts
    }
    $json = $catalog | ConvertTo-Json -Depth 10
    ConvertTo-Utf8NoBom $OutputJson $json
    Write-Output "Catalogue detecte: $OutputJson"
    Write-Output "Layouts: $(($layouts | ForEach-Object { "$($_.name):slide$($_.sourceSlide)" }) -join ', ')"
} finally {
    if ($presentation) {
        try {
            $fullName = [string]$presentation.FullName
            $alreadyOpen = $false
            if ($powerPointProcessCountBefore -gt 0) {
                try {
                    $activePp = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
                    for ($j = 1; $j -le $activePp.Presentations.Count; $j++) {
                        if ([string]$activePp.Presentations.Item($j).FullName -eq $fullName) { $alreadyOpen = $true; break }
                    }
                } catch {}
            }
            if (-not $alreadyOpen -or $powerPointProcessCountBefore -eq 0) { $presentation.Close() }
        } catch {}
    }
    if ($powerPoint -and $powerPointProcessCountBefore -eq 0) { $powerPoint.Quit() }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
