param(
    [string]$TemplatePath = (Join-Path $PSScriptRoot "ProfessionalTemplate.template.pptx"),
    [string]$InputJson = (Join-Path $PSScriptRoot "slides.example.json"),
    [string]$LayoutsJson = (Join-Path $PSScriptRoot "layouts.json"),
    [string]$IconsDir = (Join-Path $PSScriptRoot "..\..\..\..\icons"),
    [string]$OutputPath = (Join-Path (Get-Location) "Presentation.pptx"),
    [switch]$Visible
)

$ErrorActionPreference = "Stop"

$TemplatePath = (Resolve-Path -LiteralPath $TemplatePath).Path
$InputJson = (Resolve-Path -LiteralPath $InputJson).Path
$LayoutsJson = (Resolve-Path -LiteralPath $LayoutsJson).Path
$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$IconsDir = [System.IO.Path]::GetFullPath($IconsDir)

$conclusionTemplates = @(
    "La valeur se cree quand {0}"
    "L'enjeu n'est pas seulement {0}, mais {1}"
    "Le succes depend de {0}"
    "L'adoption devient possible lorsque {0}"
    "Le pilotage de {0} est la cle de la reussite"
    "{0} transforme la donnee en decision"
    "Sans {0}, pas de transformation durable"
    "L'avenir repose sur {0}"
)

$slideTypeConclus = @{
    "contexte" = "la comprehension du contexte"
    "problematique" = "la resolution du probleme"
    "question" = "la reponse apportee"
    "reponse" = "la mise en oeuvre"
    "processus" = "la standardisation du processus"
    "roles" = "la mobilisation des equipes"
    "enseignements" = "le partage des enseignements"
    "chiffres" = "la mesure de la performance"
    "bilan" = "l'amelioration continue"
    "conclusion" = "la vision partagee"
}

$slideTypeMapping = @{
    "contexte" = "context_kpi_slide"
    "problematique" = "problem_slide"
    "question" = "question_answer_slide"
    "reponse" = "question_answer_slide"
    "processus" = "process_slide"
    "roles" = "role_focus_slide"
    "enseignements" = "lessons_slide"
    "chiffres" = "context_kpi_slide"
    "bilan" = "conclusion_slide"
    "conclusion" = "conclusion_slide"
}

function ConvertTo-Hashtable($Object) {
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IEnumerable] -and $Object -isnot [string] -and $Object -isnot [System.Management.Automation.PSCustomObject]) {
        $a = @(); foreach ($item in $Object) { $a += ,(ConvertTo-Hashtable $item) }; return $a
    }
    if ($Object -is [System.Management.Automation.PSCustomObject]) {
        $h = @{}; foreach ($prop in $Object.PSObject.Properties) { $h[$prop.Name] = ConvertTo-Hashtable $prop.Value }; return $h
    }
    return $Object
}

function Get-TextShapes($Slide) {
    $shapes = @()
    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        $shape = $Slide.Shapes.Item($i)
        try {
            if ($shape.Type -eq 6) {
                for ($j = 1; $j -le $shape.GroupItems.Count; $j++) {
                    $gi = $shape.GroupItems.Item($j)
                    try { if ($gi.HasTextFrame -and $gi.TextFrame.HasText) { $shapes += [pscustomobject]@{Shape=$gi;Text=$gi.TextFrame.TextRange.Text;Top=[double]$gi.Top;Left=[double]$gi.Left} } } catch {}
                }
            }
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $shapes += [pscustomobject]@{Shape=$shape;Text=$shape.TextFrame.TextRange.Text;Top=[double]$shape.Top;Left=[double]$shape.Left}
            }
        } catch {}
    }
    return $shapes | Sort-Object Top, Left
}

function Replace-Placeholders($Slide, $Values) {
    $count = 0
    foreach ($entry in Get-TextShapes $Slide) {
        $newText = $entry.Text
        foreach ($key in $Values.Keys) {
            if ($newText.Contains($key)) {
                $val = $Values[$key]
                if ($val -is [array]) { $val = $val -join "`r" }
                $newText = $newText.Replace($key, [string]$val)
            }
        }
        if ($newText -ne $entry.Text) {
            $entry.Shape.TextFrame.TextRange.Text = $newText
            $count++
        }
    }
    return $count
}

function Clear-AllPlaceholders($Slide) {
    foreach ($entry in Get-TextShapes $Slide) {
        if ($entry.Text -match '\{\{[^}]+\}\}') {
            $newText = $entry.Text -replace '\{\{[^}]+\}\}', ''
            if ($newText -ne $entry.Text) { $entry.Shape.TextFrame.TextRange.Text = $newText.Trim() }
        }
    }
}

function Get-RemainingPlaceholders($Presentation) {
    $rem = @()
    for ($s = 1; $s -le $Presentation.Slides.Count; $s++) {
        foreach ($entry in Get-TextShapes $Presentation.Slides.Item($s)) {
            if ($entry.Text -match '\{\{[^}]+\}\}') { $rem += "Slide ${s}: $($entry.Text)" }
        }
    }
    return $rem
}

function Test-TextOverflow($Presentation) {
    $issues = @()
    for ($s = 1; $s -le $Presentation.Slides.Count; $s++) {
        foreach ($entry in Get-TextShapes $Presentation.Slides.Item($s)) {
            try {
                $frame = $entry.Shape.TextFrame
                if ($frame.TextRange.Count -gt 0 -and $frame.TextRange.BoundHeight -gt ($entry.Shape.Height * 1.2)) {
                    $txt = $entry.Text -replace "`r`n", " "
                    if ($txt.Length -gt 60) { $txt = $txt.Substring(0, 57) + "..." }
                    $issues += "Slide ${s}: texte depasse - ""$txt"""
                }
            } catch {}
        }
    }
    return $issues
}

function Reduce-Text($text, $maxWords = 12, $maxLines = 4) {
    if ([string]::IsNullOrWhiteSpace($text)) { return $text }
    $lines = $text -split "`r`n|`n"
    if ($lines.Count -gt $maxLines) { $lines = $lines[0..($maxLines-1)] }
    $result = @()
    foreach ($line in $lines) {
        $words = $line -split '\s+'
        if ($words.Count -gt $maxWords) { $words = $words[0..($maxWords-1)] }
        $result += $words -join ' '
    }
    return $result -join "`r"
}

function Select-Layout($slideSpec, $layoutsByName) {
    if ($slideSpec.ContainsKey("layout") -and $layoutsByName.ContainsKey($slideSpec.layout)) {
        return $layoutsByName[$slideSpec.layout]
    }
    if ($slideSpec.ContainsKey("type") -and $slideTypeMapping.ContainsKey($slideSpec.type)) {
        $autoLayout = $slideTypeMapping[$slideSpec.type]
        if ($layoutsByName.ContainsKey($autoLayout)) { return $layoutsByName[$autoLayout] }
    }
    return $layoutsByName["three_cards_slide"]
}

function Build-Conclusion($slideSpec, $inputData) {
    if ($slideSpec.ContainsKey("conclusion") -and -not [string]::IsNullOrWhiteSpace($slideSpec.conclusion)) {
        return $slideSpec.conclusion
    }
    $topic = ""
    if ($slideSpec.ContainsKey("type") -and $slideTypeConclus.ContainsKey($slideSpec.type)) {
        $topic = $slideTypeConclus[$slideSpec.type]
    } elseif ($slideSpec.ContainsKey("placeholders") -and $slideSpec.placeholders.ContainsKey("{{TITLE}}")) {
        $topic = $slideSpec.placeholders["{{TITLE}}"]
    }
    if ([string]::IsNullOrWhiteSpace($topic)) { $topic = "cette demarche" }
    $template = $conclusionTemplates | Get-Random
    return ($template -f $topic, "l'engagement collectif")
}

function Insert-Icon($Slide, $placeholderText, $iconName, $left, $top, $size) {
    $iconPath = Join-Path $IconsDir "$iconName.svg"
    if (-not (Test-Path -LiteralPath $iconPath)) { return }
    try {
        $Slide.Shapes.AddPicture($iconPath, 0, -1, $left, $top, $size, $size) | Out-Null
    } catch {}
}

function Get-OpenPowerPointPresentationPaths {
    $paths = @{}
    try {
        $activePp = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
        for ($i = 1; $i -le $activePp.Presentations.Count; $i++) {
            try { $fn = [string]$activePp.Presentations.Item($i).FullName; if ($fn) { $paths[$fn.ToLowerInvariant()] = $true } } catch {}
        }
    } catch {}
    return $paths
}

function Close-PresentationIfOpenedByScript($Presentation, $ExistingPaths) {
    if (-not $Presentation) { return }
    try {
        $fn = [string]$Presentation.FullName
        if (-not $fn -or -not $ExistingPaths.ContainsKey($fn.ToLowerInvariant())) { $Presentation.Close() }
    } catch {}
}

$inputData = ConvertTo-Hashtable ((Get-Content -LiteralPath $InputJson -Raw -Encoding UTF8) | ConvertFrom-Json)
$layoutData = ConvertTo-Hashtable ((Get-Content -LiteralPath $LayoutsJson -Raw -Encoding UTF8) | ConvertFrom-Json)
$layoutsByName = @{}
foreach ($layout in $layoutData.layouts) { $layoutsByName[$layout.name] = $layout }

$useTempPath = $null
Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath -Force

if ([System.IO.Path]::GetPathRoot($OutputPath) -like "*OneDrive*" -or $OutputPath -match "OneDrive") {
    $useTempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "ppt-pro-gen-" + [System.IO.Path]::GetRandomFileName() + ".pptx")
    Copy-Item -LiteralPath $OutputPath -Destination $useTempPath -Force
}

$powerPoint = $null
$presentation = $null
$ppProcBefore = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue).Count
$existingPpPaths = Get-OpenPowerPointPresentationPaths

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    try { $powerPoint.Visible = if ($Visible) { -1 } else { 0 } } catch { $powerPoint.Visible = -1 }
    try { $powerPoint.DisplayAlerts = 1 } catch { }
    $openPath = if ($useTempPath) { $useTempPath } else { $OutputPath }
    $presentation = $powerPoint.Presentations.Open($openPath, $false, $false, $Visible.IsPresent)

    $generatedSlides = @()

    $coverLayout = $layoutsByName["title_slide"]
    $cover = $presentation.Slides.Item([int]$coverLayout.sourceSlide)
    $coverValues = @{
        "{{DECK_TITLE}}" = $inputData.cover.deckTitle
        "{{DECK_SUBTITLE}}" = $inputData.cover.deckSubtitle
        "{{DATE}}" = $inputData.cover.date
        "{{AUTHORS}}" = $inputData.cover.authors
        "{{FOOTER_TITLE}}" = $inputData.cover.deckTitle
        "{{CONCLUSION}}" = ""
        "{{PAGE_NUMBER}}" = ""
    }
    $null = Replace-Placeholders $cover $coverValues
    Clear-AllPlaceholders $cover
    $generatedSlides += $cover.SlideIndex

    foreach ($slideSpec in $inputData.slides) {
        $layoutDef = Select-Layout $slideSpec $layoutsByName
        $sourceSlide = $presentation.Slides.Item([int]$layoutDef.sourceSlide)
        $dup = $sourceSlide.Duplicate()
        $newSlide = $dup.Item(1)
        $newSlide.MoveTo($presentation.Slides.Count)

        $values = @{}
        $values["{{TITLE}}"] = if ($slideSpec.placeholders."{{TITLE}}") { $slideSpec.placeholders."{{TITLE}}" } else { "" }
        $values["{{SUBTITLE}}"] = if ($slideSpec.placeholders."{{SUBTITLE}}") { $slideSpec.placeholders."{{SUBTITLE}}" } else { "" }
        $values["{{FOOTER_TITLE}}"] = $inputData.cover.deckTitle

        if ($slideSpec.ContainsKey("placeholders")) {
            foreach ($key in $slideSpec.placeholders.Keys) {
                $values[$key] = $slideSpec.placeholders[$key]
            }
        }

        $conclusion = Build-Conclusion $slideSpec $inputData
        $values["{{CONCLUSION}}"] = $conclusion

        if ($slideSpec.ContainsKey("cards")) {
            $ci = 1
            foreach ($card in $slideSpec.cards) {
                $tk = "{{CARD${ci}_TITLE}}"; $bk = "{{CARD${ci}_BODY}}"
                if ($card.ContainsKey("title")) { $values[$tk] = Reduce-Text $card.title 8 }
                if ($card.ContainsKey("body")) {
                    $bodyText = if ($card.body -is [array]) { $card.body -join "`r" } else { $card.body }
                    $values[$bk] = Reduce-Text $bodyText 12 4
                }
                $ci++
            }
        }

        if ($slideSpec.ContainsKey("kpis")) {
            $ki = 1
            foreach ($kpi in $slideSpec.kpis) {
                $values["{{KPI${ki}_VALUE}}"] = $kpi.value
                $values["{{KPI${ki}_LABEL}}"] = $kpi.label
                $ki++
            }
        }

        if ($slideSpec.ContainsKey("steps")) {
            $si = 1
            foreach ($step in $slideSpec.steps) {
                $values["{{STEP${si}}}"] = $step.title
                $values["{{STEP${si}_DESC}}"] = $step.description
                $si++
            }
        }

        if ($slideSpec.ContainsKey("lessons")) {
            $li = 1
            foreach ($lesson in $slideSpec.lessons) {
                $values["{{LESSON_${li}}}"] = $lesson
                $li++
            }
        }

        if ($slideSpec.ContainsKey("phases")) {
            $pi = 1
            foreach ($phase in $slideSpec.phases) {
                $values["{{PHASE${pi}}}"] = $phase.title
                $values["{{PHASE${pi}_DESC}}"] = $phase.description
                $pi++
            }
        }

        if ($slideSpec.ContainsKey("responsibilities")) {
            $ri = 1
            foreach ($resp in $slideSpec.responsibilities) {
                $values["{{RESP_${ri}}}"] = $resp
                $ri++
            }
        }

        if ($slideSpec.ContainsKey("takeaways")) {
            $ti = 1
            foreach ($tw in $slideSpec.takeaways) {
                $values["{{TAKEAWAY_${ti}}}"] = $tw
                $ti++
            }
        }

        if ($slideSpec.ContainsKey("answer")) { $values["{{ANSWER}}"] = $slideSpec.answer }
        if ($slideSpec.ContainsKey("question")) { $values["{{QUESTION}}"] = $slideSpec.question }
        if ($slideSpec.ContainsKey("problem")) { $values["{{PROBLEM}}"] = $slideSpec.problem }
        if ($slideSpec.ContainsKey("impacts")) { $values["{{IMPACTS}}"] = $slideSpec.impacts }
        if ($slideSpec.ContainsKey("mainMessage")) { $values["{{MAIN_MESSAGE}}"] = $slideSpec.mainMessage }
        if ($slideSpec.ContainsKey("roleName")) { $values["{{ROLE_NAME}}"] = $slideSpec.roleName }
        if ($slideSpec.ContainsKey("roleTitle")) { $values["{{ROLE_TITLE}}"] = $slideSpec.roleTitle }
        if ($slideSpec.ContainsKey("points")) {
            for ($pi = 1; $pi -le 3; $pi++) {
                if ($pi -le $slideSpec.points.Count) { $values["{{POINT_${pi}}}"] = $slideSpec.points[$pi-1] }
            }
        }

        $null = Replace-Placeholders $newSlide $values
        Clear-AllPlaceholders $newSlide
        $generatedSlides += $newSlide.SlideIndex
    }

    for ($i = $presentation.Slides.Count; $i -ge 1; $i--) {
        if ($generatedSlides -notcontains $i) {
            $presentation.Slides.Item($i).Delete()
        }
    }

    $contentCount = $presentation.Slides.Count - 1
    for ($i = 2; $i -le $presentation.Slides.Count; $i++) {
        $pageVal = "{0} / {1}" -f ($i - 1), $contentCount
        $null = Replace-Placeholders $presentation.Slides.Item($i) @{ "{{PAGE_NUMBER}}" = $pageVal }
    }

    $remaining = Get-RemainingPlaceholders $presentation
    $overflow = Test-TextOverflow $presentation

    $presentation.Save()

    if ($useTempPath) {
        Copy-Item -LiteralPath $useTempPath -Destination $OutputPath -Force
        Remove-Item -LiteralPath $useTempPath -Force -ErrorAction SilentlyContinue
    }

    Write-Output "Generation terminee: $OutputPath"
    Write-Output "Slides: $($presentation.Slides.Count) (1 cover + $contentCount contenu)"
    if ($remaining.Count -gt 0) {
        Write-Output "Placeholders restants:"
        $remaining | ForEach-Object { Write-Output "- $_" }
    }
    if ($overflow.Count -gt 0) {
        Write-Output "Alertes rendu visuel:"
        $overflow | ForEach-Object { Write-Output "- $_" }
    }
} finally {
    Close-PresentationIfOpenedByScript $presentation $existingPpPaths
    if ($powerPoint -and $ppProcBefore -eq 0) { $powerPoint.Quit() }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
