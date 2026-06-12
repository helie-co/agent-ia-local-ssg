param(
    [string]$OutputPath = (Join-Path $PSScriptRoot "ProfessionalTemplate.template.pptx"),
    [switch]$Visible
)

$ErrorActionPreference = "Stop"

$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$outputParent = Split-Path -Parent $OutputPath
if ($outputParent -and -not (Test-Path -LiteralPath $outputParent)) {
    New-Item -ItemType Directory -Path $outputParent | Out-Null
}

$slideWidth = 960
$slideHeight = 540
$margin = 36
$contentW = $slideWidth - 2 * $margin
$titleY = 24
$titleH = 40
$subtitleY = 68
$subtitleH = 28
$contentY = 108
$contentH = 340
$conclusionY = 460
$conclusionH = 40
$footerY = 510

$orange = 0x0079FF
$darkGray = 0x333333
$medGray = 0x666666
$lightBg = 0xF5F2F0
$cardBg = 0xFAFAFA
$borderLight = 0xE8E8E8
$conclusionBg = 0xFF7900
$white = 0xFFFFFF
$accentBg = 0xFFF3E8

function Add-TextBox {
    param($Slide, [double]$L, [double]$T, [double]$W, [double]$H, [string]$Text, [int]$Sz = 18, [bool]$Bold = $false, [int]$Color = 0x333333, [string]$Align = "left", [string]$FontName = "Calibri")
    $Slide.Shapes.AddTextbox(1, $L, $T, $W, $H) | Out-Null
    $s = $Slide.Shapes($Slide.Shapes.Count)
    $s.TextFrame.TextRange.Text = $Text
    $s.TextFrame.TextRange.Font.Name = $FontName
    $s.TextFrame.TextRange.Font.Size = $Sz
    $s.TextFrame.TextRange.Font.Bold = $Bold
    $s.TextFrame.TextRange.Font.Color.RGB = $Color
    if ($Align -eq "center") { $s.TextFrame.TextRange.ParagraphFormat.Alignment = 2 }
    elseif ($Align -eq "right") { $s.TextFrame.TextRange.ParagraphFormat.Alignment = 3 }
    else { $s.TextFrame.TextRange.ParagraphFormat.Alignment = 1 }
    $s.TextFrame.WordWrap = -1
    $s.TextFrame.AutoSize = 0
}

function Add-Shape {
    param($Slide, [int]$Type, [double]$L, [double]$T, [double]$W, [double]$H, [int]$FillColor = -1, [int]$LineColor = -1)
    $Slide.Shapes.AddShape($Type, $L, $T, $W, $H) | Out-Null
    $s = $Slide.Shapes($Slide.Shapes.Count)
    $s.Fill.Visible = -1
    if ($FillColor -ge 0) { $s.Fill.ForeColor.RGB = $FillColor } else { $s.Fill.Visible = 0 }
    $s.Line.Visible = -1
    if ($LineColor -ge 0) { $s.Line.ForeColor.RGB = $LineColor } else { $s.Line.Visible = 0 }
}

function Add-RoundedRect {
    param($Slide, [double]$L, [double]$T, [double]$W, [double]$H, [int]$FillColor = 0xFAFAFA, [int]$LineColor = 0xE8E8E8)
    $null = Add-Shape $Slide 5 $L $T $W $H $FillColor $LineColor
    $s = $Slide.Shapes($Slide.Shapes.Count)
    $s.Adjustments.Item(1) = 0.08
    return $s
}

function Add-SubtleShadow {
    param($Shape)
    try {
        $Shape.Shadow.Type = 1
        $Shape.Shadow.Visible = -1
        $Shape.Shadow.OffsetX = 1
        $Shape.Shadow.OffsetY = 2
        $Shape.Shadow.ForeColor.RGB = 0x000000
        $Shape.Shadow.Transparency = 0.75
        $Shape.Shadow.Size = 100
        $Shape.Shadow.Blur = 6
    } catch {}
}

function Add-ConclusionBanner {
    param($Slide, [double]$Y = $conclusionY)
    $Slide.Shapes.AddShape(1, $margin, $Y, $contentW, $conclusionH) | Out-Null
    $s = $Slide.Shapes($Slide.Shapes.Count)
    $s.Fill.Visible = -1
    $s.Fill.ForeColor.RGB = $conclusionBg
    $s.Line.Visible = 0
    $s.TextFrame.TextRange.Text = "{{CONCLUSION}}"
    $s.TextFrame.TextRange.Font.Name = "Calibri"
    $s.TextFrame.TextRange.Font.Size = 14
    $s.TextFrame.TextRange.Font.Bold = -1
    $s.TextFrame.TextRange.Font.Color.RGB = $white
    $s.TextFrame.TextRange.ParagraphFormat.Alignment = 1
    $s.TextFrame.MarginLeft = 14
    $s.TextFrame.MarginTop = 8
    $s.TextFrame.WordWrap = -1
}

function Add-Footer {
    param($Slide)
    $null = Add-TextBox $Slide $margin $footerY 400 14 "{{FOOTER_TITLE}}" 9 $false 0x999999
    $null = Add-TextBox $Slide ($slideWidth - $margin - 60) $footerY 60 14 "{{PAGE_NUMBER}}" 9 $false 0x999999 "right"
}

function Add-IconPlaceholder {
    param($Slide, [double]$L, [double]$T, [double]$Size = 32)
    $null = Add-TextBox $Slide ($L + 4) ($T + 4) ($Size - 8) ($Size - 8) "{{ICON}}" 7 $false 0xCCCCCC "center"
}

function Make-TitleSlide {
    param($Slide)
    $null = Add-Shape $Slide 1 $margin 0 ($contentW) 5 $orange -1
    $null = Add-TextBox $Slide $margin 60 $contentW 60 "{{DECK_TITLE}}" 44 $true $darkGray "center"
    $null = Add-TextBox $Slide $margin 125 $contentW 40 "{{DECK_SUBTITLE}}" 20 $false $medGray "center"
    $null = Add-Shape $Slide 1 ($slideWidth/2 - 40) 175 80 3 $orange -1
    $null = Add-TextBox $Slide $margin 200 $contentW 60 "{{DATE}}" 14 $false $medGray "center"
    $null = Add-TextBox $Slide $margin 225 $contentW 40 "{{AUTHORS}}" 12 $false $medGray "center"
    $null = Add-Footer $Slide
    $null = Add-TextBox $Slide $margin $conclusionY $contentW $conclusionH "{{CONCLUSION}}" 12 $false $white "center"
}

function Make-ContextKpiSlide {
    param($Slide)
    $null = Add-TextBox $Slide $margin $titleY $contentW $titleH "{{TITLE}}" 28 $true $darkGray
    $null = Add-TextBox $Slide $margin $subtitleY $contentW $subtitleH "{{SUBTITLE}}" 16 $false $medGray
    $kpiY = $contentY + 30
    $kpiH = 140
    $kpiW = [Math]::Floor(($contentW - 60) / 4)
    $gaps = 20
    $kpiW = [Math]::Floor(($contentW - $gaps * 3) / 4)
    for ($i = 1; $i -le 4; $i++) {
        $kpiX = $margin + ($i - 1) * ($kpiW + $gaps)
        $card = Add-RoundedRect $Slide $kpiX $kpiY $kpiW $kpiH $white $borderLight
    $null = Add-SubtleShadow $card
    $null = Add-Shape $Slide 1 ($kpiX + 4) ($kpiY + 4) ($kpiW - 8) 4 $orange -1
        $iconY = $kpiY + 14
        $iconSize = 28
    $null = Add-Shape $Slide 5 ($kpiX + $kpiW/2 - 14) $iconY $iconSize $iconSize 0xFFF3E8 -1
    $null = Add-TextBox $Slide ($kpiX + $kpiW/2 - 10) ($iconY + 2) 20 20 "{{KPI${i}_ICON}}" 9 $false 0x0079FF "center"
    $null = Add-TextBox $Slide $kpiX ($iconY + 34) $kpiW 36 "{{KPI${i}_VALUE}}" 30 $true $orange "center"
    $null = Add-TextBox $Slide ($kpiX + 4) ($iconY + 74) ($kpiW - 8) 36 "{{KPI${i}_LABEL}}" 11 $false $medGray "center"
    }
    $null = Add-ConclusionBanner $Slide
    $null = Add-Footer $Slide
}

function Make-ProblemSlide {
    param($Slide)
    $null = Add-TextBox $Slide $margin $titleY $contentW $titleH "{{TITLE}}" 28 $true $darkGray
    $leftW = [Math]::Floor($contentW / 2) - 12
    $rightW = $contentW - $leftW - 24
    $probBox = Add-RoundedRect $Slide $margin $contentY $leftW 220 $accentBg $borderLight
    $null = Add-TextBox $Slide ($margin + 12) ($contentY + 12) ($leftW - 24) 40 "Probleme" 14 $true $orange
    $null = Add-TextBox $Slide ($margin + 12) ($contentY + 52) ($leftW - 24) 150 "{{PROBLEM}}" 14 $false $darkGray
    $impactBox = Add-RoundedRect $Slide ($margin + $leftW + 24) $contentY $rightW 220 $lightBg $borderLight
    $null = Add-TextBox $Slide ($margin + $leftW + 36) ($contentY + 12) ($rightW - 24) 40 "Impacts" 14 $true $darkGray
    $null = Add-TextBox $Slide ($margin + $leftW + 36) ($contentY + 52) ($rightW - 24) 150 "{{IMPACTS}}" 13 $false $darkGray
    $null = Add-ConclusionBanner $Slide
    $null = Add-Footer $Slide
}

function Make-QuestionAnswerSlide {
    param($Slide)
    $null = Add-TextBox $Slide $margin $titleY $contentW $titleH "{{TITLE}}" 28 $true $darkGray
    $qBox = Add-RoundedRect $Slide $margin ($contentY - 20) $contentW 80 $lightBg $borderLight
    $null = Add-TextBox $Slide ($margin + 16) ($contentY - 12) ($contentW - 32) 20 "Question" 12 $true $medGray
    $null = Add-TextBox $Slide ($margin + 16) ($contentY + 12) ($contentW - 32) 36 "{{QUESTION}}" 18 $false $darkGray "center"
    $aBox = Add-RoundedRect $Slide $margin ($contentY + 80) $contentW 90 $white $orange
    $null = Add-TextBox $Slide ($margin + 16) ($contentY + 90) ($contentW - 32) 18 "Reponse cle" 12 $true $orange
    $null = Add-TextBox $Slide ($margin + 16) ($contentY + 110) ($contentW - 32) 44 "{{ANSWER}}" 24 $true $orange "center"
    $ptY = $contentY + 195
    for ($i = 1; $i -le 3; $i++) {
        $py = $ptY + ($i - 1) * 30
    $null = Add-Shape $Slide 9 $margin $py 8 8 $orange -1
    $null = Add-TextBox $Slide ($margin + 18) ($py - 4) ($contentW - 18) 24 "{{POINT_${i}}}" 14 $false $darkGray
    }
    $null = Add-ConclusionBanner $Slide
    $null = Add-Footer $Slide
}

function Make-CardsSlide {
    param($Slide, [int]$Count, [int]$Cols)
    $null = Add-TextBox $Slide $margin $titleY $contentW $titleH "{{TITLE}}" 28 $true $darkGray
    $null = Add-TextBox $Slide $margin $subtitleY $contentW $subtitleH "{{SUBTITLE}}" 16 $false $medGray
    $rows = [Math]::Ceiling([double]$Count / $Cols)
    $gapX = 20; $gapY = 20
    $cardW = [Math]::Floor(($contentW - ($Cols - 1) * $gapX) / $Cols)
    $cardH = [Math]::Floor((340 - ($rows - 1) * $gapY) / $rows)
    for ($i = 0; $i -lt $Count; $i++) {
        $col = $i % $Cols; $row = [Math]::Floor([double]$i / $Cols)
        $cx = $margin + $col * ($cardW + $gapX)
        $cy = $contentY + $row * ($cardH + $gapY)
        $num = $i + 1
        $card = Add-RoundedRect $Slide $cx $cy $cardW $cardH $white $borderLight
    $null = Add-SubtleShadow $card
    $null = Add-Shape $Slide 1 ($cx + 4) $cy ($cardW - 8) 4 $orange -1
    $null = Add-TextBox $Slide ($cx + 12) ($cy + 12) ($cardW - 24) 26 "{{CARD${num}_TITLE}}" 16 $true $darkGray
    $null = Add-TextBox $Slide ($cx + 12) ($cy + 44) ($cardW - 24) ($cardH - 56) "{{CARD${num}_BODY}}" 12 $false $darkGray
    }
    $null = Add-ConclusionBanner $Slide
    $null = Add-Footer $Slide
}

function Make-ProcessSlide {
    param($Slide, [int]$StepCount = 4)
    $null = Add-TextBox $Slide $margin $titleY $contentW $titleH "{{TITLE}}" 28 $true $darkGray
    $null = Add-TextBox $Slide $margin $subtitleY $contentW $subtitleH "{{SUBTITLE}}" 16 $false $medGray
    $stepY = $contentY + 60
    $circleD = 60
    $totalW = $contentW - 40
    $stepW = [Math]::Floor($totalW / $StepCount)
    $circleXBase = $margin + 20
    for ($i = 1; $i -le $StepCount; $i++) {
        $cx = $circleXBase + ($i - 1) * $stepW + [Math]::Floor(($stepW - $circleD) / 2)
    $null = Add-Shape $Slide 9 $cx $stepY $circleD $circleD $orange -1
    $null = Add-TextBox $Slide $cx ($stepY + 16) $circleD 28 "$i" 16 $true $white "center"
    $null = Add-TextBox $Slide ($cx - 20) ($stepY + $circleD + 8) ($circleD + 40) 24 "{{STEP${i}}}" 13 $true $darkGray "center"
    $null = Add-TextBox $Slide ($cx - 20) ($stepY + $circleD + 34) ($circleD + 40) 40 "{{STEP${i}_DESC}}" 11 $false $medGray "center"
        if ($i -lt $StepCount) {
            $arrowX = $cx + $circleD + 4; $arrowW = $stepW - $circleD - 8
            if ($arrowW -gt 12) {
    $null = Add-Shape $Slide 33 ($cx + $circleD + 4) ($stepY + 20) ($arrowW) 16 0xFFF3E8 -1
            }
        }
    }
    $null = Add-ConclusionBanner $Slide
    $null = Add-Footer $Slide
}

function Make-RoleFocusSlide {
    param($Slide)
    $null = Add-TextBox $Slide $margin $titleY $contentW $titleH "{{TITLE}}" 28 $true $darkGray
    $roleCenterX = $slideWidth / 2
    $roleY = $contentY
    $null = Add-Shape $Slide 5 ($roleCenterX - 40) $roleY 80 80 $accentBg -1
    $null = Add-TextBox $Slide ($roleCenterX - 36) ($roleY + 24) 72 30 "{{ROLE_ICON}}" 24 $false $orange "center"
    $null = Add-TextBox $Slide ($roleCenterX - 150) ($roleY + 90) 300 30 "{{ROLE_NAME}}" 22 $true $darkGray "center"
    $null = Add-TextBox $Slide ($roleCenterX - 150) ($roleY + 120) 300 24 "{{ROLE_TITLE}}" 14 $false $medGray "center"
    $respY = $roleY + 160; $respPerCol = 2; $colW = 300; $colGap = 60
    for ($i = 1; $i -le 4; $i++) {
        $col = [Math]::Floor(($i - 1) / $respPerCol)
        $row = ($i - 1) % $respPerCol
        $rx = $margin + 40 + $col * ($colW + $colGap)
        $ry = $respY + $row * 55
    $null = Add-Shape $Slide 9 $rx ($ry + 4) 12 12 $orange -1
    $null = Add-TextBox $Slide ($rx + 22) $ry ($colW - 22) 24 "{{RESP_${i}}}" 14 $false $darkGray
    }
    $null = Add-ConclusionBanner $Slide
    $null = Add-Footer $Slide
}

function Make-LessonsSlide {
    param($Slide, [int]$Count = 4)
    $null = Add-TextBox $Slide $margin $titleY $contentW $titleH "{{TITLE}}" 28 $true $darkGray
    $null = Add-TextBox $Slide $margin $subtitleY $contentW $subtitleH "{{SUBTITLE}}" 16 $false $medGray
    $lessonY = $contentY
    for ($i = 1; $i -le $Count; $i++) {
        $ly = $lessonY + ($i - 1) * 60
        $circleD2 = 36
    $null = Add-Shape $Slide 9 $margin $ly $circleD2 $circleD2 $orange -1
    $null = Add-TextBox $Slide $margin ($ly + 6) $circleD2 24 "$i" 14 $true $white "center"
    $null = Add-TextBox $Slide ($margin + 50) ($ly + 4) ($contentW - 50) 48 "{{LESSON_${i}}}" 16 $false $darkGray
    }
    $null = Add-ConclusionBanner $Slide
    $null = Add-Footer $Slide
}

function Make-AdoptionLoopSlide {
    param($Slide)
    $null = Add-TextBox $Slide $margin $titleY $contentW $titleH "{{TITLE}}" 28 $true $darkGray
    $cx = 420; $cy = 290; $radiusX = 200; $radiusY = 120
    $phases = @(@{x=$cx; y=$cy-$radiusY-40; a="top"}, @{x=$cx+$radiusX+30; y=$cy-30; a="right"}, @{x=$cx; y=$cy+$radiusY+30; a="bottom"}, @{x=$cx-$radiusX-30; y=$cy-30; a="left"})
    for ($i = 0; $i -lt 4; $i++) {
        $num = $i + 1; $p = $phases[$i]
        $pw = 160; $ph = 70; $px = $p.x - $pw/2; $py = $p.y - $ph/2
        $box = Add-RoundedRect $Slide $px $py $pw $ph $white $orange
    $null = Add-SubtleShadow $box
    $null = Add-TextBox $Slide ($px + 8) ($py + 6) ($pw - 16) 16 "{{PHASE${num}}}" 13 $true $orange "center"
    $null = Add-TextBox $Slide ($px + 8) ($py + 26) ($pw - 16) 36 "{{PHASE${num}_DESC}}" 10 $false $darkGray "center"
        if ($i -lt 3) {
            $next = $phases[$i + 1]
            $arrowX = $p.x + ($next.x - $p.x) * 0.25; $arrowY = $p.y + ($next.y - $p.y) * 0.25
            $aw = 20; $ah = 20
            if ($i -eq 0) { Add-Shape $Slide 36 ($p.x + 40) ($p.y + 10) $aw $ah $orange -1 }
            elseif ($i -eq 1) { Add-Shape $Slide 33 ($p.x + 10) ($p.y + 10) $aw $ah $orange -1 }
            elseif ($i -eq 2) { Add-Shape $Slide 34 ($p.x + 10) ($p.y - 10) $aw $ah $orange -1 }
        }
    }
    $center = Add-RoundedRect $Slide 340 250 160 80 $orange -1
    $null = Add-TextBox $Slide 340 262 160 24 "ADOPTION" 18 $true $white "center"
    $null = Add-TextBox $Slide 340 286 160 20 "Continue" 11 $false 0xFFDDCC "center"
    $null = Add-ConclusionBanner $Slide
    $null = Add-Footer $Slide
}

function Make-ConclusionSlide {
    param($Slide)
    $null = Add-Shape $Slide 1 ($margin - 36) 0 ($slideWidth) 5 $orange -1
    $null = Add-TextBox $Slide $margin 50 $contentW 50 "Pour conclure" 18 $false $medGray "center"
    $null = Add-TextBox $Slide $margin 110 $contentW 80 "{{MAIN_MESSAGE}}" 30 $true $orange "center"
    $null = Add-Shape $Slide 1 ($slideWidth/2 - 50) 200 100 3 $orange -1
    $tw = 250; $tg = 30; $tStart = ($slideWidth - (3 * $tw + 2 * $tg)) / 2
    for ($i = 1; $i -le 3; $i++) {
        $tx = $tStart + ($i - 1) * ($tw + $tg)
        Add-RoundedRect $Slide $tx 230 $tw 140 $lightBg $borderLight
    $null = Add-TextBox $Slide ($tx + 16) 250 ($tw - 32) 24 "${i}" 20 $true $orange "center"
    $null = Add-TextBox $Slide ($tx + 16) 280 ($tw - 32) 70 "{{TAKEAWAY_${i}}}" 13 $false $darkGray "center"
    }
    $null = Add-ConclusionBanner $Slide ($conclusionY - 30)
    $null = Add-Footer $Slide
}

$powerPoint = $null; $presentation = $null
try {
    $null = & {
        $script:ppPowerPoint = New-Object -ComObject PowerPoint.Application
        try { $script:ppPowerPoint.Visible = if ($Visible) { -1 } else { 0 } } catch { $script:ppPowerPoint.Visible = -1 }
        try { $script:ppPowerPoint.DisplayAlerts = 1 } catch { }
        $script:ppPresentation = $script:ppPowerPoint.Presentations.Add()
        $script:ppPresentation.PageSetup.SlideWidth = $slideWidth
        $script:ppPresentation.PageSetup.SlideHeight = $slideHeight
        $script:ppPresentation.PageSetup.SlideOrientation = 1
        $script:ppPresentation.PageSetup.SlideSize = 2
        for ($i = 1; $i -le 11; $i++) {
            [void]$script:ppPresentation.Slides.Add($i, 1)
        }
        $null = Make-TitleSlide $script:ppPresentation.Slides.Item(1)
        $null = Make-ContextKpiSlide $script:ppPresentation.Slides.Item(2)
        $null = Make-ProblemSlide $script:ppPresentation.Slides.Item(3)
        $null = Make-QuestionAnswerSlide $script:ppPresentation.Slides.Item(4)
        $null = Make-CardsSlide $script:ppPresentation.Slides.Item(5) 3 3
        $null = Make-CardsSlide $script:ppPresentation.Slides.Item(6) 4 2
        $null = Make-ProcessSlide $script:ppPresentation.Slides.Item(7) 4
        $null = Make-RoleFocusSlide $script:ppPresentation.Slides.Item(8)
        $null = Make-LessonsSlide $script:ppPresentation.Slides.Item(9) 4
        $null = Make-AdoptionLoopSlide $script:ppPresentation.Slides.Item(10)
        $null = Make-ConclusionSlide $script:ppPresentation.Slides.Item(11)
        $script:ppPresentation.SaveAs($OutputPath, 24)
        $script:ppPresentation.Saved = $true
    }
    $powerPoint = $script:ppPowerPoint
    $presentation = $script:ppPresentation
    Write-Output "Template cree: $OutputPath"
    Write-Output "Layouts: title_slide, context_kpi_slide, problem_slide, question_answer_slide, three_cards_slide, four_cards_slide, process_slide, role_focus_slide, lessons_slide, adoption_loop_slide, conclusion_slide"
} finally {
    if ($presentation) { try { $presentation.Close() } catch { } }
    if ($powerPoint) { try { $powerPoint.Quit() } catch { } }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
