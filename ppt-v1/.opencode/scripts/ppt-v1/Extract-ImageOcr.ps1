param(
    [Parameter(Mandatory = $true)]
    [string]$SourceImage,
    [string]$OutputMarkdown = (Join-Path (Join-Path (Get-Location) "ppt") "image.extracted.md")
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $SourceImage)) { throw "Image source introuvable: $SourceImage" }

$SourceImage = (Resolve-Path -LiteralPath $SourceImage).Path
$OutputMarkdown = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputMarkdown)
$outputParent = Split-Path -Parent $OutputMarkdown
if ($outputParent -and -not (Test-Path -LiteralPath $outputParent)) {
    New-Item -ItemType Directory -Path $outputParent | Out-Null
}

$supportedExtensions = @(".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff")
$extension = [System.IO.Path]::GetExtension($SourceImage).ToLowerInvariant()
if ($supportedExtensions -notcontains $extension) {
    throw "Format image non supporte pour OCR: $extension. Formats supportes: $($supportedExtensions -join ', ')"
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

$tesseractPath = Get-TesseractPath
if (-not $tesseractPath) {
    throw "tesseract.exe est introuvable. Lance /ppt-v1 --install puis relance l'extraction image."
}
$tesseractLanguage = Get-TesseractLanguage $tesseractPath

$ocrDirectory = Join-Path $outputParent "image.extracted.ocr"
if (Test-Path -LiteralPath $ocrDirectory) { Remove-Item -LiteralPath $ocrDirectory -Recurse -Force }
New-Item -ItemType Directory -Path $ocrDirectory | Out-Null

$imageCopyPath = Join-Path $ocrDirectory ("source" + $extension)
Copy-Item -LiteralPath $SourceImage -Destination $imageCopyPath -Force

$outputBase = Join-Path $ocrDirectory "source"
$outputTextPath = $outputBase + ".txt"
$null = & $tesseractPath $imageCopyPath $outputBase -l $tesseractLanguage --psm 6 2>$null

if (-not (Test-Path -LiteralPath $outputTextPath)) {
    throw "OCR image echoue: aucun fichier texte produit par Tesseract."
}

$ocrText = [System.IO.File]::ReadAllText($outputTextPath, [System.Text.Encoding]::UTF8)
$ocrLines = @(Split-TextLines $ocrText)

$title = if ($ocrLines.Count -gt 0) { $ocrLines[0] } else { "Texte extrait de l'image" }
$subtitle = if ($ocrLines.Count -gt 1) { $ocrLines[1] } else { "" }

$lines = @()
$lines += "# Reprise d'image"
$lines += ""
$lines += "Source: $SourceImage"
$lines += "OCR images: $tesseractPath"
$lines += "OCR langue: $tesseractLanguage"
$lines += ""
$lines += "## Image 1 - Texte OCR"
$lines += "Layout: classic_bullets"
$lines += ""
$lines += "Titre: $title"
$lines += "Sous-titre: $subtitle"
$lines += ""
$lines += "Texte OCR image:"
foreach ($ocrLine in $ocrLines) { $lines += "- $ocrLine" }
$lines += ""

[System.IO.File]::WriteAllLines($OutputMarkdown, $lines, [System.Text.UTF8Encoding]::new($false))
"Extraction OCR image terminee: $OutputMarkdown"
