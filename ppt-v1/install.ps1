param(
    [string]$ProjectPath = (Get-Location).Path,
    [string]$Ref = "main",
    [string]$RepositoryRawBase = "https://raw.githubusercontent.com/helie-co/agent-ia-local-ssg",
    [switch]$VerifyOnly
)

$ErrorActionPreference = "Stop"

function Resolve-ProjectPath {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Projet introuvable: $Path"
    }

    return (Resolve-Path -LiteralPath $Path).Path
}

function Copy-Or-DownloadFile {
    param(
        [string]$SourceRelativePath,
        [string]$DestinationPath
    )

    $destinationParent = Split-Path -Parent $DestinationPath
    if ($destinationParent -and -not (Test-Path -LiteralPath $destinationParent)) {
        New-Item -ItemType Directory -Path $destinationParent | Out-Null
    }

    $localSource = $null
    if ($PSScriptRoot) {
        $candidate = Join-Path $PSScriptRoot ($SourceRelativePath -replace '^ppt-v1/', '')
        if (Test-Path -LiteralPath $candidate) { $localSource = $candidate }
    }

    if ($localSource) {
        Copy-Item -LiteralPath $localSource -Destination $DestinationPath -Force
        return
    }

    $urlPath = $SourceRelativePath -replace '\\', '/'
    $url = "$RepositoryRawBase/$Ref/$urlPath"

    try {
        Invoke-WebRequest -Uri $url -OutFile $DestinationPath
    }
    catch {
        throw "Telechargement impossible: $url -> $DestinationPath. $($_.Exception.Message)"
    }
}

function Test-InstalledFiles {
    param(
        [string]$ProjectRoot,
        [array]$Files
    )

    $missing = @()
    foreach ($file in $Files) {
        $destination = Join-Path $ProjectRoot $file.Destination
        if (-not (Test-Path -LiteralPath $destination)) {
            $missing += $file.Destination
        }
    }

    if ($missing.Count -gt 0) {
        throw "Installation /ppt-v1 incomplete. Fichiers manquants: $($missing -join ', ')"
    }
}

$projectRoot = Resolve-ProjectPath $ProjectPath

$files = @(
    @{ Source = "ppt-v1/.opencode/commands/ppt-v1.md"; Destination = ".opencode/commands/ppt-v1.md" },
    @{ Source = "ppt-v1/.opencode/scripts/ppt-v1/README.md"; Destination = ".opencode/scripts/ppt-v1/README.md" },
    @{ Source = "ppt-v1/.opencode/scripts/ppt-v1/Extract-ImageOcr.ps1"; Destination = ".opencode/scripts/ppt-v1/Extract-ImageOcr.ps1" },
    @{ Source = "ppt-v1/.opencode/scripts/ppt-v1/Extract-PptContent.ps1"; Destination = ".opencode/scripts/ppt-v1/Extract-PptContent.ps1" },
    @{ Source = "ppt-v1/.opencode/scripts/ppt-v1/Get-PptLayouts.ps1"; Destination = ".opencode/scripts/ppt-v1/Get-PptLayouts.ps1" },
    @{ Source = "ppt-v1/.opencode/scripts/ppt-v1/Generate-Ppt.ps1"; Destination = ".opencode/scripts/ppt-v1/Generate-Ppt.ps1" },
    @{ Source = "ppt-v1/.opencode/scripts/ppt-v1/Initialize-TemplatePlaceholders.ps1"; Destination = ".opencode/scripts/ppt-v1/Initialize-TemplatePlaceholders.ps1" },
    @{ Source = "ppt-v1/.opencode/scripts/ppt-v1/Test-PptVisual.ps1"; Destination = ".opencode/scripts/ppt-v1/Test-PptVisual.ps1" },
    @{ Source = "ppt-v1/.opencode/scripts/ppt-v1/layouts.json"; Destination = ".opencode/scripts/ppt-v1/layouts.json" },
    @{ Source = "ppt-v1/.opencode/scripts/ppt-v1/SopraSteriaNext.layouts.json"; Destination = ".opencode/scripts/ppt-v1/SopraSteriaNext.layouts.json" },
    @{ Source = "ppt-v1/.opencode/scripts/ppt-v1/SopraSteriaNext.template.pptx"; Destination = ".opencode/scripts/ppt-v1/SopraSteriaNext.template.pptx" },
    @{ Source = "ppt-v1/.opencode/scripts/ppt-v1/slides.example.json"; Destination = ".opencode/scripts/ppt-v1/slides.example.json" }
)

if (-not $VerifyOnly) {
    foreach ($file in $files) {
        $destination = Join-Path $projectRoot $file.Destination
        Copy-Or-DownloadFile -SourceRelativePath $file.Source -DestinationPath $destination
    }
}

Test-InstalledFiles -ProjectRoot $projectRoot -Files $files

"Commande /ppt-v1 installee dans le projet: $projectRoot"
"Fichiers verifies: $($files.Count)"
"Cibles modifiees uniquement: .opencode/commands/ppt-v1.md et .opencode/scripts/ppt-v1/"
"OpenCode Desktop n'a pas ete installe ni modifie."
"Aucune configuration globale OpenCode n'a ete modifiee."
"Redemarrez OpenCode Desktop depuis ce projet pour charger la commande."
