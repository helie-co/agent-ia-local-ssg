param(
    [string]$ProjectPath = (Get-Location).Path,
    [string]$Ref = "main",
    [string]$RepositoryRawBase = "https://raw.githubusercontent.com/helie-co/agent-ia-local-ssg",
    [switch]$VerifyOnly
)

$ErrorActionPreference = "Stop"

function Resolve-ProjectPath {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Projet introuvable: $Path" }
    return (Resolve-Path -LiteralPath $Path).Path
}

function Copy-Or-DownloadFile {
    param([string]$SourceRelativePath, [string]$DestinationPath)

    $destinationParent = Split-Path -Parent $DestinationPath
    if ($destinationParent -and -not (Test-Path -LiteralPath $destinationParent)) {
        New-Item -ItemType Directory -Path $destinationParent | Out-Null
    }

    $localSource = $null
    if ($PSScriptRoot) {
        $candidate = Join-Path $PSScriptRoot ($SourceRelativePath -replace '^ppt-v2/', '')
        if (Test-Path -LiteralPath $candidate) { $localSource = $candidate }
    }

    if ($localSource) {
        Copy-Item -LiteralPath $localSource -Destination $DestinationPath -Force
        return
    }

    $urlPath = $SourceRelativePath -replace '\\', '/'
    $url = "$RepositoryRawBase/$Ref/$urlPath"
    try { Invoke-WebRequest -Uri $url -OutFile $DestinationPath }
    catch { throw "Telechargement impossible: $url -> $DestinationPath. $($_.Exception.Message)" }
}

function Test-InstalledFiles {
    param([string]$ProjectRoot, [array]$Files)

    $missing = @()
    foreach ($file in $Files) {
        $destination = Join-Path $ProjectRoot $file.Destination
        if (-not (Test-Path -LiteralPath $destination)) { $missing += $file.Destination }
    }
    if ($missing.Count -gt 0) { throw "Installation /ppt-v2 incomplete. Fichiers manquants: $($missing -join ', ')" }
}

$projectRoot = Resolve-ProjectPath $ProjectPath

$files = @(
    @{ Source = "ppt-v2/.opencode/commands/ppt-v2.md"; Destination = ".opencode/commands/ppt-v2.md" },
    @{ Source = "ppt-v2/.opencode/scripts/ppt-v2/New-CardTemplate.ps1"; Destination = ".opencode/scripts/ppt-v2/New-CardTemplate.ps1" },
    @{ Source = "ppt-v2/.opencode/scripts/ppt-v2/Get-CardLayouts.ps1"; Destination = ".opencode/scripts/ppt-v2/Get-CardLayouts.ps1" },
    @{ Source = "ppt-v2/.opencode/scripts/ppt-v2/Generate-CardPpt.ps1"; Destination = ".opencode/scripts/ppt-v2/Generate-CardPpt.ps1" },
    @{ Source = "ppt-v2/.opencode/scripts/ppt-v2/slides.example.json"; Destination = ".opencode/scripts/ppt-v2/slides.example.json" },
    @{ Source = "ppt-v2/README.md"; Destination = ".opencode/scripts/ppt-v2/README.md" }
)

if (-not $VerifyOnly) {
    foreach ($file in $files) {
        $destination = Join-Path $projectRoot $file.Destination
        Copy-Or-DownloadFile -SourceRelativePath $file.Source -DestinationPath $destination
    }
}

Test-InstalledFiles -ProjectRoot $projectRoot -Files $files

Write-Output "Commande /ppt-v2 installee dans le projet: $projectRoot"
Write-Output "Fichiers verifies: $($files.Count)"
Write-Output "Cibles: .opencode/commands/ppt-v2.md et .opencode/scripts/ppt-v2/ (6 fichiers)"
Write-Output "Redemarrez OpenCode Desktop depuis ce projet pour charger la commande."
Write-Output ""
Write-Output "Pour creer un template cards, lancez depuis OpenCode:"
Write-Output "/ppt-v2 --create-template"
