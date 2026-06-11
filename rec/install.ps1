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
        $candidate = Join-Path $PSScriptRoot ($SourceRelativePath -replace '^rec/', '')
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
        throw "Installation /rec incomplete. Fichiers manquants: $($missing -join ', ')"
    }
}

$projectRoot = Resolve-ProjectPath $ProjectPath

$files = @(
    @{ Source = "rec/.opencode/commands/rec.md"; Destination = ".opencode/commands/rec.md" },
    @{ Source = "rec/.opencode/scripts/rec/rec.ps1"; Destination = ".opencode/scripts/rec/rec.ps1" },
    @{ Source = "rec/README.md"; Destination = ".opencode/scripts/rec/README.md" }
)

if (-not $VerifyOnly) {
    foreach ($file in $files) {
        $destination = Join-Path $projectRoot $file.Destination
        Copy-Or-DownloadFile -SourceRelativePath $file.Source -DestinationPath $destination
    }
}

Test-InstalledFiles -ProjectRoot $projectRoot -Files $files

"Commande /rec installee dans le projet: $projectRoot"
"Fichiers verifies: $($files.Count)"
"Cibles modifiees uniquement: .opencode/commands/rec.md et .opencode/scripts/rec/"
"OpenCode Desktop n'a pas ete installe ni modifie."
"Aucune configuration globale OpenCode n'a ete modifiee."
"Redemarrez OpenCode Desktop depuis ce projet pour charger la commande."
""
"Pour installer les dependances (ffmpeg, whisper, modele), lancez ensuite depuis OpenCode:"
"/rec --install"
