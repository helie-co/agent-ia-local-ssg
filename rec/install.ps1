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

Write-Output "Commande /rec installee dans le projet: $projectRoot"
Write-Output "Fichiers verifies: $($files.Count)"
Write-Output ""
Write-Output "Les dependances ne sont pas installees par ce script."
Write-Output "Redemarrez OpenCode Desktop depuis ce projet, puis lancez /rec install pour installer ou reparer ffmpeg, whisper-cli, le modele whisper, et verifier Stereo Mix."
