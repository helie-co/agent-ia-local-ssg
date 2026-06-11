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

# --- Installation des dependances ---
$toolsRoot = Join-Path $env:LOCALAPPDATA 'opencode-tools'
$whisperDir = Join-Path $toolsRoot 'whisper.cpp'
$modelDir = Join-Path $whisperDir 'models'
$modelFile = Join-Path $modelDir 'ggml-small.bin'

Write-Output "=== Installation des dependances ==="

# --- ffmpeg ---
$ffmpegExe = $null

# 1) Chercher dans le PATH
$ffmpegCmd = Get-Command ffmpeg -ErrorAction SilentlyContinue
if ($ffmpegCmd) { $ffmpegExe = $ffmpegCmd.Source }

# 2) Chercher dans WinGet Packages si pas dans le PATH
if (-not $ffmpegExe) {
  $candidate = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\WinGet\Packages" -Recurse -Filter ffmpeg.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
  if ($candidate) { $ffmpegExe = $candidate }
}

# 3) Installer via winget si toujours introuvable
if (-not $ffmpegExe) {
  Write-Output "Installation de ffmpeg (winget)..."
  winget install -e --id Gyan.FFmpeg --accept-source-agreements --accept-package-agreements | Out-Null
  # Recharger le PATH depuis le systeme
  $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User')
  # Revérifier dans le PATH
  $ffmpegCmd = Get-Command ffmpeg -ErrorAction SilentlyContinue
  if ($ffmpegCmd) { $ffmpegExe = $ffmpegCmd.Source }
  # Revérifier dans WinGet Packages
  if (-not $ffmpegExe) {
    $candidate = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\WinGet\Packages" -Recurse -Filter ffmpeg.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
    if ($candidate) { $ffmpegExe = $candidate }
  }
}

if ($ffmpegExe) { Write-Output "ffmpeg: $ffmpegExe" } else { Write-Output "ffmpeg: NON installe" }

# --- whisper-cli ---
$whisperCli = Join-Path $whisperDir 'Release\whisper-cli.exe'
if (-not (Test-Path $whisperCli)) {
  Write-Output "Telechargement de whisper-cli..."
  $null = New-Item -ItemType Directory -Force -Path $whisperDir
  $zip = Join-Path $toolsRoot 'whisper-bin-x64.zip'
  try {
    Invoke-WebRequest -Uri 'https://github.com/ggml-org/whisper.cpp/releases/latest/download/whisper-bin-x64.zip' -OutFile $zip -ErrorAction Stop
    Expand-Archive -LiteralPath $zip -DestinationPath $whisperDir -Force
    Remove-Item $zip -Force
    $found = Get-ChildItem $whisperDir -Recurse -Filter 'whisper-cli.exe' -File | Where-Object { $_.FullName -match '\\Release\\' } | Select-Object -First 1 -ExpandProperty FullName
    if ($found) { $whisperCli = $found }
  } catch { Write-Output "Echec telechargement whisper-cli: $_" }
}
if (Test-Path $whisperCli) { Write-Output "whisper-cli: $whisperCli" } else { Write-Output "whisper-cli: NON installe" }

# --- modele whisper ---
if (-not (Test-Path $modelFile)) {
  Write-Output "Telechargement du modele ggml-small.bin..."
  $null = New-Item -ItemType Directory -Force -Path $modelDir
  try {
    Invoke-WebRequest -Uri 'https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-small.bin' -OutFile $modelFile -ErrorAction Stop
  } catch { Write-Output "Echec telechargement modele: $_" }
}
if (Test-Path $modelFile) {
  $size = (Get-Item $modelFile).Length
  Write-Output "modele: $modelFile ($($size / 1MB -as [int]) Mo)"
} else { Write-Output "modele: NON installe" }

# --- Stereo Mix ---
if ($ffmpegExe) {
  $output = & cmd.exe /d /c "`"$ffmpegExe`" -hide_banner -list_devices true -f dshow -i dummy 2>&1" | Out-String
  $hasStereoMix = $output -match '(?i)stereo mix|mixage stereo|what u hear'
  if ($hasStereoMix) {
    Write-Output "stereo_mix: detecte"
  } else {
    Write-Output "stereo_mix: absent"
    Write-Output "Ouverture de la fenetre Son (onglet Enregistrement)..."
    Write-Output "Activez Stereo Mix : clic droit > Afficher les peripheriques desactives > Activer"
    $null = Start-Process 'control' -ArgumentList 'mmsys.cpl,,1'
  }
}

Write-Output ""
Write-Output "Installation terminee. Redemarrez OpenCode Desktop depuis ce projet pour charger la commande."
