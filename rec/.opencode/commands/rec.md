---
description: Enregistrement ecran/audio avec transcription live (whisper.cpp), fusion, et gestion des chunks
agent: build
---

Execute la commande PowerShell ci-dessous. Retourne le texte produit par la commande dans un bloc de code markdown (``` ... ```), sans modification du contenu.

Le script `rec.ps1` gere l'enregistrement ecran/audio avec ou sans transcription live automatique.

Mode aide :

- Si la demande est exactement `/rec --help`, ne lance pas le workflow d'enregistrement.
- Lis et affiche le contenu complet de `.opencode/scripts/rec/README.md`.
- Si le fichier README est introuvable, affiche une aide courte avec les modes supportes et indique que `.opencode/scripts/rec/README.md` est manquant.

Mode (re)installation des dependances :

- Si la demande est exactement `/rec --install`, ne lance pas le workflow d'enregistrement.
- Execute le bloc PowerShell ci-dessous pour installer ou reparer les dependances : ffmpeg, whisper-cli, modele whisper, et verifier Stereo Mix.

Commande a executer pour `/rec --install` :

```powershell
$ErrorActionPreference = 'Stop'

$toolsRoot = Join-Path $env:LOCALAPPDATA 'opencode-tools'
$whisperDir = Join-Path $toolsRoot 'whisper.cpp'
$modelDir = Join-Path $whisperDir 'models'
$modelFile = Join-Path $modelDir 'ggml-small.bin'

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
  Write-Output 'Installation de ffmpeg...'
  winget install -e --id Gyan.FFmpeg --accept-source-agreements --accept-package-agreements
  $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User')
  $ffmpegCmd = Get-Command ffmpeg -ErrorAction SilentlyContinue
  if ($ffmpegCmd) { $ffmpegExe = $ffmpegCmd.Source }
  if (-not $ffmpegExe) {
    $candidate = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\WinGet\Packages" -Recurse -Filter ffmpeg.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
    if ($candidate) { $ffmpegExe = $candidate }
  }
}

if (-not $ffmpegExe) { throw 'ffmpeg introuvable apres installation' }

"ffmpeg=$ffmpegExe"

# --- whisper-cli ---
$whisperCli = Join-Path $whisperDir 'Release\whisper-cli.exe'
if (-not (Test-Path $whisperCli)) {
  Write-Output 'Telechargement de whisper-cli...'
  New-Item -ItemType Directory -Force -Path $whisperDir | Out-Null
  $zip = Join-Path $toolsRoot 'whisper-bin-x64.zip'
  Invoke-WebRequest -Uri 'https://github.com/ggml-org/whisper.cpp/releases/latest/download/whisper-bin-x64.zip' -OutFile $zip -ErrorAction Stop
  Expand-Archive -LiteralPath $zip -DestinationPath $whisperDir -Force
  Remove-Item $zip -Force
  $found = Get-ChildItem $whisperDir -Recurse -Filter 'whisper-cli.exe' -File | Where-Object { $_.FullName -match '\\Release\\' } | Select-Object -First 1 -ExpandProperty FullName
  if (-not $found) { throw 'whisper-cli.exe introuvable apres extraction' }
  $whisperCli = $found
}
"whisper-cli=$whisperCli"

# --- modele whisper ---
if (-not (Test-Path $modelFile)) {
  Write-Output 'Telechargement du modele ggml-small.bin...'
  New-Item -ItemType Directory -Force -Path $modelDir | Out-Null
  Invoke-WebRequest -Uri 'https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-small.bin' -OutFile $modelFile -ErrorAction Stop
}
if (-not (Test-Path $modelFile)) { throw 'Modele introuvable apres telechargement' }
"modele=$modelFile"

$modelSize = (Get-Item $modelFile).Length
if ($modelSize -lt 400MB) { throw "Modele whisper incomplet: $modelFile ($modelSize octets). Supprime ce fichier puis relance /rec --install." }
"modele_taille=$($modelSize) octets"

# --- Stereo Mix ---
$output = & cmd.exe /d /c "`"$ffmpegExe`" -hide_banner -list_devices true -f dshow -i dummy 2>&1" | Out-String
$hasStereoMix = $output -match '(?i)stereo mix|mixage stereo|what u hear'
if ($hasStereoMix) {
  'stereo_mix=detecte'
} else {
  'stereo_mix=absent'
  Write-Output ''
  Write-Output '=== Stereo Mix non detecte ==='
  Write-Output 'Pour enregistrer l audio systeme, active Stereo Mix :'
  Write-Output '1. La fenetre "Son" va s ouvrir (onglet Enregistrement)'
  Write-Output '2. Clic droit dans la liste vide -> Afficher les peripheriques desactives'
  Write-Output '3. Clic droit sur Stereo Mix -> Activer'
  Write-Output "4. Reviens ici et verifie avec 'ffmpeg -list_devices'"
  Start-Process 'control' -ArgumentList 'mmsys.cpl,,1'
}

Write-Output ''
Write-Output 'Installation terminee.'
Write-Output 'Redemarre OpenCode Desktop pour charger la commande /rec.'
```

Commande normale :

```powershell
$script = Join-Path (Get-Location) '.opencode\scripts\rec\rec.ps1'
$rawArgs = '$ARGUMENTS'
$output = & powershell -NoProfile -ExecutionPolicy Bypass -Command "& '$script' -RawArgs '$rawArgs'"
@($output) -join "`r`n"
```
