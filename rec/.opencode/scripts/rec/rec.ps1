param(
  [string]$Command = '',
  [string]$RawArgs = '',
  [string]$Title = '',
  [string]$File = '',
  [string]$Language = 'fr',
  [int]$Index = 0,
  [string]$Selector = '',
  [switch]$WatchTranscript,
  [int]$FfmpegPid = 0,
  [string]$Timestamp = '',
  [string]$Mode = 'audio',
  [string]$OutputFile = '',
  [switch]$BackgroundFinalize,
  [switch]$All,
  [switch]$Help
)

$Script:RecordingsDir = Join-Path (Get-Location) 'recordings'
$Script:LogsDir = Join-Path $Script:RecordingsDir 'logs'
$Script:StateFile = Join-Path $Script:RecordingsDir '.rec_state.json'
$Script:ChunkDuration = 300
$Script:ToolsRoot = Join-Path $env:LOCALAPPDATA 'opencode-tools'
$Script:WhisperDir = Join-Path $Script:ToolsRoot 'whisper.cpp'
$Script:InstallLogFile = Join-Path $Script:ToolsRoot 'rec-install.log'
$Script:InstallPidFile = Join-Path $Script:ToolsRoot 'rec-install.pid'

$script:FfmpegPath = $null
$script:WhisperCliPath = $null
$script:WhisperModelPath = $null

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class NativeMethods {
    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern uint SetThreadExecutionState(uint esFlags);
}
'@

function Start-DetachedProcess($cmdLine) {
  $startupClass = New-Object System.Management.ManagementClass("Win32_ProcessStartup")
  $startup = $startupClass.CreateInstance()
  $startup["ShowWindow"] = 0
  $workingDirectory = (Get-Location).Path
  $result = ([wmiclass]"Win32_Process").Create($cmdLine, $workingDirectory, $startup)
  if ($result.ReturnValue -ne 0) { throw "Echec creation processus (code: $($result.ReturnValue), cwd: $workingDirectory)" }
  return $result.ProcessId
}

function Get-FfmpegPath {
  if ($script:FfmpegPath) { return $script:FfmpegPath }

  $ffmpeg = Get-Command ffmpeg -ErrorAction SilentlyContinue
  if ($ffmpeg) { $script:FfmpegPath = $ffmpeg.Source; return $script:FfmpegPath }

  $wingetPath = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\WinGet\Packages" -Recurse -Filter ffmpeg.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
  if ($wingetPath) { $script:FfmpegPath = $wingetPath; return $script:FfmpegPath }

  throw "ffmpeg introuvable. Lance /rec install pour installer ou reparer les dependances."
}

function Get-WhisperCliPath {
  if ($script:WhisperCliPath) { return $script:WhisperCliPath }

  $exe = Join-Path $Script:WhisperDir 'Release\whisper-cli.exe'
  if (Test-Path $exe) { $script:WhisperCliPath = $exe; return $script:WhisperCliPath }

  $fromPath = Get-Command whisper-cli -ErrorAction SilentlyContinue
  if ($fromPath) { $script:WhisperCliPath = $fromPath.Source; return $script:WhisperCliPath }

  throw "whisper-cli introuvable. Lance /rec install pour installer ou reparer les dependances."
}

function Get-WhisperModelPath {
  if ($script:WhisperModelPath) { return $script:WhisperModelPath }

  $model = Join-Path $Script:WhisperDir 'models\ggml-small.bin'
  if (Test-Path $model) {
    $size = (Get-Item $model).Length
    if ($size -gt 400MB) { $script:WhisperModelPath = $model; return $script:WhisperModelPath }
    throw "Modele whisper incomplet (${size} octets). Supprime le fichier et relance /rec install."
  }

  throw "Modele whisper introuvable. Lance /rec install pour installer ou reparer les dependances."
}

function Get-AudioDevices {
  $ffmpeg = Get-FfmpegPath
  $output = & cmd.exe /d /c "`"$ffmpeg`" -hide_banner -list_devices true -f dshow -i dummy 2>&1" | Out-String
  $devices = @()
  foreach ($line in ($output -split "`r?`n")) {
    if ($line -match '"([^"]+)"\s+\(audio\)') {
      $devices += $matches[1]
    }
  }
  return $devices
}

function Select-AudioDevices($devices) {
  $mic = $devices | Where-Object { $_ -match '(?i)micro[ph]?one?|micro|mic' } | Select-Object -First 1
  $system = $devices | Where-Object { $_ -match '(?i)stereo mix|mixage stereo|what u hear|wave out|cable output' } | Select-Object -First 1
  return @{ mic = $mic; system = $system }
}

function Show-Devices {
  $devices = Get-AudioDevices
  $sel = Select-AudioDevices $devices
  $micFound = $false
  $sysFound = $false
  Write-Output "Peripheriques audio disponibles ($($devices.Count)):"
  foreach ($d in $devices) {
    $tag = 'OTHER'
    if ($d -eq $sel.mic) { $tag = 'MIC'; $micFound = $true }
    if ($d -eq $sel.system) { $tag = 'SYSTEM'; $sysFound = $true }
    Write-Output "  [$tag] $d"
  }
  Write-Output ""
  if (-not $micFound) { Write-Output "ATTENTION: aucun micro detecte. La transcription sera silencieuse." }
  if (-not $sysFound) { Write-Output "ATTENTION: aucun Stereo Mix detecte. Modifier les sons système > Enregistrement > Stéréo Mix > Bouton droit > Activer > OK." }
  if ($micFound -and $sysFound) { Write-Output "OK: Micro et Stereo Mix trouves, enregistrement audio+ fonctionnel." }
}

function Test-PidAlive($processId) {
  if ($processId -le 0) { return $false }
  try {
    $result = tasklist /FI "PID eq $processId" 2>&1 | Out-String
    return ($result -match [regex]::Escape("$processId")) -and ($result -notmatch 'No tasks are running')
  } catch { return $false }
}

function Get-NowStamp { return Get-Date -Format 'yyyyMMdd_HHmmss' }

function Get-NowUtcIso { return (Get-Date).ToUniversalTime().ToString('o') }

function Format-Timestamp($ts) {
  try {
    $dt = [datetime]::ParseExact($ts, 'yyyyMMdd_HHmmss', $null)
    return $dt.ToString('dd/MM HH:mm')
  } catch { return $ts }
}

function Get-StateSessions {
  if (-not (Test-Path $Script:StateFile)) { return @() }
  try {
    $raw = Get-Content $Script:StateFile -Raw -Encoding UTF8 | ConvertFrom-Json
    $sessions = @($raw.sessions)
    $alive = @($sessions | Where-Object { Test-PidAlive $_.pid })
    if ($alive.Count -ne $sessions.Count) { Set-StateSessions $alive }
    return @($alive | Sort-Object timestamp)
  } catch { return @() }
}

function Set-StateSessions($sessions) {
  $dir = Split-Path $Script:StateFile
  if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
  $payload = @{ updated_at_utc = Get-NowUtcIso; sessions = @($sessions) }
  $payload | ConvertTo-Json -Depth 10 | Set-Content $Script:StateFile -Encoding UTF8
}

function New-SessionObj($timestamp, $mode, $chunked, $processId, $outputFile, $logFile, $wakeLockPid, $liveTranscribePid) {
  return @{
    timestamp = $timestamp
    mode = $mode
    chunked = $chunked
    pid = $processId
    output_file = $outputFile
    started_at_utc = Get-NowUtcIso
    log_file = $logFile
    wake_lock_pid = $wakeLockPid
    live_transcribe_pid = $liveTranscribePid
  }
}

function Get-ChunkFiles($root, $timestamp, $mode) {
  if ($mode -eq 'video') {
    $tsFiles = @(Get-ChildItem (Join-Path $root "${timestamp}_chunk_*.ts") -ErrorAction SilentlyContinue | Where-Object { $_.Length -gt 0 } | Sort-Object Name)
    if ($tsFiles) { return $tsFiles }
    return @(Get-ChildItem (Join-Path $root "${timestamp}_chunk_*.mp4") -ErrorAction SilentlyContinue | Where-Object { $_.Length -gt 0 } | Sort-Object Name)
  }
  return @(Get-ChildItem (Join-Path $root "${timestamp}_chunk_*.mp3") -ErrorAction SilentlyContinue | Where-Object { $_.Length -gt 0 } | Sort-Object Name)
}

function Test-NoActiveSession {
  $active = @(Get-StateSessions)
  if ($active.Count -gt 0) { throw "Un enregistrement est deja actif (timestamp: $($active[0].timestamp)). Utilise 'rec stop' d'abord." }
}

function Get-SessionTranscriptionProgress($session) {
  $root = Split-Path $session.output_file -Parent
  $chunks = Get-ChunkFiles $root $session.timestamp $session.mode
  $detected = $chunks.Count
  $transcribed = 0
  $lastChunkIdx = 0

  foreach ($chunk in $chunks) {
    if ($chunk.Name -match '_chunk_(\d+)\.') { $lastChunkIdx = [math]::Max($lastChunkIdx, [int]$matches[1]) }
    $txt = [System.IO.Path]::ChangeExtension($chunk.FullName, '.txt')
    if ((Test-Path $txt) -and ((Get-Item $txt).Length -gt 0)) { $transcribed++ }
  }

  $percent = if ($detected -gt 0) { [math]::Round(($transcribed / $detected) * 100) } else { 0 }
  return @{ chunks_detected = $detected; chunks_transcribed = $transcribed; percent = $percent; last_chunk_index = $lastChunkIdx }
}

function Write-Out($msg) { Write-Output $msg }

function Invoke-WakeLockStart($ffmpegPid) {
  $tempFile = Join-Path $env:TEMP "rec_wl_$([guid]::NewGuid().ToString('N')).ps1"
  $content = @'
param($TargetPid)
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class WL {
    [DllImport("kernel32.dll")]
    public static extern uint SetThreadExecutionState(uint esFlags);
}
"@
do {
    [WL]::SetThreadExecutionState([uint32]0x80000001)
    Start-Sleep -Seconds 30
    $alive = $false
    try {
        $r = tasklist /FI "PID eq $TargetPid" 2>&1 | Out-String
        $alive = ($r -match [regex]::Escape("$TargetPid")) -and ($r -notmatch "No tasks are running")
    } catch {}
} while ($alive)
'@
  $content | Set-Content $tempFile -Encoding UTF8
  $wlCmdLine = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$tempFile`" $ffmpegPid"
  $wlPid = Start-DetachedProcess $wlCmdLine
  Start-Sleep -Milliseconds 500
  return $wlPid
}

function Start-LiveTranscriber($ffmpegPid, $timestamp, $mode, $language) {
  $scriptPath = Join-Path (Get-Location) '.opencode\scripts\rec\rec.ps1'
  $cmdLine = "& '$scriptPath' -WatchTranscript -FfmpegPid $ffmpegPid -Timestamp '$timestamp' -Mode '$mode' -Language '$language'"
  $bytes = [Text.Encoding]::Unicode.GetBytes($cmdLine)
  $encoded = [Convert]::ToBase64String($bytes)
  $ltCmdLine = "powershell.exe -NoProfile -EncodedCommand $encoded"
  return Start-DetachedProcess $ltCmdLine
}

function Build-AudioInputArgs($devices) {
  $sel = Select-AudioDevices $devices
  $argsList = New-Object Collections.Generic.List[string]
  $inputCount = 0
  if ($sel.system -and $sel.mic) {
    $argsList.AddRange([string[]]@('-f', 'dshow', '-i', "audio=$($sel.system)"))
    $argsList.AddRange([string[]]@('-f', 'dshow', '-i', "audio=$($sel.mic)"))
    $inputCount = 2
  } elseif ($sel.system) {
    $argsList.AddRange([string[]]@('-f', 'dshow', '-i', "audio=$($sel.system)"))
    $inputCount = 1
  } elseif ($sel.mic) {
    $argsList.AddRange([string[]]@('-f', 'dshow', '-i', "audio=$($sel.mic)"))
    $inputCount = 1
  } elseif ($devices.Count -gt 0) {
    $argsList.AddRange([string[]]@('-f', 'dshow', '-i', "audio=$($devices[0])"))
    $inputCount = 1
  }
  return @{ args = $argsList; count = $inputCount }
}

function Start-Recording($mode, $chunked, $title) {
  Test-NoActiveSession

  $ffmpeg = Get-FfmpegPath
  $devices = Get-AudioDevices
  if ($devices.Count -eq 0) { throw 'Aucun peripherique audio detecte.' }

  $timestamp = Get-NowStamp
  $root = $Script:RecordingsDir
  $logs = $Script:LogsDir
  New-Item -ItemType Directory -Path $root -Force | Out-Null
  New-Item -ItemType Directory -Path $logs -Force | Out-Null

  $ext = if ($mode -eq 'video') { 'mp4' } else { 'mp3' }
  $stem = $timestamp
  if ($title) { $stem = "$timestamp-$title" }
  $outputFile = Join-Path $root "$stem.$ext"

  $chunkExt = if ($mode -eq 'video' -and $chunked) { 'ts' } else { $ext }
  $chunkPattern = Join-Path $root "${timestamp}_chunk_%03d.$chunkExt"
  $logFile = Join-Path $logs "${timestamp}.log"

  $ffArgs = New-Object Collections.Generic.List[string]
  if ($mode -eq 'video') {
    $ffArgs.AddRange([string[]]@('-y', '-f', 'gdigrab', '-framerate', '15', '-i', 'desktop'))
    $audioInput = Build-AudioInputArgs $devices
    $ffArgs.AddRange($audioInput.args)
    $inputCount = $audioInput.count
    if ($inputCount -eq 2) {
      $ffArgs.AddRange([string[]]@('-filter_complex', '[1:a][2:a]amix=inputs=2[aout]', '-map', '0:v', '-map', '[aout]'))
    } elseif ($inputCount -eq 1) {
      $ffArgs.AddRange([string[]]@('-map', '0:v', '-map', '1:a'))
    }
    $ffArgs.AddRange([string[]]@('-c:v', 'libx264', '-preset', 'ultrafast', '-crf', '28', '-pix_fmt', 'yuv420p'))
    if ($inputCount -gt 0) { $ffArgs.AddRange([string[]]@('-c:a', 'aac', '-b:a', '128k')) }
    if ($chunked) {
      $ffArgs.AddRange([string[]]@('-f', 'segment', '-segment_format', 'mpegts', '-segment_time', "$($Script:ChunkDuration)", '-reset_timestamps', '1', $chunkPattern))
    } else {
      $ffArgs.Add($outputFile)
    }
  } else {
    $ffArgs.Add('-y')
    $audioInput = Build-AudioInputArgs $devices
    $ffArgs.AddRange($audioInput.args)
    if ($audioInput.count -eq 2) { $ffArgs.AddRange([string[]]@('-filter_complex', 'amix=inputs=2')) }
    $ffArgs.AddRange([string[]]@('-ac', '1', '-ar', '16000', '-c:a', 'libmp3lame', '-b:a', '96k'))
    if ($chunked) {
      $ffArgs.AddRange([string[]]@('-f', 'segment', '-segment_time', "$($Script:ChunkDuration)", '-reset_timestamps', '1', $chunkPattern))
    } else {
      $ffArgs.Add($outputFile)
    }
  }

  $flatArgs = ($ffArgs | ForEach-Object { if ($_ -match '[ \t"]') { "`"$_`"" } else { $_ } }) -join ' '
  $ffCmdLine = "`"$ffmpeg`" $flatArgs"
  $ffmpegPid = Start-DetachedProcess $ffCmdLine
  Start-Sleep -Milliseconds 500

  if (-not (Test-PidAlive $ffmpegPid)) { throw "FFmpeg s'est arrete immediatement. Log: $logFile" }

  $wakeLockPid = Invoke-WakeLockStart $ffmpegPid
  $liveTranscribePid = 0
  if ($chunked) {
    $liveTranscribePid = Start-LiveTranscriber $ffmpegPid $timestamp $mode $Language
  }

  $session = New-SessionObj $timestamp $mode $chunked $ffmpegPid $outputFile $logFile $wakeLockPid $liveTranscribePid
  $sessions = @(Get-StateSessions)
  $sessions += $session
  Set-StateSessions $sessions

  $plus = if ($chunked) { '+' } else { '' }
  Write-Output "rec demarre"
  Write-Output "- mode : ${mode}${plus}"
  Write-Output "- horodatage : $timestamp"
  Write-Output "- pid : $ffmpegPid"
  if ($liveTranscribePid -gt 0) { Write-Output "- transcription directe pid : $liveTranscribePid" }
  if ($title) { Write-Output "- titre : $title" }
  Write-Output "- sortie : $outputFile"
  Write-Output "- journal : $logFile"
  if ($chunked) { Write-Output '- note : utilise "/rec stop" pour finaliser et fusionner les chunks' }
  Write-Output '- astuce : utilise "/rec status" pour voir la progression'
}

function Stop-Pid($processId, $wakeLockPid, $liveTranscribePid) {
  if ($liveTranscribePid -gt 0) { taskkill /PID $liveTranscribePid /T /F 2>&1 | Out-Null }
  if ($wakeLockPid -gt 0) { taskkill /PID $wakeLockPid /T /F 2>&1 | Out-Null }
  taskkill /PID $processId /T 2>&1 | Out-Null
  Start-Sleep -Milliseconds 200
  for ($i = 0; $i -lt 10; $i++) { if (-not (Test-PidAlive $processId)) { break }; Start-Sleep -Milliseconds 200 }
  if (Test-PidAlive $processId) { taskkill /PID $processId /T /F 2>&1 | Out-Null }
}

function Invoke-TranscribeChunk($chunkFile, $language) {
  $ffmpeg = Get-FfmpegPath
  $whisperCli = Get-WhisperCliPath
  $model = Get-WhisperModelPath
  $inputForWhisper = $chunkFile.FullName
  $wavFile = $null

  $modelSize = (Get-Item $model).Length
  if ($modelSize -lt 400MB) { throw "Modele whisper incomplet (${modelSize} octets). Supprime le fichier et relance la transcription pour re-telecharger." }

  if ($chunkFile.Extension -in '.ts', '.mp4') {
    $wavFile = Join-Path $chunkFile.DirectoryName "$($chunkFile.BaseName)_audio.wav"
    $ffErr = & $ffmpeg -y -i $chunkFile.FullName -vn -ac 1 -ar 16000 $wavFile 2>&1
    if ($LASTEXITCODE -ne 0) { throw "Extraction audio a echoue pour $($chunkFile.Name) : $ffErr" }
    $inputForWhisper = $wavFile
  }

  $outBase = Join-Path $chunkFile.DirectoryName $chunkFile.BaseName
  $whisperErr = & $whisperCli -m $model -l $language -f $inputForWhisper -otxt -of $outBase 2>&1

  if ($wavFile -and (Test-Path $wavFile)) { Remove-Item $wavFile -Force }

  if ($LASTEXITCODE -ne 0) { throw "whisper-cli a echoue pour $($chunkFile.Name) : $whisperErr" }
}

function Read-TextFile($path) {
  if (Test-Path $path) { return (Get-Content $path -Raw -Encoding UTF8).Trim() }
  return ''
}

function Concat-Chunks($timestamp, $chunks, $finalFile) {
  $ffmpeg = Get-FfmpegPath
  if ($chunks.Count -eq 1) {
    & $ffmpeg -y -i $chunks[0].FullName -c copy -movflags +faststart $finalFile 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0 -and (Test-Path $finalFile) -and ((Get-Item $finalFile).Length -gt 0)) { return }
    Copy-Item $chunks[0].FullName $finalFile -Force
    return
  }

  $fileList = Join-Path (Split-Path $finalFile -Parent) "${timestamp}_filelist.txt"
  $lines = @($chunks | ForEach-Object { "file '$($_.FullName)'" })
  $lines -join "`n" | Set-Content $fileList -Encoding UTF8

  & $ffmpeg -y -f concat -safe 0 -i $fileList -c copy -movflags +faststart $finalFile 2>&1 | Out-Null
  if (Test-Path $fileList) { Remove-Item $fileList -Force }
  if ($LASTEXITCODE -ne 0 -and $chunks.Count -gt 0) {
    Copy-Item $chunks[0].FullName $finalFile -Force
  }
}

function Start-FinalizeAsync($session, $language) {
  $scriptPath = Join-Path (Get-Location) '.opencode\scripts\rec\rec.ps1'
  $tempFile = Join-Path $env:TEMP "rec_finalize_$([guid]::NewGuid().ToString('N')).ps1"
  $finalScript = @"
`$root = '$((Split-Path $session.output_file -Parent).Replace("'","''"))'
`$ts = '$($session.timestamp)'
`$mode = '$($session.mode)'
`$out = '$($session.output_file.Replace("'","''"))'
`$lang = '$($language)'

. '$($scriptPath.Replace("'","''"))' -BackgroundFinalize -Timestamp `$ts -Mode `$mode -OutputFile `$out -Language `$lang
Remove-Item "`$PSScriptRoot\`$([System.IO.Path]::GetFileName('$($tempFile)'))" -Force -ErrorAction SilentlyContinue
"@
  $finalScript | Set-Content $tempFile -Encoding UTF8
  $cmdLine = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$tempFile`""
  Start-DetachedProcess $cmdLine | Out-Null
}

function Do-FinalizeChunked($timestamp, $mode, $outputFile, $language) {
  $root = Split-Path $outputFile -Parent
  $chunks = Get-ChunkFiles $root $timestamp $mode
  if ($chunks.Count -eq 0) { Write-Output "Aucun chunk pour $timestamp"; return }
  $finalFile = $outputFile
  Concat-Chunks $timestamp $chunks $finalFile
  $transcriptParts = @()
  foreach ($chunk in $chunks) {
    $chunkTxt = [System.IO.Path]::ChangeExtension($chunk.FullName, '.txt')
    $text = ''
    if ((Test-Path $chunkTxt) -and ((Get-Item $chunkTxt).Length -gt 0)) {
      $text = Read-TextFile $chunkTxt
    } else {
      try { Invoke-TranscribeChunk $chunk $language; $text = Read-TextFile $chunkTxt } catch { $text = "[erreur transcription: $_]" }
    }
    if ($text) { $transcriptParts += $text }
  }
  $merged = $transcriptParts -join "`n`n"
  $txtOutput = [System.IO.Path]::ChangeExtension($finalFile, '.txt')
  $merged | Set-Content $txtOutput -Encoding UTF8
  if ((Test-Path $finalFile) -and ((Get-Item $finalFile).Length -gt 0)) {
    foreach ($chunk in $chunks) {
      Remove-Item $chunk.FullName -Force -ErrorAction SilentlyContinue
      $chunkTxt = [System.IO.Path]::ChangeExtension($chunk.FullName, '.txt')
      Remove-Item $chunkTxt -Force -ErrorAction SilentlyContinue
    }
  }
}

function Finalize-Chunked($session, $language) {
  $root = Split-Path $session.output_file -Parent
  $chunks = Get-ChunkFiles $root $session.timestamp $session.mode
  if ($chunks.Count -eq 0) { throw 'Aucun chunk trouve pour cette session.' }
  $finalFile = $session.output_file
  Concat-Chunks $session.timestamp $chunks $finalFile
  $transcriptParts = @()
  foreach ($chunk in $chunks) {
    $chunkTxt = [System.IO.Path]::ChangeExtension($chunk.FullName, '.txt')
    $text = ''
    if ((Test-Path $chunkTxt) -and ((Get-Item $chunkTxt).Length -gt 0)) {
      $text = Read-TextFile $chunkTxt
    } else {
      try { Invoke-TranscribeChunk $chunk $language; $text = Read-TextFile $chunkTxt } catch { $text = "[erreur transcription: $_]" }
    }
    if ($text) { $transcriptParts += $text }
  }
  $merged = $transcriptParts -join "`n`n"
  $txtOutput = [System.IO.Path]::ChangeExtension($finalFile, '.txt')
  $merged | Set-Content $txtOutput -Encoding UTF8
  if ((Test-Path $finalFile) -and ((Get-Item $finalFile).Length -gt 0)) {
    foreach ($chunk in $chunks) {
      Remove-Item $chunk.FullName -Force -ErrorAction SilentlyContinue
      $chunkTxt = [System.IO.Path]::ChangeExtension($chunk.FullName, '.txt')
      Remove-Item $chunkTxt -Force -ErrorAction SilentlyContinue
    }
  }
  return @{ file = $finalFile; transcript_file = $txtOutput; num_chunks = $chunks.Count }
}

function Stop-Recording($selector, $language) {
  $sessions = @(Get-StateSessions)
  if ($sessions.Count -eq 0) { throw 'Aucun enregistrement actif.' }

  $targets = @()
  if (-not $selector) { $targets = @($sessions[0]) }
  elseif ($selector -eq 'all') { $targets = $sessions }
  elseif ($selector -match '^\d+$') {
    $idx = [int]$selector
    if ($idx -lt 1 -or $idx -gt $sessions.Count) { throw "Index invalide: $idx" }
    $targets = @($sessions[$idx - 1])
  } else { throw "Selecteur invalide. Utilise: stop, stop N, ou stop all" }

  $results = @()
  foreach ($target in $targets) {
    Stop-Pid $target.pid $target.wake_lock_pid $target.live_transcribe_pid
    Start-Sleep -Seconds 1
    $result = @{ timestamp = $target.timestamp; mode = $target.mode; chunked = $target.chunked }
    if ($target.chunked) {
      $result.file = $target.output_file
      Start-FinalizeAsync $target $language
      $result.background = $true
    } else {
      $output = $target.output_file
      if (-not (Test-Path $output) -or ((Get-Item $output).Length -eq 0)) { throw "Fichier de sortie absent ou vide: $output" }
      $result.file = $output
    }
    $results += $result
  }

  $remaining = @($sessions | Where-Object { $_.timestamp -notin ($targets | ForEach-Object { $_.timestamp }) })
  Set-StateSessions $remaining

  Write-Output 'rec arret'
  foreach ($result in $results) {
    $fileStr = $result.file
    $plus = if ($result.chunked) { '+' } else { '' }
    Write-Output "- $($result.timestamp) ($($result.mode)$plus) -> $fileStr"
    if ($result.background) { Write-Output "  (fusion+transcription en cours, fichier et .txt bientot disponibles)" }
  }
}

function Get-Status {
  $sessions = @(Get-StateSessions)
  if ($sessions.Count -eq 0) { Write-Output 'aucun enregistrement actif'; return }

  Write-Output "enregistrement(s) actif(s) ($($sessions.Count))"
  for ($i = 0; $i -lt $sessions.Count; $i++) {
    $session = $sessions[$i]
    $plus = if ($session.chunked) { '+' } else { '' }
    Write-Output "$($i+1). $(Format-Timestamp $session.timestamp) | $($session.mode)$plus | pid=$($session.pid) | $($session.timestamp)"

    if (-not $session.chunked) { Write-Output '   transcription : n/a (non decoupe)'; continue }

    $progress = Get-SessionTranscriptionProgress $session
    if ($progress.chunks_detected -eq 0) { Write-Output '   transcription : attente du premier chunk...'; continue }
    Write-Output "   transcription : $($progress.chunks_transcribed)/$($progress.chunks_detected) chunks ($($progress.percent)%) | dernier chunk : $($progress.last_chunk_index)"
  }
}

function Invoke-TranscribeFile($file, $language) {
  if (-not $file) {
    $items = Get-RecordingList
    if ($items.Count -eq 0) { throw 'Aucun enregistrement disponible.' }
    $latest = $items[0]
    $file = if ($latest.video) { $latest.video } elseif ($latest.audio) { $latest.audio } else { throw 'Aucun fichier media' }
  }

  if (-not (Test-Path $file)) { throw "Fichier introuvable: $file" }

  $src = Get-Item $file
  $whisperCli = Get-WhisperCliPath
  $model = Get-WhisperModelPath
  $ffmpeg = Get-FfmpegPath

  $modelSize = (Get-Item $model).Length
  if ($modelSize -lt 400MB) { throw "Modele whisper incomplet (${modelSize} octets). Supprime le fichier et relance la transcription pour re-telecharger." }

  $whisperInput = $src.FullName
  $tempWav = $null

  if ($src.Extension -in '.mp4', '.ts', '.mov', '.mkv', '.webm') {
    Write-Output "[transcribe] extraction audio de $($src.Name)..."
    $tempWav = Join-Path $src.DirectoryName "$($src.BaseName)_audio.wav"
    $ffErr = & $ffmpeg -y -i $src.FullName -vn -ac 1 -ar 16000 $tempWav 2>&1
    if ($LASTEXITCODE -ne 0) { throw "Extraction audio a echoue : $ffErr" }
    $whisperInput = $tempWav
  }

  $inputName = Split-Path $whisperInput -Leaf
  Write-Output "[transcribe] lancement Whisper sur $inputName..."
  Write-Output "[transcribe] cette operation peut prendre plusieurs minutes..."

  $outBase = [System.IO.Path]::Combine($src.DirectoryName, $src.BaseName)
  & $whisperCli -m $model -l $language -f $whisperInput -otxt -of $outBase

  if ($tempWav -and (Test-Path $tempWav)) { Remove-Item $tempWav -Force }

  if ($LASTEXITCODE -ne 0) { throw "whisper-cli a echoue avec le code $LASTEXITCODE" }

  $txtPath = [System.IO.Path]::ChangeExtension($src.FullName, '.txt')
  if (-not (Test-Path $txtPath)) { throw "Fichier de transcription non trouve: $txtPath" }

  $text = Read-TextFile $txtPath
  Write-Output "rec transcription"
  Write-Output "- source : $($src.FullName)"
  Write-Output "- fichier texte : $txtPath"
  Write-Output "- caracteres : $($text.Length)"
}

function Get-RecordingList {
  $root = $Script:RecordingsDir
  if (-not (Test-Path $root)) { return @() }

  $grouped = @{}
  foreach ($path in (Get-ChildItem $root -File)) {
    if ($path.Extension -notin '.mp3', '.mp4', '.txt') { continue }
    if ($path.Name -match '_chunk_') { continue }
    if ($path.BaseName -match '^(\d{8}_\d{6})(?:-[a-z0-9][a-z0-9-]{0,80})?$') {
      $ts = $matches[1]
      $stem = $path.BaseName
      if (-not $grouped.ContainsKey($stem)) { $grouped[$stem] = @{ timestamp = $ts } }
      if ($path.Extension -eq '.mp3') { $grouped[$stem].audio = $path.FullName }
      elseif ($path.Extension -eq '.mp4') { $grouped[$stem].video = $path.FullName }
      else { $grouped[$stem].txt = $path.FullName }
    }
  }

  return @($grouped.Values | Sort-Object timestamp -Descending)
}

function List-Recordings {
  $items = Get-RecordingList
  if ($items.Count -eq 0) { Write-Output 'aucun enregistrement trouve'; return }

  Write-Output "enregistrements ($($items.Count)) dans $($Script:RecordingsDir)"
  for ($i = 0; $i -lt $items.Count; $i++) {
    $item = $items[$i]
    $parts = @()
    if ($item.video) { $parts += 'video' }
    if ($item.audio) { $parts += 'audio' }
    if ($item.txt) { $parts += 'texte' }
    Write-Output "$($i+1). $(Format-Timestamp $item.timestamp) | $($parts -join ', ')"
  }
}

function Find-OrphanedChunks {
  $root = $Script:RecordingsDir
  if (-not (Test-Path $root)) { return @{} }

  $groups = @{}
  foreach ($path in (Get-ChildItem $root -File)) {
    if ($path.Name -match '^(\d{8}_\d{6})_chunk_(\d{3})\.(mp3|mp4|ts)$') {
      $ts = $matches[1]
      $ext = $matches[3]
      $finalExt = if ($ext -in 'mp4', 'ts') { 'mp4' } else { 'mp3' }
      $hasFinal = $false
      foreach ($candidate in (Get-ChildItem (Join-Path $root "${ts}*.${finalExt}") -ErrorAction SilentlyContinue)) {
        if ($candidate.Name -notmatch '_chunk_') { $hasFinal = $true; break }
      }
      if (-not $hasFinal) {
        if (-not $groups.ContainsKey($ts)) { $groups[$ts] = @() }
        $groups[$ts] += $path
      }
    }
  }

  $sorted = @{}
  foreach ($key in ($groups.Keys | Sort-Object -Descending)) { $sorted[$key] = @($groups[$key] | Sort-Object Name) }
  return $sorted
}

function Recover-Chunks($selector, $language) {
  $groups = Find-OrphanedChunks
  if ($groups.Count -eq 0) {
    Write-Output 'aucun chunk orphelin'
    return
  }

  $keys = @($groups.Keys | Sort-Object -Descending)

  if (-not $selector) {
    Write-Output "groupes de chunks orphelins ($($groups.Count))"
    for ($i = 0; $i -lt $keys.Count; $i++) {
      $ts = $keys[$i]
      $chunks = $groups[$ts]
      $ext = if ($chunks.Count -gt 0) { $chunks[0].Extension.TrimStart('.') } else { '?' }
      Write-Output "$($i+1). $ts | $($chunks.Count) chunks | $ext"
    }
    Write-Output "utilise : rec recover N ou rec recover all"
    return
  }

  $targets = @()
  if ($selector -eq 'all') { $targets = $keys }
  elseif ($selector -match '^\d+$') {
    $idx = [int]$selector
    if ($idx -lt 1 -or $idx -gt $keys.Count) { throw "Index invalide: $idx" }
    $targets = @($keys[$idx - 1])
  } else { throw "Selecteur invalide. Utilise: recover, recover N, ou recover all" }

  if ($targets.Count -eq 0) { Write-Output 'Aucune cible.'; return }

  $results = @()
  foreach ($ts in $targets) {
    $chunks = $groups[$ts]
    $mode = if ($chunks[0].Extension -in '.mp4', '.ts') { 'video' } else { 'audio' }
    $finalExt = if ($mode -eq 'video') { '.mp4' } else { '.mp3' }
    $finalFile = Join-Path (Split-Path $chunks[0].FullName -Parent) "${ts}${finalExt}"

    Concat-Chunks $ts $chunks $finalFile

    $transcripts = @()
    foreach ($chunk in $chunks) {
      $chunkTxt = [System.IO.Path]::ChangeExtension($chunk.FullName, '.txt')
      $text = ''
      if ((Test-Path $chunkTxt) -and ((Get-Item $chunkTxt).Length -gt 0)) {
        $text = Read-TextFile $chunkTxt
      } else {
        try { Invoke-TranscribeChunk $chunk $language; $text = Read-TextFile $chunkTxt } catch { $text = "[erreur: $_]" }
      }
      if ($text) { $transcripts += $text }
    }

    $merged = $transcripts -join "`n`n"
    $txtOutput = [System.IO.Path]::ChangeExtension($finalFile, '.txt')
    $merged | Set-Content $txtOutput -Encoding UTF8

    if ((Test-Path $finalFile) -and ((Get-Item $finalFile).Length -gt 0)) {
      foreach ($chunk in $chunks) {
        Remove-Item $chunk.FullName -Force -ErrorAction SilentlyContinue
        $chunkTxt = [System.IO.Path]::ChangeExtension($chunk.FullName, '.txt')
        Remove-Item $chunkTxt -Force -ErrorAction SilentlyContinue
      }
    }

    $results += @{ timestamp = $ts; file = $finalFile; transcript_file = $txtOutput; num_chunks = $chunks.Count }
  }

    Write-Output "groupes recuperes : $($results.Count)"
  foreach ($item in $results) {
    Write-Output "- $($item.timestamp) | $($item.num_chunks) chunks -> $($item.file)"
    Write-Output "  transcription : $($item.transcript_file)"
  }
}

function Clean-Chunks($allDirs) {
  $root = $Script:RecordingsDir
  $totalRemoved = 0
  $groupsCleaned = 0
  $dirsScanned = 0

  $dirsToScan = if ($allDirs) {
    @(Get-ChildItem (Split-Path $Script:RecordingsDir -Parent) -Directory -Filter 'recordings' -Recurse -ErrorAction SilentlyContinue)
  } else { @($root) }

  foreach ($dir in $dirsToScan) {
    if (-not (Test-Path $dir)) { continue }
    $dirsScanned++
    $chunksByTs = @{}
    foreach ($path in (Get-ChildItem $dir -File)) {
      if ($path.Name -match '^(\d{8}_\d{6})_chunk_(\d{3})\.(mp3|mp4|ts)$') {
        $ts = $matches[1]
        if (-not $chunksByTs.ContainsKey($ts)) { $chunksByTs[$ts] = @() }
        $chunksByTs[$ts] += $path
      }
    }
    foreach ($entry in $chunksByTs.GetEnumerator()) {
      $chunks = $entry.Value
      $ext = $chunks[0].Extension.TrimStart('.')
      $finalExt = if ($ext -in 'mp4', 'ts') { 'mp4' } else { 'mp3' }
      $finalFile = $null
      foreach ($candidate in (Get-ChildItem (Join-Path $dir "${ts}*.${finalExt}") -ErrorAction SilentlyContinue)) {
        if ($candidate.Name -notmatch '_chunk_') { $finalFile = $candidate; break }
      }
      if ($finalFile -and ((Get-Item $finalFile).Length -gt 0)) {
        foreach ($chunk in $chunks) {
          Remove-Item $chunk.FullName -Force -ErrorAction SilentlyContinue
          $txtFile = [System.IO.Path]::ChangeExtension($chunk.FullName, '.txt')
          Remove-Item $txtFile -Force -ErrorAction SilentlyContinue
          $wavFile = Join-Path $chunk.DirectoryName "$($chunk.BaseName)_audio.wav"
          Remove-Item $wavFile -Force -ErrorAction SilentlyContinue
        }
        $totalRemoved += $chunks.Count
        $groupsCleaned++
      }
    }
  }

  if ($totalRemoved -eq 0) { Write-Output 'aucun chunk a nettoyer'; return }
  Write-Output "rec nettoyage"
  if ($allDirs) { Write-Output "- repertoires parcourus : $dirsScanned" }
  Write-Output "- groupes nettoyes : $groupsCleaned"
  Write-Output "- chunks supprimes : $totalRemoved"
}

function Start-WatchTranscript($ffmpegPid, $timestamp, $mode, $language) {
  $root = $Script:RecordingsDir
  $processed = @{}
  $lastNewChunkAt = (Get-Date)
  $pollSeconds = 3
  $minChunkAgeSeconds = 2
  $idleExitSeconds = 20

  [NativeMethods]::SetThreadExecutionState([uint32]0x80000001) | Out-Null

  try {
    while ($true) {
      $chunks = Get-ChunkFiles $root $timestamp $mode
      $now = Get-Date
      $newChunkSeen = $false

      foreach ($chunk in $chunks) {
        $key = $chunk.FullName
        if ($processed.ContainsKey($key)) { continue }
        $age = ($now - $chunk.LastWriteTime).TotalSeconds
        if ($age -lt $minChunkAgeSeconds) { continue }

        $chunkTxt = [System.IO.Path]::ChangeExtension($chunk.FullName, '.txt')
        if (-not (Test-Path $chunkTxt) -or ((Get-Item $chunkTxt).Length -eq 0)) {
          try { Invoke-TranscribeChunk $chunk $language } catch { continue }
        }

        $processed[$key] = $true
        $newChunkSeen = $true
      }

      if ($newChunkSeen) {
        $lastNewChunkAt = $now
        $chunkTexts = @()
        foreach ($chunk in $chunks) {
          $chunkTxt = [System.IO.Path]::ChangeExtension($chunk.FullName, '.txt')
          if ((Test-Path $chunkTxt) -and ((Get-Item $chunkTxt).Length -gt 0)) {
            $chunkTexts += Read-TextFile $chunkTxt
          }
        }
        $progressPath = Join-Path $root "${timestamp}_progress.txt"
        ($chunkTexts -join "`n`n") | Set-Content $progressPath -Encoding UTF8
      }

      [NativeMethods]::SetThreadExecutionState([uint32]0x80000001) | Out-Null

      if (-not (Test-PidAlive $ffmpegPid)) {
        if (($now - $lastNewChunkAt).TotalSeconds -ge $idleExitSeconds) { break }
      }

      Start-Sleep -Seconds $pollSeconds
    }
  } finally {
    [NativeMethods]::SetThreadExecutionState([uint32]0x80000000) | Out-Null
  }
}

function Show-Usage {
  Write-Output @'
Utilisation : /rec [commande] [options]

Commandes :
  help                    Afficher cette aide (/rec sans argument fait pareil)
  video [--titre T]       Enregistrement ecran (non decoupe)
  video+ [--titre T]      Enregistrement ecran + transcription directe
  audio [--titre T]       Enregistrement audio (non decoupe)
  audio+ [--titre T]      Enregistrement audio + transcription directe
  stop [--langue fr]      Arreter l'enregistrement et finaliser
  status                  Afficher la progression
  transcribe [--fichier F] Transcrire un fichier
  devices                 Lister les peripheriques audio
  list                    Lister les enregistrements
  recover [N|all]         Recuperer des chunks orphelins
  clean [--all]           Nettoyer les chunks complets
  install                 Lancer l'installation des dependances en arriere-plan
  install-status          Afficher la progression de l'installation
  install --window        Lancer l'installation dans une fenetre PowerShell
'@
}

function Download-FileWithProgress {
  param(
    [string]$Uri,
    [string]$OutFile,
    [string]$Label
  )

  $parent = Split-Path -Parent $OutFile
  if ($parent -and -not (Test-Path $parent)) { New-Item -ItemType Directory -Force -Path $parent | Out-Null }

  Write-Output "$Label : demarrage du telechargement"
  Write-Output "$Label : source=$Uri"
  Write-Output "$Label : destination=$OutFile"

  $request = [System.Net.HttpWebRequest]::Create($Uri)
  $request.UserAgent = 'opencode-rec-installer'
  $response = $request.GetResponse()
  try {
    $total = [int64]$response.ContentLength
    if ($total -gt 0) { Write-Output "$Label : taille=$total octets" }

    $inputStream = $response.GetResponseStream()
    $outputStream = [System.IO.File]::Open($OutFile, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    try {
      $buffer = New-Object byte[] (1024 * 1024)
      $downloaded = [int64]0
      $nextLog = [int64](10MB)
      while (($read = $inputStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
        $outputStream.Write($buffer, 0, $read)
        $downloaded += $read
        if ($downloaded -ge $nextLog) {
          if ($total -gt 0) {
            $percent = [math]::Round(($downloaded * 100.0) / $total, 1)
            Write-Output "$Label : $downloaded / $total octets ($percent%)"
          } else {
            Write-Output "$Label : $downloaded octets telecharges"
          }
          $nextLog = $downloaded + [int64](10MB)
        }
      }
      Write-Output "$Label : telechargement termine ($downloaded octets)"
    } finally {
      if ($outputStream) { $outputStream.Dispose() }
      if ($inputStream) { $inputStream.Dispose() }
    }
  } finally {
    if ($response) { $response.Dispose() }
  }
}

function Install-Dependencies {
  $ErrorActionPreference = 'Stop'

  Write-Output '=== Installation /rec ==='
  Write-Output 'Verification des dependances: ffmpeg, whisper-cli, modele whisper, Stereo Mix.'

  $toolsRoot = Join-Path $env:LOCALAPPDATA 'opencode-tools'
  $whisperDir = Join-Path $toolsRoot 'whisper.cpp'
  $modelDir = Join-Path $whisperDir 'models'
  $modelFile = Join-Path $modelDir 'ggml-small.bin'

  Write-Output "tools_root=$toolsRoot"

  $ffmpegExe = $null
  Write-Output ''
  Write-Output '[1/4] Verification de ffmpeg...'
  $ffmpegCmd = Get-Command ffmpeg -ErrorAction SilentlyContinue
  if ($ffmpegCmd) {
    $ffmpegExe = $ffmpegCmd.Source
    Write-Output 'ffmpeg deja disponible dans le PATH.'
  }

  if (-not $ffmpegExe) {
    Write-Output 'ffmpeg absent du PATH. Recherche dans les packages winget locaux...'
    $candidate = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\WinGet\Packages" -Recurse -Filter ffmpeg.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
    if ($candidate) {
      $ffmpegExe = $candidate
      Write-Output 'ffmpeg trouve dans les packages winget locaux.'
    }
  }

  if (-not $ffmpegExe) {
    Write-Output 'Installation de ffmpeg via winget...'
    winget install -e --id Gyan.FFmpeg --accept-source-agreements --accept-package-agreements
    Write-Output 'winget termine. Rafraichissement du PATH...'
    $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User')
    $ffmpegCmd = Get-Command ffmpeg -ErrorAction SilentlyContinue
    if ($ffmpegCmd) { $ffmpegExe = $ffmpegCmd.Source }
    if (-not $ffmpegExe) {
      Write-Output 'ffmpeg non trouve dans le PATH. Nouvelle recherche dans les packages winget locaux...'
      $candidate = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\WinGet\Packages" -Recurse -Filter ffmpeg.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
      if ($candidate) { $ffmpegExe = $candidate }
    }
  }

  if (-not $ffmpegExe) { throw 'ffmpeg introuvable apres installation' }
  Write-Output "ffmpeg=$ffmpegExe"

  Write-Output ''
  Write-Output '[2/4] Verification de whisper-cli...'
  $whisperCli = Join-Path $whisperDir 'Release\whisper-cli.exe'
  if (-not (Test-Path $whisperCli)) {
    Write-Output 'whisper-cli absent. Telechargement de whisper-cli...'
    New-Item -ItemType Directory -Force -Path $whisperDir | Out-Null
    $zip = Join-Path $toolsRoot 'whisper-bin-x64.zip'
    Download-FileWithProgress -Uri 'https://github.com/ggml-org/whisper.cpp/releases/latest/download/whisper-bin-x64.zip' -OutFile $zip -Label 'whisper-cli'
    Write-Output 'Extraction de whisper-cli...'
    Expand-Archive -LiteralPath $zip -DestinationPath $whisperDir -Force
    Remove-Item $zip -Force
    $found = Get-ChildItem $whisperDir -Recurse -Filter 'whisper-cli.exe' -File | Where-Object { $_.FullName -match '\\Release\\' } | Select-Object -First 1 -ExpandProperty FullName
    if (-not $found) { throw 'whisper-cli.exe introuvable apres extraction' }
    $whisperCli = $found
  } else {
    Write-Output 'whisper-cli deja installe.'
  }
  Write-Output "whisper-cli=$whisperCli"

  Write-Output ''
  Write-Output '[3/4] Verification du modele whisper ggml-small.bin...'
  if ((Test-Path $modelFile) -and ((Get-Item $modelFile).Length -lt 400MB)) {
    Write-Output 'Modele present mais incomplet. Suppression puis nouveau telechargement.'
    Remove-Item $modelFile -Force
  }

  if (-not (Test-Path $modelFile)) {
    Write-Output 'Modele absent. Telechargement du modele ggml-small.bin...'
    Write-Output 'Cette etape peut prendre plusieurs minutes selon la connexion.'
    New-Item -ItemType Directory -Force -Path $modelDir | Out-Null
    Download-FileWithProgress -Uri 'https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-small.bin' -OutFile $modelFile -Label 'modele whisper'
  } else {
    Write-Output 'Modele deja present.'
  }
  if (-not (Test-Path $modelFile)) { throw 'Modele introuvable apres telechargement' }
  Write-Output "modele=$modelFile"

  $modelSize = (Get-Item $modelFile).Length
  if ($modelSize -lt 400MB) { throw "Modele whisper incomplet: $modelFile ($modelSize octets). Supprime ce fichier puis relance /rec install." }
  Write-Output "modele_taille=$($modelSize) octets"

  Write-Output ''
  Write-Output '[4/4] Verification de Stereo Mix...'
  $output = & cmd.exe /d /c "`"$ffmpegExe`" -hide_banner -list_devices true -f dshow -i dummy 2>&1" | Out-String
  $hasStereoMix = $output -match '(?i)stereo mix|mixage stereo|what u hear'
  if ($hasStereoMix) {
    Write-Output 'stereo_mix=detecte'
  } else {
    Write-Output 'stereo_mix=absent'
    Write-Output ''
    Write-Output '=== Stereo Mix non detecte ==='
    Write-Output 'Pour enregistrer l audio systeme :'
    Write-Output 'Modifier les sons système > Enregistrement > Stéréo Mix > Bouton droit > Activer > OK'
  }

  Write-Output ''
  Write-Output 'Installation terminee.'
  Write-Output 'Redemarre OpenCode Desktop pour charger la commande /rec.'
}

function Start-InstallAsync($openWindow) {
  New-Item -ItemType Directory -Force -Path $Script:ToolsRoot | Out-Null

  if (Test-Path $Script:InstallPidFile) {
    $oldPidText = (Get-Content $Script:InstallPidFile -Raw -ErrorAction SilentlyContinue).Trim()
    if ($oldPidText -match '^\d+$' -and (Test-PidAlive ([int]$oldPidText))) {
      Write-Output 'Installation /rec deja en cours.'
      Write-Output "- pid : $oldPidText"
      Write-Output "- log : $Script:InstallLogFile"
      Write-Output 'Utilise /rec install-status pour suivre la progression.'
      return
    }
  }

  $scriptPath = Join-Path (Get-Location) '.opencode\scripts\rec\rec.ps1'
  $escapedScript = $scriptPath.Replace("'", "''")
  $escapedLog = $Script:InstallLogFile.Replace("'", "''")
  $escapedPid = $Script:InstallPidFile.Replace("'", "''")
  $cmd = @"
`$ErrorActionPreference = 'Stop'
try {
  & '$escapedScript' -Command install-run *>&1 | Tee-Object -FilePath '$escapedLog'
} catch {
  `$message = "INSTALLATION /rec EN ERREUR : `$(`$_.Exception.Message)"
  `$message | Tee-Object -FilePath '$escapedLog' -Append
  exit 1
} finally {
  Remove-Item '$escapedPid' -Force -ErrorAction SilentlyContinue
}
"@
  $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($cmd))

  if ($openWindow) {
    $process = Start-Process powershell.exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-NoExit', '-EncodedCommand', $encoded) -PassThru
    $installPid = $process.Id
  } else {
    $installPid = Start-DetachedProcess "powershell.exe -NoProfile -ExecutionPolicy Bypass -EncodedCommand $encoded"
  }

  Set-Content $Script:InstallPidFile $installPid -Encoding UTF8

  Write-Output 'Installation /rec lancee.'
  Write-Output "- pid : $installPid"
  Write-Output "- log : $Script:InstallLogFile"
  if ($openWindow) { Write-Output '- affichage : fenetre PowerShell separee' }
  Write-Output 'Utilise /rec install-status pour suivre la progression dans OpenCode.'
}

function Show-InstallStatus {
  $pidText = ''
  if (Test-Path $Script:InstallPidFile) { $pidText = (Get-Content $Script:InstallPidFile -Raw -ErrorAction SilentlyContinue).Trim() }
  $running = $false
  if ($pidText -match '^\d+$') { $running = Test-PidAlive ([int]$pidText) }

  Write-Output 'Statut installation /rec'
  if ($pidText) { Write-Output "- pid : $pidText" }
  Write-Output "- en cours : $running"
  Write-Output "- log : $Script:InstallLogFile"

  $ffmpeg = Get-Command ffmpeg -ErrorAction SilentlyContinue
  if ($ffmpeg) { Write-Output "- ffmpeg : $($ffmpeg.Source)" } else { Write-Output '- ffmpeg : absent du PATH courant' }

  $whisperCli = Join-Path $Script:WhisperDir 'Release\whisper-cli.exe'
  if (Test-Path $whisperCli) { Write-Output "- whisper-cli : $whisperCli" } else { Write-Output '- whisper-cli : absent' }

  $modelFile = Join-Path $Script:WhisperDir 'models\ggml-small.bin'
  if (Test-Path $modelFile) {
    $size = (Get-Item $modelFile).Length
    $percent = [math]::Round(($size * 100.0) / 487601967, 1)
    Write-Output "- modele : $size / 487601967 octets ($percent%)"
  } else {
    Write-Output '- modele : absent'
  }

  if (Test-Path $Script:InstallLogFile) {
    Write-Output ''
    Write-Output 'Dernieres lignes du log :'
    $lines = @(Get-Content $Script:InstallLogFile -Encoding UTF8 -ErrorAction SilentlyContinue)
    foreach ($line in ($lines | Select-Object -Last 80)) { Write-Output $line }
  } else {
    Write-Output ''
    Write-Output 'Aucun log disponible.'
  }
}

if ($Help) { Show-Usage; return }

if ($WatchTranscript) {
  if ($Language -eq '') { $Language = 'fr' }
  Start-WatchTranscript $FfmpegPid $Timestamp $Mode $Language
  return
}

if ($BackgroundFinalize) {
  Do-FinalizeChunked $Timestamp $Mode $OutputFile $Language
  return
}

if ($RawArgs) {
  $tokens = @()
  $current = ''
  $inQuote = $false
  $quoteChar = ''
  foreach ($ch in $RawArgs.ToCharArray()) {
    if ($inQuote) {
      if ($ch -eq $quoteChar) { $inQuote = $false; continue }
      $current += $ch
    } elseif ($ch -in '"', "'") {
      $inQuote = $true
      $quoteChar = $ch
    } elseif ($ch -eq ' ') {
      if ($current) { $tokens += $current; $current = '' }
    } else {
      $current += $ch
    }
  }
  if ($current) { $tokens += $current }

  if ($tokens.Count -gt 0) {
    if ($tokens[0] -eq '--help') { $Command = 'help' }
    else { $Command = $tokens[0] }
    $i = 1
    while ($i -lt $tokens.Count) {
      switch -Wildcard ($tokens[$i]) {
        '--title' { $i++; if ($i -lt $tokens.Count) { $Title = $tokens[$i] } }
        '--titre' { $i++; if ($i -lt $tokens.Count) { $Title = $tokens[$i] } }
        '--file'  { $i++; if ($i -lt $tokens.Count) { $File = $tokens[$i] } }
        '--fichier' { $i++; if ($i -lt $tokens.Count) { $File = $tokens[$i] } }
        '--language' { $i++; if ($i -lt $tokens.Count) { $Language = $tokens[$i] } }
        '--langue' { $i++; if ($i -lt $tokens.Count) { $Language = $tokens[$i] } }
        '--all'   { $All = $true }
        default   { if (-not $Selector) { $Selector = $tokens[$i] } }
      }
      $i++
    }
  }
}

if (-not $Command -and $args.Count -gt 0) { $Command = $args[0] }

switch ($Command) {
  '' { Show-Usage; return }
  'help' { Show-Usage; return }
  'install' { Start-InstallAsync ($Selector -eq '--window') }
  'install-run' { Install-Dependencies }
  'install-status' { Show-InstallStatus }
  'video' { Start-Recording 'video' $false $Title }
  'video+' { Start-Recording 'video' $true $Title }
  'audio' { Start-Recording 'audio' $false $Title }
  'audio+' { Start-Recording 'audio' $true $Title }
  'stop' { Stop-Recording $Selector $Language }
  'status' { Get-Status }
  'devices' { Show-Devices }
  'transcribe' { Invoke-TranscribeFile $File $Language }
  'list' { List-Recordings }
  'recover' { Recover-Chunks $Selector $Language }
  'clean' { Clean-Chunks $All }
  default { Write-Output "Commande inconnue: $Command"; Show-Usage; return 1 }
}
