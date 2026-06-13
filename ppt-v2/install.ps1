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

function Copy-DirectoryRecursive {
    param([string]$SourceRelativeDir, [string]$DestDir)
    if ($PSScriptRoot) {
        $localDir = Join-Path $PSScriptRoot ($SourceRelativeDir -replace '^ppt-v2/', '')
        if (Test-Path -LiteralPath $localDir) {
            if (-not (Test-Path -LiteralPath $DestDir)) { New-Item -ItemType Directory -Path $DestDir | Out-Null }
            foreach ($item in Get-ChildItem -LiteralPath $localDir) {
                $destItem = Join-Path $DestDir $item.Name
                if ($item.PSIsContainer) { Copy-DirectoryRecursive -SourceRelativeDir ($SourceRelativeDir + "/" + $item.Name) -DestDir $destItem }
                else { Copy-Item -LiteralPath $item.FullName -Destination $destItem -Force }
            }
            return
        }
    }
    $urlDir = $SourceRelativeDir -replace '\\', '/'
    $metaUrl = "$RepositoryRawBase/$Ref/$urlDir"
    try {
        $listing = Invoke-WebRequest -Uri $metaUrl -UseBasicParsing
        Write-Output "Telechargement repertoire via archive..."
    }
    catch {
        $branch = $Ref
        $archiveUrl = "$RepositoryRawBase/$branch/ppt-v2-icons.zip"
        try {
            Invoke-WebRequest -Uri $archiveUrl -OutFile "$env:TEMP\ppt-v2-icons.zip"
            Expand-Archive -Path "$env:TEMP\ppt-v2-icons.zip" -DestinationPath $DestDir -Force
        } catch { throw "Impossible de telecharger les icones" }
    }
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
    @{ Source = "ppt-v2/.opencode/scripts/ppt-v2/generate.py"; Destination = ".opencode/scripts/ppt-v2/generate.py" },
    @{ Source = "ppt-v2/.opencode/scripts/ppt-v2/theme.py"; Destination = ".opencode/scripts/ppt-v2/theme.py" },
    @{ Source = "ppt-v2/.opencode/scripts/ppt-v2/card.py"; Destination = ".opencode/scripts/ppt-v2/card.py" },
    @{ Source = "ppt-v2/.opencode/scripts/ppt-v2/icons.py"; Destination = ".opencode/scripts/ppt-v2/icons.py" },
    @{ Source = "ppt-v2/.opencode/scripts/ppt-v2/footer.py"; Destination = ".opencode/scripts/ppt-v2/footer.py" },
    @{ Source = "ppt-v2/.opencode/scripts/ppt-v2/conclusion.py"; Destination = ".opencode/scripts/ppt-v2/conclusion.py" },
    @{ Source = "ppt-v2/.opencode/scripts/ppt-v2/decor.py"; Destination = ".opencode/scripts/ppt-v2/decor.py" },
    @{ Source = "ppt-v2/.opencode/scripts/ppt-v2/layouts.py"; Destination = ".opencode/scripts/ppt-v2/layouts.py" },
    @{ Source = "ppt-v2/.opencode/scripts/ppt-v2/slidebuilder.py"; Destination = ".opencode/scripts/ppt-v2/slidebuilder.py" },
    @{ Source = "ppt-v2/.opencode/scripts/ppt-v2/parser.py"; Destination = ".opencode/scripts/ppt-v2/parser.py" },
    @{ Source = "ppt-v2/.opencode/scripts/ppt-v2/quality.py"; Destination = ".opencode/scripts/ppt-v2/quality.py" },
    @{ Source = "ppt-v2/.opencode/scripts/ppt-v2/__init__.py"; Destination = ".opencode/scripts/ppt-v2/__init__.py" },
    @{ Source = "ppt-v2/.opencode/scripts/ppt-v2/slides.example.json"; Destination = ".opencode/scripts/ppt-v2/slides.example.json" },
    @{ Source = "ppt-v2/icons"; Destination = ".opencode/scripts/ppt-v2/icons" },
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
Write-Output "Fichiers Python installes: 12 modules + icones"
Write-Output ""

$pythonPaths = @(
    "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
    "${env:ProgramFiles}\Python312\python.exe"
)
$found = $false
foreach ($p in $pythonPaths) {
    if (Test-Path $p) {
        Write-Output "Python trouve: $p"
        & $p -m pip show python-pptx 2>$null | Out-Null
        if ($LASTEXITCODE -eq 0) {
            Write-Output "python-pptx deja installe"
        } else {
            Write-Output "Installation de python-pptx..."
            & $p -m pip install python-pptx 2>&1
        }
        $found = $true
        break
    }
}
if (-not $found) {
    Write-Output "ATTENTION: Python 3.12 introuvable. Installez Python 3.12+ puis: pip install python-pptx"
    Write-Output "Telechargement: https://www.python.org/downloads/"
}

Write-Output ""
Write-Output "Pour utiliser la commande, redemarrez OpenCode Desktop depuis ce projet."
Write-Output "Exemple: /ppt-v2 Genere 3 slides sur l'IA generative"
Write-Output "         /ppt-v2 --strict Genere 5 slides sur la data"
