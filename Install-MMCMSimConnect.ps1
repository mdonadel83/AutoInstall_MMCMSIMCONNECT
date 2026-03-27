<#
.SYNOPSIS
    MMCM SimConnect - Installer automatico da GitHub v2.3
.DESCRIPTION
    Scarica e installa il plugin MMCM SimConnect in SimHub.
    Copia l'intera struttura da Release/net48 (sottocartelle, exe, dll, tutto).
    USA download ZIP diretto (NO API, NO rate limit).
.NOTES
    Repository: https://github.com/mdonadel83/MMCMSimConnect
#>

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"

# ============================================
# CONFIGURAZIONE
# ============================================
$RepoOwner = "mdonadel83"
$RepoName = "MMCMSimConnect"
$BranchName = "main"
$SourcePath = "Release\net48"

# Download diretto (non usa l'API, nessun rate limit)
$ZipDownloadUrl = "https://github.com/$RepoOwner/$RepoName/archive/refs/heads/$BranchName.zip"

$TempFolder = Join-Path $env:TEMP "MMCMSimConnect_Install"

# ============================================
# FUNZIONI
# ============================================

function Write-Header {
    Clear-Host
    Write-Host ""
    Write-Host "  ==========================================================" -ForegroundColor Cyan
    Write-Host "  |     MMCM SimConnect - Installer v2.3 (Full Tree)       |" -ForegroundColor Cyan
    Write-Host "  |     Plugin per SimHub - No Rate Limit                  |" -ForegroundColor Cyan
    Write-Host "  ==========================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Write-Step {
    param([string]$Message)
    Write-Host "  [*] $Message" -ForegroundColor Yellow
}

function Write-Success {
    param([string]$Message)
    Write-Host "  [OK] $Message" -ForegroundColor Green
}

function Write-ErrorCustom {
    param([string]$Message)
    Write-Host "  [X] $Message" -ForegroundColor Red
}

function Find-SimHub {
    Write-Step "Ricerca cartella SimHub..."
    
    try {
        $regPath = Get-ItemProperty -Path "HKCU:\Software\SimHub" -ErrorAction SilentlyContinue
        if ($regPath.InstallDirectory -and (Test-Path $regPath.InstallDirectory)) {
            return $regPath.InstallDirectory
        }
    } catch {}
    
    try {
        $regPath = Get-ItemProperty -Path "HKLM:\Software\SimHub" -ErrorAction SilentlyContinue
        if ($regPath.InstallDirectory -and (Test-Path $regPath.InstallDirectory)) {
            return $regPath.InstallDirectory
        }
    } catch {}
    
    try {
        $regPath = Get-ItemProperty -Path "HKLM:\Software\WOW6432Node\SimHub" -ErrorAction SilentlyContinue
        if ($regPath.InstallDirectory -and (Test-Path $regPath.InstallDirectory)) {
            return $regPath.InstallDirectory
        }
    } catch {}
    
    $commonPaths = @(
        "$env:LOCALAPPDATA\SimHub",
        "C:\Program Files (x86)\SimHub",
        "C:\Program Files\SimHub",
        "D:\SimHub",
        "E:\SimHub"
    )
    
    foreach ($path in $commonPaths) {
        if (Test-Path (Join-Path $path "SimHubWPF.exe")) {
            return $path
        }
    }
    
    return $null
}

function Download-BranchZip {
    param(
        [string]$Url,
        [string]$DestinationZip
    )
    
    Write-Step "Download ZIP del repository da GitHub..."
    Write-Host "      URL: $Url" -ForegroundColor Gray
    
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $Url -OutFile $DestinationZip -UseBasicParsing -UserAgent "MMCMSimConnect-Installer"
        $ProgressPreference = 'Continue'
        
        $fileSize = (Get-Item $DestinationZip).Length
        $fileSizeKB = [math]::Round($fileSize / 1024, 1)
        Write-Success "ZIP scaricato: $fileSizeKB KB"
    }
    catch {
        throw "Errore durante il download dello ZIP: $_"
    }
}

function Find-SourceFolder {
    param(
        [string]$ExtractedRoot,
        [string]$SourceSubPath
    )
    
    $topDirs = Get-ChildItem -Path $ExtractedRoot -Directory
    $actualRoot = $ExtractedRoot
    
    if ($topDirs.Count -eq 1) {
        $actualRoot = $topDirs[0].FullName
        Write-Host "      Root repo: $($topDirs[0].Name)" -ForegroundColor Gray
    }
    
    # Strategia 1: Percorso esatto
    $exactPath = Join-Path $actualRoot $SourceSubPath
    if (Test-Path $exactPath) {
        Write-Host "      Trovato: $SourceSubPath" -ForegroundColor Gray
        return $exactPath
    }
    
    # Strategia 2: Cerca ricorsivamente net48
    $found = Get-ChildItem -Path $actualRoot -Directory -Recurse -Filter "net48" | Select-Object -First 1
    if ($found) {
        Write-Host "      Trovato net48 in: $($found.FullName.Substring($ExtractedRoot.Length))" -ForegroundColor Gray
        return $found.FullName
    }
    
    # Strategia 3: Se root ha binari, usa root
    $hasBinaries = Get-ChildItem -Path $actualRoot -File | 
                   Where-Object { $_.Extension -in ".dll", ".exe" } | 
                   Select-Object -First 1
    
    if ($hasBinaries) {
        Write-Host "      File binari nella root, uso root come sorgente" -ForegroundColor Gray
        return $actualRoot
    }
    
    return $null
}

function Copy-DirectoryTree {
    param(
        [string]$SourceDir,
        [string]$DestDir
    )
    
    $count = 0
    
    # Estensioni da escludere
    $excludeExtensions = @(".cs", ".csproj", ".sln", ".suo", ".user", ".gitignore", ".gitattributes", ".md")
    
    # Cartelle da escludere
    $excludeFolders = @(".git", ".vs", ".github", "obj", "bin", "node_modules", "packages", ".idea")
    
    # Copia file nella directory corrente
    Get-ChildItem -Path $SourceDir -File | ForEach-Object {
        $ext = $_.Extension.ToLower()
        
        if ($ext -notin $excludeExtensions -and -not $_.Name.StartsWith(".")) {
            $destPath = Join-Path $DestDir $_.Name
            
            $destDirPath = Split-Path $destPath -Parent
            if (-not (Test-Path $destDirPath)) {
                New-Item -ItemType Directory -Path $destDirPath -Force | Out-Null
            }
            
            Copy-Item -Path $_.FullName -Destination $destPath -Force
            Write-Host "      Copiato: $($_.Name)" -ForegroundColor Gray
            $count++
        }
    }
    
    # Ricorsione nelle sottocartelle
    Get-ChildItem -Path $SourceDir -Directory | ForEach-Object {
        $folderName = $_.Name
        
        if ($folderName -notin $excludeFolders -and -not $folderName.StartsWith(".")) {
            $destSubDir = Join-Path $DestDir $folderName
            
            if (-not (Test-Path $destSubDir)) {
                New-Item -ItemType Directory -Path $destSubDir -Force | Out-Null
            }
            
            $subCount = Copy-DirectoryTree -SourceDir $_.FullName -DestDir $destSubDir
            
            if ($subCount -gt 0) {
                Write-Host "      Cartella: $folderName/ - $subCount file" -ForegroundColor DarkGray
            }
            
            $count += $subCount
        }
    }
    
    return $count
}

# ============================================
# MAIN
# ============================================

try {
    Write-Header
    
    # Trova SimHub
    $simhubPath = Find-SimHub
    
    if (-not $simhubPath) {
        Write-Host ""
        Write-Host "  SimHub non trovato automaticamente." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Inserisci il percorso della cartella SimHub:" -ForegroundColor White
        Write-Host "  (es: C:\Program Files\SimHub)" -ForegroundColor Gray
        Write-Host ""
        $simhubPath = Read-Host "  Percorso"
        
        if (-not (Test-Path $simhubPath)) {
            throw "Il percorso specificato non esiste!"
        }
    }
    
    Write-Success "SimHub trovato: $simhubPath"
    
    # Verifica SimHub
    $simhubExe = Join-Path $simhubPath "SimHubWPF.exe"
    if (-not (Test-Path $simhubExe)) {
        Write-Host ""
        Write-Host "  ATTENZIONE: SimHubWPF.exe non trovato." -ForegroundColor Yellow
        $confirm = Read-Host "  Continuare comunque? (S/N)"
        if ($confirm -ne "S" -and $confirm -ne "s") {
            Write-Host "  Installazione annullata." -ForegroundColor Yellow
            Read-Host "  Premi INVIO per chiudere"
            exit 0
        }
    }
    
    # Verifica SimHub non in esecuzione
    $simhubProcess = Get-Process -Name "SimHubWPF" -ErrorAction SilentlyContinue
    if ($simhubProcess) {
        Write-Host ""
        Write-Host "  ATTENZIONE: SimHub e in esecuzione!" -ForegroundColor Yellow
        Write-Host "  Chiudi SimHub prima di continuare." -ForegroundColor Yellow
        $confirm = Read-Host "  Premi INVIO quando SimHub e chiuso (o A per annullare)"
        if ($confirm -eq "A" -or $confirm -eq "a") {
            Read-Host "  Premi INVIO per chiudere"
            exit 0
        }
        
        $simhubProcess = Get-Process -Name "SimHubWPF" -ErrorAction SilentlyContinue
        if ($simhubProcess) {
            throw "SimHub e ancora in esecuzione."
        }
    }
    
    Write-Host ""
    Write-Host "  ==========================================================" -ForegroundColor DarkGray
    Write-Host "  I file verranno installati in:" -ForegroundColor White
    Write-Host "  $simhubPath" -ForegroundColor Cyan
    Write-Host "  ==========================================================" -ForegroundColor DarkGray
    Write-Host ""
    
    $confirm = Read-Host "  Procedere con l installazione? (S/N)"
    if ($confirm -ne "S" -and $confirm -ne "s") {
        Read-Host "  Premi INVIO per chiudere"
        exit 0
    }
    
    Write-Host ""
    
    # Pulisci temp
    if (Test-Path $TempFolder) {
        Remove-Item -Path $TempFolder -Recurse -Force
    }
    New-Item -ItemType Directory -Path $TempFolder -Force | Out-Null
    
    # === DOWNLOAD ZIP ===
    $zipPath = Join-Path $TempFolder "repo.zip"
    $extractPath = Join-Path $TempFolder "extracted"
    
    Download-BranchZip -Url $ZipDownloadUrl -DestinationZip $zipPath
    
    # === ESTRAZIONE ===
    Write-Step "Estrazione ZIP..."
    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
    Write-Success "Estrazione completata"
    
    # === TROVA CARTELLA SORGENTE ===
    Write-Step "Ricerca cartella Release/net48..."
    $sourceFolder = Find-SourceFolder -ExtractedRoot $extractPath -SourceSubPath $SourcePath
    
    if (-not $sourceFolder) {
        throw "Cartella sorgente non trovata nello ZIP! Verificare che esista $SourcePath nel repository."
    }
    
    Write-Success "Sorgente: $sourceFolder"
    
    # === MOSTRA RIEPILOGO ===
    Write-Host ""
    $allFiles = Get-ChildItem -Path $sourceFolder -Recurse -File
    $allDirs = Get-ChildItem -Path $sourceFolder -Recurse -Directory
    
    $totalFiles = $allFiles.Count
    $totalDirs = $allDirs.Count + 1
    Write-Host "  Contenuto da installare: $totalFiles file in $totalDirs cartelle" -ForegroundColor White
    
    $topItems = Get-ChildItem -Path $sourceFolder
    foreach ($item in $topItems) {
        if ($item.PSIsContainer) {
            $subFileCount = (Get-ChildItem -Path $item.FullName -Recurse -File).Count
            Write-Host "      [DIR] $($item.Name)/ - $subFileCount file" -ForegroundColor Gray
        } else {
            $sizeKB = [math]::Round($item.Length / 1024, 1)
            Write-Host "      [FILE] $($item.Name) - $sizeKB KB" -ForegroundColor Gray
        }
    }
    
    Write-Host ""
    
    # === COPIA INTERA STRUTTURA ===
    Write-Step "Installazione in SimHub (con sottocartelle)..."
    
    $copiedCount = Copy-DirectoryTree -SourceDir $sourceFolder -DestDir $simhubPath
    
    Write-Success "Installazione completata - $copiedCount file copiati"
    
    # Pulizia
    Write-Step "Pulizia file temporanei..."
    Remove-Item -Path $TempFolder -Recurse -Force -ErrorAction SilentlyContinue
    Write-Success "Pulizia completata"
    
    # Messaggio finale
    Write-Host ""
    Write-Host "  ==========================================================" -ForegroundColor Green
    Write-Host "  |                                                        |" -ForegroundColor Green
    Write-Host "  |   INSTALLAZIONE COMPLETATA CON SUCCESSO!               |" -ForegroundColor Green
    Write-Host "  |                                                        |" -ForegroundColor Green
    Write-Host "  |   Avvia SimHub e attiva il plugin MMCM SimConnect      |" -ForegroundColor Green
    Write-Host "  |   dal menu dei plugin.                                 |" -ForegroundColor Green
    Write-Host "  |                                                        |" -ForegroundColor Green
    Write-Host "  ==========================================================" -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host ""
    Write-ErrorCustom "ERRORE: $_"
    Write-Host ""
    
    if (Test-Path $TempFolder) {
        Remove-Item -Path $TempFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Read-Host "  Premi INVIO per chiudere"