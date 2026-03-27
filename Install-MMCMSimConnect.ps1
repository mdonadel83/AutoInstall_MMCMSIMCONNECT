<#
.SYNOPSIS
    MMCM SimConnect - Installer automatico da GitHub (VERSIONE CORRETTA)
.DESCRIPTION
    Scarica l'ultima versione del plugin da GitHub e la installa in SimHub.
    USA IL DOWNLOAD ZIP DEL BRANCH (1 sola richiesta HTTP) anziché la Contents API
    per evitare il rate limit di GitHub (60 req/ora per utenti non autenticati).
.NOTES
    Repository: https://github.com/mdonadel83/MMCMSimConnect
    Fix: Risolto errore "API rate limit exceeded" 
#>

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"

# ============================================
# CONFIGURAZIONE
# ============================================
$RepoOwner = "mdonadel83"
$RepoName = "MMCMSimConnect"
$BranchName = "main"
$SourcePath = "Release/net48"  # Cartella da cui copiare i file

# URL per scaricare l'intero branch come ZIP (NON usa l'API, nessun rate limit!)
$ZipDownloadUrl = "https://github.com/$RepoOwner/$RepoName/archive/refs/heads/$BranchName.zip"

$TempFolder = Join-Path $env:TEMP "MMCMSimConnect_Install"

# ============================================
# FUNZIONI
# ============================================

function Write-Header {
    Clear-Host
    Write-Host ""
    Write-Host "  ==========================================================" -ForegroundColor Cyan
    Write-Host "  |         MMCM SimConnect - Installer v2.1 (Fixed)       |" -ForegroundColor Cyan
    Write-Host "  |         Plugin per SimHub                              |" -ForegroundColor Cyan
    Write-Host "  |         (No API Rate Limit)                            |" -ForegroundColor Cyan
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
    
    # Metodo 1: Registro HKCU
    try {
        $regPath = Get-ItemProperty -Path "HKCU:\Software\SimHub" -ErrorAction SilentlyContinue
        if ($regPath.InstallDirectory -and (Test-Path $regPath.InstallDirectory)) {
            return $regPath.InstallDirectory
        }
    } catch {}
    
    # Metodo 2: Registro HKLM
    try {
        $regPath = Get-ItemProperty -Path "HKLM:\Software\SimHub" -ErrorAction SilentlyContinue
        if ($regPath.InstallDirectory -and (Test-Path $regPath.InstallDirectory)) {
            return $regPath.InstallDirectory
        }
    } catch {}
    
    # Metodo 3: Registro HKLM WOW6432Node
    try {
        $regPath = Get-ItemProperty -Path "HKLM:\Software\WOW6432Node\SimHub" -ErrorAction SilentlyContinue
        if ($regPath.InstallDirectory -and (Test-Path $regPath.InstallDirectory)) {
            return $regPath.InstallDirectory
        }
    } catch {}
    
    # Metodo 4: Percorsi comuni
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
        
        # Usa Invoke-WebRequest con progress bar
        $ProgressPreference = 'SilentlyContinue'  # Velocizza il download
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

function Extract-AndCopyFiles {
    param(
        [string]$ZipPath,
        [string]$ExtractPath,
        [string]$SourceSubPath,
        [string]$DestinationFolder
    )
    
    Write-Step "Estrazione ZIP..."
    
    # Estrai lo ZIP
    Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
    
    # GitHub crea una cartella tipo "MMCMSimConnect-main" dentro lo ZIP
    $extractedRoot = Get-ChildItem -Path $ExtractPath -Directory | Select-Object -First 1
    
    if (-not $extractedRoot) {
        throw "Nessuna cartella trovata nello ZIP estratto!"
    }
    
    Write-Host "      Cartella estratta: $($extractedRoot.Name)" -ForegroundColor Gray
    
    # Cerca la cartella sorgente (Release/net48)
    $sourceFolder = Join-Path $extractedRoot.FullName $SourceSubPath
    
    if (-not (Test-Path $sourceFolder)) {
        # Fallback: cerca ricorsivamente
        Write-Host "      Percorso '$SourceSubPath' non trovato direttamente, ricerca..." -ForegroundColor Yellow
        
        $found = Get-ChildItem -Path $extractedRoot.FullName -Directory -Recurse | 
                 Where-Object { $_.FullName -like "*$($SourceSubPath.Replace('/', '\'))*" } |
                 Select-Object -First 1
        
        if ($found) {
            $sourceFolder = $found.FullName
            Write-Host "      Trovato: $sourceFolder" -ForegroundColor Gray
        } else {
            # Se non troviamo Release/net48, usa la root (potrebbe essere un repo con struttura diversa)
            Write-Host "      '$SourceSubPath' non trovato. Cerco file DLL nella root..." -ForegroundColor Yellow
            $sourceFolder = $extractedRoot.FullName
        }
    }
    
    Write-Success "Sorgente: $sourceFolder"
    
    # Copia i file nella destinazione
    Write-Step "Copia file in SimHub..."
    $copiedCount = 0
    
    Get-ChildItem -Path $sourceFolder -Recurse -File | ForEach-Object {
        $relativePath = $_.FullName.Substring($sourceFolder.Length + 1)
        $destPath = Join-Path $DestinationFolder $relativePath
        
        $destDir = Split-Path $destPath -Parent
        if (-not (Test-Path $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }
        
        Copy-Item -Path $_.FullName -Destination $destPath -Force
        Write-Host "      Copiato: $relativePath" -ForegroundColor Gray
        $copiedCount++
    }
    
    return $copiedCount
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
    
    # Verifica che sia SimHub
    $simhubExe = Join-Path $simhubPath "SimHubWPF.exe"
    if (-not (Test-Path $simhubExe)) {
        Write-Host ""
        Write-Host "  ATTENZIONE: SimHubWPF.exe non trovato in questa cartella." -ForegroundColor Yellow
        $confirm = Read-Host "  Continuare comunque? (S/N)"
        if ($confirm -ne "S" -and $confirm -ne "s") {
            Write-Host "  Installazione annullata." -ForegroundColor Yellow
            Read-Host "  Premi INVIO per chiudere"
            exit 0
        }
    }
    
    # Verifica che SimHub non sia in esecuzione
    $simhubProcess = Get-Process -Name "SimHubWPF" -ErrorAction SilentlyContinue
    if ($simhubProcess) {
        Write-Host ""
        Write-Host "  ATTENZIONE: SimHub e in esecuzione!" -ForegroundColor Yellow
        Write-Host "  Chiudi SimHub prima di continuare." -ForegroundColor Yellow
        Write-Host ""
        $confirm = Read-Host "  Premi INVIO quando SimHub e chiuso (o A per annullare)"
        if ($confirm -eq "A" -or $confirm -eq "a") {
            Write-Host "  Installazione annullata." -ForegroundColor Yellow
            Read-Host "  Premi INVIO per chiudere"
            exit 0
        }
        
        $simhubProcess = Get-Process -Name "SimHubWPF" -ErrorAction SilentlyContinue
        if ($simhubProcess) {
            throw "SimHub e ancora in esecuzione. Chiudilo e riprova."
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
        Write-Host "  Installazione annullata." -ForegroundColor Yellow
        Read-Host "  Premi INVIO per chiudere"
        exit 0
    }
    
    Write-Host ""
    
    # Pulisci cartella temporanea
    if (Test-Path $TempFolder) {
        Remove-Item -Path $TempFolder -Recurse -Force
    }
    New-Item -ItemType Directory -Path $TempFolder -Force | Out-Null
    
    # ====== METODO CORRETTO: Scarica ZIP intero (1 sola richiesta HTTP) ======
    $zipPath = Join-Path $TempFolder "repo.zip"
    $extractPath = Join-Path $TempFolder "extracted"
    
    Download-BranchZip -Url $ZipDownloadUrl -DestinationZip $zipPath
    
    $copiedCount = Extract-AndCopyFiles -ZipPath $zipPath `
                                        -ExtractPath $extractPath `
                                        -SourceSubPath $SourcePath `
                                        -DestinationFolder $simhubPath
    
    Write-Success "Installazione completata ($copiedCount file)"
    
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