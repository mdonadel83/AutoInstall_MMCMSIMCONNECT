<#
.SYNOPSIS
    MMCM SimConnect - Installer automatico da GitHub
.DESCRIPTION
    Scarica l'ultima versione del plugin da GitHub e la installa in SimHub
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
$SourcePath = "Release/net48"
$GitHubApiUrl = "https://api.github.com/repos/$RepoOwner/$RepoName/contents/$SourcePath`?ref=$BranchName"
$TempFolder = Join-Path $env:TEMP "MMCMSimConnect_Install"

# ============================================
# FUNZIONI
# ============================================

function Write-Header {
    Clear-Host
    Write-Host ""
    Write-Host "  ==========================================================" -ForegroundColor Cyan
    Write-Host "  |         MMCM SimConnect - Installer v2.0               |" -ForegroundColor Cyan
    Write-Host "  |         Plugin per SimHub                              |" -ForegroundColor Cyan
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

function Get-GitHubFiles {
    param([string]$ApiUrl)
    
    Write-Step "Connessione a GitHub..."
    
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        $headers = @{
            "User-Agent" = "MMCMSimConnect-Installer"
            "Accept" = "application/vnd.github.v3+json"
        }
        
        $response = Invoke-RestMethod -Uri $ApiUrl -Headers $headers -Method Get
        return $response
    }
    catch {
        throw "Errore durante la connessione a GitHub: $_"
    }
}

function Download-File {
    param(
        [string]$Url,
        [string]$Destination
    )
    
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "MMCMSimConnect-Installer")
        $webClient.DownloadFile($Url, $Destination)
    }
    catch {
        throw "Errore download: $_"
    }
}

function Download-GitHubFolder {
    param(
        [string]$ApiUrl,
        [string]$DestinationFolder
    )
    
    $files = Get-GitHubFiles -ApiUrl $ApiUrl
    $downloadedFiles = @()
    
    foreach ($item in $files) {
        if ($item.type -eq "file") {
            $fileName = $item.name
            $downloadUrl = $item.download_url
            $destPath = Join-Path $DestinationFolder $fileName
            
            Write-Host "      Scarico: $fileName" -ForegroundColor Gray
            Download-File -Url $downloadUrl -Destination $destPath
            $downloadedFiles += $destPath
        }
        elseif ($item.type -eq "dir") {
            $subFolder = Join-Path $DestinationFolder $item.name
            New-Item -ItemType Directory -Path $subFolder -Force | Out-Null
            $subFiles = Download-GitHubFolder -ApiUrl $item.url -DestinationFolder $subFolder
            $downloadedFiles += $subFiles
        }
    }
    
    return $downloadedFiles
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
    
    # Crea cartella temporanea
    if (Test-Path $TempFolder) {
        Remove-Item -Path $TempFolder -Recurse -Force
    }
    New-Item -ItemType Directory -Path $TempFolder -Force | Out-Null
    
    # Scarica i file da GitHub
    Write-Step "Download dei file da GitHub..."
    $downloadedFiles = Download-GitHubFolder -ApiUrl $GitHubApiUrl -DestinationFolder $TempFolder
    Write-Success "Download completato ($($downloadedFiles.Count) file)"
    
    # Copia i file in SimHub
    Write-Step "Installazione in SimHub..."
    $copiedCount = 0
    
    Get-ChildItem -Path $TempFolder -Recurse | ForEach-Object {
        if (-not $_.PSIsContainer) {
            $relativePath = $_.FullName.Substring($TempFolder.Length + 1)
            $destPath = Join-Path $simhubPath $relativePath
            
            $destDir = Split-Path $destPath -Parent
            if (-not (Test-Path $destDir)) {
                New-Item -ItemType Directory -Path $destDir -Force | Out-Null
            }
            
            Copy-Item -Path $_.FullName -Destination $destPath -Force
            Write-Host "      Copiato: $relativePath" -ForegroundColor Gray
            $copiedCount++
        }
    }
    
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
