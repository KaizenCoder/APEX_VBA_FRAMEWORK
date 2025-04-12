# Script d'installation pour Node.js et cursor-tools
# ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ exÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cuter en tant qu'administrateur

# VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification des droits administrateur
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "Ce script doit ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Âªtre exÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cutÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© en tant qu'administrateur." -ForegroundColor Red
    Write-Host "Tentative de relancement en tant qu'administrateur..." -ForegroundColor Yellow
    
    # Tentative de relancement du script en tant qu'administrateur
    $scriptPath = $MyInvocation.MyCommand.Definition
    $arguments = "-ExecutionPolicy Bypass -File `"$scriptPath`""
    
    try {
        Start-Process powershell.exe -Verb RunAs -ArgumentList $arguments
        # Quitter ce processus aprÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s avoir lancÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© la version admin
        exit
    } catch {
        Write-Host "Impossible de relancer automatiquement en tant qu'administrateur." -ForegroundColor Red
        Write-Host "Veuillez relancer PowerShell en tant qu'administrateur manuellement." -ForegroundColor Red
        exit 1
    }
}

# DÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©finir l'encodage en UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

Write-Host "DÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©marrage de l'installation..." -ForegroundColor Green

# VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier si Node.js est dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©jÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  installÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©
$nodeVersion = node -v 2>$null
if ($nodeVersion) {
    Write-Host "Node.js est dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©jÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  installÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© (version $nodeVersion)" -ForegroundColor Yellow
} else {
    Write-Host "TÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©lÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©chargement de Node.js v20.11.1..." -ForegroundColor Cyan
    $nodeUrl = "https://nodejs.org/dist/v20.11.1/node-v20.11.1-x64.msi"
    $nodeInstaller = "$env:TEMP\node-installer.msi"
    
    try {
        Invoke-WebRequest -Uri $nodeUrl -OutFile $nodeInstaller
        Write-Host "Installation de Node.js..." -ForegroundColor Cyan
        Start-Process msiexec.exe -ArgumentList "/i `"$nodeInstaller`" /quiet" -Wait
        Remove-Item $nodeInstaller -Force
        Write-Host "Node.js a ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©tÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© installÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© avec succÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s!" -ForegroundColor Green
    } catch {
        Write-Host "Erreur lors de l'installation de Node.js: $_" -ForegroundColor Red
        exit 1
    }
}

# VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier si npm est disponible
$npmVersion = npm -v 2>$null
if (-not $npmVersion) {
    Write-Host "Erreur: npm n'est pas disponible aprÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s l'installation de Node.js" -ForegroundColor Red
    exit 1
}

# CrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation du script temporaire pour la deuxiÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨me ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©tape
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$tempScript = @'
# VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification de l'installation de Node.js
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

Write-Host "VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification de l'installation de Node.js..." -ForegroundColor Yellow
try {
    $nodeVersion = node --version
    $npmVersion = npm --version
    Write-Host "Node.js $nodeVersion installÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© avec succÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s" -ForegroundColor Green
    Write-Host "npm $npmVersion installÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© avec succÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s" -ForegroundColor Green
} catch {
    Write-Host "Erreur lors de la vÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification de Node.js: $_" -ForegroundColor Red
    exit
}

# Installation de cursor-tools (via vibe-tools)
Write-Host "Installation de cursor-tools..." -ForegroundColor Yellow
try {
    npm install -g vibe-tools
    Write-Host "cursor-tools installÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© avec succÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s" -ForegroundColor Green
} catch {
    Write-Host "Erreur lors de l'installation de cursor-tools: $_" -ForegroundColor Red
    exit
}

# CrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation du fichier de configuration
Write-Host "CrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation du fichier de configuration..." -ForegroundColor Yellow
$configContent = @"
# Configuration cursor-tools
# Ajoutez vos clÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©s API ici
ANTHROPIC_API_KEY=your_key_here
OPENAI_API_KEY=your_key_here
# Chemin vers les logs Cursor (optionnel)
CURSOR_LOGS_PATH=%APPDATA%\Cursor\User\workspaceStorage
"@

try {
    $configContent | Out-File -FilePath ".cursor-tools.env" -Encoding utf8
    Write-Host "Fichier de configuration crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© avec succÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s" -ForegroundColor Green
} catch {
    Write-Host "Erreur lors de la crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation du fichier de configuration: $_" -ForegroundColor Red
}

# Ajout de vibe-tools au PATH
Write-Host "VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification du PATH pour vibe-tools..." -ForegroundColor Yellow
$npmBinPath = npm config get prefix
$vibePath = Join-Path -Path $npmBinPath -ChildPath "vibe-tools.cmd"

if (Test-Path $vibePath) {
    Write-Host "vibe-tools trouvÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© au chemin: $vibePath" -ForegroundColor Green
} else {
    Write-Host "vibe-tools.cmd non trouvÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© au chemin attendu. VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifiez l'installation." -ForegroundColor Red
}

Write-Host "`nInstallation terminÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e !" -ForegroundColor Green
Write-Host "N'oubliez pas de configurer vos clÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©s API dans le fichier .cursor-tools.env" -ForegroundColor Yellow
Write-Host "Pour utiliser cursor-tools, redÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©marrez votre terminal puis exÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cutez: vibe <commande>" -ForegroundColor Cyan
'@

# Sauvegarde du script temporaire
$tempScriptPath = Join-Path -Path $scriptPath -ChildPath "temp_install.ps1"
$tempScript | Out-File -FilePath $tempScriptPath -Encoding utf8

# Notification ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  l'utilisateur
Write-Host "RedÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©marrage de PowerShell pour finaliser l'installation..." -ForegroundColor Yellow
Write-Host "Le script va continuer dans une nouvelle fenÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Âªtre PowerShell..." -ForegroundColor Yellow
Start-Sleep -Seconds 3

# ExÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cuter le script temporaire dans une nouvelle fenÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Âªtre avec le chemin courant
$currentPath = (Get-Location).Path
Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$tempScriptPath`" -WorkingDirectory `"$currentPath`"" 