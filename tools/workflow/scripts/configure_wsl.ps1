# Script PowerShell pour configurer WSL pour le projet APEX VBA Framework
# Ce script aide aÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â  raÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©soudre les problaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â¨mes d'accaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â¨s aux fichiers et de configuration Git

# Force l'encodage UTF-8 pour l'affichage
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

# Fonction pour afficher les sections
function Write-Section {
    param ([string]$Title)
    Write-Host "`n=== $Title ===" -ForegroundColor Cyan
}

# VaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©rification de WSL
Write-Section "VaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©rification de WSL"
$wslStatus = wsl --status

if ($LASTEXITCODE -ne 0) {
    Write-Host "WSL n'est pas correctement installaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©. Veuillez installer WSL2." -ForegroundColor Red
    exit 1
}

Write-Host "WSL est installaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â© et fonctionnel." -ForegroundColor Green
Write-Host $wslStatus

# VaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©rification de la distribution Ubuntu
Write-Section "VaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©rification de la distribution Ubuntu"
$distrosList = wsl --list
$hasUbuntu = $distrosList -match "Ubuntu-22.04"

if (-not $hasUbuntu) {
    Write-Host "Ubuntu-22.04 n'est pas installaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©. Installation en cours..." -ForegroundColor Yellow
    wsl --install -d Ubuntu-22.04
    Write-Host "Ubuntu-22.04 a aÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©taÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â© installaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©. Veuillez complaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©ter la configuration en creeaa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©ant un utilisateur." -ForegroundColor Green
    Write-Host "Relancez ce script apraÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â¨s avoir configuraÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â© Ubuntu-22.04." -ForegroundColor Yellow
    exit 0
} else {
    Write-Host "Ubuntu-22.04 est installaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©." -ForegroundColor Green
}

# VaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©rification que la distribution fonctionne
try {
    $testResult = wsl --distribution Ubuntu-22.04 -- echo "WSL fonctionne correctement"
    Write-Host "Test de la distribution: $testResult" -ForegroundColor Green
} catch {
    Write-Host "Erreur lors du test de la distribution: $_" -ForegroundColor Red
    exit 1
}

# Configuration de WSL
Write-Section "Configuration de WSL.conf"

$wslConfContent = @"
[boot]
systemd=true

[automount]
enabled = true
options = "metadata,umask=22,fmask=11"
mountFsTab = false

[interop]
enabled = true
appendWindowsPath = true
"@

# CraÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©ation du fichier temporaire
$tempFile = New-TemporaryFile
$wslConfContent | Out-File -FilePath $tempFile -Encoding utf8

# Copie du fichier vers WSL
Write-Host "Copie de la configuration WSL..." -ForegroundColor Yellow

# CraÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©er un fichier de script bash temporaire pour copier le fichier avec sudo
$bashScript = @"
#!/bin/bash
cat "$($tempFile.FullName | wsl --distribution Ubuntu-22.04 -- wslpath -u)" | sudo tee /etc/wsl.conf > /dev/null
echo "Configuration WSL mise aÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â  jour avec succesaa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â¨s."
"@

$tempBashScript = New-TemporaryFile
$bashScript | Out-File -FilePath "$($tempBashScript.FullName).sh" -Encoding utf8

# ExaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©cution du script bash
Write-Host "Vous devrez saisir votre mot de passe WSL pour continuer:" -ForegroundColor Yellow
wsl --distribution Ubuntu-22.04 -- bash "$($tempBashScript.FullName | wsl --distribution Ubuntu-22.04 -- wslpath -u).sh"

# Nettoyage des fichiers temporaires
Remove-Item -Path $tempFile -Force
Remove-Item -Path "$($tempBashScript.FullName).sh" -Force

# VaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©rification de Git
Write-Section "Configuration de Git"

Write-Host "VaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©rification de la configuration Git actuelle dans WSL..." -ForegroundColor Yellow
$gitConfig = wsl --distribution Ubuntu-22.04 -- git config --list

# VaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©rifier si le nom d'utilisateur et l'email sont configuraÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©s
$hasUserName = $gitConfig -match "user.name="
$hasUserEmail = $gitConfig -match "user.email="

if (-not $hasUserName -or -not $hasUserEmail) {
    Write-Host "Configuration de Git dans WSL..." -ForegroundColor Yellow
    
    $gitUserName = Read-Host "Entrez votre nom pour Git (laissez vide pour 'APEX Framework Developer')"
    if ([string]::IsNullOrWhiteSpace($gitUserName)) {
        $gitUserName = "APEX Framework Developer"
    }
    
    $gitUserEmail = Read-Host "Entrez votre email pour Git (laissez vide pour 'apex.framework@example.com')"
    if ([string]::IsNullOrWhiteSpace($gitUserEmail)) {
        $gitUserEmail = "apex.framework@example.com"
    }
    
    wsl --distribution Ubuntu-22.04 -- git config --global user.name "$gitUserName"
    wsl --distribution Ubuntu-22.04 -- git config --global user.email "$gitUserEmail"
    wsl --distribution Ubuntu-22.04 -- git config --global core.autocrlf input
    
    Write-Host "Git configuraÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â© avec succesaa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â¨s dans WSL." -ForegroundColor Green
} else {
    Write-Host "Git est daÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©jaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â  configuraÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â© dans WSL." -ForegroundColor Green
}

# Test d'accaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â¨s aux fichiers
Write-Section "Test d'accaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â¨s aux fichiers"

$testResult = wsl --distribution Ubuntu-22.04 -- touch /mnt/d/Dev/Apex_VBA_FRAMEWORK/wsl_test_file 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "Test d'aÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©criture raÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©ussi." -ForegroundColor Green
    wsl --distribution Ubuntu-22.04 -- rm /mnt/d/Dev/Apex_VBA_FRAMEWORK/wsl_test_file
} else {
    Write-Host "ProblaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â¨me d'accaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â¨s aux fichiers: $testResult" -ForegroundColor Red
    Write-Host "Vous devrez redaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©marrer WSL pour appliquer les changements." -ForegroundColor Yellow
}

# RedaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©marrage de WSL
Write-Section "RedaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©marrage de WSL"
Write-Host "Pour appliquer toutes les modifications, WSL doit aÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Âªtre redaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©marraÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©." -ForegroundColor Yellow
$restart = Read-Host "Voulez-vous redaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©marrer WSL maintenant? (O/N)"

if ($restart -eq "O" -or $restart -eq "o") {
    Write-Host "RedaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©marrage de WSL..." -ForegroundColor Yellow
    wsl --shutdown
    Start-Sleep -Seconds 3
    wsl --distribution Ubuntu-22.04 -- echo "WSL redaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©marraÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â© avec succesaa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â¨s."
    
    # Test apraÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â¨s redaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©marrage
    $testResult = wsl --distribution Ubuntu-22.04 -- touch /mnt/d/Dev/Apex_VBA_FRAMEWORK/wsl_test_file 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Test d'aÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©criture apraÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â¨s redaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©marrage raÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©ussi." -ForegroundColor Green
        wsl --distribution Ubuntu-22.04 -- rm /mnt/d/Dev/Apex_VBA_FRAMEWORK/wsl_test_file
    } else {
        Write-Host "ProblaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â¨me persistant d'accaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â¨s aux fichiers apraÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â¨s redaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©marrage: $testResult" -ForegroundColor Red
    }
} else {
    Write-Host "N'oubliez pas de redaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©marrer WSL plus tard avec 'wsl --shutdown'." -ForegroundColor Yellow
}

Write-Section "RaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©sumaÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©"
Write-Host "Configuration WSL termineaa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢aaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â 'aÃƒÆ’"Â "Ã¢â€žÂ¢"ÃƒÆ’...Ãƒâ€šÃ‚Â¡eÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©e. Consultez le guide docs/WSL_SETUP_GUIDE.md pour plus d'informations." -ForegroundColor Green 