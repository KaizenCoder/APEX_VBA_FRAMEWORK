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
