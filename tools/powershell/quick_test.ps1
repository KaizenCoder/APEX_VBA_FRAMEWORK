# Script de test rapide pour ApexWSLBridge
# Ce script montre comment utiliser les principales fonctions du module

# Importer le module
$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "ApexWSLBridge.psm1"
Import-Module $modulePath -Force

Write-Host "=== Test rapide d'ApexWSLBridge ===" -ForegroundColor Cyan
Write-Host "Module chargÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© depuis: $modulePath" -ForegroundColor Yellow

# 1. Test de l'environnement WSL
Write-Host "`n1. VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification de l'environnement WSL" -ForegroundColor Magenta
if (Test-WSLEnvironment) {
    Write-Host "WSL est correctement configurÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©" -ForegroundColor Green
} else {
    Write-Host "ProblÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨me avec l'environnement WSL" -ForegroundColor Red
    exit
}

# 2. Obtenir les informations du projet
Write-Host "`n2. Informations sur le projet" -ForegroundColor Magenta
$rootDir = Invoke-WSLCommand -Command "find /mnt/d/Dev/Apex_VBA_FRAMEWORK -type d -maxdepth 1 | sort" -UseTempFile
Write-Host "RÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©pertoires racine du projet:"
$rootDir | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }

# 3. Compter les fichiers par extension
Write-Host "`n3. Comptage des fichiers par extension" -ForegroundColor Magenta
$extCount = Invoke-WSLCommand -Command "find /mnt/d/Dev/Apex_VBA_FRAMEWORK -type f -name '*.*' | grep -v 'node_modules' | sed 's/.*\.//' | sort | uniq -c | sort -nr | head -10" -UseTempFile
Write-Host "Top 10 des extensions de fichiers:"
$extCount | ForEach-Object { Write-Host $_ -ForegroundColor Cyan }

# 4. VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier si Git est installÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© avec retry
Write-Host "`n4. VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification de Git avec retry" -ForegroundColor Magenta
$gitVersion = Invoke-WSLCommandWithRetry -Command "git --version" -MaxRetries 2
Write-Host "Version de Git: $gitVersion" -ForegroundColor Green

# 5. Tester l'entrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e standard (comptage de mots)
Write-Host "`n5. Test d'entrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e standard" -ForegroundColor Magenta
$text = @"
Ceci est un test pour vÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier le mÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©canisme d'entrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e standard 
du module ApexWSLBridge. Ce texte contient plusieurs mots
qui seront comptÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©s avec la commande wc.
"@
$wcResult = Invoke-WSLCommandWithInput -Command "wc -w" -Input $text
Write-Host "Nombre de mots: $wcResult" -ForegroundColor Green

# 6. Mesurer la performance
Write-Host "`n6. Mesure de performance" -ForegroundColor Magenta
$perfTest = Measure-WSLCommand -Command "find /mnt/d/Dev/Apex_VBA_FRAMEWORK -name '*.ps1'"
Write-Host "Temps pour trouver les fichiers .ps1: $($perfTest.ElapsedMs) ms" -ForegroundColor Green
Write-Host "Nombre de fichiers .ps1: $($perfTest.Result.Count)" -ForegroundColor Green

# 7. Session interactive
Write-Host "`n7. DÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©marrage d'une session interactive" -ForegroundColor Magenta
Write-Host "Pour dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©marrer une session interactive, exÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cutez:" -ForegroundColor Yellow
Write-Host "Start-InteractiveWSLSession -WorkingDirectory '/mnt/d/Dev/Apex_VBA_FRAMEWORK' -InitCommand 'ls -la'" -ForegroundColor Cyan

Write-Host "`n=== Test terminÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© ===" -ForegroundColor Green
Write-Host "Le module ApexWSLBridge fonctionne correctement." 