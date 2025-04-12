# Script de test pour le module ApexWSLBridge
# Ce script permet de vÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier le bon fonctionnement des fonctions du module

# Importation du module (force le rechargement s'il est dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©jÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  importÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©)
$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "ApexWSLBridge.psm1"
Import-Module -Force $modulePath

# Fonction pour afficher les rÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©sultats de test
function Write-TestResult {
    param (
        [string]$TestName,
        [bool]$Success,
        [object]$Result = $null
    )
    
    if ($Success) {
        Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€¦Ã¢â‚¬Å“ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ TEST RÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â°USSI: $TestName" -ForegroundColor Green
    } else {
        Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€šÃ‚ÂÃƒâ€¦Ã¢â‚¬â„¢ TEST ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â°CHOUÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â°: $TestName" -ForegroundColor Red
    }
    
    if ($null -ne $Result) {
        Write-Host "   RÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©sultat: " -NoNewline
        Write-Host $Result -ForegroundColor Cyan
    }
    
    Write-Host ""
}

# Effacer l'ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cran
Clear-Host

Write-Host "=== Tests du module ApexWSLBridge ===" -ForegroundColor Magenta
Write-Host "Module chargÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© depuis: $modulePath" -ForegroundColor Yellow
Write-Host "Date du test: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Yellow
Write-Host "=======================================" -ForegroundColor Magenta
Write-Host ""

# Test 1: Environnement WSL
Write-Host "ÃƒÆ’Ã‚Â°Ãƒâ€¦Ã‚Â¸Ãƒâ€šÃ‚Â§Ãƒâ€šÃ‚Âª TEST 1: Environnement WSL" -ForegroundColor Cyan
$testResult = Test-WSLEnvironment
Write-TestResult -TestName "VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification de l'environnement WSL" -Success ($null -ne $testResult) -Result $testResult

# Test 2: Montage de disques
Write-Host "ÃƒÆ’Ã‚Â°Ãƒâ€¦Ã‚Â¸Ãƒâ€šÃ‚Â§Ãƒâ€šÃ‚Âª TEST 2: Montage de disques" -ForegroundColor Cyan
$testResult = Get-WSLMountStatus -Drive "d"
$mountSuccess = $null -ne $testResult -and $testResult.HasMetadata
Write-TestResult -TestName "VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification du montage du disque D:" -Success $mountSuccess -Result $testResult.RawInfo

# Test 3: Commande simple
Write-Host "ÃƒÆ’Ã‚Â°Ãƒâ€¦Ã‚Â¸Ãƒâ€šÃ‚Â§Ãƒâ€šÃ‚Âª TEST 3: Commande simple" -ForegroundColor Cyan
$testResult = Invoke-WSLCommand -Command "ls -la /mnt/d/Dev/Apex_VBA_FRAMEWORK | head -n 5" -LogOutput
Write-TestResult -TestName "ExÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cution d'une commande simple" -Success ($null -ne $testResult) -Result ($testResult | Out-String)

# Test 4: Commande avec fichier temporaire
Write-Host "ÃƒÆ’Ã‚Â°Ãƒâ€¦Ã‚Â¸Ãƒâ€šÃ‚Â§Ãƒâ€šÃ‚Âª TEST 4: Commande avec fichier temporaire" -ForegroundColor Cyan
$testResult = Invoke-WSLCommand -Command "ls -la /mnt/d/Dev/Apex_VBA_FRAMEWORK | head -n 5" -UseTempFile
Write-TestResult -TestName "ExÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cution avec fichier temporaire" -Success ($null -ne $testResult) -Result ($testResult | Out-String)

# Test 5: Commande avec retry
Write-Host "ÃƒÆ’Ã‚Â°Ãƒâ€¦Ã‚Â¸Ãƒâ€šÃ‚Â§Ãƒâ€šÃ‚Âª TEST 5: Commande avec retry" -ForegroundColor Cyan
$testResult = Invoke-WSLCommandWithRetry -Command "ls -la /mnt/d/Dev/Apex_VBA_FRAMEWORK | head -n 5" -MaxRetries 2
Write-TestResult -TestName "ExÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cution avec retry" -Success ($null -ne $testResult) -Result "Commande exÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cutÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e avec succÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s"

# Test 6: Commande avec entrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e
Write-Host "ÃƒÆ’Ã‚Â°Ãƒâ€¦Ã‚Â¸Ãƒâ€šÃ‚Â§Ãƒâ€šÃ‚Âª TEST 6: Commande avec entrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e" -ForegroundColor Cyan
$testInput = "Test d'entrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e pour la commande WSL"
$testResult = Invoke-WSLCommandWithInput -Command "wc -w" -Input $testInput
$wordCount = [int]$testResult.Trim()
Write-TestResult -TestName "Commande avec entrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e" -Success ($wordCount -gt 0) -Result "Nombre de mots: $wordCount"

# Test 7: Mesure de performance
Write-Host "ÃƒÆ’Ã‚Â°Ãƒâ€¦Ã‚Â¸Ãƒâ€šÃ‚Â§Ãƒâ€šÃ‚Âª TEST 7: Mesure de performance" -ForegroundColor Cyan
$testResult = Measure-WSLCommand -Command "ls -la /mnt/d/Dev/Apex_VBA_FRAMEWORK"
Write-TestResult -TestName "Mesure de performance" -Success ($testResult.ElapsedMs -ge 0) -Result "Temps d'exÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cution: $($testResult.ElapsedMs) ms"

# Test 8: Fichier batch
Write-Host "ÃƒÆ’Ã‚Â°Ãƒâ€¦Ã‚Â¸Ãƒâ€šÃ‚Â§Ãƒâ€šÃ‚Âª TEST 8: ExÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cution de commandes en batch" -ForegroundColor Cyan

# CrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation d'un fichier batch temporaire
$batchContent = @"
# Commandes batch de test
echo 'Test de batch'
ls -la /mnt/d/Dev/Apex_VBA_FRAMEWORK | head -n 3
pwd
"@

$tempBatchFile = [System.IO.Path]::GetTempFileName()
$batchContent | Out-File -FilePath $tempBatchFile -Encoding utf8

$testResult = Start-WSLBatchFromFile -FilePath $tempBatchFile
$success = $testResult.Success -and $testResult.Results.Count -gt 0

# Nettoyage
Remove-Item -Path $tempBatchFile -Force

Write-TestResult -TestName "ExÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cution de commandes en batch" -Success $success -Result "Nombre de rÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©sultats: $($testResult.Results.Count)"

# RÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©sumÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© des tests
Write-Host "=======================================" -ForegroundColor Magenta
Write-Host "Tous les tests ont ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©tÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© exÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cutÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©s." -ForegroundColor Yellow
Write-Host "Consultez les logs pour plus de dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©tails: $($script:LogPath)" -ForegroundColor Yellow
Write-Host "=======================================" -ForegroundColor Magenta 