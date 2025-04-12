#==========================================================================
# Script : FixClassFiles.ps1
# Version : 1.0
# Purpose : Correction des fichiers .cls pour enlever les commentaires devant VERSION 1.0 CLASS
# Date : 12/04/2025
#==========================================================================

# Affichage du titre
Write-Host "===== CORRECTION DES FICHIERS DE CLASSE VBA =====" -ForegroundColor Cyan
Write-Host ""

# --- Variables ---
$LogFile = "fix_classes_log.txt"
$SourceFolders = @(
    "apex-core",
    "apex-metier",
    "apex-ui"
)
$BackupFolder = "classes_backup"

# --- Initialisation ---
Write-Host "[INFO] Initialisation..." -ForegroundColor Green
if (Test-Path $LogFile) { Remove-Item $LogFile -Force }
"DÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©but de la correction: $(Get-Date)" | Out-File $LogFile

# --- CrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation du dossier de sauvegarde ---
Write-Host "[INFO] CrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation du dossier de sauvegarde..." -ForegroundColor Green
if (-not (Test-Path $BackupFolder)) {
    New-Item -Path $BackupFolder -ItemType Directory | Out-Null
}

# --- Obtenir la liste des fichiers de classe ---
Write-Host "[INFO] Recherche des fichiers de classe (.cls)..." -ForegroundColor Green
$ClassFiles = Get-ChildItem -Path $SourceFolders -Filter "*.cls" -Recurse | Select-Object -ExpandProperty FullName
$TotalFiles = $ClassFiles.Count
Write-Host "[INFO] $TotalFiles fichiers de classe trouvÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©s" -ForegroundColor Green
"Fichiers de classe trouvÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©s: $TotalFiles" | Out-File $LogFile -Append

# --- Correction des fichiers ---
Write-Host "[INFO] Correction des fichiers..." -ForegroundColor Green
$CorrectedCount = 0
$AlreadyOkCount = 0
$ErrorCount = 0

foreach ($file in $ClassFiles) {
    try {
        # Lire le contenu du fichier
        $content = Get-Content -Path $file -Raw -Encoding UTF8
        $fileName = Split-Path -Path $file -Leaf
        
        # VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier si le problÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨me existe
        if ($content -match "'\s*VERSION\s+1\.0\s+CLASS") {
            # CrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©er une sauvegarde
            $backupPath = Join-Path -Path $BackupFolder -ChildPath $fileName
            Copy-Item -Path $file -Destination $backupPath -Force
            
            # Corriger le contenu
            $correctedContent = $content -replace "'\s*(VERSION\s+1\.0\s+CLASS)", '$1'
            
            # ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â°crire le contenu corrigÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©
            $correctedContent | Out-File -FilePath $file -Encoding UTF8
            
            $CorrectedCount++
            Write-Host "   Fichier corrigÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $fileName" -ForegroundColor Green
            "Fichier corrigÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $file" | Out-File $LogFile -Append
        } else {
            $AlreadyOkCount++
            "Fichier dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©jÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  correct: $file" | Out-File $LogFile -Append
        }
    } catch {
        $ErrorCount++
        Write-Host "   [ERREUR] ProblÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨me avec le fichier $fileName : $($_.Exception.Message)" -ForegroundColor Red
        "ERREUR: ProblÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨me avec le fichier $file : $($_.Exception.Message)" | Out-File $LogFile -Append
    }
    
    # Afficher la progression
    $Progress = [Math]::Round((($CorrectedCount + $AlreadyOkCount + $ErrorCount) / $TotalFiles) * 100)
    Write-Host "`rProgression: $Progress% ($($CorrectedCount + $AlreadyOkCount + $ErrorCount)/$TotalFiles)" -NoNewline -ForegroundColor Yellow
}

Write-Host "`n[INFO] Correction terminÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e" -ForegroundColor Green
Write-Host "   $CorrectedCount fichiers corrigÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©s" -ForegroundColor Green
Write-Host "   $AlreadyOkCount fichiers dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©jÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  corrects" -ForegroundColor Green
Write-Host "   $ErrorCount fichiers avec erreurs" -ForegroundColor Red

"Fichiers corrigÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©s: $CorrectedCount" | Out-File $LogFile -Append
"Fichiers dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©jÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  corrects: $AlreadyOkCount" | Out-File $LogFile -Append
"Fichiers avec erreurs: $ErrorCount" | Out-File $LogFile -Append

# --- Conseil pour recrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©er l'add-in ---
Write-Host ""
Write-Host "[INFO] Prochaine ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©tape : RecrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©er l'add-in" -ForegroundColor Yellow
Write-Host "Pour crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©er l'add-in avec les fichiers corrigÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©s, exÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cutez : ./tools/CreateApexAddIn.ps1" -ForegroundColor White

# --- Rapport final ---
Write-Host "[INFO] Finalisation..." -ForegroundColor Green
"Fin de la correction: $(Get-Date)" | Out-File $LogFile -Append

Write-Host ""
Write-Host "===== FIN DE LA CORRECTION =====" -ForegroundColor Cyan 