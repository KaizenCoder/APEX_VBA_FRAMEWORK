#==========================================================================
# Script : ExportRelease.ps1
# Version : 1.0
# Purpose : Exportation des versions de release du framework APEX
# Date : 10/04/2025
#==========================================================================

# Affichage du titre
Write-Host "===== EXPORTATION DE LA RELEASE APEX FRAMEWORK =====" -ForegroundColor Cyan
Write-Host ""

# --- Variables ---
$Version = "1.0.0"
$ReleaseDir = "release"
$DestinationDir = "\\serveur\partage\APEX_Framework\releases"
$LogFile = "export_log.txt"

# --- Initialisation ---
Write-Host "[INFO] Initialisation de l'exportation..." -ForegroundColor Green
if (Test-Path $LogFile) { Remove-Item $LogFile -Force }
"DÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©but de l'exportation: $(Get-Date)" | Out-File $LogFile

# --- VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification des prÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©requis ---
Write-Host "[INFO] VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification des prÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©requis..." -ForegroundColor Green
if (-not (Test-Path "$ReleaseDir\ApexFramework_v$Version.zip")) {
    Write-Host "[ERREUR] Le fichier de release n'existe pas: $ReleaseDir\ApexFramework_v$Version.zip" -ForegroundColor Red
    "ERREUR: Fichier de release non trouvÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $ReleaseDir\ApexFramework_v$Version.zip" | Out-File $LogFile -Append
    Write-Host "ExÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cutez d'abord BuildRelease.bat pour crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©er la release." -ForegroundColor Yellow
    exit 1
}

# --- CrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation des rÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©pertoires de destination si nÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cessaire ---
Write-Host "[INFO] VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification des rÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©pertoires de destination..." -ForegroundColor Green
if (-not (Test-Path $DestinationDir)) {
    try {
        New-Item -Path $DestinationDir -ItemType Directory -Force | Out-Null
        "RÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©pertoire de destination crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $DestinationDir" | Out-File $LogFile -Append
    }
    catch {
        Write-Host "[ERREUR] Impossible de crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©er le rÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©pertoire de destination: $DestinationDir" -ForegroundColor Red
        "ERREUR: Impossible de crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©er le rÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©pertoire: $($_.Exception.Message)" | Out-File $LogFile -Append
        exit 1
    }
}

# --- Copie du fichier de release ---
Write-Host "[INFO] Copie du fichier de release vers la destination..." -ForegroundColor Green
try {
    Copy-Item "$ReleaseDir\ApexFramework_v$Version.zip" -Destination "$DestinationDir\ApexFramework_v$Version.zip" -Force
    "Fichier copiÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $ReleaseDir\ApexFramework_v$Version.zip -> $DestinationDir\ApexFramework_v$Version.zip" | Out-File $LogFile -Append
}
catch {
    Write-Host "[ERREUR] ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â°chec de la copie du fichier de release: $($_.Exception.Message)" -ForegroundColor Red
    "ERREUR: ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â°chec de la copie: $($_.Exception.Message)" | Out-File $LogFile -Append
    exit 1
}

# --- CrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation du fichier de mÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©tadonnÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©es ---
Write-Host "[INFO] CrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation du fichier de mÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©tadonnÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©es..." -ForegroundColor Green
$MetadataFile = "$DestinationDir\ApexFramework_v$Version.metadata"
@"
Version: $Version
Date d'exportation: $(Get-Date -Format "dd/MM/yyyy")
Heure d'exportation: $(Get-Date -Format "HH:mm:ss")
Taille du fichier: $([Math]::Round((Get-Item "$ReleaseDir\ApexFramework_v$Version.zip").Length / 1MB, 2)) MB
"@ | Out-File $MetadataFile

# --- Mise ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  jour du fichier de versions disponibles ---
Write-Host "[INFO] Mise ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  jour du fichier des versions disponibles..." -ForegroundColor Green
$VersionsFile = "$DestinationDir\versions_disponibles.txt"
if (-not (Test-Path $VersionsFile)) {
    "# APEX FRAMEWORK - VERSIONS DISPONIBLES" | Out-File $VersionsFile
    "# Format: Version | Date | Taille" | Out-File $VersionsFile -Append
    "# ----------------------------------" | Out-File $VersionsFile -Append
}

"v$Version | $(Get-Date -Format "dd/MM/yyyy") | $([Math]::Round((Get-Item "$ReleaseDir\ApexFramework_v$Version.zip").Length / 1MB, 2)) MB" | Out-File $VersionsFile -Append

# --- VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification finale ---
Write-Host "[INFO] VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification finale..." -ForegroundColor Green
if (Test-Path "$DestinationDir\ApexFramework_v$Version.zip") {
    Write-Host "[SUCCÃƒÆ’Ã†â€™Ãƒâ€¹Ã¢â‚¬Â S] L'exportation s'est terminÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e avec succÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s." -ForegroundColor Green
    Write-Host "Fichier exportÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $DestinationDir\ApexFramework_v$Version.zip" -ForegroundColor White
    "Export rÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ussi vers: $DestinationDir\ApexFramework_v$Version.zip" | Out-File $LogFile -Append
}
else {
    Write-Host "[ERREUR] L'exportation a ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©chouÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©. Le fichier n'est pas prÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©sent ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  la destination." -ForegroundColor Red
    "ERREUR: VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification finale ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©chouÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e. Fichier non prÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©sent ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  destination." | Out-File $LogFile -Append
    exit 1
}

# --- Rapport final ---
Write-Host "[INFO] Finalisation..." -ForegroundColor Green
"Fin de l'exportation: $(Get-Date)" | Out-File $LogFile -Append

Write-Host ""
Write-Host "===== FIN DE L'EXPORTATION =====" -ForegroundColor Cyan
