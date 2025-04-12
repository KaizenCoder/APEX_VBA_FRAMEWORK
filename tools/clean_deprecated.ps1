# clean_deprecated.ps1
# Script de nettoyage des fichiers .DEPRECATED pour Apex Framework

Write-Host "====================================================" -ForegroundColor Cyan
Write-Host " APEX Framework - Nettoyage des fichiers .DEPRECATED" -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host ""

$count = 0
$files = Get-ChildItem -Path . -Recurse -Filter "*.DEPRECATED" -File

if ($files.Count -eq 0) {
    Write-Host "Aucun fichier .DEPRECATED trouvÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©." -ForegroundColor Yellow
} else {
    foreach ($file in $files) {
        Write-Host "Suppression: $($file.FullName)" -ForegroundColor Gray
        Remove-Item -Path $file.FullName -Force
        $count++
    }
    
    Write-Host ""
    Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€¦Ã¢â‚¬Å“ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ $count fichiers .DEPRECATED supprimÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©s avec succÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s." -ForegroundColor Green
}

Write-Host ""
Write-Host "Appuyez sur une touche pour continuer..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") 