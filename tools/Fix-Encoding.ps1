# Fix-Encoding.ps1
# Script pour convertir tous les fichiers PowerShell en UTF-8 sans BOM
# Ce script resout les problemes d'encodage qui causent des erreurs dans PowerShell

$root = "D:\Dev\Apex_VBA_FRAMEWORK"
$files = Get-ChildItem -Path $root -Recurse -Include *.ps1

# Compteurs pour le rapport
$filesWithBom = 0
$filesWithoutBom = 0
$filesProcessed = 0

# Fonction pour verifier si un fichier contient un BOM UTF-8
function Test-BOM {
    param($Path)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    return ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
}

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "CORRECTION DE L'ENCODAGE DES FICHIERS POWERSHELL" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Recherche des fichiers PowerShell..." -ForegroundColor Yellow
Write-Host "Nombre de fichiers trouves: $($files.Count)" -ForegroundColor Yellow
Write-Host ""

foreach ($file in $files) {
    $hasBom = Test-BOM -Path $file.FullName
    
    if ($hasBom) {
        Write-Host "[+] Reencodage (BOM detecte): $($file.FullName)" -ForegroundColor Yellow
        $filesWithBom++
    } else {
        Write-Host "[v] Verification (Sans BOM): $($file.FullName)" -ForegroundColor Gray
        $filesWithoutBom++
    }
    
    # Reencoder tous les fichiers pour s'assurer qu'ils sont tous en UTF-8 sans BOM
    $content = Get-Content -Path $file.FullName -Raw
    $utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($file.FullName, $content, $utf8NoBomEncoding)
    
    $filesProcessed++
}

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "RAPPORT D'ENCODAGE" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "Fichiers traites: $filesProcessed" -ForegroundColor White
Write-Host "Fichiers qui avaient un BOM: $filesWithBom" -ForegroundColor Yellow
Write-Host "Fichiers deja sans BOM: $filesWithoutBom" -ForegroundColor Green
Write-Host ""
Write-Host "[OK] Tous les fichiers .ps1 ont ete reencodees en UTF-8 sans BOM." -ForegroundColor Green
Write-Host ""
Write-Host "Pour verifier l'encodage d'un fichier specifique, utilisez:" -ForegroundColor Cyan
Write-Host "PS> .\tools\Test-FileEncoding.ps1 'chemin\vers\fichier.ps1'" -ForegroundColor White 