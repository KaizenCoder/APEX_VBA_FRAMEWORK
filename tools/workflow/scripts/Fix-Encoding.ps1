# Fix-Encoding.ps1
# Script pour corriger l'encodage des fichiers PowerShell

$files = Get-ChildItem -Path . -Filter "*.ps1"
foreach ($file in $files) {
    $content = Get-Content -Path $file.FullName -Raw
    $utf8NoBOM = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($file.FullName, $content, $utf8NoBOM)
    Write-Host "Encodage corrige pour : $($file.Name)"
} 