# Script de migration des logs
$ErrorActionPreference = "Stop"

# Chemins source et destination
$oldPath = "D:\Dev\Apex_VBA_FRAMEWORK\src\Tools\Logger\reports\sessions"
$newPath = "D:\Dev\Apex_VBA_FRAMEWORK\tools\workflow\sessions"

# Création du dossier destination si nécessaire
if (-not (Test-Path $newPath)) {
    New-Item -Path $newPath -ItemType Directory -Force
}

try {
    # Copie des fichiers
    if (Test-Path $oldPath) {
        Get-ChildItem -Path $oldPath -Filter "*.md" | ForEach-Object {
            $destFile = Join-Path $newPath $_.Name
            if (Test-Path $destFile) {
                # Si le fichier existe, on fusionne le contenu
                $oldContent = Get-Content $_.FullName -Raw
                $newContent = Get-Content $destFile -Raw
                $mergedContent = "$oldContent`n`n$newContent"
                Set-Content -Path $destFile -Value $mergedContent -Encoding UTF8
                Write-Host "Fusion du fichier : $($_.Name)"
            }
            else {
                # Sinon, on copie simplement le fichier
                Copy-Item $_.FullName -Destination $destFile
                Write-Host "Copie du fichier : $($_.Name)"
            }
        }
        Write-Host "Migration terminée avec succès"
    }
    else {
        Write-Host "Aucun ancien log à migrer"
    }
}
catch {
    Write-Error "Erreur lors de la migration : $_"
    exit 1
} 