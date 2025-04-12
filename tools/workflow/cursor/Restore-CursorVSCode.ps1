# =============================================================================
# Script de restauration de l'intégration VS Code/Cursor
# =============================================================================
#
# .SYNOPSIS
#   Restaure l'intégration VS Code/Cursor depuis une sauvegarde.
#
# .DESCRIPTION
#   Ce script permet de restaurer l'intégration depuis une sauvegarde, incluant :
#   - Restauration des configurations
#   - Réinstallation des extensions
#   - Validation de la restauration
#
# .PARAMETER BackupPath
#   Chemin vers la sauvegarde à restaurer.
#   Si non spécifié, utilise la dernière sauvegarde disponible.
#
# .PARAMETER Force
#   Force la restauration sans confirmation.
#
# .PARAMETER RestoreExtensions
#   Réinstalle également les extensions sauvegardées.
#
# .EXAMPLE
#   .\Restore-CursorVSCode.ps1
#   Restaure depuis la dernière sauvegarde.
#
# .EXAMPLE
#   .\Restore-CursorVSCode.ps1 -BackupPath "path/to/backup" -Force
#   Restaure depuis une sauvegarde spécifique sans confirmation.
#
# .NOTES
#   Version     : 1.0
#   Auteur      : APEX Framework Team
#   Création    : 11/04/2024
#   Mise à jour : 11/04/2024
#
# .LINK
#   https://github.com/org/repo/wiki/Restoration
#
# =============================================================================

# Script de restauration de l'intégration VS Code avec Cursor
[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$BackupPath,
    [switch]$Force,
    [switch]$RestoreExtensions
)

function Find-LatestBackup {
    Write-Host "`n🔍 Recherche de la dernière sauvegarde..." -ForegroundColor Cyan
    
    $backups = Get-ChildItem "tools/workflow/cursor/backup" -Directory |
    Where-Object { $_.Name -like "vscode_*" } |
    Sort-Object CreationTime -Descending
    
    if ($backups.Count -eq 0) {
        throw "Aucune sauvegarde trouvée"
    }
    
    $latest = $backups[0]
    Write-Host "  Trouvé: $($latest.Name)" -ForegroundColor Gray
    return $latest.FullName
}

function Restore-VSCodeConfiguration {
    param (
        [string]$SourcePath
    )
    
    Write-Host "`n📦 Restauration de la configuration VS Code..." -ForegroundColor Cyan
    
    # Création du dossier .vscode si nécessaire
    if (-not (Test-Path ".vscode")) {
        New-Item -ItemType Directory -Path ".vscode" -Force | Out-Null
    }
    
    # Restauration des fichiers
    $vscodeFiles = @(
        "settings.json",
        "tasks.json",
        "keybindings.json",
        "extensions.json"
    )
    
    foreach ($file in $vscodeFiles) {
        $sourcePath = Join-Path $SourcePath $file
        $targetPath = Join-Path ".vscode" $file
        
        if (Test-Path $sourcePath) {
            Copy-Item $sourcePath $targetPath -Force
            Write-Host "  Restauré: $file" -ForegroundColor Gray
        }
    }
}

function Restore-Extensions {
    Write-Host "`n📦 Restauration des extensions..." -ForegroundColor Cyan
    
    $extensionsPath = ".vscode/extensions.json"
    if (Test-Path $extensionsPath) {
        $extensions = Get-Content $extensionsPath -Raw | ConvertFrom-Json
        
        foreach ($ext in $extensions.recommendations) {
            Write-Host "  Installation: $ext" -ForegroundColor Gray
            code --install-extension $ext
        }
    }
}

function Test-Restoration {
    Write-Host "`n🔍 Validation de la restauration..." -ForegroundColor Cyan
    $errors = @()
    
    # 1. Vérification des fichiers
    $vscodeFiles = @(
        "settings.json",
        "tasks.json",
        "keybindings.json",
        "extensions.json"
    )
    
    foreach ($file in $vscodeFiles) {
        if (-not (Test-Path ".vscode/$file")) {
            $errors += "Fichier manquant: $file"
        }
    }
    
    # 2. Vérification des paramètres Cursor
    if (Test-Path ".vscode/settings.json") {
        $settings = Get-Content ".vscode/settings.json" -Raw | ConvertFrom-Json
        if (-not ($settings.PSObject.Properties.Name -contains "cursor.rules.enabled")) {
            $errors += "Paramètres Cursor manquants"
        }
    }
    
    if ($errors.Count -gt 0) {
        Write-Warning "❌ Problèmes détectés:"
        $errors | ForEach-Object { Write-Warning "  - $_" }
        return $false
    }
    
    Write-Host "✅ Restauration validée" -ForegroundColor Green
    return $true
}

# Exécution principale
try {
    Write-Host "==================================================="
    Write-Host "     RESTAURATION DE L'INTÉGRATION VS CODE         "
    Write-Host "==================================================="
    
    # Détermination du chemin de sauvegarde
    $backupDir = if ($BackupPath) {
        $BackupPath
    }
    else {
        Find-LatestBackup
    }
    
    if (-not $Force) {
        $response = Read-Host "Voulez-vous restaurer la configuration depuis $backupDir ? (O/N)"
        if ($response -ne "O") {
            Write-Host "Restauration annulée" -ForegroundColor Yellow
            exit 0
        }
    }
    
    # Restauration
    Restore-VSCodeConfiguration -SourcePath $backupDir
    
    if ($RestoreExtensions) {
        Restore-Extensions
    }
    
    # Validation
    if (Test-Restoration) {
        Write-Host "`n✨ Restauration terminée avec succès" -ForegroundColor Green
        Write-Host "Note: Redémarrez VS Code pour appliquer tous les changements" -ForegroundColor Yellow
    }
    else {
        throw "Erreurs lors de la validation de la restauration"
    }
}
catch {
    Write-Error "❌ Erreur lors de la restauration: $_"
    exit 1
} 