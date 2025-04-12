# =============================================================================
# Script de désinstallation de l'intégration VS Code/Cursor
# =============================================================================
#
# .SYNOPSIS
#   Désinstalle proprement l'intégration entre VS Code et Cursor.
#
# .DESCRIPTION
#   Ce script effectue une désinstallation complète de l'intégration, incluant :
#   - Suppression des configurations
#   - Désinstallation des extensions
#   - Nettoyage des fichiers temporaires
#   - Sauvegarde optionnelle
#
# .PARAMETER Force
#   Force la désinstallation sans confirmation.
#
# .PARAMETER KeepSettings
#   Conserve les paramètres VS Code personnalisés.
#
# .PARAMETER KeepExtensions
#   Conserve les extensions installées.
#
# .PARAMETER NoBackup
#   Ne crée pas de sauvegarde avant la désinstallation.
#
# .EXAMPLE
#   .\Uninstall-CursorVSCode.ps1
#   Désinstalle avec confirmation et sauvegarde.
#
# .EXAMPLE
#   .\Uninstall-CursorVSCode.ps1 -Force -NoBackup
#   Désinstalle sans confirmation ni sauvegarde.
#
# .NOTES
#   Version     : 1.0
#   Auteur      : APEX Framework Team
#   Création    : 11/04/2024
#   Mise à jour : 11/04/2024
#
# .LINK
#   https://github.com/org/repo/wiki/Uninstallation
#
# =============================================================================

# Script de désinstallation de l'intégration VS Code avec Cursor
[CmdletBinding()]
param (
    [switch]$Force,
    [switch]$KeepSettings,
    [switch]$KeepExtensions,
    [switch]$NoBackup
)

function Backup-VSCodeConfiguration {
    Write-Host "`n📦 Sauvegarde de la configuration VS Code..." -ForegroundColor Cyan
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupDir = "tools/workflow/cursor/backup/vscode_$timestamp"
    
    # Création du dossier de backup
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    
    # Sauvegarde des fichiers VS Code
    $vscodeFiles = @(
        "settings.json",
        "tasks.json",
        "keybindings.json",
        "extensions.json"
    )
    
    foreach ($file in $vscodeFiles) {
        $sourcePath = Join-Path ".vscode" $file
        if (Test-Path $sourcePath) {
            Copy-Item $sourcePath (Join-Path $backupDir $file)
            Write-Host "  Sauvegardé: $file" -ForegroundColor Gray
        }
    }
    
    Write-Host "✅ Configuration sauvegardée dans: $backupDir" -ForegroundColor Green
    return $backupDir
}

function Remove-CursorSettings {
    Write-Host "`n🗑️ Suppression des paramètres Cursor..." -ForegroundColor Cyan
    
    $settingsPath = ".vscode/settings.json"
    if (Test-Path $settingsPath) {
        $settings = Get-Content $settingsPath -Raw | ConvertFrom-Json
        
        # Liste des paramètres Cursor à supprimer
        $cursorSettings = @(
            "cursor.rules.enabled",
            "cursor.rules.validateOnSave",
            "cursor.rules.validateOnType",
            "workspaceInit.tasks",
            "powershell.scriptAnalysis.settingsPath"
        )
        
        # Suppression des paramètres
        foreach ($setting in $cursorSettings) {
            if ($settings.PSObject.Properties.Name -contains $setting) {
                $settings.PSObject.Properties.Remove($setting)
                Write-Host "  Supprimé: $setting" -ForegroundColor Gray
            }
        }
        
        # Sauvegarde des modifications
        $settings | ConvertTo-Json -Depth 10 | Set-Content $settingsPath
    }
}

function Remove-CursorTasks {
    Write-Host "`n🗑️ Suppression des tâches Cursor..." -ForegroundColor Cyan
    
    $tasksPath = ".vscode/tasks.json"
    if (Test-Path $tasksPath) {
        $tasks = Get-Content $tasksPath -Raw | ConvertFrom-Json
        
        # Filtrer les tâches non-Cursor
        $tasks.tasks = @($tasks.tasks | Where-Object { $_.label -notlike "Cursor:*" })
        
        # Supprimer les inputs Cursor
        if ($tasks.inputs) {
            $tasks.inputs = @($tasks.inputs | Where-Object { $_.id -notlike "cursor*" })
        }
        
        # Sauvegarde des modifications
        $tasks | ConvertTo-Json -Depth 10 | Set-Content $tasksPath
        Write-Host "  Tâches Cursor supprimées" -ForegroundColor Gray
    }
}

function Remove-CursorKeybindings {
    Write-Host "`n🗑️ Suppression des raccourcis Cursor..." -ForegroundColor Cyan
    
    $keybindingsPath = ".vscode/keybindings.json"
    if (Test-Path $keybindingsPath) {
        $keybindings = Get-Content $keybindingsPath -Raw | ConvertFrom-Json
        
        # Filtrer les keybindings non-Cursor
        $newKeybindings = @($keybindings | Where-Object {
                -not ($_.args -like "Cursor:*")
            })
        
        # Sauvegarde des modifications
        $newKeybindings | ConvertTo-Json -Depth 10 | Set-Content $keybindingsPath
        Write-Host "  Raccourcis Cursor supprimés" -ForegroundColor Gray
    }
}

function Remove-CursorExtensions {
    if (-not $KeepExtensions) {
        Write-Host "`n🗑️ Suppression des extensions Cursor..." -ForegroundColor Cyan
        
        $extensionsPath = ".vscode/extensions.json"
        if (Test-Path $extensionsPath) {
            $extensions = Get-Content $extensionsPath -Raw | ConvertFrom-Json
            
            # Liste des extensions Cursor
            $cursorExtensions = @(
                "ms-vscode.powershell",
                "usernamehw.errorlens",
                "gruntfuggly.todo-tree"
            )
            
            # Filtrer les extensions
            $extensions.recommendations = @($extensions.recommendations | Where-Object {
                    $ext = $_
                    -not ($cursorExtensions -contains $ext)
                })
            
            # Sauvegarde des modifications
            $extensions | ConvertTo-Json -Depth 10 | Set-Content $extensionsPath
            Write-Host "  Extensions Cursor supprimées de la configuration" -ForegroundColor Gray
            
            # Désinstallation des extensions
            foreach ($ext in $cursorExtensions) {
                code --uninstall-extension $ext
                Write-Host "  Désinstallé: $ext" -ForegroundColor Gray
            }
        }
    }
}

function Test-Uninstallation {
    Write-Host "`n🔍 Validation de la désinstallation..." -ForegroundColor Cyan
    $errors = @()
    
    # 1. Vérification des paramètres
    if (Test-Path ".vscode/settings.json") {
        $settings = Get-Content ".vscode/settings.json" -Raw | ConvertFrom-Json
        if ($settings.PSObject.Properties.Name -like "cursor.*") {
            $errors += "Paramètres Cursor toujours présents"
        }
    }
    
    # 2. Vérification des tâches
    if (Test-Path ".vscode/tasks.json") {
        $tasks = Get-Content ".vscode/tasks.json" -Raw | ConvertFrom-Json
        if ($tasks.tasks | Where-Object { $_.label -like "Cursor:*" }) {
            $errors += "Tâches Cursor toujours présentes"
        }
    }
    
    # 3. Vérification des raccourcis
    if (Test-Path ".vscode/keybindings.json") {
        $keybindings = Get-Content ".vscode/keybindings.json" -Raw | ConvertFrom-Json
        if ($keybindings | Where-Object { $_.args -like "Cursor:*" }) {
            $errors += "Raccourcis Cursor toujours présents"
        }
    }
    
    if ($errors.Count -gt 0) {
        Write-Warning "❌ Problèmes détectés:"
        $errors | ForEach-Object { Write-Warning "  - $_" }
        return $false
    }
    
    Write-Host "✅ Désinstallation validée" -ForegroundColor Green
    return $true
}

# Exécution principale
try {
    Write-Host "==================================================="
    Write-Host "     DÉSINSTALLATION DE L'INTÉGRATION VS CODE      "
    Write-Host "==================================================="
    
    if (-not $Force) {
        $response = Read-Host "Voulez-vous vraiment désinstaller l'intégration VS Code avec Cursor ? (O/N)"
        if ($response -ne "O") {
            Write-Host "Désinstallation annulée" -ForegroundColor Yellow
            exit 0
        }
    }
    
    # Backup si nécessaire
    if (-not $NoBackup) {
        $backupDir = Backup-VSCodeConfiguration
    }
    
    # Désinstallation
    if (-not $KeepSettings) {
        Remove-CursorSettings
        Remove-CursorTasks
        Remove-CursorKeybindings
    }
    Remove-CursorExtensions
    
    # Validation
    if (Test-Uninstallation) {
        Write-Host "`n✨ Désinstallation terminée avec succès" -ForegroundColor Green
        Write-Host "Note: Redémarrez VS Code pour appliquer tous les changements" -ForegroundColor Yellow
    }
    else {
        throw "Erreurs lors de la validation de la désinstallation"
    }
}
catch {
    Write-Error "❌ Erreur lors de la désinstallation: $_"
    exit 1
} 