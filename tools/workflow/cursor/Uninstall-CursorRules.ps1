# =============================================================================
# Script de désinstallation des règles Cursor
# =============================================================================
#
# .SYNOPSIS
#   Désinstalle les règles Cursor de l'environnement PowerShell.
#
# .DESCRIPTION
#   Ce script effectue une désinstallation complète des règles Cursor :
#   - Suppression des hooks PowerShell
#   - Nettoyage des fichiers de session
#   - Suppression des logs
#   - Restauration de la configuration VS Code
#   - Nettoyage de l'environnement
#
# .PARAMETER RemoveSessionFiles
#   Supprime les fichiers de session (.cursor-session-*.json).
#
# .PARAMETER Force
#   Force la désinstallation sans demander de confirmation.
#
# .PARAMETER KeepBackup
#   Conserve les fichiers de sauvegarde après la désinstallation.
#
# .EXAMPLE
#   .\Uninstall-CursorRules.ps1
#   Désinstalle les règles avec confirmation.
#
# .EXAMPLE
#   .\Uninstall-CursorRules.ps1 -Force -RemoveSessionFiles
#   Désinstalle les règles et supprime les sessions sans confirmation.
#
# .INPUTS
#   [switch] RemoveSessionFiles
#   [switch] Force
#   [switch] KeepBackup
#
# .OUTPUTS
#   [PSObject] Résultat de la désinstallation avec :
#   - Status : État de la désinstallation
#   - RemovedFiles : Liste des fichiers supprimés
#   - Warnings : Avertissements éventuels
#
# .NOTES
#   Version     : 1.0
#   Auteur      : APEX Framework Team
#   Création    : 11/04/2024
#   Mise à jour : 11/04/2024
#   Prérequis   :
#   - PowerShell 5.1 ou supérieur
#   - Droits administrateur
#
# .LINK
#   https://github.com/org/repo/wiki/Uninstallation
#
# .COMPONENT
#   APEX VBA Framework
#
# =============================================================================

# Validation des prérequis
#requires -Version 5.1
#requires -RunAsAdministrator

# Script de désinstallation des règles Cursor
[CmdletBinding()]
param (
    [switch]$RemoveSessionFiles,
    [switch]$Force,
    [switch]$KeepBackup
)

function Remove-CursorComponents {
    Write-Host "`n🔄 Suppression des composants Cursor..." -ForegroundColor Cyan
    
    # 1. Désinstallation des hooks
    . "$PSScriptRoot\Unregister-CursorHooks.ps1"
    
    # 2. Nettoyage des fichiers de session
    if ($RemoveSessionFiles) {
        Write-Host "`n🗑️ Suppression des fichiers de session..." -ForegroundColor Cyan
        Get-ChildItem -Path (Get-Location) -Filter ".cursor-session-*.json" | 
        ForEach-Object {
            Remove-Item $_.FullName -Force:$Force
            Write-Host "  Supprimé: $($_.Name)" -ForegroundColor Gray
        }
    }
    
    # 3. Nettoyage des logs
    if (Test-Path "logs/cursor") {
        Write-Host "`n🗑️ Nettoyage des logs..." -ForegroundColor Cyan
        Remove-Item "logs/cursor" -Recurse -Force:$Force
    }
    
    # 4. Restauration de la configuration VS Code
    $vscodePath = ".vscode/settings.json"
    if (Test-Path $vscodePath) {
        Write-Host "`n🔄 Restauration de la configuration VS Code..." -ForegroundColor Cyan
        $settings = Get-Content $vscodePath -Raw | ConvertFrom-Json
        
        # Suppression des configurations Cursor
        @(
            "workspaceInit.tasks",
            "powershell.scriptAnalysis.settingsPath",
            "powershell.debugging.createTemporaryIntegratedConsole",
            "powershell.integratedConsole.suppressStartupBanner",
            "powershell.integratedConsole.focusConsoleOnExecute",
            "powershell.startAutomatically",
            "powershell.enableProfileLoading"
        ) | ForEach-Object {
            if ($settings.PSObject.Properties.Name -contains $_) {
                $settings.PSObject.Properties.Remove($_)
                Write-Host "  Supprimé: $_" -ForegroundColor Gray
            }
        }
        
        $settings | ConvertTo-Json -Depth 10 | Set-Content $vscodePath
    }
}

function Remove-CursorEnvironment {
    Write-Host "`n🧹 Nettoyage de l'environnement..." -ForegroundColor Cyan
    
    # 1. Suppression des variables d'environnement
    @(
        'CURSOR_WORKSPACE',
        'CURSOR_RULES_LOADED'
    ) | ForEach-Object {
        if (Test-Path "env:$_") {
            Remove-Item "env:$_"
            Write-Host "  Variable supprimée: $_" -ForegroundColor Gray
        }
    }
    
    # 2. Nettoyage des dossiers temporaires
    if (-not $KeepBackup) {
        $backupDir = "tools/workflow/cursor/backup"
        if (Test-Path $backupDir) {
            Remove-Item $backupDir -Recurse -Force:$Force
            Write-Host "  Backups supprimés" -ForegroundColor Gray
        }
    }
}

function Test-Uninstallation {
    Write-Host "`n🔍 Validation de la désinstallation..." -ForegroundColor Cyan
    $errors = @()
    
    # 1. Vérification des hooks
    if (Get-Content $PROFILE.CurrentUserAllHosts | Select-String "Hook Cursor Rules") {
        $errors += "Hooks toujours présents dans le profil PowerShell"
    }
    
    # 2. Vérification des variables d'environnement
    if ($env:CURSOR_WORKSPACE -or $env:CURSOR_RULES_LOADED) {
        $errors += "Variables d'environnement toujours présentes"
    }
    
    # 3. Vérification de la configuration VS Code
    if (Test-Path ".vscode/settings.json") {
        $settings = Get-Content ".vscode/settings.json" -Raw | ConvertFrom-Json
        if ($settings.PSObject.Properties.Name -contains "workspaceInit.tasks") {
            $errors += "Configuration VS Code toujours présente"
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
    Write-Host "     DÉSINSTALLATION DES RÈGLES CURSOR             "
    Write-Host "==================================================="
    
    if (-not $Force) {
        $response = Read-Host "Voulez-vous vraiment désinstaller les règles Cursor ? (O/N)"
        if ($response -ne "O") {
            Write-Host "Désinstallation annulée" -ForegroundColor Yellow
            exit 0
        }
    }
    
    # Désinstallation
    Remove-CursorComponents
    Remove-CursorEnvironment
    
    # Validation
    if (Test-Uninstallation) {
        Write-Host "`n✨ Désinstallation terminée avec succès" -ForegroundColor Green
        Write-Host "Note: Redémarrez votre terminal pour appliquer tous les changements" -ForegroundColor Yellow
    }
    else {
        throw "Erreurs lors de la validation de la désinstallation"
    }
}
catch {
    Write-Error "❌ Erreur lors de la désinstallation: $_"
    exit 1
} 