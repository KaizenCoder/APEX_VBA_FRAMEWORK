# =============================================================================
# Script d'installation des règles Cursor
# =============================================================================
#
# .SYNOPSIS
#   Installe et configure les règles Cursor dans l'environnement PowerShell.
#
# .DESCRIPTION
#   Ce script installe et configure les règles Cursor dans l'environnement :
#   - Installation des hooks PowerShell
#   - Configuration de VS Code
#   - Initialisation de l'environnement
#   - Création du fichier de session
#   - Validation de l'installation
#
# .PARAMETER Force
#   Force l'installation sans demander de confirmation.
#
# .PARAMETER NoBackup
#   Ne crée pas de sauvegarde avant l'installation.
#
# .EXAMPLE
#   .\Install-CursorRules.ps1
#   Installe les règles avec sauvegarde et confirmation.
#
# .EXAMPLE
#   .\Install-CursorRules.ps1 -Force -NoBackup
#   Installe les règles sans sauvegarde ni confirmation.
#
# .INPUTS
#   [switch] Force
#   [switch] NoBackup
#
# .OUTPUTS
#   [PSObject] Résultat de l'installation avec :
#   - Status : État de l'installation
#   - BackupPath : Chemin de la sauvegarde (si créée)
#   - Components : Liste des composants installés
#
# .NOTES
#   Version     : 1.0
#   Auteur      : APEX Framework Team
#   Création    : 11/04/2024
#   Mise à jour : 11/04/2024
#   Prérequis   :
#   - PowerShell 5.1 ou supérieur
#   - VS Code installé
#   - Git installé
#
# .LINK
#   https://github.com/org/repo/wiki/Installation
#
# .COMPONENT
#   APEX VBA Framework
#
# =============================================================================

# Validation des prérequis
#requires -Version 5.1
#requires -RunAsAdministrator

[CmdletBinding()]
param (
    [switch]$Force,
    [switch]$NoBackup
)

function Backup-ExistingConfiguration {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupDir = "tools/workflow/cursor/backup/$timestamp"
    
    # Création du dossier de backup
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    
    # Sauvegarde des fichiers existants
    if (Test-Path $PROFILE.CurrentUserAllHosts) {
        Copy-Item $PROFILE.CurrentUserAllHosts "$backupDir/profile.ps1"
    }
    if (Test-Path ".vscode/settings.json") {
        Copy-Item ".vscode/settings.json" "$backupDir/settings.json"
    }
    
    Write-Host "✅ Configuration sauvegardée dans: $backupDir" -ForegroundColor Green
}

function Install-Prerequisites {
    # Vérification PowerShell
    $psVersion = $PSVersionTable.PSVersion
    if ($psVersion.Major -lt 5) {
        throw "PowerShell 5.1 ou supérieur requis. Version actuelle: $psVersion"
    }
    
    # Création des dossiers nécessaires
    $folders = @(
        "tools/workflow/cursor",
        ".vscode",
        "logs/cursor"
    )
    foreach ($folder in $folders) {
        if (-not (Test-Path $folder)) {
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
        }
    }
}

function Install-CursorComponents {
    # 1. Installation des hooks PowerShell
    Write-Host "`n📦 Installation des hooks PowerShell..." -ForegroundColor Cyan
    . "$PSScriptRoot\Register-CursorHooks.ps1"
    
    # 2. Configuration VS Code
    Write-Host "`n📦 Configuration de VS Code..." -ForegroundColor Cyan
    $vscodePath = ".vscode/settings.json"
    $settings = @{
        "powershell.scriptAnalysis.settingsPath"                = "./.cursor-rules.d/config/PSScriptAnalyzerSettings.psd1"
        "powershell.debugging.createTemporaryIntegratedConsole" = $true
        "powershell.integratedConsole.suppressStartupBanner"    = $true
        "powershell.integratedConsole.focusConsoleOnExecute"    = $false
        "powershell.startAutomatically"                         = $true
        "powershell.enableProfileLoading"                       = $true
        "workspaceInit.tasks"                                   = @(
            @{
                "name"    = "Initialize Cursor Rules"
                "command" = "powershell"
                "args"    = @("-NoProfile", "-Command", "& {. `$env:CURSOR_WORKSPACE\tools\workflow\scripts\Register-CursorHooks.ps1}")
                "runOn"   = @("workspaceOpen", "folderOpen")
            }
        )
    }
    $settings | ConvertTo-Json -Depth 10 | Set-Content $vscodePath
    
    # 3. Initialisation de l'environnement
    Write-Host "`n📦 Initialisation de l'environnement..." -ForegroundColor Cyan
    $env:CURSOR_WORKSPACE = (Get-Location).Path
    $env:CURSOR_RULES_LOADED = $true
    
    # 4. Création du fichier de session initial
    $sessionFile = ".cursor-session-$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    @{
        workspace          = $env:CURSOR_WORKSPACE
        timestamp          = (Get-Date).ToString('o')
        installation_date  = (Get-Date).ToString('o')
        powershell_version = $PSVersionTable.PSVersion.ToString()
    } | ConvertTo-Json > $sessionFile
}

function Test-Installation {
    Write-Host "`n🔍 Validation de l'installation..." -ForegroundColor Cyan
    
    # 1. Vérification des fichiers
    $requiredFiles = @(
        $PROFILE.CurrentUserAllHosts,
        ".vscode/settings.json",
        "tools/workflow/cursor/Register-CursorHooks.ps1",
        "tools/workflow/cursor/Unregister-CursorHooks.ps1"
    )
    
    foreach ($file in $requiredFiles) {
        if (-not (Test-Path $file)) {
            Write-Warning "❌ Fichier manquant: $file"
            return $false
        }
    }
    
    # 2. Vérification des variables d'environnement
    if (-not $env:CURSOR_WORKSPACE -or -not $env:CURSOR_RULES_LOADED) {
        Write-Warning "❌ Variables d'environnement non configurées"
        return $false
    }
    
    # 3. Test des hooks
    if (-not (Get-Content $PROFILE.CurrentUserAllHosts | Select-String "Hook Cursor Rules")) {
        Write-Warning "❌ Hooks non installés"
        return $false
    }
    
    Write-Host "✅ Installation validée" -ForegroundColor Green
    return $true
}

# Exécution principale
try {
    Write-Host "==================================================="
    Write-Host "     INSTALLATION DES RÈGLES CURSOR                 "
    Write-Host "==================================================="
    
    # Backup si demandé
    if (-not $NoBackup) {
        Backup-ExistingConfiguration
    }
    
    # Installation
    Install-Prerequisites
    Install-CursorComponents
    
    # Validation
    if (Test-Installation) {
        Write-Host "`n✨ Installation terminée avec succès" -ForegroundColor Green
        Write-Host "Note: Redémarrez votre terminal pour activer les hooks" -ForegroundColor Yellow
    }
    else {
        throw "Erreurs lors de la validation de l'installation"
    }
}
catch {
    Write-Error "❌ Erreur lors de l'installation: $_"
    exit 1
} 