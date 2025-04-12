# =============================================================================
# Script de configuration de l'intégration VS Code/Cursor
# =============================================================================
#
# .SYNOPSIS
#   Configure et maintient l'intégration entre VS Code et Cursor sur la même machine.
#
# .DESCRIPTION
#   Ce script configure l'intégration complète entre VS Code et Cursor, incluant :
#   - Configuration du workspace partagé
#   - Synchronisation des extensions
#   - Configuration du débogueur
#   - Configuration du terminal intégré
#   - Synchronisation en temps réel des configurations
#
# .PARAMETER Force
#   Force la configuration sans confirmation.
#
# .PARAMETER EnableSharing
#   Active toutes les fonctionnalités de partage.
#
# .PARAMETER WatchMode
#   Active la surveillance en temps réel des changements de configuration.
#
# .EXAMPLE
#   .\Configure-IDEIntegration.ps1
#   Configure l'intégration avec les paramètres par défaut.
#
# .EXAMPLE
#   .\Configure-IDEIntegration.ps1 -Force -EnableSharing
#   Configure l'intégration complète sans confirmation.
#
# .EXAMPLE
#   .\Configure-IDEIntegration.ps1 -WatchMode
#   Configure l'intégration et surveille les changements en temps réel.
#
# .NOTES
#   Version     : 2.0
#   Auteur      : APEX Framework Team
#   Création    : 11/04/2024
#   Mise à jour : 12/04/2024
#
# .LINK
#   https://github.com/org/repo/wiki/IDE-Integration
#
# =============================================================================

# Script de configuration de l'intégration VS Code et Cursor
[CmdletBinding()]
param (
    [switch]$Force,
    [switch]$EnableSharing,
    [switch]$WatchMode
)

# Fonction de journalisation améliorée
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Debug', 'Success')]
        [string]$Level = 'Info',
        [switch]$NoNewline
    )
    
    $colors = @{
        'Info'    = 'Cyan'
        'Warning' = 'Yellow'
        'Error'   = 'Red'
        'Debug'   = 'Gray'
        'Success' = 'Green'
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Affichage console
    Write-Host $logMessage -ForegroundColor $colors[$Level] -NoNewline:$NoNewline
    
    # Journalisation dans un fichier
    $logPath = Join-Path $PSScriptRoot "ide_integration.log"
    Add-Content -Path $logPath -Value $logMessage
}

function Set-SharedWorkspace {
    Write-Host "`n🔄 Configuration du workspace partagé..." -ForegroundColor Cyan
    
    # Vérification et configuration du workspace
    if (-not $env:CURSOR_WORKSPACE) {
        $env:CURSOR_WORKSPACE = $PWD.Path
        Write-Log "CURSOR_WORKSPACE non défini, utilisation du répertoire courant: $($PWD.Path)" -Level Warning
    }
    
    # Configuration du partage de workspace
    $sharedConfig = @{
        "shared.workspace.enabled"       = $true
        "shared.workspace.path"          = $env:CURSOR_WORKSPACE
        "shared.extensions.sync"         = $true
        "shared.terminal.integration"    = $true
        "shared.debugging.allowCrossIDE" = $true
    }
    
    # Mise à jour VS Code
    $vscodePath = ".vscode/settings.json"
    
    # Création du dossier .vscode s'il n'existe pas
    if (-not (Test-Path ".vscode")) {
        New-Item -ItemType Directory -Path ".vscode" | Out-Null
        Write-Log "Création du dossier .vscode" -Level Info
    }
    
    # Création ou mise à jour du fichier settings.json
    if (Test-Path $vscodePath) {
        $settings = Get-Content $vscodePath -Raw | ConvertFrom-Json
    }
    else {
        $settings = [PSCustomObject]@{}
        Write-Log "Création d'un nouveau fichier settings.json" -Level Info
    }
    
    foreach ($key in $sharedConfig.Keys) {
        Add-Member -InputObject $settings -NotePropertyName $key -NotePropertyValue $sharedConfig[$key] -Force
    }
    $settings | ConvertTo-Json -Depth 10 | Set-Content $vscodePath
    
    # Configuration Cursor
    $cursorConfigDir = ".cursor-rules.d"
    if (-not (Test-Path $cursorConfigDir)) {
        New-Item -ItemType Directory -Path $cursorConfigDir | Out-Null
        Write-Log "Création du dossier $cursorConfigDir" -Level Info
    }
    
    $cursorConfig = Join-Path $cursorConfigDir "ide_integration.json"
    $sharedConfig | ConvertTo-Json -Depth 10 | Set-Content $cursorConfig
    
    Write-Log "Configuration du workspace partagé terminée" -Level Success
}

function Set-ExtensionSync {
    Write-Host "`n🔄 Configuration de la synchronisation des extensions..." -ForegroundColor Cyan
    
    # Liste des extensions partagées
    $sharedExtensions = @(
        "ms-vscode.powershell",
        "eamodio.gitlens",
        "usernamehw.errorlens",
        "gruntfuggly.todo-tree"
    )
    
    # Configuration VS Code
    $extensionsPath = ".vscode/extensions.json"
    $extensions = @{
        "recommendations"         = $sharedExtensions
        "unwantedRecommendations" = @()
    }
    $extensions | ConvertTo-Json | Set-Content $extensionsPath
    
    # Installation dans les deux IDE
    foreach ($ext in $sharedExtensions) {
        Write-Host "  Installation: $ext" -ForegroundColor Gray
        code --install-extension $ext
        # Cursor utilise les mêmes extensions que VS Code
    }
}

function Set-DebuggerIntegration {
    Write-Host "`n🔄 Configuration de l'intégration du débogueur..." -ForegroundColor Cyan
    
    $debugConfig = @{
        "debug.allowBreakpointsEverywhere" = $true
        "debug.showInStatusBar"            = "always"
        "debug.toolBarLocation"            = "floating"
        "debug.internalConsoleOptions"     = "openOnSessionStart"
    }
    
    # Configuration VS Code
    $launchPath = ".vscode/launch.json"
    $launch = @{
        "version"        = "0.2.0"
        "configurations" = @(
            @{
                "name"    = "Shared PowerShell Debug"
                "type"    = "PowerShell"
                "request" = "launch"
                "script"  = "${file}"
                "cwd"     = "${workspaceFolder}"
            }
        )
    }
    $launch | ConvertTo-Json -Depth 10 | Set-Content $launchPath
    
    # Mise à jour settings.json
    $vscodePath = ".vscode/settings.json"
    if (Test-Path $vscodePath) {
        $settings = Get-Content $vscodePath -Raw | ConvertFrom-Json
        foreach ($key in $debugConfig.Keys) {
            Add-Member -InputObject $settings -NotePropertyName $key -NotePropertyValue $debugConfig[$key] -Force
        }
        $settings | ConvertTo-Json -Depth 10 | Set-Content $vscodePath
    }
}

function Set-TerminalIntegration {
    Write-Host "`n🔄 Configuration de l'intégration du terminal..." -ForegroundColor Cyan
    
    $terminalConfig = @{
        "terminal.integrated.defaultProfile.windows" = "PowerShell"
        "terminal.integrated.profiles.windows"       = @{
            "PowerShell" = @{
                "path" = "pwsh.exe"
                "icon" = "terminal-powershell"
            }
        }
        "terminal.integrated.env.windows"            = @{
            "CURSOR_WORKSPACE"    = $env:CURSOR_WORKSPACE
            "CURSOR_RULES_LOADED" = "true"
        }
    }
    
    # Mise à jour settings.json
    $vscodePath = ".vscode/settings.json"
    if (Test-Path $vscodePath) {
        $settings = Get-Content $vscodePath -Raw | ConvertFrom-Json
        foreach ($key in $terminalConfig.Keys) {
            Add-Member -InputObject $settings -NotePropertyName $key -NotePropertyValue $terminalConfig[$key] -Force
        }
        $settings | ConvertTo-Json -Depth 10 | Set-Content $vscodePath
    }
}

function Test-Integration {
    Write-Host "`n🔍 Validation de l'intégration..." -ForegroundColor Cyan
    $errors = @()
    
    # 1. Vérification des fichiers de configuration
    $requiredFiles = @(
        ".vscode/settings.json",
        ".vscode/launch.json",
        ".cursor-rules.d/ide_integration.json"
    )
    
    foreach ($file in $requiredFiles) {
        if (-not (Test-Path $file)) {
            $errors += "Fichier manquant: $file"
        }
    }
    
    # 2. Vérification des paramètres de partage
    if (Test-Path ".vscode/settings.json") {
        $settings = Get-Content ".vscode/settings.json" -Raw | ConvertFrom-Json
        if (-not $settings."shared.workspace.enabled") {
            $errors += "Partage de workspace non activé"
        }
    }
    
    if ($errors.Count -gt 0) {
        Write-Warning "❌ Problèmes détectés:"
        $errors | ForEach-Object { Write-Warning "  - $_" }
        return $false
    }
    
    Write-Host "✅ Intégration validée" -ForegroundColor Green
    return $true
}

function Watch-ConfigurationChanges {
    Write-Log "Démarrage de la surveillance des changements..." -Level Info
    
    $paths = @{
        VSCode = ".vscode"
        Cursor = ".cursor-rules.d"
    }
    
    $fsw = New-Object System.IO.FileSystemWatcher
    $fsw.Path = $paths.VSCode
    $fsw.Filter = "*.json"
    $fsw.IncludeSubdirectories = $false
    
    # Création de l'action pour les événements
    $eventAction = {
        param($source, $e)
        
        $eventFile = $e.Name
        $eventPath = $e.FullPath
        $eventType = $e.ChangeType
        
        Write-Log "Changement détecté : $eventType - $eventFile" -Level Debug
        
        switch ($eventFile) {
            "settings.json" {
                Set-SharedWorkspace
                Write-Log "Configuration du workspace mise à jour" -Level Info
            }
            "launch.json" {
                Set-DebuggerIntegration
                Write-Log "Configuration du débogueur mise à jour" -Level Info
            }
            "extensions.json" {
                Set-ExtensionSync
                Write-Log "Configuration des extensions mise à jour" -Level Info
            }
        }
    }
    
    # Enregistrement des événements
    try {
        [void](Register-ObjectEvent -InputObject $fsw -EventName Created -Action $eventAction)
        [void](Register-ObjectEvent -InputObject $fsw -EventName Changed -Action $eventAction)
        [void](Register-ObjectEvent -InputObject $fsw -EventName Deleted -Action $eventAction)
        [void](Register-ObjectEvent -InputObject $fsw -EventName Renamed -Action $eventAction)
        
        # Activation de la surveillance
        $fsw.EnableRaisingEvents = $true
        
        Write-Log "Surveillance active. Appuyez sur Ctrl+C pour arrêter." -Level Info
        
        while ($WatchMode) {
            Start-Sleep -Seconds 1
        }
    }
    catch {
        Write-Log "Erreur lors de la surveillance : $_" -Level Error
    }
    finally {
        $fsw.EnableRaisingEvents = $false
        $fsw.Dispose()
        Get-EventSubscriber | Unregister-Event
    }
}

# Exécution principale
try {
    Write-Log "==================================================="
    Write-Log "     CONFIGURATION DE L'INTÉGRATION VS CODE/CURSOR  "
    Write-Log "==================================================="
    
    if (-not $Force) {
        $response = Read-Host "Voulez-vous configurer l'intégration VS Code/Cursor ? (O/N)"
        if ($response -ne "O") {
            Write-Log "Configuration annulée" -Level Warning
            exit 0
        }
    }
    
    # Configuration initiale
    Set-SharedWorkspace
    Set-ExtensionSync
    Set-DebuggerIntegration
    Set-TerminalIntegration
    
    # Validation
    if (Test-Integration) {
        Write-Log "Configuration terminée avec succès" -Level Success
        Write-Log "Note: Redémarrez les deux IDE pour appliquer tous les changements" -Level Warning
        
        # Mode surveillance si activé
        if ($WatchMode) {
            Write-Log "Démarrage du mode surveillance..." -Level Info
            Watch-ConfigurationChanges
        }
    }
    else {
        throw "Erreurs lors de la validation de l'intégration"
    }
}
catch {
    Write-Log "Erreur lors de la configuration: $_" -Level Error
    exit 1
}
finally {
    if (-not $WatchMode) {
        Write-Log "Fin du processus de configuration" -Level Info
    }
} 