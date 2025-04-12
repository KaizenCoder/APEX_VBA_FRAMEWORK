# =============================================================================
# Script d'initialisation des règles Cursor
# =============================================================================
#
# .SYNOPSIS
#   Initialise et configure les règles Cursor pour l'environnement de développement.
#
# .DESCRIPTION
#   Ce script configure l'environnement de développement avec les règles Cursor :
#   - Création de la structure de dossiers .cursor-rules.d
#   - Configuration des règles de base
#   - Initialisation des hooks Git
#   - Configuration des variables d'environnement
#   - Mise en place des règles de validation
#   - Configuration des extensions recommandées
#
# .PARAMETER Force
#   Force l'initialisation sans demander de confirmation.
#
# .PARAMETER ConfigPath
#   Chemin vers un fichier de configuration personnalisé.
#   Par défaut : ".cursor-rules.d/config.json"
#
# .PARAMETER Environment
#   Environnement cible (Development, Production, Test).
#   Par défaut : "Development"
#
# .PARAMETER SkipGitHooks
#   Ne pas initialiser les hooks Git.
#
# .PARAMETER EnableTelemetry
#   Active la télémétrie pour le suivi des performances.
#
# .EXAMPLE
#   .\Initialize-CursorRules.ps1
#   Initialise les règles avec les paramètres par défaut.
#
# .EXAMPLE
#   .\Initialize-CursorRules.ps1 -Force -Environment Production
#   Initialise les règles pour l'environnement de production sans confirmation.
#
# .EXAMPLE
#   .\Initialize-CursorRules.ps1 -ConfigPath "custom/config.json" -SkipGitHooks
#   Initialise avec une configuration personnalisée sans les hooks Git.
#
# .INPUTS
#   [string] ConfigPath
#   [string] Environment
#   [switch] Force
#   [switch] SkipGitHooks
#   [switch] EnableTelemetry
#
# .OUTPUTS
#   [PSObject] Résultat de l'initialisation avec :
#   - Status : État de l'initialisation
#   - ConfigPath : Chemin de la configuration utilisée
#   - Hooks : Liste des hooks installés
#   - Extensions : Liste des extensions configurées
#
# .NOTES
#   Version     : 1.0
#   Auteur      : APEX Framework Team
#   Création    : 11/04/2024
#   Mise à jour : 11/04/2024
#   Prérequis   :
#   - PowerShell 5.1 ou supérieur
#   - Git 2.30.0 ou supérieur
#   - VS Code 1.60.0 ou supérieur
#   - Cursor installé et configuré
#
# .LINK
#   https://github.com/org/repo/wiki/Cursor-Rules
#
# .COMPONENT
#   APEX VBA Framework
#
# =============================================================================

# Validation des prérequis
#requires -Version 5.1
# Suppression de la contrainte administrateur

# Script d'initialisation des règles Cursor
[CmdletBinding()]
param (
    [string]$RulesPath = ".cursor-rules",
    [switch]$Force,
    [switch]$NoBackup
)

function Initialize-RulesDirectory {
    param (
        [string]$Path
    )
    
    Write-Host "`n📁 Initialisation du répertoire des règles..." -ForegroundColor Cyan
    
    # Création du répertoire s'il n'existe pas
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
        Write-Host "  Créé: $Path" -ForegroundColor Gray
    }
    
    # Structure de base des règles
    $baseStructure = @{
        "general"           = @{
            "description" = "Règles générales applicables à tous les fichiers"
            "rules"       = @()
        }
        "language-specific" = @{
            "description" = "Règles spécifiques aux langages"
            "rules"       = @{
                "powershell" = @()
                "javascript" = @()
                "typescript" = @()
                "python"     = @()
            }
        }
        "project-specific"  = @{
            "description" = "Règles spécifiques au projet"
            "rules"       = @()
        }
    }
    
    # Création des fichiers de règles
    $baseStructure.Keys | ForEach-Object {
        $categoryPath = Join-Path $Path "$_.json"
        if (-not (Test-Path $categoryPath) -or $Force) {
            $baseStructure[$_] | ConvertTo-Json -Depth 10 | Set-Content $categoryPath
            Write-Host "  Créé: $categoryPath" -ForegroundColor Gray
        }
    }
}

function Initialize-Configuration {
    Write-Host "`n⚙️ Configuration de l'environnement..." -ForegroundColor Cyan
    
    # 1. Configuration VS Code
    $vscodePath = ".vscode/settings.json"
    if (-not (Test-Path ".vscode")) {
        New-Item -ItemType Directory -Path ".vscode" | Out-Null
    }
    
    $vscodeSettings = @{
        "workspaceInit.tasks"                                   = @("Initialize-CursorRules.ps1")
        "powershell.scriptAnalysis.settingsPath"                = "PSScriptAnalyzerSettings.psd1"
        "powershell.debugging.createTemporaryIntegratedConsole" = $false
        "powershell.integratedConsole.suppressStartupBanner"    = $true
        "powershell.integratedConsole.focusConsoleOnExecute"    = $false
        "powershell.startAutomatically"                         = $true
        "powershell.enableProfileLoading"                       = $true
    }
    
    if (Test-Path $vscodePath) {
        $existingSettings = Get-Content $vscodePath -Raw | ConvertFrom-Json
        $mergedSettings = Merge-Objects $existingSettings $vscodeSettings
        $mergedSettings | ConvertTo-Json -Depth 10 | Set-Content $vscodePath
    }
    else {
        $vscodeSettings | ConvertTo-Json -Depth 10 | Set-Content $vscodePath
    }
    
    # 2. Création du fichier de configuration PSScriptAnalyzer
    $analyzerSettings = @{
        "IncludeRules" = @(
            "PSAvoidUsingCmdletAliases",
            "PSAvoidUsingWriteHost",
            "PSUseApprovedVerbs"
        )
        "ExcludeRules" = @()
        "Rules"        = @{
            "PSAvoidUsingCmdletAliases" = @{
                "allowList" = @()
            }
        }
    }
    
    $analyzerSettings | ConvertTo-Json -Depth 10 | 
    Set-Content "PSScriptAnalyzerSettings.psd1"
}

function Merge-Objects {
    param (
        $Object1,
        $Object2
    )
    
    $merged = $Object1.PsObject.Copy()
    foreach ($property in $Object2.PSObject.Properties) {
        if (-not $merged.PSObject.Properties[$property.Name]) {
            $merged | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
        }
    }
    return $merged
}

function Test-Initialization {
    Write-Host "`n🔍 Validation de l'initialisation..." -ForegroundColor Cyan
    $errors = @()
    
    # 1. Vérification de la structure des règles
    if (-not (Test-Path $RulesPath)) {
        $errors += "Répertoire des règles non créé"
    }
    else {
        @("general.json", "language-specific.json", "project-specific.json") | ForEach-Object {
            if (-not (Test-Path (Join-Path $RulesPath $_))) {
                $errors += "Fichier de règles manquant: $_"
            }
        }
    }
    
    # 2. Vérification de la configuration VS Code
    if (-not (Test-Path ".vscode/settings.json")) {
        $errors += "Configuration VS Code manquante"
    }
    
    # 3. Vérification de la configuration PSScriptAnalyzer
    if (-not (Test-Path "PSScriptAnalyzerSettings.psd1")) {
        $errors += "Configuration PSScriptAnalyzer manquante"
    }
    
    if ($errors.Count -gt 0) {
        Write-Warning "❌ Problèmes détectés:"
        $errors | ForEach-Object { Write-Warning "  - $_" }
        return $false
    }
    
    Write-Host "✅ Initialisation validée" -ForegroundColor Green
    return $true
}

# Exécution principale
try {
    Write-Host "==================================================="
    Write-Host "     INITIALISATION DES RÈGLES CURSOR              "
    Write-Host "==================================================="
    
    # Sauvegarde si nécessaire
    if (-not $NoBackup) {
        $backupDir = "tools/workflow/cursor/backup/init_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
        
        if (Test-Path $RulesPath) {
            Copy-Item $RulesPath $backupDir -Recurse
        }
        if (Test-Path ".vscode") {
            Copy-Item ".vscode" $backupDir -Recurse
        }
        Write-Host "📦 Sauvegarde créée dans: $backupDir" -ForegroundColor Gray
    }
    
    # Initialisation
    Initialize-RulesDirectory -Path $RulesPath
    Initialize-Configuration
    
    # Validation
    if (Test-Initialization) {
        Write-Host "`n✨ Initialisation terminée avec succès" -ForegroundColor Green
        Write-Host "Les règles Cursor sont prêtes à être utilisées" -ForegroundColor Yellow
    }
    else {
        throw "Erreurs lors de la validation de l'initialisation"
    }
}
catch {
    Write-Error "❌ Erreur lors de l'initialisation: $_"
    exit 1
}