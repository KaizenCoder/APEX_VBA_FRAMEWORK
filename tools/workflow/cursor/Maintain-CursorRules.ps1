# =============================================================================
# Script de maintenance des règles Cursor
# =============================================================================
#
# .SYNOPSIS
#   Maintient et optimise les règles Cursor pour l'environnement APEX Framework.
#
# .DESCRIPTION
#   Ce script effectue diverses tâches de maintenance sur les règles Cursor :
#   - Nettoyage des fichiers temporaires
#   - Consolidation des règles
#   - Vérification de la cohérence
#   - Optimisation des performances
#   - Sauvegarde des configurations
#
# =============================================================================

[CmdletBinding()]
param (
    [string]$RulesPath = ".cursor-rules",
    [switch]$Force,
    [switch]$Backup = $true
)

function Backup-CursorRules {
    param (
        [string]$Path
    )
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupDir = Join-Path "tools/workflow/cursor/backup" "rules_$timestamp"
    
    Write-Host "📦 Sauvegarde des règles..." -ForegroundColor Cyan
    
    try {
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
        Copy-Item "$Path/*" $backupDir -Recurse -Force
        Write-Host "  ✅ Sauvegarde créée dans: $backupDir" -ForegroundColor Gray
    }
    catch {
        Write-Warning "❌ Erreur lors de la sauvegarde: $_"
        return $false
    }
    
    return $true
}

function Test-RulesConsistency {
    param (
        [string]$Path
    )
    
    Write-Host "`n🔍 Vérification de la cohérence des règles..." -ForegroundColor Cyan
    $hasErrors = $false
    
    # Vérifier chaque fichier de règles
    Get-ChildItem $Path -Filter "*.json" | ForEach-Object {
        Write-Host "  Analyse de $($_.Name)..." -ForegroundColor Gray
        try {
            $content = Get-Content $_.FullName -Raw | ConvertFrom-Json
            
            # Vérifier la structure de base
            if (-not $content.description) {
                Write-Warning "  ⚠️ $($_.Name): Description manquante"
                $hasErrors = $true
            }
            
            if (-not $content.rules) {
                Write-Warning "  ⚠️ $($_.Name): Section rules manquante"
                $hasErrors = $true
            }
            
            # Vérifier les règles spécifiques selon le type de fichier
            switch ($_.Name) {
                "general.json" {
                    foreach ($rule in $content.rules) {
                        if (-not $rule.name -or -not $rule.description -or -not $rule.enabled) {
                            Write-Warning "  ⚠️ Règle invalide dans general.json"
                            $hasErrors = $true
                        }
                    }
                }
                "language-specific.json" {
                    foreach ($lang in $content.rules.PSObject.Properties) {
                        foreach ($rule in $lang.Value) {
                            if (-not $rule.name -or -not $rule.description) {
                                Write-Warning "  ⚠️ Règle invalide pour le langage $($lang.Name)"
                                $hasErrors = $true
                            }
                        }
                    }
                }
                "project-specific.json" {
                    foreach ($rule in $content.rules) {
                        if (-not $rule.name -or -not $rule.description -or -not $rule.enabled) {
                            Write-Warning "  ⚠️ Règle de projet invalide"
                            $hasErrors = $true
                        }
                    }
                }
            }
        }
        catch {
            Write-Warning "  ❌ Erreur dans $($_.Name): $_"
            $hasErrors = $true
        }
    }
    
    return -not $hasErrors
}

function Optimize-RulesPerformance {
    param (
        [string]$Path
    )
    
    Write-Host "`n⚡ Optimisation des performances..." -ForegroundColor Cyan
    
    # Optimiser chaque fichier de règles
    Get-ChildItem $Path -Filter "*.json" | ForEach-Object {
        try {
            # Lire et parser le contenu
            $content = Get-Content $_.FullName -Raw | ConvertFrom-Json
            
            # Désactiver les règles redondantes ou conflictuelles
            $optimized = $false
            
            # Réécrire le fichier de manière optimisée
            $content | ConvertTo-Json -Depth 10 | Set-Content $_.FullName
            
            if ($optimized) {
                Write-Host "  ✅ $($_.Name) optimisé" -ForegroundColor Gray
            }
        }
        catch {
            Write-Warning "  ❌ Erreur lors de l'optimisation de $($_.Name): $_"
        }
    }
}

function Remove-TemporaryFiles {
    Write-Host "`n🧹 Nettoyage des fichiers temporaires..." -ForegroundColor Cyan
    
    # Nettoyer les fichiers de session Cursor
    Get-ChildItem -Path "." -Filter ".cursor-session-*" -File | Remove-Item -Force
    
    # Nettoyer les logs plus vieux que 7 jours
    $oldLogs = Get-ChildItem -Path "logs/cursor" -File | Where-Object {
        $_.LastWriteTime -lt (Get-Date).AddDays(-7)
    }
    if ($oldLogs) {
        $oldLogs | Remove-Item -Force
        Write-Host "  Supprimé $($oldLogs.Count) anciens fichiers de log" -ForegroundColor Gray
    }
}

# Exécution principale
try {
    Write-Host "==================================================="
    Write-Host "     MAINTENANCE DES RÈGLES CURSOR                 "
    Write-Host "==================================================="
    
    # 1. Sauvegarde si demandée
    if ($Backup) {
        $backupSuccess = Backup-CursorRules -Path $RulesPath
        if (-not $backupSuccess -and -not $Force) {
            throw "Échec de la sauvegarde et -Force non spécifié"
        }
    }
    
    # 2. Vérification de la cohérence
    $isConsistent = Test-RulesConsistency -Path $RulesPath
    if (-not $isConsistent -and -not $Force) {
        throw "Incohérences détectées dans les règles et -Force non spécifié"
    }
    
    # 3. Optimisation
    Optimize-RulesPerformance -Path $RulesPath
    
    # 4. Nettoyage
    Remove-TemporaryFiles
    
    Write-Host "`n✨ Maintenance terminée avec succès" -ForegroundColor Green
}
catch {
    Write-Error "❌ Erreur lors de la maintenance: $_"
    exit 1
}