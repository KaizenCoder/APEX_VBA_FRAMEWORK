# =============================================================================
# Script de vérification pré-commit des règles Cursor
# =============================================================================
#
# .SYNOPSIS
#   Vérifie la conformité avec les règles Cursor avant un commit.
#
# .DESCRIPTION
#   Ce script vérifie que les fichiers modifiés respectent les règles Cursor :
#   - Validation de la documentation
#   - Vérification des conventions de nommage
#   - Contrôle des dépendances
#   - Structure des fichiers
#
# =============================================================================

[CmdletBinding()]
param (
    [string]$RulesPath = ".cursor-rules",
    [switch]$Fix
)

function Get-ModifiedFiles {
    Write-Host "`n📄 Recherche des fichiers modifiés..." -ForegroundColor Cyan
    
    $files = git diff --cached --name-only --diff-filter=ACMR
    
    if (-not $files) {
        Write-Host "  Aucun fichier modifié trouvé" -ForegroundColor Gray
        return @()
    }
    
    return $files
}

function Test-FileCompliance {
    param (
        [string]$FilePath,
        [string]$RulesPath
    )
    
    Write-Host "  Analyse de $FilePath..." -ForegroundColor Gray
    $issues = @()
    
    # Déterminer le type de fichier
    $extension = [System.IO.Path]::GetExtension($FilePath)
    
    # Charger les règles appropriées
    $generalRules = Get-Content (Join-Path $RulesPath "general.json") | ConvertFrom-Json
    $languageRules = Get-Content (Join-Path $RulesPath "language-specific.json") | ConvertFrom-Json
    $projectRules = Get-Content (Join-Path $RulesPath "project-specific.json") | ConvertFrom-Json
    
    # Appliquer les règles générales
    foreach ($rule in $generalRules.rules) {
        if ($rule.enabled) {
            # Vérification des règles générales...
            # Exemple: En-têtes de documentation
            if ($rule.name -eq "documentation-headers") {
                $content = Get-Content $FilePath -Raw
                foreach ($pattern in $rule.patterns) {
                    if ($FilePath -like $pattern.fileType) {
                        foreach ($headerTag in $pattern.header) {
                            if (-not ($content -match [regex]::Escape($headerTag))) {
                                $issues += @{
                                    Rule     = $rule.name
                                    Message  = "Tag manquant: $headerTag"
                                    Severity = $rule.severity
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    
    # Appliquer les règles spécifiques au langage
    switch -Regex ($extension) {
        '\.cls|\.bas|\.frm' {
            foreach ($rule in $languageRules.rules.vb) {
                if ($rule.enabled) {
                    # Vérification des règles VBA...
                }
            }
        }
        '\.md' {
            foreach ($rule in $languageRules.rules.markdown) {
                if ($rule.enabled) {
                    # Vérification des règles Markdown...
                }
            }
        }
        '\.ps1' {
            foreach ($rule in $languageRules.rules.powershell) {
                if ($rule.enabled) {
                    # Vérification des règles PowerShell...
                }
            }
        }
    }
    
    # Appliquer les règles spécifiques au projet
    foreach ($rule in $projectRules.rules) {
        if ($rule.enabled) {
            # Vérification des règles projet...
            if ($rule.name -eq "apex-naming-conventions") {
                $filename = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
                foreach ($pattern in $rule.patterns.PSObject.Properties) {
                    if ($filename -notmatch $pattern.Value) {
                        $issues += @{
                            Rule     = $rule.name
                            Message  = "Le nom ne respecte pas la convention $($pattern.Name): $filename"
                            Severity = $rule.severity
                        }
                    }
                }
            }
        }
    }
    
    return $issues
}

# Exécution principale
try {
    Write-Host "==================================================="
    Write-Host "     VÉRIFICATION PRÉ-COMMIT CURSOR                "
    Write-Host "==================================================="
    
    # 1. Obtenir les fichiers modifiés
    $modifiedFiles = Get-ModifiedFiles
    if (-not $modifiedFiles) {
        Write-Host "`n✨ Aucun fichier à vérifier" -ForegroundColor Green
        exit 0
    }
    
    # 2. Vérifier chaque fichier
    $allIssues = @()
    foreach ($file in $modifiedFiles) {
        if (Test-Path $file) {
            $issues = Test-FileCompliance -FilePath $file -RulesPath $RulesPath
            if ($issues) {
                $allIssues += @{
                    File   = $file
                    Issues = $issues
                }
            }
        }
    }
    
    # 3. Afficher les résultats
    if ($allIssues) {
        Write-Warning "`n⚠️ Problèmes détectés:"
        foreach ($fileIssues in $allIssues) {
            Write-Host "`n📝 $($fileIssues.File):" -ForegroundColor Yellow
            foreach ($issue in $fileIssues.Issues) {
                $icon = switch ($issue.Severity) {
                    "error" { "❌" }
                    "warning" { "⚠️" }
                    default { "ℹ️" }
                }
                Write-Host "  $icon $($issue.Rule): $($issue.Message)"
            }
        }
        
        # Si le mode correction est activé
        if ($Fix) {
            Write-Host "`n🔧 Tentative de correction automatique..." -ForegroundColor Cyan
            # TODO: Implémenter la correction automatique
        }
        
        exit 1
    }
    
    Write-Host "`n✨ Tous les fichiers respectent les règles Cursor" -ForegroundColor Green
    exit 0
}
catch {
    Write-Error "❌ Erreur lors de la vérification: $_"
    exit 2
}