# =============================================================================
# Script d'analyse des résultats des tests d'intégration VS Code/Cursor
# =============================================================================
#
# .SYNOPSIS
#   Analyse les résultats des tests d'intégration entre VS Code et Cursor.
#
# .DESCRIPTION
#   Ce script analyse les résultats des tests d'intégration entre VS Code et Cursor.
#   Il génère des rapports détaillés et des métriques de performance.
#
# .PARAMETER ResultsPath
#   Chemin vers le dossier contenant les résultats des tests.
#   Par défaut : "tests/results"
#
# .PARAMETER GenerateReport
#   Génère un rapport détaillé au format Markdown.
#
# .PARAMETER ExportExcel
#   Exporte les résultats au format Excel (non implémenté).
#
# .EXAMPLE
#   .\Analyze-IntegrationResults.ps1
#   Analyse les résultats avec les paramètres par défaut.
#
# .EXAMPLE
#   .\Analyze-IntegrationResults.ps1 -GenerateReport
#   Analyse les résultats et génère un rapport Markdown.
#
# .NOTES
#   Version     : 1.0
#   Auteur      : APEX Framework Team
#   Création    : 11/04/2024
#   Mise à jour : 11/04/2024
#
# .LINK
#   https://github.com/org/repo/wiki/Integration-Testing
#
# =============================================================================

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$ResultsPath = "tests/results",
    [switch]$GenerateReport,
    [switch]$ExportExcel
)

function Import-TestResults {
    param (
        [string]$Path
    )
    
    Write-Host "`n📊 Importation des résultats de test..." -ForegroundColor Cyan
    
    $results = @{
        Workspace  = @()
        Extensions = @()
        Debugger   = @()
        Terminal   = @()
        Summary    = @{
            TotalTests  = 0
            PassedTests = 0
            FailedTests = 0
            SuccessRate = 0
        }
    }
    
    # Récupération des derniers fichiers de résultats
    $latestResults = Get-ChildItem $Path -Filter "*.json" | 
    Group-Object { $_.Name.Split('_')[0] } | 
    ForEach-Object { $_.Group | Sort-Object LastWriteTime -Descending | Select-Object -First 1 }
    
    foreach ($file in $latestResults) {
        $content = Get-Content $file.FullName -Raw | ConvertFrom-Json
        $testType = $file.Name.Split('_')[0]
        
        switch ($testType) {
            "Workspace" { $results.Workspace = $content }
            "Extensions" { $results.Extensions = $content }
            "Debugger" { $results.Debugger = $content }
            "Terminal" { $results.Terminal = $content }
        }
    }
    
    return $results
}

function Analyze-TestResults {
    param (
        [hashtable]$Results
    )
    
    Write-Host "`n🔍 Analyse des résultats..." -ForegroundColor Cyan
    
    $analysis = @{
        Summary = @{
            TotalTests  = 0
            PassedTests = 0
            FailedTests = 0
            SuccessRate = 0
        }
        Details = @{
            Workspace  = @{
                Status = "Non testé"
                Issues = @()
            }
            Extensions = @{
                Status = "Non testé"
                Issues = @()
            }
            Debugger   = @{
                Status = "Non testé"
                Issues = @()
            }
            Terminal   = @{
                Status = "Non testé"
                Issues = @()
            }
        }
    }
    
    # Analyse Workspace
    if ($Results.Workspace) {
        $wsSuccess = 0
        foreach ($test in $Results.Workspace) {
            $analysis.Summary.TotalTests++
            if ($test.Success) {
                $wsSuccess++
                $analysis.Summary.PassedTests++
            }
            else {
                $analysis.Details.Workspace.Issues += $test.Message
                $analysis.Summary.FailedTests++
            }
        }
        $analysis.Details.Workspace.Status = if ($wsSuccess -eq $Results.Workspace.Count) { "✅ OK" } else { "❌ Échec" }
    }
    
    # Analyse Extensions
    if ($Results.Extensions) {
        $extSuccess = 0
        foreach ($test in $Results.Extensions) {
            $analysis.Summary.TotalTests++
            if ($test.VSCode -and $test.Cursor) {
                $extSuccess++
                $analysis.Summary.PassedTests++
            }
            else {
                $analysis.Details.Extensions.Issues += "Extension $($test.Extension) non synchronisée"
                $analysis.Summary.FailedTests++
            }
        }
        $analysis.Details.Extensions.Status = if ($extSuccess -eq $Results.Extensions.Count) { "✅ OK" } else { "❌ Échec" }
    }
    
    # Analyse Debugger
    if ($Results.Debugger) {
        $dbgSuccess = 0
        foreach ($test in $Results.Debugger) {
            $analysis.Summary.TotalTests++
            if ($test.Success) {
                $dbgSuccess++
                $analysis.Summary.PassedTests++
            }
            else {
                $analysis.Details.Debugger.Issues += $test.Message
                $analysis.Summary.FailedTests++
            }
        }
        $analysis.Details.Debugger.Status = if ($dbgSuccess -eq $Results.Debugger.Count) { "✅ OK" } else { "❌ Échec" }
    }
    
    # Analyse Terminal
    if ($Results.Terminal) {
        $termSuccess = 0
        foreach ($test in $Results.Terminal) {
            $analysis.Summary.TotalTests++
            if ($test.Exists) {
                $termSuccess++
                $analysis.Summary.PassedTests++
            }
            else {
                $analysis.Details.Terminal.Issues += "Variable $($test.Variable) manquante"
                $analysis.Summary.FailedTests++
            }
        }
        $analysis.Details.Terminal.Status = if ($termSuccess -eq $Results.Terminal.Count) { "✅ OK" } else { "❌ Échec" }
    }
    
    # Calcul du taux de succès
    if ($analysis.Summary.TotalTests -gt 0) {
        $analysis.Summary.SuccessRate = [math]::Round(($analysis.Summary.PassedTests / $analysis.Summary.TotalTests) * 100, 2)
    }
    
    return $analysis
}

function Export-AnalysisReport {
    param (
        [hashtable]$Analysis,
        [string]$OutputPath = "tests/reports"
    )
    
    Write-Host "`n📝 Génération du rapport..." -ForegroundColor Cyan
    
    # Création du dossier de rapports
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $reportPath = Join-Path $OutputPath "integration_report_$timestamp.md"
    
    $report = @"
# Rapport d'Intégration VS Code/Cursor
Date: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")

## 📊 Résumé
- Tests totaux: $($Analysis.Summary.TotalTests)
- Tests réussis: $($Analysis.Summary.PassedTests)
- Tests échoués: $($Analysis.Summary.FailedTests)
- Taux de succès: $($Analysis.Summary.SuccessRate)%

## 🔍 Détails

### Workspace
Status: $($Analysis.Details.Workspace.Status)
$(if ($Analysis.Details.Workspace.Issues) {
    "`nProblèmes détectés:`n" + ($Analysis.Details.Workspace.Issues | ForEach-Object { "- $_`n" })
})

### Extensions
Status: $($Analysis.Details.Extensions.Status)
$(if ($Analysis.Details.Extensions.Issues) {
    "`nProblèmes détectés:`n" + ($Analysis.Details.Extensions.Issues | ForEach-Object { "- $_`n" })
})

### Débogueur
Status: $($Analysis.Details.Debugger.Status)
$(if ($Analysis.Details.Debugger.Issues) {
    "`nProblèmes détectés:`n" + ($Analysis.Details.Debugger.Issues | ForEach-Object { "- $_`n" })
})

### Terminal
Status: $($Analysis.Details.Terminal.Status)
$(if ($Analysis.Details.Terminal.Issues) {
    "`nProblèmes détectés:`n" + ($Analysis.Details.Terminal.Issues | ForEach-Object { "- $_`n" })
})

## 📋 Recommandations
$(if ($Analysis.Summary.FailedTests -gt 0) {
@"
1. Vérifier la configuration du workspace
2. Valider la synchronisation des extensions
3. Tester le débogueur croisé
4. Contrôler les variables d'environnement
"@
} else {
    "✅ Aucune action requise - Tous les tests sont passés"
})
"@
    
    $report | Out-File $reportPath -Encoding UTF8
    Write-Host "  Rapport généré: $reportPath" -ForegroundColor Green
    
    return $reportPath
}

# Exécution principale
try {
    Write-Host "==================================================="
    Write-Host "     ANALYSE DES RÉSULTATS D'INTÉGRATION           "
    Write-Host "==================================================="
    
    # Import des résultats
    $testResults = Import-TestResults -Path $ResultsPath
    
    # Analyse
    $analysis = Analyze-TestResults -Results $testResults
    
    # Affichage du résumé
    Write-Host "`n📊 Résumé de l'analyse" -ForegroundColor Yellow
    Write-Host "===================="
    Write-Host "Tests totaux : $($analysis.Summary.TotalTests)"
    Write-Host "Tests réussis: $($analysis.Summary.PassedTests)" -ForegroundColor Green
    Write-Host "Tests échoués: $($analysis.Summary.FailedTests)" -ForegroundColor Red
    Write-Host "Taux de succès: $($analysis.Summary.SuccessRate)%" -ForegroundColor $(if ($analysis.Summary.SuccessRate -ge 80) { "Green" } else { "Yellow" })
    
    # Génération du rapport si demandé
    if ($GenerateReport) {
        $reportPath = Export-AnalysisReport -Analysis $analysis
    }
    
    # Export Excel si demandé
    if ($ExportExcel) {
        # TODO: Implémenter l'export Excel
        Write-Warning "Export Excel non implémenté"
    }
    
    # Statut de sortie
    if ($analysis.Summary.FailedTests -gt 0) {
        Write-Warning "❌ Des problèmes ont été détectés. Consultez le rapport pour plus de détails."
        exit 1
    }
    else {
        Write-Host "`n✨ Validation réussie" -ForegroundColor Green
    }
}
catch {
    Write-Error "❌ Erreur lors de l'analyse: $_"
    exit 1
} 