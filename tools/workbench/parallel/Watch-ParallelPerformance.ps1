# =============================================================================
# 🧭 Session de travail – 2024-04-14
# =============================================================================

<#
.SYNOPSIS
    Monitore les performances des processus parallèles (VSCode, Cursor, Excel)

.DESCRIPTION
    Surveille en continu les métriques de performance des processus VSCode, Cursor et Excel,
    incluant l'utilisation CPU, mémoire et les métriques spécifiques VBA.
    Génère des alertes si les seuils sont dépassés et sauvegarde les rapports.

.NOTES
    Version     : 1.2
    Author      : APEX Framework
    Created     : 2024-04-14
    Updated     : 2024-04-14
#>

#Requires -Version 5.1
#Requires -RunAsAdministrator

[CmdletBinding()]
param (
    [int]$IntervalSeconds = 30,
    [int]$CpuThreshold = 80,
    [int]$MemoryThreshold = 2048,
    [string]$MonitoringPath = (Join-Path $PSScriptRoot "../../../monitoring")
)

# ==============================================================================
# 🎯 Objectif(s)
# ==============================================================================
# - Monitorer les performances des processus parallèles
# - Générer des alertes en cas de dépassement de seuils
# - Sauvegarder les métriques pour analyse

# ==============================================================================
# 📌 Suivi des tâches
# ==============================================================================
<#
| Tâche | Module | Statut | Commentaire |
|-------|--------|--------|-------------|
| Monitoring CPU/RAM | Performance | ✅ | Implémenté |
| Métriques VBA | Excel | ✅ | Ajouté |
| Alertes | Monitoring | ✅ | Configuré |
#>

# ==============================================================================
# 🔄 Initialisation
# ==============================================================================
$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'

# Vérification et création du dossier de monitoring
try {
    if (-not (Test-Path $MonitoringPath)) {
        Write-Verbose "📁 Création du dossier de monitoring : $MonitoringPath"
        New-Item -ItemType Directory -Path $MonitoringPath -Force -ErrorAction Stop | Out-Null
        Write-Verbose "✅ Dossier de monitoring créé avec succès"
    }
    
    # Test des permissions d'écriture
    $testFile = Join-Path $MonitoringPath "test.tmp"
    try {
        [System.IO.File]::WriteAllText($testFile, "Test")
        Remove-Item $testFile -Force
        Write-Verbose "✅ Permissions d'écriture vérifiées"
    }
    catch {
        throw "❌ Permissions insuffisantes sur le dossier de monitoring : $_"
    }
}
catch {
    Write-Error "❌ Erreur lors de l'initialisation du dossier de monitoring : $_"
    exit 1
}

$alertLogPath = Join-Path $MonitoringPath "performance_alerts.log"

# ==============================================================================
# 📋 Fonctions
# ==============================================================================

function Get-VbaMetric {
    param (
        [string]$MetricName
    )
    
    try {
        # Récupération des métriques depuis le fichier de monitoring VBA
        $vbaMetricsPath = Join-Path $MonitoringPath "vba_metrics.json"
        if (Test-Path $vbaMetricsPath) {
            $vbaMetrics = Get-Content $vbaMetricsPath | ConvertFrom-Json
            return $vbaMetrics.$MetricName
        }
        return 0
    }
    catch {
        Write-Warning "Impossible de récupérer la métrique VBA $MetricName : $_"
        return 0
    }
}

function Get-ProcessMetrics {
    param (
        [string]$ProcessName
    )
    
    try {
        $process = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
        if (-not $process) {
            return @{
                IsRunning    = $false
                CpuPercent   = 0
                MemoryMB     = 0
                ThreadCount  = 0
                HandleCount  = 0
                ProcessCount = 0
            }
        }

        # Calcul CPU %
        $cpuPercent = [math]::Round(($process | Measure-Object -Property CPU -Sum).Sum, 2)
        
        # Métriques de base
        $metrics = @{
            IsRunning    = $true
            CpuPercent   = $cpuPercent
            MemoryMB     = [math]::Round($process.WorkingSet64 / 1MB, 2)
            ThreadCount  = ($process | Measure-Object -Property Threads -Sum).Sum
            HandleCount  = ($process | Measure-Object -Property Handles -Sum).Sum
            ProcessCount = ($process | Measure-Object).Count
        }

        # Métriques spécifiques Excel/VBA si applicable
        if ($ProcessName -eq "EXCEL") {
            $metrics += @{
                VbaCallTime    = Get-VbaMetric -MetricName "CallTime"
                VbaMemoryUsage = Get-VbaMetric -MetricName "MemoryUsage"
                VbaLastCall    = Get-VbaMetric -MetricName "LastCall"
                VbaErrorCount  = Get-VbaMetric -MetricName "ErrorCount"
            }
        }

        return $metrics
    }
    catch {
        Write-Error "Erreur lors de la collecte des métriques pour $ProcessName : $_"
        return $null
    }
}

# ==============================================================================
# 🚀 Exécution principale
# ==============================================================================
try {
    Write-Verbose "🔄 Démarrage du monitoring des performances..."
    
    while ($true) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $report = @{
            Timestamp = $timestamp
            VSCode    = Get-ProcessMetrics -ProcessName "Code"
            Cursor    = Get-ProcessMetrics -ProcessName "Cursor"
            Excel     = Get-ProcessMetrics -ProcessName "EXCEL"
        }

        # Vérification des seuils et alertes
        foreach ($app in @("VSCode", "Cursor", "Excel")) {
            $metrics = $report[$app]
            if ($metrics.IsRunning) {
                if ($metrics.CpuPercent -gt $CpuThreshold) {
                    $alert = "[$timestamp] ⚠️ $app : CPU élevé ($($metrics.CpuPercent)%)"
                    Add-Content -Path $alertLogPath -Value $alert
                }
                if ($metrics.MemoryMB -gt $MemoryThreshold) {
                    $alert = "[$timestamp] ⚠️ $app : Mémoire élevée ($($metrics.MemoryMB) MB)"
                    Add-Content -Path $alertLogPath -Value $alert
                }
            }
        }

        # Sauvegarde du rapport
        $reportPath = Join-Path $MonitoringPath "performance_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
        $report | ConvertTo-Json -Depth 10 | Out-File $reportPath

        Start-Sleep -Seconds $IntervalSeconds
    }
}
catch {
    Write-Error "❌ Erreur dans le monitoring : $_"
    exit 1
}

# ==============================================================================
# ✅ Clôture de session
# ==============================================================================
Write-Verbose "✨ Script terminé avec succès"
exit 0 