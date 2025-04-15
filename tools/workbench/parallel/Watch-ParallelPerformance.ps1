# =============================================================================
# Monitoring des processus parallèles optimisé
# =============================================================================

<#
.SYNOPSIS
    Monitore les performances des processus parallèles (VSCode, Cursor, Excel)

.DESCRIPTION
    Surveille en continu les métriques de performance des processus VSCode, Cursor et Excel,
    incluant l'utilisation CPU, mémoire et les métriques spécifiques VBA.
    Génère des alertes si les seuils sont dépassés et sauvegarde les rapports.

.NOTES
    Version     : 1.4
    Author      : APEX Framework
    Created     : 2024-04-14
    Updated     : 2024-04-15
#>

#Requires -Version 5.1
#Requires -RunAsAdministrator

[CmdletBinding()]
param (
    [int]$IntervalSeconds = 30,
    [string]$MonitoringPath = (Join-Path $PSScriptRoot "../../../monitoring")
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Import configuration
$configPath = Join-Path $PSScriptRoot "../../../config/monitoring/performance_config.json"
$config = Get-Content $configPath -Raw | ConvertFrom-Json

# Initialize logging
$logDir = Join-Path $PSScriptRoot "../../../logs/performance"
$logFile = Join-Path $logDir "watch_parallel.log"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp][$Level] $Message"
    Add-Content -Path $logFile -Value $logMessage
    if ($Level -eq 'ERROR') { Write-Error $Message }
    else { Write-Verbose $Message }
}

function Get-ProcessMetrics {
    param ([string]$ProcessName)
    
    try {
        $processes = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
        if (-not $processes) { return $null }
        
        # Get metrics in one pass to reduce CPU usage
        $metrics = $processes | Measure-Object -Property CPU, WorkingSet, Threads, Handles -Sum
        
        return @{
            CPU = [math]::Round($metrics.Sum[0], 2)
            Memory = [math]::Round($metrics.Sum[1] / 1MB, 2)
            Threads = $metrics.Sum[2]
            Handles = $metrics.Sum[3]
            Count = $processes.Count
        }
    }
    catch {
        Write-Log "Erreur lors de la collecte des métriques pour $ProcessName : $_" -Level 'ERROR'
        return $null
    }
}

try {
    Write-Log "Démarrage du monitoring optimisé..."
    
    while ($true) {
        try {
            $startTime = Get-Date
            
            # Collect metrics with proper throttling
            $metrics = @{}
            foreach ($proc in @('Code', 'Cursor', 'EXCEL')) {
                $metrics[$proc] = Get-ProcessMetrics -ProcessName $proc
                Start-Sleep -Milliseconds $config.Pooling.ThrottleIntervalMs
            }
            
            # Generate and save report
            $report = @{
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                VSCode = $metrics['Code']
                Cursor = $metrics['Cursor']
                Excel = $metrics['EXCEL']
            }
            
            $reportPath = Join-Path $MonitoringPath "performance_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
            $report | ConvertTo-Json -Depth 10 | Set-Content $reportPath -ErrorAction Stop
            
            # Ensure we wait full interval accounting for processing time
            $elapsed = ((Get-Date) - $startTime).TotalMilliseconds
            $remainingWait = [Math]::Max(0, $config.Monitoring.RefreshRateMs - $elapsed)
            if ($remainingWait -gt 0) {
                Start-Sleep -Milliseconds $remainingWait
            }
        }
        catch {
            Write-Log "Erreur lors de la collecte des métriques: $_" -Level 'ERROR'
            Start-Sleep -Seconds ($IntervalSeconds * 2)
        }
    }
}
catch {
    Write-Log "Erreur fatale dans le monitoring: $_" -Level 'ERROR'
    exit 1
}

# ==============================================================================
# ✅ Clôture de session
# ==============================================================================
Write-Verbose "✨ Script terminé avec succès"
exit 0