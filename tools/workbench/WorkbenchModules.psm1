# =============================================================================
# APEX Workbench - Module Principal
# =============================================================================

# Variables globales
$script:modulePath = $PSScriptRoot
$script:logFile = Join-Path $PSScriptRoot "../logs/workbench.log"
$script:monitoringPath = Join-Path $PSScriptRoot "../monitoring"

# Fonctions de logging
function Initialize-LogDirectory {
    $logDir = Split-Path $script:logFile -Parent
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
}

function Write-WorkbenchLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    
    Initialize-LogDirectory
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] $Message"
    
    # Écriture dans le fichier
    Add-Content -Path $script:logFile -Value $logEntry -Encoding UTF8
    
    # Affichage console avec couleur
    $color = switch ($Level) {
        "INFO" { "White" }
        "WARNING" { "Yellow" }
        "ERROR" { "Red" }
    }
    Write-Host $logEntry -ForegroundColor $color
}

# Fonctions de performance
function Get-ProcessPerformance {
    param([string]$ProcessName)
    
    Get-Process | Where-Object { $_.ProcessName -like "*$ProcessName*" } | 
    Select-Object ProcessName, CPU, WorkingSet, @{
        Name       = 'MemoryMB'
        Expression = { [math]::Round($_.WorkingSet / 1MB, 2) }
    }
}

function Update-PerformanceMetrics {
    # Métriques Cursor
    $cursorPerf = Get-ProcessPerformance "cursor"
    if ($cursorPerf) {
        $script:performanceMetrics["Cursor"] = @{
            CPU        = $cursorPerf.CPU
            Memory     = $cursorPerf.MemoryMB
            LastUpdate = Get-Date
        }
    }
    
    # Métriques VSCode
    $vscodePerf = Get-ProcessPerformance "code"
    if ($vscodePerf) {
        $script:performanceMetrics["VSCode"] = @{
            CPU        = $vscodePerf.CPU
            Memory     = $vscodePerf.MemoryMB
            LastUpdate = Get-Date
        }
    }
}

function Test-PerformanceThresholds {
    foreach ($agent in $script:performanceMetrics.Keys) {
        $metrics = $script:performanceMetrics[$agent]
        
        # Vérification CPU
        if ($metrics.CPU -gt $script:thresholds.CpuUsagePercent) {
            Write-WorkbenchLog "Alerte CPU élevé pour $agent : $($metrics.CPU)%" "WARNING"
        }
        
        # Vérification Mémoire
        if ($metrics.Memory -gt $script:thresholds.MemoryUsageMB) {
            Write-WorkbenchLog "Alerte mémoire élevée pour $agent : $($metrics.Memory)MB" "WARNING"
        }
        
        # Vérification Temps de réponse
        $responseTime = ((Get-Date) - $metrics.LastUpdate).TotalMilliseconds
        if ($responseTime -gt $script:thresholds.ResponseTimeMS) {
            Write-WorkbenchLog "Alerte temps de réponse pour $agent : $($responseTime)ms" "WARNING"
        }
    }
}

function Export-PerformanceReport {
    if (-not (Test-Path $script:monitoringPath)) {
        New-Item -ItemType Directory -Path $script:monitoringPath -Force | Out-Null
        Write-WorkbenchLog "Création du répertoire monitoring: $script:monitoringPath" "INFO"
    }
    
    $reportPath = Join-Path $script:monitoringPath "performance_report.json"
    $report = @{
        Timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Metrics    = $script:performanceMetrics
        Thresholds = $script:thresholds
    }
    
    $report | ConvertTo-Json -Depth 10 | Set-Content $reportPath
    Write-WorkbenchLog "Rapport de performance exporté: $reportPath" "INFO"
}

function Start-PerformanceMonitoring {
    Write-WorkbenchLog "Démarrage de la surveillance des performances" "INFO"
    
    while ($true) {
        try {
            Update-PerformanceMetrics
            Test-PerformanceThresholds
            
            # Export du rapport toutes les 5 minutes
            $timeSinceLastCheck = (Get-Date) - $script:lastCheck
            if ($timeSinceLastCheck.TotalMinutes -ge 5) {
                Export-PerformanceReport
                $script:lastCheck = Get-Date
            }
            
            Start-Sleep -Seconds 30
        }
        catch {
            Write-WorkbenchLog "Erreur dans la surveillance des performances: $_" "ERROR"
            Start-Sleep -Seconds 60  # Délai plus long en cas d'erreur
        }
    }
}

# Initialisation des variables globales
$script:lastCheck = Get-Date
$script:performanceMetrics = @{}
$script:thresholds = @{
    CpuUsagePercent = 30
    MemoryUsageMB   = 1000
    ResponseTimeMS  = 2000
}

# Export des fonctions
Export-ModuleMember -Function @(
    'Write-WorkbenchLog',
    'Start-PerformanceMonitoring'
) 