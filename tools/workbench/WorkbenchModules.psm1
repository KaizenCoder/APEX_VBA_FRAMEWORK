# =============================================================================
# APEX Workbench - Module Principal
# =============================================================================

# Import des modules requis
Import-Module (Join-Path $PSScriptRoot "../powershell/PerformanceMonitoring.psm1") -Force
Import-Module (Join-Path $PSScriptRoot "../../config/monitoring/performance_config.psm1") -Force

# Variables globales
$script:modulePath = $PSScriptRoot
$script:logFile = Join-Path $PSScriptRoot "../logs/workbench.log"
$script:monitoringPath = Join-Path $PSScriptRoot "../monitoring"
$script:config = Get-PerformanceConfig
$script:lastCheck = Get-Date
$script:performanceMetrics = @{}

# Fonctions de logging
function Initialize-LogDirectory {
    if (-not (Test-Path $script:logFile)) {
        $logDir = Split-Path $script:logFile -Parent
        if (-not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
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
    
    $processes = Get-Process | Where-Object { $_.ProcessName -like "*$ProcessName*" }
    if ($processes) {
        return @{
            ProcessName  = $ProcessName
            CPU          = ($processes | Measure-Object -Property CPU -Sum).Sum
            WorkingSet   = ($processes | Measure-Object -Property WorkingSet -Sum).Sum
            MemoryMB     = [math]::Round(($processes | Measure-Object -Property WorkingSet -Sum).Sum / 1MB, 2)
            ThreadCount  = ($processes | Measure-Object -Property Threads -Sum).Sum
            HandleCount  = ($processes | Measure-Object -Property Handles -Sum).Sum
            ProcessCount = ($processes | Measure-Object).Count
        }
    }
    return $null
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
        $thresholds = $script:config.Thresholds
        
        # Vérification CPU
        if ($metrics.CPU -gt $thresholds.CpuUsagePercent) {
            Write-WorkbenchLog "Alerte CPU élevé pour $agent : $($metrics.CPU)%" "WARNING"
        }
        
        # Vérification Mémoire
        if ($metrics.Memory -gt $thresholds.MemoryUsageMB) {
            Write-WorkbenchLog "Alerte mémoire élevée pour $agent : $($metrics.Memory)MB" "WARNING"
        }
        
        # Vérification Temps de réponse
        $responseTime = ((Get-Date) - $metrics.LastUpdate).TotalMilliseconds
        if ($responseTime -gt $thresholds.ResponseTimeMS) {
            Write-WorkbenchLog "Alerte temps de réponse pour $agent : $($responseTime)ms" "WARNING"
        }
    }
}

function Export-PerformanceReport {
    if (-not (Test-Path $script:monitoringPath)) {
        New-Item -ItemType Directory -Path $script:monitoringPath -Force | Out-Null
        Write-WorkbenchLog "Création du répertoire monitoring: $script:monitoringPath" "INFO"
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $reportPath = Join-Path $script:monitoringPath "performance_report_$timestamp.json"
    $report = @{
        Timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Metrics    = $script:performanceMetrics
        Thresholds = $script:config.Thresholds
        Config     = $script:config.Monitoring
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
            
            # Export du rapport selon la configuration
            $timeSinceLastCheck = (Get-Date) - $script:lastCheck
            if ($timeSinceLastCheck.TotalMilliseconds -ge $script:config.Monitoring.RefreshRateMs) {
                Export-PerformanceReport
                $script:lastCheck = Get-Date
            }
            
            # Pause selon la configuration
            Start-Sleep -Milliseconds $script:config.Monitoring.RefreshRateMs
        }
        catch {
            Write-WorkbenchLog "Erreur dans la surveillance des performances: $_" "ERROR"
            Start-Sleep -Seconds 60  # Délai plus long en cas d'erreur
        }
    }
}

# Export des fonctions
Export-ModuleMember -Function @(
    'Write-WorkbenchLog',
    'Start-PerformanceMonitoring',
    'Get-ProcessPerformance',
    'Update-PerformanceMetrics',
    'Test-PerformanceThresholds',
    'Export-PerformanceReport'
)