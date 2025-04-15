. (Join-Path $PSScriptRoot "../common/Write-WorkbenchLog.ps1")

# Configuration du module
$script:ModulePath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:RootPath = (Get-Item $script:ModulePath).Parent.Parent.Parent.FullName
$script:monitoringPath = Join-Path $script:RootPath "monitoring"
$script:lastCheck = Get-Date
$script:performanceMetrics = @{}
$script:thresholds = @{
    CpuUsagePercent = 150
    MemoryUsageMB   = 200
    ResponseTimeMS  = 5000
}

# Fonctions de performance
function Initialize-MonitoringDirectory {
    if (-not (Test-Path $script:monitoringPath)) {
        New-Item -ItemType Directory -Path $script:monitoringPath -Force | Out-Null
        Write-WorkbenchLog "Création du répertoire monitoring: $script:monitoringPath" "INFO"
    }
    
    # Rotation des logs (garder les 5 derniers)
    Get-ChildItem -Path $script:monitoringPath -Filter "performance_report*.json" | 
    Sort-Object LastWriteTime -Descending | 
    Select-Object -Skip 5 | 
    Remove-Item -Force
}

function Export-PerformanceReport {
    Initialize-MonitoringDirectory
    
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $reportPath = Join-Path $script:monitoringPath "performance_report_$timestamp.json"
    
    $report = @{
        Timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Metrics    = $script:performanceMetrics
        Thresholds = $script:thresholds
        Trends     = @{
            CpuTrend    = Get-CpuTrend
            MemoryTrend = Get-MemoryTrend
        }
    }
    
    $report | ConvertTo-Json -Depth 10 | Set-Content $reportPath
    Write-WorkbenchLog "Rapport de performance exporté: $reportPath" "INFO"
}

function Get-CpuTrend {
    $lastReports = Get-ChildItem -Path $script:monitoringPath -Filter "performance_report*.json" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 3 |
    ForEach-Object { Get-Content $_.FullName | ConvertFrom-Json }
    
    if ($lastReports.Count -ge 3) {
        $trend = ($lastReports[0].Metrics.CpuUsage - $lastReports[2].Metrics.CpuUsage) / 3
        return [math]::Round($trend, 2)
    }
    return 0
}

function Get-MemoryTrend {
    $lastReports = Get-ChildItem -Path $script:monitoringPath -Filter "performance_report*.json" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 3 |
    ForEach-Object { Get-Content $_.FullName | ConvertFrom-Json }
    
    if ($lastReports.Count -ge 3) {
        $trend = ($lastReports[0].Metrics.MemoryUsage - $lastReports[2].Metrics.MemoryUsage) / 3
        return [math]::Round($trend, 2)
    }
    return 0
}