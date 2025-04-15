# =============================================================================
# Module de monitoring de performance pour APEX Framework
# =============================================================================

$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'

# Configuration du module
$script:ModulePath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:RootPath = (Get-Item $script:ModulePath).Parent.Parent.FullName
$script:LogsPath = Join-Path -Path $script:RootPath -ChildPath "logs\performance"
$script:DefaultLogFile = Join-Path -Path $script:LogsPath -ChildPath "vscode_performance.log"
$script:HistoryFile = Join-Path -Path $script:LogsPath -ChildPath "performance_history.json"
$script:ConfigFile = Join-Path -Path $script:RootPath -ChildPath "config\monitoring\performance_config.json"

# Seuils de performance standardisés
$script:StandardThresholds = @{
    CpuUsagePercent = 2      # Objectif < 2%
    MemoryUsageMB   = 100      # Objectif < 100MB
    ResponseTimeMS  = 300     # Objectif < 300ms
}

function Initialize-LogEnvironment {
    if (-not (Test-Path -Path $script:LogsPath)) {
        New-Item -Path $script:LogsPath -ItemType Directory -Force | Out-Null
    }
    
    if (-not (Test-Path -Path $script:ConfigFile)) {
        $defaultConfig = @{
            LogLevel          = "INFO"
            MaxLogSizeMB      = 10
            EnableRotation    = $true
            KeepLogDays       = 7
            MonitoringEnabled = $true
            RefreshRateMs     = 5000
            Thresholds        = $script:StandardThresholds
        }
        $defaultConfig | ConvertTo-Json | Set-Content -Path $script:ConfigFile -Encoding UTF8
    }
}

function Write-PerformanceLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'DEBUG')]
        [string]$Level,
        
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter()]
        [string]$Module = 'VSCode',
        
        [Parameter()]
        [hashtable]$Metrics
    )
    
    Initialize-LogEnvironment
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $metricsJson = if ($Metrics) { " " + ($Metrics | ConvertTo-Json -Compress) } else { "" }
    $logEntry = "[$timestamp][$Level][$Module] $Message$metricsJson"
    
    Add-Content -Path $script:DefaultLogFile -Value $logEntry -Encoding UTF8
    
    switch ($Level) {
        'ERROR' { Write-Host $logEntry -ForegroundColor Red }
        'WARNING' { Write-Host $logEntry -ForegroundColor Yellow }
        'INFO' { Write-Verbose $logEntry }
        'DEBUG' { Write-Debug $logEntry }
    }
    
    if ($Metrics -and $Level -ne 'DEBUG') {
        Add-PerformanceHistory -Level $Level -Module $Module -Metrics $Metrics
    }
}

function Add-PerformanceHistory {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Level,
        
        [Parameter(Mandatory = $true)]
        [string]$Module,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$Metrics
    )
    
    $history = @()
    if (Test-Path -Path $script:HistoryFile) {
        try {
            $history = Get-Content -Path $script:HistoryFile -Raw | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            Write-Warning "Impossible de charger l'historique. Un nouveau fichier sera créé."
        }
    }
    
    if ($history.Count -gt 1000) {
        $history = $history | Select-Object -Skip ($history.Count - 1000)
    }
    
    $entry = [PSCustomObject]@{
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Level     = $Level
        Module    = $Module
        Metrics   = $Metrics
    }
    
    $history += $entry
    $history | ConvertTo-Json | Set-Content -Path $script:HistoryFile -Encoding UTF8
}

function Invoke-LogRotation {
    [CmdletBinding()]
    param()
    
    try {
        $config = Get-Content -Path $script:ConfigFile -Raw | ConvertFrom-Json
        
        if (-not $config.EnableRotation) {
            return
        }
        
        $logFile = $script:DefaultLogFile
        if (-not (Test-Path -Path $logFile)) {
            return
        }
        
        $logFileInfo = Get-Item -Path $logFile
        $maxSizeBytes = $config.MaxLogSizeMB * 1MB
        
        if ($logFileInfo.Length -gt $maxSizeBytes) {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $archivePath = Join-Path -Path $script:LogsPath -ChildPath "archive"
            
            if (-not (Test-Path -Path $archivePath)) {
                New-Item -Path $archivePath -ItemType Directory -Force | Out-Null
            }
            
            $archiveFile = Join-Path -Path $archivePath -ChildPath "vscode_performance_$timestamp.log"
            Move-Item -Path $logFile -Destination $archiveFile
            Write-Verbose "Rotation du log effectuée: $archiveFile"
            
            # Nettoyage des anciens logs
            $oldLogs = Get-ChildItem -Path $archivePath -Filter "*.log" | 
            Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$config.KeepLogDays) }
            
            foreach ($oldLog in $oldLogs) {
                Remove-Item -Path $oldLog.FullName -Force
                Write-Verbose "Ancien log supprimé: $($oldLog.FullName)"
            }
        }
    }
    catch {
        Write-Error "Erreur lors de la rotation des logs: $_"
    }
}

function Get-SystemPerformanceMetrics {
    [CmdletBinding()]
    param()
    
    $metrics = @{}
    
    try {
        $cpuLoad = (Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction Stop).CounterSamples.CookedValue
        $metrics.Add("CPU", [math]::Round($cpuLoad, 2))
        
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        $memoryUsed = $os.TotalVisibleMemorySize - $os.FreePhysicalMemory
        $metrics.Add("MemoryMB", [math]::Round($memoryUsed / 1024, 2))
        $metrics.Add("MemoryPct", [math]::Round($memoryUsed * 100 / $os.TotalVisibleMemorySize, 2))
        
        $vsCodeProcesses = Get-Process | Where-Object { $_.Name -like "*code*" }
        if ($vsCodeProcesses) {
            $vscodeMemory = ($vsCodeProcesses | Measure-Object -Property WorkingSet -Sum).Sum / 1MB
            $metrics.Add("VSCodeMemoryMB", [math]::Round($vscodeMemory, 2))
        }
        
        $cursorProcesses = Get-Process | Where-Object { $_.Name -like "*cursor*" }
        if ($cursorProcesses) {
            $cursorMemory = ($cursorProcesses | Measure-Object -Property WorkingSet -Sum).Sum / 1MB
            $metrics.Add("CursorMemoryMB", [math]::Round($cursorMemory, 2))
        }
    }
    catch {
        Write-Warning "Erreur lors de la collecte des métriques: $_"
    }
    
    return $metrics
}

# Export des fonctions
Export-ModuleMember -Function @(
    'Write-PerformanceLog',
    'Initialize-LogEnvironment',
    'Get-SystemPerformanceMetrics',
    'Invoke-LogRotation'
) -Variable @(
    'StandardThresholds'
)