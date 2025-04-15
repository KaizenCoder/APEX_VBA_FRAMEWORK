# =============================================================================
# Module de configuration pour le monitoring de performance
# =============================================================================

# Configuration standardisée
$script:PerformanceConfig = @{
    Monitoring = @{
        Enabled       = $true
        RefreshRateMs = 5000
        LogRotation   = @{
            Enabled   = $true
            MaxSizeMB = 10
            KeepDays  = 7
        }
    }
    Thresholds = @{
        CpuUsagePercent = 2      # Objectif < 2%
        MemoryUsageMB   = 100      # Objectif < 100MB
        ResponseTimeMS  = 300      # Objectif < 300ms
    }
    Pooling    = @{
        MaxWorkers  = 4
        QueueSize   = 100
        TaskTimeout = 30000      # 30 secondes
    }
    Logging    = @{
        Level  = "INFO"
        Path   = "logs/performance"
        Format = "[{timestamp}][{level}][{module}] {message}"
    }
}

function Get-PerformanceConfig {
    return $script:PerformanceConfig
}

function Set-PerformanceThreshold {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        
        [Parameter(Mandatory = $true)]
        [double]$Value
    )
    
    if ($script:PerformanceConfig.Thresholds.ContainsKey($Name)) {
        $script:PerformanceConfig.Thresholds[$Name] = $Value
        return $true
    }
    return $false
}

function Initialize-PerformanceConfig {
    $configPath = Join-Path $PSScriptRoot "performance_config.json"
    
    if (Test-Path $configPath) {
        try {
            $savedConfig = Get-Content $configPath -Raw | ConvertFrom-Json
            
            # Mise à jour de la configuration avec les valeurs sauvegardées
            foreach ($key in $savedConfig.PSObject.Properties.Name) {
                if ($script:PerformanceConfig.ContainsKey($key)) {
                    $script:PerformanceConfig[$key] = $savedConfig.$key
                }
            }
        }
        catch {
            Write-Warning "Impossible de charger la configuration. Utilisation des valeurs par défaut."
        }
    }
    
    # Sauvegarde de la configuration
    $script:PerformanceConfig | ConvertTo-Json -Depth 10 | Set-Content $configPath
}

# Export des fonctions
Export-ModuleMember -Function Get-PerformanceConfig, Set-PerformanceThreshold, Initialize-PerformanceConfig -Variable PerformanceConfig