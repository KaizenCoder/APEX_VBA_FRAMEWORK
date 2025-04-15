. (Join-Path $PSScriptRoot "../common/Write-WorkbenchLog.ps1")

$script:processCache = @{}
$script:lastCleanup = Get-Date

function Get-ProcessMetrics {
    param (
        [string]$ProcessName
    )
    
    try {
        $processes = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
        if (-not $processes) {
            return $null
        }

        # Get metrics efficiently in one pass
        $metrics = $processes | Measure-Object CPU, WorkingSet, Threads -Sum
        return @{
            CPU     = [math]::Round($metrics.Sum[0], 2)
            Memory  = [math]::Round($metrics.Sum[1] / 1MB, 2)
            Threads = $metrics.Sum[2]
        }
    }
    catch {
        Write-WorkbenchLog "Erreur métrique $ProcessName : $_" "ERROR"
        return $null
    }
}

function Optimize-VSCodeProcess {
    $vsCodeMetrics = Get-ProcessMetrics "Code"
    if ($vsCodeMetrics.CPU -gt 150) {
        Write-WorkbenchLog "Optimisation VSCode - CPU élevé: $($vsCodeMetrics.CPU)%" "WARNING"
        
        # Clear VS Code caches if needed
        $appData = $env:APPDATA
        $cachePaths = @(
            "$appData\Code\Cache",
            "$appData\Code\CachedData",
            "$appData\Code\CachedExtensions"
        )
        
        foreach ($path in $cachePaths) {
            if (Test-Path $path) {
                Get-ChildItem -Path $path -File | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-7) } | Remove-Item -Force
            }
        }
        
        Write-WorkbenchLog "Cache VSCode nettoyé" "INFO"
    }
}

function Monitor-ProcessHealth {
    param(
        [int]$IntervalSeconds = 30
    )
    
    $configPath = Join-Path $PSScriptRoot "../../../config/monitoring/performance_config.json"
    $config = Get-Content $configPath -Raw | ConvertFrom-Json
    
    while ($true) {
        try {
            $vsCodeMetrics = Get-ProcessMetrics "Code"
            $cursorMetrics = Get-ProcessMetrics "Cursor"
            
            # Validate against thresholds from config
            if ($vsCodeMetrics) {
                if ($vsCodeMetrics.CPU -gt $config.Thresholds.CpuUsagePercent -or 
                    $vsCodeMetrics.Memory -gt $config.Thresholds.MemoryUsageMB) {
                    Optimize-VSCodeProcess
                }
            }
            
            if ($cursorMetrics) {
                if ($cursorMetrics.Memory -gt $config.Thresholds.MemoryUsageMB) {
                    Write-WorkbenchLog "Alerte mémoire Cursor: $($cursorMetrics.Memory)MB" "WARNING"
                    # Clean up orphaned Cursor processes if any
                    Get-Process | Where-Object { $_.ProcessName -like "*cursor*" -and $_.StartTime -lt (Get-Date).AddHours(-4) } | 
                        ForEach-Object { 
                            try { 
                                $_.Kill()
                                Write-WorkbenchLog "Processus Cursor orphelin terminé: $($_.Id)" "INFO"
                            } catch { 
                                Write-WorkbenchLog "Impossible de terminer le processus: $_" "ERROR" 
                            }
                        }
                }
            }
            
            # Periodic cleanup
            if (((Get-Date) - $script:lastCleanup).TotalHours -ge 1) {
                [System.GC]::Collect()
                $script:lastCleanup = Get-Date
            }
            
            Start-Sleep -Seconds $IntervalSeconds
        }
        catch {
            Write-WorkbenchLog "Erreur monitoring: $_" "ERROR"
            Start-Sleep -Seconds ($IntervalSeconds * 2)
        }
    }
}

Export-ModuleMember -Function Monitor-ProcessHealth, Get-ProcessMetrics