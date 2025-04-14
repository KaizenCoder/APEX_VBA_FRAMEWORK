. (Join-Path $PSScriptRoot "../common/Write-WorkbenchLog.ps1")

function Get-ProcessMetrics {
    param (
        [string]$ProcessName
    )
    
    $processes = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if ($processes) {
        $metrics = @{
            CPU     = ($processes | Measure-Object CPU -Sum).Sum
            Memory  = ($processes | Measure-Object WorkingSet -Sum).Sum / 1MB
            Threads = ($processes | Measure-Object Threads -Sum).Sum
        }
        return $metrics
    }
    return $null
}

function Optimize-VSCodeProcess {
    $vsCodeMetrics = Get-ProcessMetrics "Code"
    if ($vsCodeMetrics.CPU -gt 150) {
        Write-WorkbenchLog "Optimisation VSCode - CPU élevé: $($vsCodeMetrics.CPU)%" "WARNING"
        
        # Collecte des extensions actives
        $extensionsPath = Join-Path $env:APPDATA "Code/User/globalStorage"
        $activeExtensions = Get-ChildItem $extensionsPath -Directory | 
        Where-Object { Test-Path (Join-Path $_.FullName "state.vscdb") }
        
        # Désactivation temporaire des extensions non essentielles
        foreach ($ext in $activeExtensions) {
            if ($ext.Name -notmatch "(ms-vscode|cursor)") {
                $statePath = Join-Path $ext.FullName "state.vscdb"
                if (Test-Path $statePath) {
                    Rename-Item $statePath "$statePath.bak" -Force
                    Write-WorkbenchLog "Extension désactivée: $($ext.Name)" "INFO"
                }
            }
        }
        
        # Nettoyage du cache
        $cachePath = Join-Path $env:APPDATA "Code/Cache"
        if (Test-Path $cachePath) {
            Get-ChildItem $cachePath -File | Remove-Item -Force
            Write-WorkbenchLog "Cache VSCode nettoyé" "INFO"
        }
    }
}

function Monitor-ProcessHealth {
    while ($true) {
        $vsCodeMetrics = Get-ProcessMetrics "Code"
        $cursorMetrics = Get-ProcessMetrics "Cursor"
        
        if ($vsCodeMetrics) {
            if ($vsCodeMetrics.CPU -gt 150 -or $vsCodeMetrics.Memory -gt 1000) {
                Optimize-VSCodeProcess
            }
        }
        
        if ($cursorMetrics) {
            if ($cursorMetrics.Memory -gt 200) {
                Write-WorkbenchLog "Alerte mémoire Cursor: $($cursorMetrics.Memory)MB" "WARNING"
            }
        }
        
        Start-Sleep -Seconds 30
    }
}

Export-ModuleMember -Function Monitor-ProcessHealth 