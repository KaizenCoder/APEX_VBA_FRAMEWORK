# =============================================================================
# APEX Workbench - Agent Cursor
# =============================================================================

# Import des fonctions communes
Import-Module (Join-Path $PSScriptRoot "../common/Write-WorkbenchLog.psm1") -Force

# Variables globales
$script:lastActivity = Get-Date
$script:cursorProcesses = @()
$script:cursorLogs = @()
$script:sessionMetrics = @{
    SessionStartTime    = Get-Date
    ActiveMinutes       = 0
    FileEdits           = 0
    PromptCount         = 0
    AverageResponseTime = 0
    TotalResponses      = 0
    TotalResponseTime   = 0
    Errors              = 0
}
$script:monitoringPath = Join-Path $PSScriptRoot "../../../monitoring"

function Initialize-Monitoring {
    if (-not (Test-Path $script:monitoringPath)) {
        New-Item -ItemType Directory -Path $script:monitoringPath -Force | Out-Null
        Write-WorkbenchLog "Répertoire de monitoring créé: $script:monitoringPath" "INFO"
    }
}

function Get-CursorProcesses {
    Get-Process | Where-Object { $_.ProcessName -like "*cursor*" }
}

function Update-CursorActivity {
    $script:lastActivity = Get-Date
    $processes = Get-CursorProcesses
    
    # Détection des nouveaux processus
    foreach ($process in $processes) {
        if ($script:cursorProcesses.Id -notcontains $process.Id) {
            Write-WorkbenchLog "Nouveau processus Cursor détecté: $($process.Id)" "INFO"
            $script:cursorProcesses += $process
        }
    }
    
    # Détection des processus terminés
    $script:cursorProcesses = $script:cursorProcesses | Where-Object {
        $currentProcess = $_
        if ($processes.Id -notcontains $currentProcess.Id) {
            Write-WorkbenchLog "Processus Cursor terminé: $($currentProcess.Id)" "INFO"
            return $false
        }
        return $true
    }
}

function Get-CursorLogs {
    $logPath = Join-Path $env:APPDATA "Cursor\logs"
    if (Test-Path $logPath) {
        Get-ChildItem $logPath -Filter "*.log" | 
        Sort-Object LastWriteTime -Descending | 
        Select-Object -First 5
    }
}

function Watch-CursorLogs {
    $newLogs = Get-CursorLogs
    
    # Détection des nouveaux logs
    foreach ($log in $newLogs) {
        if ($script:cursorLogs.FullName -notcontains $log.FullName) {
            Write-WorkbenchLog "Nouveau log Cursor détecté: $($log.Name)" "INFO"
            $script:cursorLogs += $log
            
            # Analyse du contenu du log
            $content = Get-Content $log.FullName -Tail 20
            foreach ($line in $content) {
                # Détection des activités liées au code
                if ($line -match "file saved|file opened|edit") {
                    $script:sessionMetrics.FileEdits++
                }
                # Détection des prompts
                if ($line -match "prompt|query|asked") {
                    $script:sessionMetrics.PromptCount++
                }
                # Détection des temps de réponse
                if ($line -match "response time: (\d+)") {
                    $responseTime = [int]$matches[1]
                    $script:sessionMetrics.TotalResponses++
                    $script:sessionMetrics.TotalResponseTime += $responseTime
                    $script:sessionMetrics.AverageResponseTime = $script:sessionMetrics.TotalResponseTime / $script:sessionMetrics.TotalResponses
                }
                # Détection des erreurs
                if ($line -match "error|warning|fail") {
                    Write-WorkbenchLog "Alerte dans le log $($log.Name): $line" "WARNING"
                    $script:sessionMetrics.Errors++
                }
            }
        }
    }
}

function Update-SessionMetrics {
    # Mise à jour des minutes actives
    $script:sessionMetrics.ActiveMinutes = ([math]::Round(((Get-Date) - $script:sessionMetrics.SessionStartTime).TotalMinutes))
    
    # Sauvegarde des métriques
    $metricsPath = Join-Path $script:monitoringPath "cursor_metrics.json"
    $script:sessionMetrics | ConvertTo-Json | Set-Content -Path $metricsPath -Encoding UTF8
    
    Write-WorkbenchLog "Métriques Cursor mises à jour: $($script:sessionMetrics.PromptCount) prompts, $($script:sessionMetrics.FileEdits) modifications" "INFO"
}

function Start-CursorMonitoring {
    Write-WorkbenchLog "Démarrage de la surveillance Cursor" "INFO"
    Initialize-Monitoring
    
    while ($true) {
        try {
            Update-CursorActivity
            Watch-CursorLogs
            
            # Vérification de l'inactivité
            $inactiveTime = (Get-Date) - $script:lastActivity
            if ($inactiveTime.TotalMinutes -gt 30) {
                Write-WorkbenchLog "Cursor inactif depuis $([math]::Round($inactiveTime.TotalMinutes)) minutes" "WARNING"
            }
            
            # Mise à jour des métriques toutes les 5 minutes
            if ($script:sessionMetrics.ActiveMinutes % 5 -eq 0) {
                Update-SessionMetrics
            }
            
            Start-Sleep -Seconds 10
        }
        catch {
            Write-WorkbenchLog "Erreur dans la surveillance Cursor: $_" "ERROR"
            Start-Sleep -Seconds 30  # Délai plus long en cas d'erreur
        }
    }
}

Export-ModuleMember -Function Start-CursorMonitoring, Get-CursorLogs, Update-SessionMetrics 