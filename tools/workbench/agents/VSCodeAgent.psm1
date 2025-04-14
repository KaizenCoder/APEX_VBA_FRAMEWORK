# =============================================================================
# APEX Workbench - Agent VSCode
# =============================================================================

# Import des fonctions communes
. (Join-Path $PSScriptRoot "../common/Write-WorkbenchLog.ps1")

# Variables globales
$script:lastActivity = Get-Date
$script:vscodeProcesses = @()
$script:extensionUpdates = @()

function Get-VSCodeProcesses {
    Get-Process | Where-Object { $_.ProcessName -like "*code*" }
}

function Update-VSCodeActivity {
    $script:lastActivity = Get-Date
    $processes = Get-VSCodeProcesses
    
    # Détection des nouveaux processus
    foreach ($process in $processes) {
        if ($script:vscodeProcesses.Id -notcontains $process.Id) {
            Write-WorkbenchLog "Nouveau processus VSCode détecté: $($process.Id)" "INFO"
            $script:vscodeProcesses += $process
        }
    }
    
    # Détection des processus terminés
    $script:vscodeProcesses = $script:vscodeProcesses | Where-Object {
        $currentProcess = $_
        if ($processes.Id -notcontains $currentProcess.Id) {
            Write-WorkbenchLog "Processus VSCode terminé: $($currentProcess.Id)" "INFO"
            return $false
        }
        return $true
    }
}

function Check-ExtensionUpdates {
    $extensionsPath = Join-Path $env:USERPROFILE ".vscode\extensions"
    if (Test-Path $extensionsPath) {
        $currentExtensions = Get-ChildItem $extensionsPath -Directory
        
        foreach ($ext in $currentExtensions) {
            $packagePath = Join-Path $ext.FullName "package.json"
            if (Test-Path $packagePath) {
                try {
                    $package = Get-Content $packagePath -Raw | ConvertFrom-Json
                    $extId = "$($package.publisher).$($package.name)"
                    
                    if ($script:extensionUpdates -notcontains $extId) {
                        Write-WorkbenchLog "Extension VSCode détectée: $extId v$($package.version)" "INFO"
                        $script:extensionUpdates += $extId
                    }
                }
                catch {
                    Write-WorkbenchLog "Erreur lors de la lecture de l'extension $($ext.Name): $_" "ERROR"
                }
            }
        }
    }
}

function Watch-WorkspaceChanges {
    $workspacePath = Join-Path $env:USERPROFILE "workspace"
    if (Test-Path $workspacePath) {
        $changes = Get-ChildItem $workspacePath -Recurse |
        Where-Object { $_.LastWriteTime -gt $script:lastActivity }
            
        foreach ($change in $changes) {
            Write-WorkbenchLog "Modification détectée: $($change.FullName)" "INFO"
        }
    }
}

function Start-VSCodeMonitoring {
    Write-WorkbenchLog "Démarrage de la surveillance VSCode" "INFO"
    
    while ($true) {
        try {
            Update-VSCodeActivity
            Check-ExtensionUpdates
            Watch-WorkspaceChanges
            
            # Vérification de l'inactivité
            $inactiveTime = (Get-Date) - $script:lastActivity
            if ($inactiveTime.TotalMinutes -gt 30) {
                Write-WorkbenchLog "VSCode inactif depuis $($inactiveTime.TotalMinutes) minutes" "WARNING"
            }
            
            Start-Sleep -Seconds 10
        }
        catch {
            Write-WorkbenchLog "Erreur dans la surveillance VSCode: $_" "ERROR"
            Start-Sleep -Seconds 30  # Délai plus long en cas d'erreur
        }
    }
}

Export-ModuleMember -Function Start-VSCodeMonitoring 