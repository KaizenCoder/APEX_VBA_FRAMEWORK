# =============================================================================
# Script de planification de l'agent documentaire pour APEX Framework
# =============================================================================
#
# .SYNOPSIS
#   Configure une tâche planifiée pour exécuter l'agent documentaire
#
# .DESCRIPTION
#   Ce script crée une tâche planifiée Windows qui exécute l'agent documentaire
#   à intervalles réguliers pour vérifier la conformité de la documentation
#
# .PARAMETER Interval
#   Intervalle d'exécution : 'Daily', 'Weekly', 'Monthly'
#
# .PARAMETER Time
#   Heure d'exécution (format HH:mm)
#
# .PARAMETER GenerateReport
#   Indique si l'agent doit générer un rapport
#
# .PARAMETER AutoFix
#   Indique si l'agent doit corriger automatiquement les problèmes détectés
#
# .EXAMPLE
#   .\schedule_doc_agent.ps1 -Interval Daily -Time 09:00 -GenerateReport
#   Planifie l'exécution quotidienne de l'agent documentaire à 9h00 avec génération de rapport
#
# =============================================================================

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [ValidateSet('Daily', 'Weekly', 'Monthly')]
    [string]$Interval = 'Daily',
    
    [Parameter(Mandatory = $false)]
    [string]$Time = "01:00",
    
    [Parameter(Mandatory = $false)]
    [switch]$GenerateReport = $true,
    
    [Parameter(Mandatory = $false)]
    [switch]$AutoFix = $false,
    
    [Parameter(Mandatory = $false)]
    [string]$TaskName = "APEX_DocumentationAgent"
)

# Vérification des droits administrateur
function Test-Administrator {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Administrator)) {
    Write-Warning "Ce script nécessite des droits d'administrateur pour créer une tâche planifiée."
    Write-Warning "Veuillez relancer le script en tant qu'administrateur."
    exit 1
}

# Chemin du répertoire de travail
$workingDir = Join-Path $PSScriptRoot "..\..\.."
$workingDir = Resolve-Path $workingDir

# Chemin du script Python
$scriptPath = Join-Path $PSScriptRoot "doc_agent.py"
$configPath = Join-Path $PSScriptRoot "doc_guidelines.json"

# Construction de la commande Python
$pythonArgs = @(
    $scriptPath,
    "--target", "`"$workingDir`"",
    "--config", "`"$configPath`""
)

if ($GenerateReport) {
    $reportPath = Join-Path $workingDir "reports\doc_compliance_%date:~-4,4%%date:~-7,2%%date:~-10,2%.md"
    $pythonArgs += "--report"
    $pythonArgs += "`"$reportPath`""
}

if ($AutoFix) {
    $pythonArgs += "--fix"
}

$pythonCommand = "python " + ($pythonArgs -join " ")

# Configuration de la tâche planifiée
$taskAction = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c $pythonCommand" -WorkingDirectory $workingDir

# Définition du déclencheur selon l'intervalle spécifié
switch ($Interval) {
    'Daily' {
        $taskTrigger = New-ScheduledTaskTrigger -Daily -At $Time
    }
    'Weekly' {
        $taskTrigger = New-ScheduledTaskTrigger -Weekly -At $Time -DaysOfWeek Monday
    }
    'Monthly' {
        $taskTrigger = New-ScheduledTaskTrigger -Monthly -At $Time -DaysOfMonth 1
    }
}

# Paramètres de la tâche
$taskSettings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -WakeToRun -StartWhenAvailable

# Créer ou mettre à jour la tâche planifiée
$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

if ($existingTask) {
    # Mettre à jour la tâche existante
    Set-ScheduledTask -TaskName $TaskName -Action $taskAction -Trigger $taskTrigger -Settings $taskSettings
    Write-Host "✅ Tâche planifiée '$TaskName' mise à jour avec succès." -ForegroundColor Green
}
else {
    # Créer une nouvelle tâche
    Register-ScheduledTask -TaskName $TaskName -Action $taskAction -Trigger $taskTrigger -Settings $taskSettings -Description "Agent de vérification de documentation APEX Framework"
    Write-Host "✅ Tâche planifiée '$TaskName' créée avec succès." -ForegroundColor Green
}

# Afficher les détails de la tâche
Write-Host "`nDétails de la tâche planifiée:" -ForegroundColor Cyan
Write-Host "  Nom: $TaskName" -ForegroundColor Gray
Write-Host "  Intervalle: $Interval" -ForegroundColor Gray
Write-Host "  Heure d'exécution: $Time" -ForegroundColor Gray
Write-Host "  Génération de rapport: $GenerateReport" -ForegroundColor Gray
Write-Host "  Correction automatique: $AutoFix" -ForegroundColor Gray
Write-Host "  Commande: $pythonCommand" -ForegroundColor Gray

Write-Host "`nRemarque: La tâche s'exécutera avec les droits de l'utilisateur actuellement connecté." -ForegroundColor Yellow