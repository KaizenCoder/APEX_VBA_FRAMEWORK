# Start-ApexSession.ps1
# Script pour demarrer une nouvelle session de developpement APEX VBA Framework
# Interface simplifiee pour New-SessionLog.ps1

param (
    [Parameter(Mandatory=$false)]
    [string]$Title = "",
    
    [Parameter(Mandatory=$false)]
    [string[]]$Objectives = @()
)

# Determination du repertoire de base du projet
$projectRoot = "D:\Dev\Apex_VBA_FRAMEWORK"
$scriptDirectory = Join-Path -Path $projectRoot -ChildPath "tools\workflow\scripts"

# Chemin absolu vers le script principal
$sessionLogScript = Join-Path -Path $scriptDirectory -ChildPath "New-SessionLog.ps1"

if (-not (Test-Path $sessionLogScript)) {
    Write-Host "Erreur: Script New-SessionLog.ps1 introuvable a : $sessionLogScript" -ForegroundColor Red
    exit 1
}

# Ajouter le repertoire des scripts au chemin de recherche
$env:PATH += ";$scriptDirectory"

# Afficher un en-tete
Clear-Host
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host "           DEMARRAGE D'UNE SESSION APEX VBA            " -ForegroundColor Cyan
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host ""

# Mode interactif si aucun titre n'est fourni
if ([string]::IsNullOrWhiteSpace($Title)) {
    Write-Host "Donnez un titre a votre session:" -ForegroundColor Yellow
    $Title = Read-Host
    
    if ([string]::IsNullOrWhiteSpace($Title)) {
        $Title = "Session de travail - $(Get-Date -Format 'dd MMMM yyyy')"
        Write-Host "Titre par defaut utilise: $Title" -ForegroundColor Gray
    }
}

# Demander les objectifs si non fournis
if ($Objectives.Count -eq 0) {
    Write-Host "`nEntrez les objectifs de cette session (terminez par une ligne vide):" -ForegroundColor Yellow
    
    $objectiveInput = "dummy"
    $objectivesList = @()
    
    while (-not [string]::IsNullOrWhiteSpace($objectiveInput)) {
        $objectiveInput = Read-Host
        if (-not [string]::IsNullOrWhiteSpace($objectiveInput)) {
            $objectivesList += $objectiveInput
        }
    }
    
    $Objectives = $objectivesList
    
    if ($Objectives.Count -eq 0) {
        Write-Host "Aucun objectif specifie." -ForegroundColor Gray
    } else {
        Write-Host "`nObjectifs enregistres:" -ForegroundColor Green
        foreach ($obj in $Objectives) {
            Write-Host "- $obj" -ForegroundColor Green
        }
    }
}

# Lancer le script de creation de session
Write-Host "`nCreation de la session..." -ForegroundColor Cyan

Import-Module $sessionLogScript -Force
$result = New-ApexSession -Title $Title -Objectives $Objectives

if ($result) {
    Write-Host "`nSession demarree avec succes!" -ForegroundColor Green
    Write-Host "`nCommandes disponibles:" -ForegroundColor Yellow
    Write-Host "- Pour ajouter une tache: Add-TaskToSession -Name 'Nom' -Module 'Module'" -ForegroundColor Gray
    Write-Host "- Pour terminer la session: Complete-ApexSession" -ForegroundColor Gray
} else {
    Write-Host "Erreur lors du demarrage de la session" -ForegroundColor Red
    exit 1
} 