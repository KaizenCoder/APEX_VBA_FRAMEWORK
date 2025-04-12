# Apex.SessionManager.psm1
# Module de gestion des sessions de developpement APEX VBA Framework

#Requires -Version 5.1

# Force l'encodage UTF-8
$script:OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Variables globales du module
$script:projectRoot = "D:\Dev\Apex_VBA_FRAMEWORK"
$script:logsRoot = Join-Path -Path $projectRoot -ChildPath "tools\workflow\logs"
$script:sessionsDir = Join-Path -Path $script:logsRoot -ChildPath "sessions"
$script:templatePath = Join-Path -Path $projectRoot -ChildPath "tools\workflow\templates\session_log_template.md"
$script:currentSession = $null

# Creation des repertoires necessaires
if (-not (Test-Path $script:logsRoot)) { New-Item -Path $script:logsRoot -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $script:sessionsDir)) { New-Item -Path $script:sessionsDir -ItemType Directory -Force | Out-Null }

<#
.SYNOPSIS
    Cree une nouvelle session de developpement APEX.
.DESCRIPTION
    Cette fonction cree une nouvelle session de developpement avec un ID unique,
    un titre et des objectifs. Elle genere un fichier Markdown structure pour
    suivre l'avancement de la session.
.PARAMETER Title
    Le titre de la session. Si non specifie, utilise la date du jour.
.PARAMETER Objectives
    Liste des objectifs de la session.
.EXAMPLE
    New-ApexSession -Title "Refactoring module de logging" -Objectives @("Implementer nouveau format", "Tests unitaires")
#>
function New-ApexSession {
    [CmdletBinding()]
    param (
        [string]$Title = "",
        [string[]]$Objectives = @()
    )
    
    # Generer ID unique
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $sessionId = "session_$timestamp"
    $sessionPath = Join-Path -Path $script:sessionsDir -ChildPath "$sessionId.md"
    
    # Titre par defaut si non specifie
    if ([string]::IsNullOrWhiteSpace($Title)) {
        $Title = "Session de travail - $(Get-Date -Format 'dd MMMM yyyy')"
    }
    
    # Creer contenu du log
    $date = Get-Date -Format "dd MMMM yyyy"
    $content = @"
# Session de travail - $date

## Objectifs de la session
"@

    if ($Objectives.Count -gt 0) {
        foreach ($obj in $Objectives) {
            $content += "`n- [ ] $obj"
        }
    } else {
        $content += "`n- [ ] Objectif 1`n- [ ] Objectif 2"
    }

    $content += @"

## Suivi des taches

| Tache | Module | Statut | Commentaire |
|-------|---------|--------|-------------|
| | | | |

## Prompts IA utilises

| Heure | Agent | Prompt |
|-------|-------|--------|
| | | |

## Tests effectues

## Modules modifies

## Commits prevus

## Bilan de session

"@

    # Ecrire fichier
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($sessionPath, $content, $utf8NoBom)
    
    # Stocker session courante
    $script:currentSession = @{
        Id = $sessionId
        Path = $sessionPath
        StartTime = Get-Date
        Title = $Title
        Objectives = $Objectives
    }
    
    Write-Host "Session '$Title' creee: $sessionPath" -ForegroundColor Green
    return $sessionId
}

<#
.SYNOPSIS
    Ajoute une tache a la session en cours.
.DESCRIPTION
    Ajoute une nouvelle tache a la session specifiee ou a la session courante.
.PARAMETER Name
    Nom de la tache.
.PARAMETER Module
    Module concerne par la tache.
.PARAMETER Status
    Etat de la tache (En cours, Termine, etc.).
.PARAMETER Comment
    Commentaire optionnel sur la tache.
#>
function Add-TaskToSession {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Name,
        [string]$Module = "",
        [ValidateSet("En cours", "Termine", "Abandonne", "Bloque")]
        [string]$Status = "En cours",
        [string]$Comment = ""
    )
    
    if ($null -eq $script:currentSession) {
        Write-Error "Aucune session active. Demarrez une session avec New-ApexSession."
        return
    }
    
    $sessionPath = $script:currentSession.Path
    $content = Get-Content -Path $sessionPath -Raw
    
    # Ajouter la tache
    $taskLine = "| $Name | $Module | $Status | $Comment |"
    $content = $content -replace "\| \| \| \| \|", "$taskLine`n| | | | |"
    
    # Sauvegarder
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($sessionPath, $content, $utf8NoBom)
    
    Write-Host "Tache ajoutee: $Name" -ForegroundColor Green
}

<#
.SYNOPSIS
    Termine la session de developpement en cours.
.DESCRIPTION
    Finalise la session en ajoutant un resume et la duree totale.
.PARAMETER Summary
    Resume de la session.
#>
function Complete-ApexSession {
    [CmdletBinding()]
    param (
        [string]$Summary = ""
    )
    
    if ($null -eq $script:currentSession) {
        Write-Error "Aucune session active a terminer."
        return
    }
    
    $sessionPath = $script:currentSession.Path
    $content = Get-Content -Path $sessionPath -Raw
    
    # Calculer duree
    $duration = "{0:hh\:mm\:ss}" -f ((Get-Date) - $script:currentSession.StartTime)
    
    # Ajouter duree et resume
    $content = $content -replace "(# Session de travail.*?\n)", "`$1`n**Duree de la session**: $duration`n"
    if (-not [string]::IsNullOrWhiteSpace($Summary)) {
        $content = $content -replace "(## Bilan de session\n\n).*?$", "`$1$Summary`n"
    }
    
    # Sauvegarder
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($sessionPath, $content, $utf8NoBom)
    
    Write-Host "Session terminee: $($script:currentSession.Title)" -ForegroundColor Green
    Write-Host "Duree: $duration" -ForegroundColor Cyan
    
    $script:currentSession = $null
}

<#
.SYNOPSIS
    Retourne les informations sur la session en cours.
#>
function Get-CurrentSession {
    [CmdletBinding()]
    param()
    return $script:currentSession
}

# Exporter les fonctions publiques
Export-ModuleMember -Function @(
    'New-ApexSession',
    'Add-TaskToSession',
    'Complete-ApexSession',
    'Get-CurrentSession'
) 