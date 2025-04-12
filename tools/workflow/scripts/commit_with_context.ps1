# Script de commit avec contexte pour APEX VBA Framework
# Ce script permet d'enrichir les commits Git avec des informations contextuelles

# Trouver le module ApexWSLBridge de façon flexible
$modulePath = $null
$possiblePaths = @(
    # Chemin relatif standard
    (Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath "..\..\powershell\ApexWSLBridge.psm1"),
    # Chemin absolu direct
    "D:\Dev\Apex_VBA_FRAMEWORK\tools\powershell\ApexWSLBridge.psm1",
    # Recherche dans tools
    (Get-ChildItem -Path "D:\Dev\Apex_VBA_FRAMEWORK\tools" -Recurse -Filter "ApexWSLBridge.psm1" -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName)
)

# Essayer chaque chemin possible
foreach ($path in $possiblePaths) {
    if ($path -and (Test-Path $path)) {
        $modulePath = $path
        break
    }
}

# Importer le module si trouvé
if ($modulePath) {
    Import-Module $modulePath -Force
    $wslBridgeAvailable = $true
    Write-Host "Module ApexWSLBridge chargé depuis: $modulePath" -ForegroundColor Green
} else {
    $wslBridgeAvailable = $false
    Write-Host "Module ApexWSLBridge non trouvé. Les commandes WSL pourraient être moins fiables." -ForegroundColor Yellow
}

function Get-CommitMessage {
    param (
        [string]$DefaultType = "feat",
        [string]$DefaultScope = "core"
    )
    
    Write-Host "`n=== Création d'un message de commit standardisé ===`n" -ForegroundColor Cyan
    
    # Affichage des types disponibles
    Write-Host "Types de commit disponibles:" -ForegroundColor Yellow
    Write-Host "  feat     - Nouvelle fonctionnalité"
    Write-Host "  fix      - Correction de bug"
    Write-Host "  docs     - Documentation"
    Write-Host "  style    - Formatage, semicolons, etc."
    Write-Host "  refactor - Refactorisation du code"
    Write-Host "  perf     - Améliorations de performance"
    Write-Host "  test     - Tests"
    Write-Host "  build    - Système de build, dépendances"
    Write-Host "  ci       - Intégration continue"
    Write-Host "  chore    - Tâches diverses"
    
    # Saisie du type de commit
    $commitType = Read-Host "`nType du commit [$DefaultType]"
    if ([string]::IsNullOrWhiteSpace($commitType)) {
        $commitType = $DefaultType
    }
    
    # Saisie du scope
    $commitScope = Read-Host "Scope du commit (module concerné) [$DefaultScope]"
    if ([string]::IsNullOrWhiteSpace($commitScope)) {
        $commitScope = $DefaultScope
    }
    
    # Saisie du titre
    $commitTitle = ""
    while ([string]::IsNullOrWhiteSpace($commitTitle)) {
        $commitTitle = Read-Host "Titre du commit (obligatoire)"
        if ([string]::IsNullOrWhiteSpace($commitTitle)) {
            Write-Host "Le titre est obligatoire." -ForegroundColor Red
        }
    }
    
    # Saisie du corps
    Write-Host "`nCorps du message (détails, contexte, etc.) - Terminez par une ligne vide:" -ForegroundColor Yellow
    $commitBody = @()
    $line = " "
    while (-not [string]::IsNullOrEmpty($line)) {
        $line = Read-Host
        if (-not [string]::IsNullOrEmpty($line)) {
            $commitBody += $line
        }
    }
    
    # Saisie des breaking changes
    Write-Host "`nBreaking changes (laissez vide s'il n'y en a pas):" -ForegroundColor Yellow
    $breakingChanges = Read-Host
    
    # Saisie des références
    Write-Host "`nRéférences (issues, PRs, etc. - ex: #123, #456):" -ForegroundColor Yellow
    $references = Read-Host
    
    # Construction du message
    $message = "$commitType($commitScope): $commitTitle"
    
    if ($commitBody.Count -gt 0) {
        $message += "`n`n" + ($commitBody -join "`n")
    }
    
    if (-not [string]::IsNullOrWhiteSpace($breakingChanges)) {
        $message += "`n`nBREAKING CHANGE: $breakingChanges"
    }
    
    if (-not [string]::IsNullOrWhiteSpace($references)) {
        $message += "`n`n$references"
    }
    
    return $message
}

function Get-GitStatus {
    if ($wslBridgeAvailable) {
        return Invoke-WSLCommand -Command "cd /mnt/d/Dev/Apex_VBA_FRAMEWORK && git status --porcelain" -UseTempFile
    } else {
        return & git status --porcelain
    }
}

function Get-GitDiff {
    if ($wslBridgeAvailable) {
        return Invoke-WSLCommand -Command "cd /mnt/d/Dev/Apex_VBA_FRAMEWORK && git diff --staged" -UseTempFile
    } else {
        return & git diff --staged
    }
}

function Invoke-GitCommit {
    param (
        [string]$Message
    )
    
    # Écrire le message dans un fichier temporaire
    $tempFile = [System.IO.Path]::GetTempFileName()
    $Message | Out-File -FilePath $tempFile -Encoding utf8
    
    if ($wslBridgeAvailable) {
        $wslPath = Invoke-WSLCommand -Command "wslpath -u '$tempFile'"
        $result = Invoke-WSLCommand -Command "cd /mnt/d/Dev/Apex_VBA_FRAMEWORK && git commit -F $wslPath" -UseTempFile
    } else {
        $result = & git commit -F $tempFile
    }
    
    # Nettoyage
    Remove-Item -Path $tempFile -Force
    
    return $result
}

function Update-SessionLog {
    param (
        [string]$CommitMessage,
        [string]$CommitDiff
    )
    
    $sessionLogDir = Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath "..\logs\sessions"
    $templatePath = Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath "..\templates\session_log_template.md"
    
    # Créer le répertoire s'il n'existe pas
    if (-not (Test-Path $sessionLogDir)) {
        New-Item -Path $sessionLogDir -ItemType Directory -Force | Out-Null
    }
    
    # Nom du fichier de log
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $sessionLogPath = Join-Path -Path $sessionLogDir -ChildPath "session_$timestamp.md"
    
    # Contenu du log
    $logContent = @"
# Session de travail APEX VBA Framework - $((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))

## Informations de session
- **Date**: $((Get-Date).ToString("yyyy-MM-dd"))
- **Heure**: $((Get-Date).ToString("HH:mm:ss"))
- **Utilisateur**: $env:USERNAME
- **Poste**: $env:COMPUTERNAME

## Message de commit
```
$CommitMessage
```

## Modifications
```diff
$CommitDiff
```

## Notes additionnelles
<!-- Ajoutez ici des notes personnelles, des idées, ou des commentaires sur la session -->

## Tâches à réaliser prochainement
- [ ] Tâche 1
- [ ] Tâche 2
"@
    
    # Écriture du fichier de log
    $logContent | Out-File -FilePath $sessionLogPath -Encoding utf8
    
    return $sessionLogPath
}

# Vaérification des changements
$gitStatus = Get-GitStatus
if ([string]::IsNullOrWhiteSpace($gitStatus)) {
    Write-Host "Aucun changement à committer." -ForegroundColor Yellow
    exit 0
}

# Affichage des changements
Write-Host "`n=== Changements à committer ===`n" -ForegroundColor Cyan
Write-Host $gitStatus

# Saisie du message de commit
$commitMessage = Get-CommitMessage

# Raécupaération du diff
$commitDiff = Get-GitDiff

# Création du log de session
$sessionLogPath = Update-SessionLog -CommitMessage $commitMessage -CommitDiff $commitDiff

# Commit
Write-Host "`nCommit en cours..." -ForegroundColor Yellow
$commitResult = Invoke-GitCommit -Message $commitMessage
Write-Host $commitResult -ForegroundColor Green

# Affichage du chemin vers le log de session
Write-Host "`nSession log créé à: $sessionLogPath" -ForegroundColor Cyan
Write-Host "Commit terminé avec succès!" -ForegroundColor Green