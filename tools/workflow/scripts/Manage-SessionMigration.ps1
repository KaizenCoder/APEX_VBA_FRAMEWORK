# Manage-SessionMigration.ps1
<#
.SYNOPSIS
    Gestion de la migration et validation des sessions de développement APEX.

.DESCRIPTION
    Ce script facilite l'utilisation des outils de migration et validation des sessions
    en fournissant une interface PowerShell conviviale.

.PARAMETER Action
    L'action à effectuer : 
    - Simulate : Simule la migration
    - Migrate : Effectue la migration
    - Validate : Valide la migration
    - All : Effectue la migration puis la validation

.PARAMETER Force
    Force la migration même en cas d'erreurs de validation

.PARAMETER NoBackup
    Désactive la sauvegarde automatique

.PARAMETER Cleanup
    Nettoie l'ancienne structure après migration

.EXAMPLE
    .\Manage-SessionMigration.ps1 -Action Simulate
    .\Manage-SessionMigration.ps1 -Action Migrate -Force -Cleanup
    .\Manage-SessionMigration.ps1 -Action All -NoBackup
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('Simulate', 'Migrate', 'Validate', 'All')]
    [string]$Action,

    [switch]$Force,
    [switch]$NoBackup,
    [switch]$Cleanup
)

# Configuration
$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'

# Chemins
$scriptPath = $PSScriptRoot
$pythonScripts = @{
    Migration = Join-Path $scriptPath "migrate_sessions.py"
    Validation = Join-Path $scriptPath "validate_migration.py"
}

function Test-PythonInstallation {
    try {
        $pythonVersion = python --version
        Write-Verbose "Python détecté : $pythonVersion"
        return $true
    }
    catch {
        Write-Error "Python n'est pas installé ou n'est pas dans le PATH"
        return $false
    }
}

function Test-ScriptRequirements {
    $requirements = @(
        @{
            Path = $pythonScripts.Migration
            Name = "Script de migration"
        },
        @{
            Path = $pythonScripts.Validation
            Name = "Script de validation"
        }
    )

    foreach ($req in $requirements) {
        if (-not (Test-Path $req.Path)) {
            Write-Error "$($req.Name) non trouvé : $($req.Path)"
            return $false
        }
    }
    return $true
}

function Invoke-SessionMigration {
    param (
        [switch]$DryRun
    )

    $arguments = @()
    if ($DryRun) { $arguments += "--dry-run" }
    if ($Force) { $arguments += "--force" }
    if ($NoBackup) { $arguments += "--no-backup" }
    if ($Cleanup) { $arguments += "--cleanup" }

    Write-Verbose "Exécution de la migration avec les arguments : $arguments"
    $process = Start-Process -FilePath "python" -ArgumentList (@($pythonScripts.Migration) + $arguments) -Wait -NoNewWindow -PassThru
    return $process.ExitCode -eq 0
}

function Invoke-SessionValidation {
    Write-Verbose "Exécution de la validation"
    $process = Start-Process -FilePath "python" -ArgumentList $pythonScripts.Validation -Wait -NoNewWindow -PassThru
    return $process.ExitCode -eq 0
}

# Vérification des prérequis
if (-not (Test-PythonInstallation)) { exit 1 }
if (-not (Test-ScriptRequirements)) { exit 1 }

# Exécution selon l'action demandée
$success = $true
switch ($Action) {
    'Simulate' {
        Write-Host "🔍 Simulation de la migration..." -ForegroundColor Cyan
        $success = Invoke-SessionMigration -DryRun
    }
    'Migrate' {
        Write-Host "🚀 Exécution de la migration..." -ForegroundColor Cyan
        $success = Invoke-SessionMigration
    }
    'Validate' {
        Write-Host "✔️ Validation de la migration..." -ForegroundColor Cyan
        $success = Invoke-SessionValidation
    }
    'All' {
        Write-Host "🚀 Migration et validation..." -ForegroundColor Cyan
        $success = Invoke-SessionMigration
        if ($success) {
            Write-Host "✔️ Validation post-migration..." -ForegroundColor Cyan
            $success = Invoke-SessionValidation
        }
    }
}

# Affichage du résultat
if ($success) {
    Write-Host "`n✅ Opération terminée avec succès" -ForegroundColor Green
}
else {
    Write-Host "`n❌ L'opération a échoué" -ForegroundColor Red
    Write-Host "📝 Consultez les logs pour plus de détails" -ForegroundColor Yellow
    exit 1
}

# Affichage des rapports disponibles
$reports = @(
    "migration_report.md",
    "validation_report.md",
    "session_migration.log",
    "session_validation.log"
) | Where-Object { Test-Path (Join-Path $scriptPath "..") }

if ($reports) {
    Write-Host "`n📊 Rapports disponibles :" -ForegroundColor Cyan
    foreach ($report in $reports) {
        Write-Host "   - $report"
    }
} 