# Register-ApexScripts.ps1
# Script pour enregistrer les commandes APEX dans PowerShell
# Permet d'executer les commandes depuis n'importe quel repertoire

# Force l'encodage UTF-8 sans BOM
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Verifier si le script est execute en tant qu'administrateur
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "Ce script doit aÃƒÆ’"Å¡Ãƒâ€šÃ‚Âªtre execute en tant qu'administrateur pour modifier le profil PowerShell." -ForegroundColor Red
    Write-Host "Veuillez redemarrer ce script avec les droits d'administrateur." -ForegroundColor Red
    exit 1
}

# Chemins
$projectRoot = "D:\Dev\Apex_VBA_FRAMEWORK"
$modulesDir = Join-Path -Path $projectRoot -ChildPath "modules"
$userModulesPath = Join-Path -Path ([Environment]::GetFolderPath("MyDocuments")) -ChildPath "WindowsPowerShell\Modules"

# Creer le dossier des modules utilisateur si necessaire
if (-not (Test-Path $userModulesPath)) {
    New-Item -Path $userModulesPath -ItemType Directory -Force | Out-Null
}

# Copier le module Apex.SessionManager
$moduleSource = Join-Path -Path $modulesDir -ChildPath "Apex.SessionManager"
$moduleDest = Join-Path -Path $userModulesPath -ChildPath "Apex.SessionManager"

if (Test-Path $moduleDest) {
    Remove-Item -Path $moduleDest -Recurse -Force
}
Copy-Item -Path $moduleSource -Destination $userModulesPath -Recurse -Force

# Importer le module
Import-Module -Name Apex.SessionManager -Force -ErrorAction SilentlyContinue

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "INSTALLATION Raaa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â°USSIE" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Pour commencer aÃƒÆ’"Å¡Ãƒâ€šÃ‚Â  utiliser les commandes APEX:" -ForegroundColor Yellow
Write-Host "1. Demarrer une session: New-ApexSession" -ForegroundColor White
Write-Host "2. Ajouter une teche: Add-TaskToSession" -ForegroundColor White
Write-Host "3. Terminer une session: Complete-ApexSession" -ForegroundColor White
Write-Host ""
Write-Host "Ces commandes sont maintenant disponibles dans toutes vos sessions PowerShell!" -ForegroundColor Green

# Enregistrement du script d'initialisation Cursor
Write-Host "Enregistrement du script Initialize-CursorSession..."
$cursorInitScript = @{
    Name = "Initialize-CursorSession"
    Path = "tools/workflow/scripts/Initialize-CursorSession.ps1"
    Description = "Automatisation de l'initialisation des règles Cursor"
    RequiredBefore = @("development", "review", "commit")
    Version = "1.0"
}
Register-ApexScript @cursorInitScript 