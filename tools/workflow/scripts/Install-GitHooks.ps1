# Install-GitHooks.ps1
# Installation des hooks Git pour le framework APEX
# Référence: chat_050 (2024-04-11 16:50)
# Source: chat_049 (Pipeline validation)

[CmdletBinding()]
param()

# Configuration
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Vérification de l'existence du dossier .git
if (-not (Test-Path ".git")) {
    Write-Error "❌ Dossier .git non trouvé. Exécutez ce script depuis la racine du projet."
    exit 1
}

# Création du dossier hooks si nécessaire
$hooksDir = ".git/hooks"
if (-not (Test-Path $hooksDir)) {
    New-Item -ItemType Directory -Path $hooksDir | Out-Null
    Write-Host "📁 Dossier hooks créé: $hooksDir"
}

# Copie du hook pre-commit
$sourceHook = "tools/workflow/scripts/pre-commit"
$targetHook = "$hooksDir/pre-commit"

try {
    Copy-Item -Path $sourceHook -Destination $targetHook -Force
    # Rendre le fichier exécutable sous Unix
    if ($IsLinux -or $IsMacOS) {
        chmod +x $targetHook
    }
    Write-Host "✅ Hook pre-commit installé avec succès"
} catch {
    Write-Error "❌ Erreur lors de l'installation du hook: $_"
    exit 1
}

# Test du pipeline
Write-Host "`n🔍 Test du pipeline de validation..."
try {
    & "tools/workflow/scripts/Start-EncodingPipeline.ps1"
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✅ Pipeline de validation OK"
    } else {
        Write-Warning "⚠️ Le pipeline a détecté des problèmes. Utilisez -Fix pour corriger."
    }
} catch {
    Write-Error "❌ Erreur lors du test du pipeline: $_"
    exit 1
}

Write-Host "`n✨ Installation terminée" 