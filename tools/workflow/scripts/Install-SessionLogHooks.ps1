# Install-SessionLogHooks.ps1
<#
.SYNOPSIS
    Installe les hooks Git pour la validation des logs de session
.DESCRIPTION
    Configure les hooks Git pre-commit pour valider automatiquement
    le format et l'encodage des fichiers de log de session
#>

[CmdletBinding()]
param(
    [Parameter()]
    [switch]$Force
)

# Chemins
$scriptPath = $PSScriptRoot
$gitRoot = git rev-parse --show-toplevel 2>$null
$hookPath = Join-Path $gitRoot ".git\hooks"
$preCommitPath = Join-Path $hookPath "pre-commit"

# Vérification de Git
if (-not $gitRoot) {
    Write-Error "Ce script doit être exécuté dans un dépôt Git."
    exit 1
}

# Création du hook pre-commit
$hookContent = @"
#!/bin/sh
# Pre-commit hook pour la validation des logs de session

# Récupération des fichiers modifiés
changed_files=\$(git diff --cached --name-only --diff-filter=ACM "*.md" | grep -i "logs/sessions/")

if [ -n "\$changed_files" ]; then
    echo "🔍 Validation des logs de session modifiés..."
    
    # Exécution du script de validation
    pwsh -NoProfile -ExecutionPolicy Bypass -File "$scriptPath\Test-SessionLogFormat.ps1" -Path "\$changed_files"
    
    if [ \$? -ne 0 ]; then
        echo "❌ La validation des logs a échoué. Veuillez corriger les erreurs avant de commiter."
        exit 1
    fi
fi

exit 0
"@

# Installation du hook
try {
    # Création du dossier hooks si nécessaire
    if (-not (Test-Path $hookPath)) {
        New-Item -ItemType Directory -Path $hookPath | Out-Null
    }
    
    # Vérification si le hook existe déjà
    if ((Test-Path $preCommitPath) -and -not $Force) {
        Write-Error "Le hook pre-commit existe déjà. Utilisez -Force pour le remplacer."
        exit 1
    }
    
    # Écriture du hook
    $hookContent | Out-File -FilePath $preCommitPath -Encoding utf8 -Force
    
    # Rendre le script exécutable sous Unix
    if ($IsLinux -or $IsMacOS) {
        chmod +x $preCommitPath
    }
    
    Write-Host "✅ Hook pre-commit installé avec succès." -ForegroundColor Green
    Write-Host "Les logs de session seront automatiquement validés lors des commits."
}
catch {
    Write-Error "Erreur lors de l'installation du hook : $_"
    exit 1
} 