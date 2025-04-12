# Installation des hooks Git pour le projet APEX Framework
# Référence: chat_038 (2024-04-11 16:30) - Correction encodage des scripts
# Source: chat_037 (2024-04-11 16:15) - Installation des hooks Git

# Vérifier que le dossier des hooks existe
$hooksTarget = ".git/hooks"
$hooksSource = "tools/workflow/scripts/hooks"

if (-not (Test-Path $hooksTarget)) {
    New-Item -ItemType Directory -Path $hooksTarget | Out-Null
}

# Copier les hooks
$hooks = @("pre-commit")
foreach ($hook in $hooks) {
    $sourceHook = Join-Path $hooksSource $hook
    $targetHook = Join-Path $hooksTarget $hook

    if (Test-Path $sourceHook) {
        Copy-Item $sourceHook $targetHook -Force

        # Sous Windows, on ne peut pas rendre les fichiers exécutables
        if ($IsLinux -or $IsMacOS) {
            chmod +x $targetHook
        }
        Write-Host "✅ Hook $hook installé" -ForegroundColor Green
    } else {
        Write-Host "⚠️ Hook $hook introuvable dans $hooksSource" -ForegroundColor Yellow
    }
}

# Créer le modèle de message de commit
git config commit.template ".gitmessage"

Write-Host "`nConfiguration de Git terminée:" -ForegroundColor Cyan
Write-Host "- Hooks Git installés dans $hooksTarget" -ForegroundColor White
Write-Host "- Template de commit configuré" -ForegroundColor White

Write-Host "`n🎉 Installation terminée avec succès!" -ForegroundColor Green