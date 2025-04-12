function Unregister-CursorHooks {
    [CmdletBinding()]
    param (
        [switch]$RemoveSessionFiles,
        [switch]$Force
    )

    Write-Host "🔄 Désinstallation des hooks Cursor..." -ForegroundColor Cyan

    # 1. Nettoyage du profil PowerShell
    $profilePath = $PROFILE.CurrentUserAllHosts
    if (Test-Path $profilePath) {
        $content = Get-Content $profilePath -Raw
        if ($content -match "(?ms)# Hook Cursor Rules.*?'@") {
            $newContent = $content -replace "(?ms)# Hook Cursor Rules.*?'@\r?\n", ""
            Set-Content -Path $profilePath -Value $newContent
            Write-Host "✅ Hooks supprimés du profil PowerShell" -ForegroundColor Green
        }
    }

    # 2. Suppression des variables d'environnement
    $envVars = @(
        'CURSOR_WORKSPACE',
        'CURSOR_RULES_LOADED'
    )
    foreach ($var in $envVars) {
        if (Test-Path "env:$var") {
            Remove-Item "env:$var"
            Write-Host "✅ Variable d'environnement $var supprimée" -ForegroundColor Green
        }
    }

    # 3. Nettoyage des fichiers de session
    if ($RemoveSessionFiles) {
        $sessionFiles = Get-ChildItem -Path (Get-Location) -Filter ".cursor-session-*.json"
        if ($sessionFiles) {
            $sessionFiles | Remove-Item -Force:$Force
            Write-Host "✅ Fichiers de session supprimés" -ForegroundColor Green
        }
    }

    # 4. Restauration des paramètres VS Code
    $vscodePath = ".vscode/settings.json"
    if (Test-Path $vscodePath) {
        $settings = Get-Content $vscodePath -Raw | ConvertFrom-Json
        
        # Suppression des configurations Cursor
        if ($settings.PSObject.Properties.Name -contains "workspaceInit.tasks") {
            $settings.PSObject.Properties.Remove("workspaceInit.tasks")
        }
        
        # Mise à jour du fichier
        $settings | ConvertTo-Json -Depth 10 | Set-Content $vscodePath
        Write-Host "✅ Configuration VS Code restaurée" -ForegroundColor Green
    }

    Write-Host "`n✨ Désinstallation terminée" -ForegroundColor Green
    Write-Host "Note: Redémarrez votre terminal pour appliquer tous les changements" -ForegroundColor Yellow
}

# Exécution avec confirmation
if ($Force -or (Read-Host "Voulez-vous désinstaller les hooks Cursor ? (O/N)") -eq 'O') {
    Unregister-CursorHooks -RemoveSessionFiles
} 