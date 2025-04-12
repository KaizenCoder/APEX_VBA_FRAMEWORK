# Script pont PowerShell pour Cursor
# Référence: chat_048 (2024-04-11 16:35)
# Source: chat_047 (Correction encodage)

# Force l'encodage UTF-8 sans BOM pour la sortie
[System.Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

# Affiche l'environnement PowerShell
Write-Output "=== Environnement PowerShell ==="
Write-Output "Version PowerShell: $($PSVersionTable.PSVersion)"
Write-Output "Nom de l'ordinateur: $env:COMPUTERNAME"
Write-Output "Utilisateur actuel: $env:USERNAME"
Write-Output "Répertoire actuel: $PWD"
Write-Output "Encodage PowerShell: UTF-8"
Write-Output "=== Fin de l'environnement ==="

# Exécute la commande si fournie
if ($args.Count -gt 1 -and $args[0] -eq "-Command") {
    Write-Output "`nExécution de la commande: $($args[1])"
    try {
        Invoke-Expression $args[1]
    }
    catch {
        Write-Error "Erreur lors de l'exécution: $_"
        exit 1
    }
} 