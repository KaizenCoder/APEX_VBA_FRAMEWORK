# Script de nettoyage des anciens logs
$ErrorActionPreference = "Stop"

# Chemin à supprimer
$oldPath = "D:\Dev\Apex_VBA_FRAMEWORK\src\Tools\Logger\reports\sessions"

try {
    # Arrêt de tous les processus qui pourraient bloquer le dossier
    $processes = @("python", "pythonw", "cursor")
    foreach ($proc in $processes) {
        Get-Process | Where-Object { $_.ProcessName -like "*$proc*" } | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 3

    # Forcer la libération des handles
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    # Suppression du dossier avec retries
    $maxAttempts = 3
    $attempt = 0
    $success = $false
    
    while (-not $success -and $attempt -lt $maxAttempts) {
        $attempt++
        try {
            if (Test-Path $oldPath) {
                Remove-Item -Path $oldPath -Recurse -Force
                $success = $true
                Write-Host "Ancien dossier de logs supprimé avec succès"
            }
            else {
                Write-Host "Le dossier n'existe pas ou a déjà été supprimé"
                $success = $true
            }
        }
        catch {
            Write-Host "Tentative $attempt échouée, nouvelle tentative dans 2 secondes..."
            Start-Sleep -Seconds 2
        }
    }
    
    if (-not $success) {
        throw "Impossible de supprimer le dossier après $maxAttempts tentatives"
    }
}
catch {
    Write-Error "Erreur lors du nettoyage : $_"
    exit 1
} 