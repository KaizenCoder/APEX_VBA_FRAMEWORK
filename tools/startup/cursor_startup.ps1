# Script de démarrage pour Cursor
$ErrorActionPreference = "Stop"

# Chemins
$workspaceRoot = "D:\Dev\Apex_VBA_FRAMEWORK"
$loggerPath = Join-Path $workspaceRoot "src\Tools\Logger\logger_gui.py"
$venvPath = Join-Path $workspaceRoot "src\Tools\Logger\.venv"
$activateScript = Join-Path $venvPath "Scripts\Activate.ps1"
$pythonPath = Join-Path $venvPath "Scripts\python.exe"

try {
    # Vérification de l'existence des fichiers
    if (-not (Test-Path $loggerPath)) {
        Write-Error "Interface graphique non trouvée : $loggerPath"
        exit 1
    }

    # Activation de l'environnement virtuel si nécessaire
    if (Test-Path $activateScript) {
        . $activateScript
    }

    # Vérification si l'interface est déjà lancée
    $existingProcess = Get-Process | Where-Object { $_.ProcessName -eq "python" -and $_.CommandLine -like "*logger_gui.py*" }
    if ($existingProcess) {
        Write-Host "Interface déjà en cours d'exécution"
        exit 0
    }

    # Lancement de l'interface graphique
    Start-Process -FilePath $pythonPath -ArgumentList $loggerPath -WindowStyle Hidden -NoNewWindow
    Write-Host "Interface de journalisation lancée avec succès"
}
catch {
    Write-Error "Erreur lors du lancement de l'interface : $_"
    exit 1
} 