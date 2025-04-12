# powershell_bridge.ps1
# Ce script sert d'interface entre l'assistant IA et PowerShell
# Il permet d'exécuter des commandes et de récupérer leurs résultats

param (
    [Parameter(Mandatory=$false)]
    [string]$Command = ""
)

function Show-SystemInfo {
    Write-Host "`n=== INFORMATIONS SYSTÈME ===" -ForegroundColor Cyan
    Write-Host "Version PowerShell : $($PSVersionTable.PSVersion)"
    Write-Host "Nom de l'ordinateur : $env:COMPUTERNAME"
    Write-Host "Utilisateur actuel : $env:USERNAME"
    Write-Host "Répertoire actuel : $(Get-Location)"
    Write-Host "================================`n"
}

function Execute-Command {
    param (
        [string]$CommandToExecute
    )
    
    try {
        Write-Host "`n>>> Exécution de: $CommandToExecute" -ForegroundColor Yellow
        Write-Host "------------------------" -ForegroundColor Yellow
        
        # Utilisation d'Invoke-Expression pour exécuter la commande
        $result = Invoke-Expression -Command $CommandToExecute
        
        # Afficher le résultat
        $result
        
        Write-Host "------------------------" -ForegroundColor Yellow
        Write-Host "Commande exécutée avec succès.`n" -ForegroundColor Green
    }
    catch {
        Write-Host "ERREUR: $_" -ForegroundColor Red
        Write-Host "------------------------`n" -ForegroundColor Yellow
    }
}

# Affichage initial
Write-Host "`n=====================================" -ForegroundColor Magenta
Write-Host "  PONT POWERSHELL - ASSISTANT IA" -ForegroundColor Magenta
Write-Host "=====================================" -ForegroundColor Magenta

# Afficher les informations système
Show-SystemInfo

# Si une commande est fournie en paramètre, l'exécuter
if ($Command -ne "") {
    Execute-Command -CommandToExecute $Command
}
else {
    Write-Host "Mode interactif. Utilisez ce script avec le paramètre -Command pour exécuter une commande spécifique." -ForegroundColor Cyan
    Write-Host "Exemple: .\powershell_bridge.ps1 -Command 'Get-Process | Select-Object -First 5'" -ForegroundColor Cyan
}

Write-Host "`nPont PowerShell prêt à l'emploi!" -ForegroundColor Green