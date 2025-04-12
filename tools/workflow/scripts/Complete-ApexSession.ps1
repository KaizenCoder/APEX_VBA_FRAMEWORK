# Complete-ApexSession.ps1
# Script pour terminer une session de dÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©veloppement APEX VBA Framework
# Interface simplifiÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©e pour New-SessionLog.ps1

param (
    [Parameter(Mandatory=$false)]
    [string]$SessionId = "",
    
    [Parameter(Mandatory=$false)]
    [string]$Summary = ""
)

# DÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©termination du rÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©pertoire de base du projet
$projectRoot = "D:\Dev\Apex_VBA_FRAMEWORK"
$scriptDirectory = Join-Path -Path $projectRoot -ChildPath "tools\workflow\scripts"

# Chemin absolu vers le script principal
$sessionLogScript = Join-Path -Path $scriptDirectory -ChildPath "New-SessionLog.ps1"

if (-not (Test-Path $sessionLogScript)) {
    Write-Host "Erreur: Script New-SessionLog.ps1 introuvable a: $sessionLogScript" -ForegroundColor Red
    exit 1
}

# Afficher un en-tÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Âªte
Clear-Host
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host "           TERMINER UNE SESSION APEX VBA               " -ForegroundColor Cyan
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host ""

# Demander un rÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©sumÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â© si non fourni
if ([string]::IsNullOrWhiteSpace($Summary)) {
    Write-Host "Entrez un resume de la session (terminez par une ligne vide):" -ForegroundColor Yellow
    
    $lines = @()
    $line = "dummy"
    
    while (-not [string]::IsNullOrWhiteSpace($line)) {
        $line = Read-Host
        if (-not [string]::IsNullOrWhiteSpace($line)) {
            $lines += $line
        }
    }
    
    if ($lines.Count -gt 0) {
        $Summary = $lines -join "`n"
    } else {
        $Summary = "Session terminee sans commentaire."
    }
}

# Lancer le script de crÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©ation de session
Write-Host "`nFinalisation de la session..." -ForegroundColor Cyan

# PrÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©parer la commande
$command = "& `"$sessionLogScript`" -Action Complete"

if (-not [string]::IsNullOrWhiteSpace($SessionId)) {
    $command += " -SessionId `"$SessionId`""
}

# Ajouter le rÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©sumÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â© comme un bloc de texte via un fichier temporaire
$tempFile = [System.IO.Path]::GetTempFileName()
$Summary | Out-File -FilePath $tempFile -Encoding utf8

$command += " -Summary (Get-Content -Path `"$tempFile`" -Raw)"

# ExÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©cuter la commande
try {
    $result = Invoke-Expression $command
    
    # Afficher un rÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©sumÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©
    Write-Host "`n=======================================================" -ForegroundColor Cyan
    Write-Host "Session terminee le: $(Get-Date -Format 'dd/MM/yyyy a HH:mm')" -ForegroundColor White
    Write-Host "=======================================================" -ForegroundColor Cyan
    
    Write-Host "`nResume de la session:" -ForegroundColor Yellow
    Write-Host $Summary -ForegroundColor White
    
    Write-Host "`nVous pouvez maintenant proceder a un commit de vos modifications." -ForegroundColor Green
} catch {
    Write-Host "Erreur lors de la finalisation de la session: $_" -ForegroundColor Red
} finally {
    # Nettoyer les fichiers temporaires
    if (Test-Path $tempFile) {
        Remove-Item -Path $tempFile -Force
    }
} 