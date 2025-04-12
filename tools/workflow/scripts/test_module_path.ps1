# Script pour vÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©rifier le chemin du module ApexWSLBridge

$scriptRoot = $PSScriptRoot
$workflowRoot = Split-Path -Parent $PSScriptRoot
$modulePath = Join-Path -Path $workflowRoot -ChildPath "..\..\powershell\ApexWSLBridge.psm1"
$modulePath2 = Join-Path -Path (Split-Path -Parent (Split-Path -Parent $workflowRoot)) -ChildPath "tools\powershell\ApexWSLBridge.psm1"

Write-Host "VÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©rification des chemins du module ApexWSLBridge:" -ForegroundColor Cyan
Write-Host "1. PSScriptRoot = $scriptRoot" -ForegroundColor Yellow
Write-Host "2. Parent (workflow) = $workflowRoot" -ForegroundColor Yellow
Write-Host "3. Chemin calculÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â© du module = $modulePath" -ForegroundColor Yellow
Write-Host "4. Chemin alternatif = $modulePath2" -ForegroundColor Yellow

Write-Host "`nTest d'existence:" -ForegroundColor Cyan
if (Test-Path $modulePath) {
    Write-Host "Le chemin calculÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â© existe: $modulePath" -ForegroundColor Green
} else {
    Write-Host "Le chemin calculÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â© n'existe PAS: $modulePath" -ForegroundColor Red
}

if (Test-Path $modulePath2) {
    Write-Host "Le chemin alternatif existe: $modulePath2" -ForegroundColor Green
} else {
    Write-Host "Le chemin alternatif n'existe PAS: $modulePath2" -ForegroundColor Red
}

# Recherche du module dans tout le rÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©pertoire tools
Write-Host "`nRecherche du module dans tout le rÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©pertoire tools:" -ForegroundColor Cyan
$toolsRoot = Join-Path -Path (Split-Path -Parent (Split-Path -Parent $workflowRoot)) -ChildPath "tools"
$foundModules = Get-ChildItem -Path $toolsRoot -Recurse -Filter "ApexWSLBridge.psm1" -ErrorAction SilentlyContinue

if ($foundModules.Count -gt 0) {
    Write-Host "Modules trouvÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©s:" -ForegroundColor Green
    foreach ($module in $foundModules) {
        Write-Host "- $($module.FullName)" -ForegroundColor Green
    }
} else {
    Write-Host "Aucun module ApexWSLBridge.psm1 trouvÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â© dans $toolsRoot" -ForegroundColor Red
}

# Tester le chargement du module s'il est trouvÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©
if ($foundModules.Count -gt 0) {
    $moduleToTest = $foundModules[0].FullName
    Write-Host "`nTest de chargement du module $moduleToTest" -ForegroundColor Cyan
    try {
        Import-Module $moduleToTest -Force
        Write-Host "Module chargÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â© avec succÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â¨s!" -ForegroundColor Green
        
        # VÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©rification des fonctions exportÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©es
        $commands = Get-Command -Module ApexWSLBridge
        Write-Host "Fonctions disponibles: $($commands.Count)" -ForegroundColor Yellow
        foreach ($cmd in $commands) {
            Write-Host "- $($cmd.Name)" -ForegroundColor Gray
        }
    } catch {
        Write-Host "Erreur lors du chargement du module: $_" -ForegroundColor Red
    }
} 