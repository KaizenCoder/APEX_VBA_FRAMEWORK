# =========================================================================
# install_xlwings.ps1
# Description : Script d'installation et configuration de xlwings pour APEX Framework
# Auteur      : ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â°quipe APEX Framework
# Date        : 2025-04-15
# Version     : 1.0
# =========================================================================

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  INSTALLATION XLWINGS POUR APEX FRAMEWORK   " -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier les permissions d'administrateur
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€¦Ã‚Â¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã‚Â¯Ãƒâ€šÃ‚Â¸Ãƒâ€šÃ‚Â Ce script ne s'exÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cute pas en tant qu'administrateur." -ForegroundColor Yellow
    Write-Host "   Certaines opÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rations pourraient ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©chouer." -ForegroundColor Yellow
    Write-Host ""
}

# VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier l'installation de Python
Write-Host "ÃƒÆ’Ã‚Â°Ãƒâ€¦Ã‚Â¸ÃƒÂ¢Ã¢â€šÂ¬Ã‚ÂÃƒâ€šÃ‚Â VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification de Python..." -ForegroundColor Cyan
try {
    $pythonVersion = python --version
    Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€¦Ã¢â‚¬Å“ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ Python dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©tectÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© : $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€šÃ‚ÂÃƒâ€¦Ã¢â‚¬â„¢ Python n'est pas installÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© ou n'est pas dans le PATH" -ForegroundColor Red
    Write-Host "   Veuillez installer Python 3.8+ depuis https://www.python.org/" -ForegroundColor Red
    exit 1
}

# VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier pip
Write-Host "ÃƒÆ’Ã‚Â°Ãƒâ€¦Ã‚Â¸ÃƒÂ¢Ã¢â€šÂ¬Ã‚ÂÃƒâ€šÃ‚Â VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification de pip..." -ForegroundColor Cyan
try {
    $pipVersion = pip --version
    Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€¦Ã¢â‚¬Å“ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ pip dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©tectÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© : $pipVersion" -ForegroundColor Green
} catch {
    Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€šÃ‚ÂÃƒâ€¦Ã¢â‚¬â„¢ pip n'est pas installÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© ou n'est pas dans le PATH" -ForegroundColor Red
    Write-Host "   Essai d'installation automatique de pip..." -ForegroundColor Yellow
    
    try {
        python -m ensurepip --upgrade
        Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€¦Ã¢â‚¬Å“ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ pip installÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© avec succÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s" -ForegroundColor Green
    } catch {
        Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€šÃ‚ÂÃƒâ€¦Ã¢â‚¬â„¢ ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â°chec de l'installation de pip" -ForegroundColor Red
        exit 1
    }
}

# Mettre ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  jour pip
Write-Host "ÃƒÆ’Ã‚Â°Ãƒâ€¦Ã‚Â¸ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œÃƒâ€šÃ‚Â¦ Mise ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  jour de pip..." -ForegroundColor Cyan
python -m pip install --upgrade pip
Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€¦Ã¢â‚¬Å“ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ pip mis ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  jour" -ForegroundColor Green

# Installer les dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©pendances requises
Write-Host "ÃƒÆ’Ã‚Â°Ãƒâ€¦Ã‚Â¸ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œÃƒâ€šÃ‚Â¦ Installation des dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©pendances..." -ForegroundColor Cyan
$requirements = @("pandas", "openpyxl")
foreach ($pkg in $requirements) {
    Write-Host "   Installation de $pkg..." -ForegroundColor Cyan
    pip install $pkg
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€¦Ã‚Â¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã‚Â¯Ãƒâ€šÃ‚Â¸Ãƒâ€šÃ‚Â ProblÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨me lors de l'installation de $pkg" -ForegroundColor Yellow
    }
}
Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€¦Ã¢â‚¬Å“ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ DÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©pendances installÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©es" -ForegroundColor Green

# Installer xlwings
Write-Host "ÃƒÆ’Ã‚Â°Ãƒâ€¦Ã‚Â¸ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œÃƒâ€šÃ‚Â¦ Installation de xlwings..." -ForegroundColor Cyan
pip install --upgrade xlwings

if ($LASTEXITCODE -ne 0) {
    Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€šÃ‚ÂÃƒâ€¦Ã¢â‚¬â„¢ ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â°chec de l'installation de xlwings" -ForegroundColor Red
    exit 1
} else {
    # VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier la version de xlwings
    $xlwingsVersion = python -c "import xlwings; print(xlwings.__version__)"
    Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€¦Ã¢â‚¬Å“ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ xlwings $xlwingsVersion installÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© avec succÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s" -ForegroundColor Green
}

# Installer l'add-in Excel
Write-Host "ÃƒÆ’Ã‚Â°Ãƒâ€¦Ã‚Â¸ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œÃƒâ€šÃ‚Â¦ Installation de l'add-in Excel..." -ForegroundColor Cyan
python -c "from xlwings.cli import main; main()" addin install

if ($LASTEXITCODE -ne 0) {
    Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€¦Ã‚Â¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã‚Â¯Ãƒâ€šÃ‚Â¸Ãƒâ€šÃ‚Â ProblÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨me lors de l'installation de l'add-in xlwings." -ForegroundColor Yellow
    Write-Host "   Vous devrez peut-ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Âªtre l'installer manuellement." -ForegroundColor Yellow
} else {
    # VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier l'installation de l'add-in
    $addinStatus = python -c "from xlwings.cli import main; main()" addin status
    Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€¦Ã¢â‚¬Å“ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ Add-in xlwings installÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©" -ForegroundColor Green
    Write-Host "   $addinStatus" -ForegroundColor Green
}

# Test xlwings
Write-Host "ÃƒÆ’Ã‚Â°Ãƒâ€¦Ã‚Â¸Ãƒâ€šÃ‚Â§Ãƒâ€šÃ‚Âª Test de fonctionnement de xlwings..." -ForegroundColor Cyan
$testScriptPath = Join-Path $PSScriptRoot "test_xlwings.py"

if (Test-Path $testScriptPath) {
    $runTest = Read-Host "Voulez-vous exÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cuter le script de test xlwings? (o/n)"
    if ($runTest -eq "o") {
        Write-Host "ExÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cution du test..." -ForegroundColor Cyan
        
        try {
            & python $testScriptPath
            Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€¦Ã¢â‚¬Å“ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ Test xlwings terminÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©" -ForegroundColor Green
        } catch {
            Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€šÃ‚ÂÃƒâ€¦Ã¢â‚¬â„¢ Erreur lors du test xlwings: $_" -ForegroundColor Red
        }
    }
} else {
    Write-Host "ÃƒÆ’Ã‚Â¢Ãƒâ€¦Ã‚Â¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã‚Â¯Ãƒâ€šÃ‚Â¸Ãƒâ€šÃ‚Â Script de test non trouvÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $testScriptPath" -ForegroundColor Yellow
}

# RÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©sumÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©
Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "          INSTALLATION TERMINÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â°E              " -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Python      : $pythonVersion" -ForegroundColor White
Write-Host "pip         : $pipVersion" -ForegroundColor White
Write-Host "xlwings     : $xlwingsVersion" -ForegroundColor White
Write-Host ""
Write-Host "Add-in Excel: InstallÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© dans le dossier XLSTART" -ForegroundColor White
Write-Host ""
Write-Host "Pour utiliser xlwings dans vos scripts Python :" -ForegroundColor Green
Write-Host "import xlwings as xw" -ForegroundColor Green
Write-Host ""
Write-Host "Pour vÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier l'installation :" -ForegroundColor Yellow
Write-Host "python -c `"from xlwings.cli import main; main()`" addin status" -ForegroundColor Yellow
Write-Host ""
Write-Host "Pour plus d'informations, consultez :" -ForegroundColor Cyan
Write-Host "- docs/Components/XLWings_Integration.md" -ForegroundColor Cyan
Write-Host "- docs/AI_ONBOARDING_GUIDE.md" -ForegroundColor Cyan
Write-Host "- https://docs.xlwings.org/" -ForegroundColor Cyan
Write-Host "" 