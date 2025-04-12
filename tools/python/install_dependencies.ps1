# DÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©finir l'encodage en UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

Write-Host "DÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©marrage de l'installation des dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©pendances..." -ForegroundColor Green

# Configuration des variables d'environnement
Write-Host "Recherche de Python 3.12..." -ForegroundColor Cyan

# Chemins possibles pour Python 3.12
$possiblePaths = @(
    "C:\Users\$env:USERNAME\AppData\Local\Programs\Python\Python312",
    "C:\Program Files\Python312",
    "C:\Python312"
)

$pythonPath = $null

foreach ($path in $possiblePaths) {
    if (Test-Path "$path\python.exe") {
        $pythonPath = $path
        break
    }
}

# VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier si Python est trouvÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© via commande 'py'
if (-not $pythonPath) {
    try {
        $pyVersion = py -3.12 --version
        if ($pyVersion -match "Python 3.12") {
            $pythonPathOutput = py -3.12 -c "import sys; print(sys.executable)"
            $pythonPath = Split-Path -Parent $pythonPathOutput
            Write-Host "Python 3.12 trouvÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© via launcher py: $pythonPath" -ForegroundColor Green
        }
    } catch {
        # Py launcher non disponible ou Python 3.12 non disponible via launcher
    }
}

# VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier si python est disponible dans PATH
if (-not $pythonPath) {
    try {
        $pythonVersionCheck = python -V
        if ($pythonVersionCheck -match "Python 3.12") {
            $pythonPathOutput = python -c "import sys; print(sys.executable)"
            $pythonPath = Split-Path -Parent $pythonPathOutput
            Write-Host "Python 3.12 trouvÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© dans PATH: $pythonPath" -ForegroundColor Green
        }
    } catch {
        # Python non disponible dans PATH
    }
}

if (-not $pythonPath) {
    Write-Host "Python 3.12 n'a pas ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©tÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© trouvÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©. Chemins vÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifiÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©s:" -ForegroundColor Red
    foreach ($path in $possiblePaths) {
        Write-Host " - $path" -ForegroundColor Red
    }
    Write-Host "Veuillez installer Python 3.12 ou spÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cifier manuellement le chemin dans ce script." -ForegroundColor Red
    exit 1
}

Write-Host "Python trouvÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© : $pythonPath" -ForegroundColor Green

# Ajouter Python au PATH systÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨me
$currentPath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
if ($currentPath -notlike "*$pythonPath*") {
    $newPath = "$currentPath;$pythonPath;$pythonPath\Scripts"
    [Environment]::SetEnvironmentVariable("PATH", $newPath, "Machine")
    Write-Host "Python ajoutÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© au PATH systÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨me." -ForegroundColor Green
}

# Ajouter PYTHONPATH
$pythonPathEnv = [Environment]::GetEnvironmentVariable("PYTHONPATH", "Machine")
if (-not $pythonPathEnv) {
    [Environment]::SetEnvironmentVariable("PYTHONPATH", $pythonPath, "Machine")
    Write-Host "PYTHONPATH configurÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©." -ForegroundColor Green
}

# Recharger les variables d'environnement
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")

Write-Host "Variables d'environnement configurÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©es avec succÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s." -ForegroundColor Green

# Configurer le cache pip
$pipConfigPath = "$env:APPDATA\pip"
if (-not (Test-Path $pipConfigPath)) {
    New-Item -Path $pipConfigPath -ItemType Directory -Force | Out-Null
}

$pipConfigContent = @"
[global]
cache-dir=C:/ApexEnv/Python_Config/pip_cache
"@

$pipConfigContent | Out-File -FilePath "$pipConfigPath\pip.ini" -Encoding utf8 -Force
Write-Host "Configuration PIP crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e avec succÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s." -ForegroundColor Green

# VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier si pip est installÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©
try {
    $pipVersion = & "$pythonPath\python.exe" -m pip --version
    Write-Host "pip version : $pipVersion" -ForegroundColor Green
} catch {
    Write-Host "pip n'est pas installÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©. Tentative d'installation..." -ForegroundColor Yellow
    try {
        & "$pythonPath\python.exe" -m ensurepip --default-pip
        $pipVersion = & "$pythonPath\python.exe" -m pip --version
        Write-Host "pip installÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© avec succÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s : $pipVersion" -ForegroundColor Green
    } catch {
        Write-Host "Erreur lors de l'installation de pip : $_" -ForegroundColor Red
        exit 1
    }
}

# Mettre ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  jour pip
Write-Host "Mise ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  jour de pip..." -ForegroundColor Cyan
try {
    & "$pythonPath\python.exe" -m pip install --upgrade pip
} catch {
    Write-Host "Erreur lors de la mise ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  jour de pip : $_" -ForegroundColor Red
    exit 1
}

# Installer les dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©pendances
Write-Host "Installation des dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©pendances..." -ForegroundColor Cyan
try {
    & "$pythonPath\python.exe" -m pip install -r requirements.txt
} catch {
    Write-Host "Erreur lors de l'installation des dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©pendances : $_" -ForegroundColor Red
    exit 1
}

# VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier l'installation
Write-Host "VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification de l'installation..." -ForegroundColor Cyan
$packages = @(
    "pandas",
    "matplotlib",
    "seaborn",
    "plotly",
    "jupyter",
    "pytest",
    "pytest-benchmark"
)

foreach ($package in $packages) {
    try {
        $installed = & "$pythonPath\python.exe" -m pip show $package
        Write-Host "$package est installÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©" -ForegroundColor Green
    } catch {
        Write-Host "$package n'est pas installÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© correctement" -ForegroundColor Red
        exit 1
    }
}

Write-Host "Installation terminÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e avec succÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s!" -ForegroundColor Green
Write-Host "Vous pouvez maintenant exÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cuter les scripts d'analyse des logs IA." -ForegroundColor Green 