# Script PowerShell pour faciliter l'interaction avec WSL
# Permet d'exÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©cuter des commandes WSL de maniÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â¨re fiable

# Force l'encodage UTF-8 pour l'affichage
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

# Fonction principale pour exÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©cuter des commandes WSL
function Invoke-WSLCommand {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Command,
        
        [Parameter(Mandatory=$false)]
        [string]$Distribution = "Ubuntu-22.04"
    )
    
    try {
        $result = wsl --distribution $Distribution -- bash -c "$Command"
        return $result
    }
    catch {
        Write-Host "Erreur lors de l'exÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©cution de la commande WSL: $_" -ForegroundColor Red
        return $null
    }
}

# Fonction pour vÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©rifier l'accessibilitÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â© des fichiers
function Test-WSLFileAccess {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path,
        
        [Parameter(Mandatory=$false)]
        [string]$Distribution = "Ubuntu-22.04"
    )
    
    $wslPath = Convert-WindowsToWSLPath -WindowsPath $Path
    $result = Invoke-WSLCommand -Command "if [ -f '$wslPath' ]; then echo 'FILE_EXISTS'; elif [ -d '$wslPath' ]; then echo 'DIR_EXISTS'; else echo 'NOT_FOUND'; fi"
    
    return $result.Trim()
}

# Fonction pour convertir un chemin Windows en chemin WSL
function Convert-WindowsToWSLPath {
    param (
        [Parameter(Mandatory=$true)]
        [string]$WindowsPath
    )
    
    # Conversion simple pour les lecteurs standards
    if ($WindowsPath -match "^([A-Za-z]):(.*)$") {
        $drive = $matches[1].ToLower()
        $path = $matches[2] -replace "\\", "/"
        return "/mnt/$drive$path"
    }
    
    # Si c'est dÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©jÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â  un chemin WSL, le retourner tel quel
    if ($WindowsPath -match "^/mnt/[a-z]/.*$") {
        return $WindowsPath
    }
    
    # Sinon, essayer de le convertir avec WSL
    try {
        $wslPath = Invoke-WSLCommand -Command "wslpath '$WindowsPath'"
        return $wslPath.Trim()
    }
    catch {
        Write-Host "Impossible de convertir le chemin Windows en chemin WSL: $WindowsPath" -ForegroundColor Red
        return $null
    }
}

# Fonction pour crÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©er un fichier dans WSL
function New-WSLFile {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path,
        
        [Parameter(Mandatory=$true)]
        [string]$Content,
        
        [Parameter(Mandatory=$false)]
        [string]$Distribution = "Ubuntu-22.04"
    )
    
    $wslPath = Convert-WindowsToWSLPath -WindowsPath $Path
    
    # CrÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©er un fichier temporaire
    $tempFile = New-TemporaryFile
    $Content | Out-File -FilePath $tempFile -Encoding utf8
    
    # Convertir le chemin temporaire en chemin WSL
    $tempPath = $tempFile.FullName
    $wslTempPath = Invoke-WSLCommand -Command "wslpath '$tempPath'"
    
    # Copier le contenu dans WSL
    $result = Invoke-WSLCommand -Command "cat '$wslTempPath' > '$wslPath'"
    
    # Nettoyage
    Remove-Item -Path $tempFile -Force
    
    return Test-WSLFileAccess -Path $Path
}

# Fonction pour rendre un fichier exÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©cutable
function Set-WSLExecutable {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path,
        
        [Parameter(Mandatory=$false)]
        [string]$Distribution = "Ubuntu-22.04"
    )
    
    $wslPath = Convert-WindowsToWSLPath -WindowsPath $Path
    $result = Invoke-WSLCommand -Command "chmod +x '$wslPath'"
    
    # VÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©rifier si le fichier est maintenant exÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©cutable
    $check = Invoke-WSLCommand -Command "if [ -x '$wslPath' ]; then echo 'EXECUTABLE'; else echo 'NOT_EXECUTABLE'; fi"
    
    return $check.Trim()
}

# Tester les fonctions
Write-Host "Test des fonctions WSL Bridge:" -ForegroundColor Cyan
$testFile = "D:\Dev\Apex_VBA_FRAMEWORK\test_wsl_bridge.txt"
$testContent = "Test de crÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©ation de fichier via WSL Bridge`nLigne 2`nLigne 3"

Write-Host "CrÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©ation d'un fichier de test..." -ForegroundColor Yellow
$result = New-WSLFile -Path $testFile -Content $testContent
Write-Host "RÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©sultat: $result" -ForegroundColor Green

Write-Host "VÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©rification de l'existence du fichier..." -ForegroundColor Yellow
$result = Test-WSLFileAccess -Path $testFile
Write-Host "RÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â©sultat: $result" -ForegroundColor Green

Write-Host "WSL Bridge est prÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Âªt ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢ÃƒÆ’Ã†â€™"Ã‚Â aa"Å¡Ã‚Â¬a"Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€šÃ‚Â aa"Ã…Â¡Ãƒâ€šÃ‚Â¬a"Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢ÃƒÆ’"Â "Ã¢â€žÂ¢aa"Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’...Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ "â„¢aa"Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™"Ã…Â¡ÃƒÆ’"Å¡Ãƒâ€šÃ‚Â  l'emploi." -ForegroundColor Green 