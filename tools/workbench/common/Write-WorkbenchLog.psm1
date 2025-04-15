# =============================================================================
# APEX Workbench - Module de Logging Commun
# =============================================================================

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'

$script:ModulePath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:RootPath = (Get-Item $script:ModulePath).Parent.Parent.Parent.FullName
$script:LogPath = Join-Path $script:RootPath "logs\workbench"
$script:LogFile = Join-Path $script:LogPath "workbench.log"

function Write-WorkbenchLog {
    [CmdletBinding()]
    param
    (
        [Parameter()]
        [string]
        $Message,
        
        [Parameter()]
        [ValidateSet('INFO', 'WARNING', 'ERROR')]
        [string]
        $Level = 'INFO'
    )

    # Créer le répertoire des logs s'il n'existe pas
    if (-not (Test-Path $script:LogPath)) {
        New-Item -ItemType Directory -Path $script:LogPath -Force | Out-Null
    }

    # Format du message
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Level] $Message"

    # Écriture dans le fichier
    Add-Content -Path $script:LogFile -Value $logMessage -Encoding UTF8

    # Affichage console avec couleur
    $color = switch ($Level) {
        'INFO' { 'White' }
        'WARNING' { 'Yellow' }
        'ERROR' { 'Red' }
    }
    Write-Host $logMessage -ForegroundColor $color
}

# Export de la fonction
Export-ModuleMember -Function Write-WorkbenchLog