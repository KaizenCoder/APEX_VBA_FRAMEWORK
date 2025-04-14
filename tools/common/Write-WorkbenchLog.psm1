# Module de logging pour le Workbench APEX
# Version: 1.0
# Date: 2024-04-14

$script:logPath = Join-Path $PSScriptRoot "../../logs"
if (-not (Test-Path $script:logPath)) {
    New-Item -ItemType Directory -Path $script:logPath -Force | Out-Null
}

function Write-WorkbenchLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )

    # Format du message
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Level] $Message"

    # Fichier de log principal
    $logFile = Join-Path $script:logPath "workbench.log"

    # Écriture dans le fichier
    Add-Content -Path $logFile -Value $logMessage

    # Affichage console avec couleur
    $color = switch ($Level) {
        "INFO" { "White" }
        "WARNING" { "Yellow" }
        "ERROR" { "Red" }
    }
    Write-Host $logMessage -ForegroundColor $color
}

# Export de la fonction
Export-ModuleMember -Function Write-WorkbenchLog 