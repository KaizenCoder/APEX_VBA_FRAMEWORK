# =============================================================================
# APEX Workbench - Module de Logging Commun
# =============================================================================

$script:logFile = Join-Path $PSScriptRoot "../../logs/workbench.log"
$script:logLevels = @{
    "INFO"    = 0
    "WARNING" = 1
    "ERROR"   = 2
}

function Initialize-LogDirectory {
    $logDir = Split-Path $script:logFile -Parent
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
}

function Write-WorkbenchLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    
    Initialize-LogDirectory
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] $Message"
    
    # Écriture dans le fichier
    Add-Content -Path $script:logFile -Value $logEntry -Encoding UTF8
    
    # Affichage console avec couleur
    $color = switch ($Level) {
        "INFO" { "White" }
        "WARNING" { "Yellow" }
        "ERROR" { "Red" }
    }
    Write-Host $logEntry -ForegroundColor $color
}

# Export de la fonction
Export-ModuleMember -Function Write-WorkbenchLog 