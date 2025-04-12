# Add-TaskToSession.ps1
# Script pour ajouter une tache a une session de developpement APEX VBA Framework
# Interface simplifiee pour New-SessionLog.ps1

param (
    [Parameter(Mandatory=$true)]
    [string]$Name,
    
    [Parameter(Mandatory=$true)]
    [string]$Module,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("aÃƒâ€šÃ‚ÂÃƒâ€šÃ‚Â³ En cours", "aÃƒâ€¦"... Termine", "aÃƒâ€šÃ‚ÂÃƒâ€¦"â„¢ Abandonne", "aÃƒâ€¦Ã‚Â¡Ãƒâ€šÃ‚Â iÃƒâ€šÃ‚Â¸Ãƒâ€šÃ‚Â Bloque")]
    [string]$Status = "aÃƒâ€šÃ‚ÂÃƒâ€šÃ‚Â³ En cours",
    
    [Parameter(Mandatory=$false)]
    [string]$Comment = ""
)

# Importer le module principal
$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "New-SessionLog.ps1"
Import-Module $scriptPath -Force

# Ajouter la tache
$task = @{
    Name = $Name
    Module = $Module
    Status = $Status
    Comment = $Comment
}

Add-TaskToSession -Task $task 