# Fix-ApexEncoding.ps1
<#
.SYNOPSIS
    Corrige les problemes d'encodage typiques dans les scripts PowerShell
.DESCRIPTION
    Parcourt recursivement un dossier et corrige les caracteres mal encodes
.NOTES
    Le fichier est reenregistre en UTF-8 sans BOM.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory=$true)]
    [string]$Path,
    
    [Parameter()]
    [string[]]$Extensions = @("*.ps1", "*.psm1", "*.psd1"),
    
    [Parameter()]
    [switch]$Recursive,
    
    [Parameter()]
    [switch]$Backup,
    
    [Parameter()]
    [switch]$Validate
)

# Import des modules
Import-Module (Join-Path $PSScriptRoot "modules\EncodingCore.psm1")

function Process-File {
    param([string]$FilePath)
    
    Write-Host "Traitement: $FilePath"
    
    try {
        # Backup si demandé
        if ($Backup) {
            $backupPath = Backup-File $FilePath
            Write-Verbose "Backup créé: $backupPath"
        }
        
        # Lecture avec détection d'encodage
        try {
            $content = [System.IO.File]::ReadAllText($FilePath)
        }
        catch {
            $bytes = [System.IO.File]::ReadAllBytes($FilePath)
            $encodings = @('UTF8', '1252', '850')
            foreach ($enc in $encodings) {
                try {
                    $content = [System.Text.Encoding]::GetEncoding($enc).GetString($bytes)
                    break
                }
                catch { continue }
            }
        }
        
        # Application des patterns
        $patterns = Get-EncodingPatterns
        $originalContent = $content
        
        foreach ($p in $patterns) {
            $content = $content -replace $p.Pattern, $p.Replacement
        }
        
        # Si aucun changement, skip
        if ($content -eq $originalContent) {
            Write-Verbose "Aucune modification nécessaire pour $FilePath"
            return
        }
        
        # Validation si demandée
        if ($Validate) {
            $valid = Test-PowerShellSyntax $content
            if (-not $valid) {
                Write-Error "Validation échouée pour $FilePath"
                return
            }
        }
        
        # Sauvegarde
        if ($PSCmdlet.ShouldProcess($FilePath, "Correction d'encodage")) {
            [System.IO.File]::WriteAllText($FilePath, $content, [System.Text.UTF8Encoding]::new($false))
            Write-Host "Terminé: $FilePath" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Erreur lors du traitement de $FilePath : $_"
    }
}

# Traitement principal
try {
    Write-Host "Démarrage de la correction d'encodage..."
    Write-Host "Dossier cible: $Path"
    
    # Récupération des fichiers
    $files = Get-ChildItem -Path $Path -Include $Extensions -Recurse:$Recursive
    
    foreach ($file in $files) {
        Process-File $file.FullName
    }
    
    Write-Host "Correction d'encodage terminée avec succès." -ForegroundColor Green
}
catch {
    Write-Error "Erreur globale: $_"
} 