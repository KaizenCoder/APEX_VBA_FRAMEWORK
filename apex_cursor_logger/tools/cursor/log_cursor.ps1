# Script de journalisation pour Cursor
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)][string]$Prompt,
    [Parameter(Mandatory = $true)][string]$Response,
    [Parameter(Mandatory = $false)][string]$Agent = "Claude-3",
    [Parameter(Mandatory = $false)][string]$Note = "+"
)

# Configuration initiale de l'encodage PowerShell
$PSDefaultParameterValues['*:Encoding'] = 'utf8'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[System.Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[System.Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)

# Fonction de validation d'encodage améliorée
function Test-StringEncoding {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][string]$Source
    )
    
    try {
        # Test d'encodage UTF-8
        $utf8 = [System.Text.UTF8Encoding]::new($false)
        $bytes = $utf8.GetBytes($Text)
        $decoded = $utf8.GetString($bytes)
        
        # Vérification des caractères spéciaux
        $specialChars = [regex]::Matches($Text, '[\x{1F300}-\x{1F9FF}]')
        if ($specialChars.Count -gt 0) {
            Write-Host "ℹ️ Émojis détectés dans $Source : $($specialChars.Count) trouvés" -ForegroundColor Cyan
        }
        
        if ($decoded -ne $Text) {
            Write-Warning "⚠️ Problème d'encodage détecté dans $Source"
            return $false
        }
        return $true
    }
    catch {
        Write-Warning "❌ Erreur lors de la validation de l'encodage de $Source : $_"
        return $false
    }
}

# Fonction de nettoyage des caractères spéciaux
function Format-SpecialCharacters {
    param([string]$Text)
    
    # Remplacement des caractères problématiques connus
    $replacements = @{
        '—' = '-'      # Em dash
        '–' = '-'      # En dash
        ''' = ''''     # Smart quotes left
        ''' = ''''     # Smart quotes right
        '‟' = '"'      # Double quotes left
        '"' = '"'      # Double quotes right
        '…' = '...'    # Ellipsis
    }
    
    $result = $Text
    foreach ($key in $replacements.Keys) {
        $result = $result.Replace($key, $replacements[$key])
    }
    
    return $result
}

# Validation des paramètres
$encodingValid = @(
    (Test-StringEncoding -Text $Prompt -Source "Prompt"),
    (Test-StringEncoding -Text $Response -Source "Response"),
    (Test-StringEncoding -Text $Note -Source "Note")
) -notcontains $false

if (-not $encodingValid) {
    Write-Warning "⚠️ Certains paramètres contiennent des caractères non valides en UTF-8"
}

# Validation et correction de l'encodage du script
try {
    $encodingScript = Join-Path (Split-Path -Parent $PSScriptRoot) "workflow/scripts/Fix-ApexEncoding.ps1"
    if (Test-Path $encodingScript) {
        & $encodingScript -Path $MyInvocation.MyCommand.Path -Force
    }
}
catch {
    Write-Warning "❌ Impossible de corriger l'encodage : $_"
}

# Obtention du chemin du script
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonScript = Join-Path $ScriptDir "cursor-autolog.py"

# Vérification de l'existence du script Python
if (-not (Test-Path $PythonScript)) {
    throw "❌ Script Python non trouvé : $PythonScript"
}

try {
    # Configuration de l'environnement Python
    $env:PYTHONIOENCODING = "utf-8"
    $env:PYTHONLEGACYWINDOWSSTDIO = "utf-8"
    $env:PYTHONUTF8 = "1"
    
    # Nettoyage et échappement des paramètres
    $cleanPrompt = Format-SpecialCharacters -Text $Prompt
    $cleanResponse = Format-SpecialCharacters -Text $Response
    $cleanNote = Format-SpecialCharacters -Text $Note
    
    $escapedPrompt = [Management.Automation.WildcardPattern]::Escape($cleanPrompt)
    $escapedResponse = [Management.Automation.WildcardPattern]::Escape($cleanResponse)
    $escapedNote = [Management.Automation.WildcardPattern]::Escape($cleanNote)
    
    Write-Host "[ℹ] Journalisation de l'interaction..." -ForegroundColor Cyan
    
    # Exécution du script Python
    $Result = & python $PythonScript $escapedPrompt $escapedResponse $Agent $escapedNote
    
    # Affichage du résultat
    if ($LASTEXITCODE -eq 0) {
        Write-Host "[✓] Journalisation réussie" -ForegroundColor Green
        exit 0
    }
    else {
        Write-Error "❌ Erreur lors de la journalisation : $Result"
        exit 1
    }
}
catch {
    Write-Error "❌ Erreur PowerShell : $_"
    exit 1
}
