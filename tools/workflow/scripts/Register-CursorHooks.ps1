# Script d'enregistrement des hooks Cursor
function Register-CursorHooks {
    [CmdletBinding()]
    param()

    # Chemin du profil PowerShell
    $profilePath = $PROFILE.CurrentUserAllHosts
    
    # Création du profil si n'existe pas
    if (-not (Test-Path $profilePath)) {
        New-Item -Path $profilePath -ItemType File -Force
    }

    # Hook à ajouter
    $hookContent = @'
# Hook Cursor Rules
function Global:Initialize-CursorEnvironment {
    $cursorRulesPath = ".cursor-rules"
    if (Test-Path $cursorRulesPath) {
        $workspace = (Get-Location).Path
        $env:CURSOR_WORKSPACE = $workspace
        $env:CURSOR_RULES_LOADED = $false
        
        # Création du fichier de session
        $sessionFile = ".cursor-session-$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
        @{
            workspace = $workspace
            timestamp = (Get-Date).ToString('o')
            rules_version = (Get-Content $cursorRulesPath | Select-String "Version: ").ToString()
        } | ConvertTo-Json > $sessionFile

        # Chargement des règles
        Write-Host "🔄 Chargement des règles Cursor..." -ForegroundColor Cyan
        Get-Content $cursorRulesPath | Out-Null
        $env:CURSOR_RULES_LOADED = $true
        
        # Validation de l'environnement
        & "$workspace\tools\workflow\scripts\Test-CursorRules.ps1" -Quiet
    }
}

# Auto-initialisation au changement de répertoire
$Global:PWD_Previous = $PWD
function Global:Watch-Location {
    if ($PWD.Path -ne $Global:PWD_Previous) {
        $Global:PWD_Previous = $PWD.Path
        Initialize-CursorEnvironment
    }
}

# Hook de prompt PowerShell
function Global:prompt {
    Watch-Location
    "PS $($executionContext.SessionState.Path.CurrentLocation)$('>' * ($nestedPromptLevel + 1)) "
}
'@

    # Ajout du hook au profil
    if (-not (Get-Content $profilePath | Select-String "Hook Cursor Rules")) {
        Add-Content -Path $profilePath -Value "`n$hookContent"
        Write-Host "✅ Hooks Cursor installés dans le profil PowerShell" -ForegroundColor Green
    }
    else {
        Write-Host "ℹ️ Hooks Cursor déjà installés" -ForegroundColor Yellow
    }
}

# Installation des hooks
Register-CursorHooks 