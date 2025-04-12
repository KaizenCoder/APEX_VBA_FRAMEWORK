function Initialize-CursorSession {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$WorkspacePath = (Get-Location).Path
    )

    Write-Host "==================================================="
    Write-Host "     INITIALISATION AUTOMATIQUE CURSOR RULES        "
    Write-Host "==================================================="

    # 1. Lecture et validation du fichier .cursor-rules
    Write-Host "`n1. Lecture des règles Cursor..."
    $cursorRules = Get-Content -Path (Join-Path $WorkspacePath ".cursor-rules") -Raw
    if (-not $cursorRules) {
        throw "Erreur: Impossible de lire .cursor-rules"
    }
    Write-Host "✅ Règles APEX Framework lues"

    # 2. Vérification et correction de l'encodage
    Write-Host "`n2. Validation de l'encodage..."
    & "$WorkspacePath\tools\Fix-Encoding.ps1"

    # 3. Consultation des sessions
    Write-Host "`n3. Consultation des sessions..."
    $today = Get-Date -Format "yyyy_MM_dd"
    $sessionsPath = Join-Path $WorkspacePath "tools\workflow\sessions"
    $todaySessions = Get-ChildItem -Path $sessionsPath -Recurse -File | 
                    Where-Object { $_.Name -match $today }
    
    if ($todaySessions) {
        foreach ($session in $todaySessions) {
            Write-Host "   - Lecture: $($session.Name)"
            Get-Content $session.FullName | Out-Null
        }
    }
    Write-Host "✅ Sessions prioritaires consultées"

    # 4. Vérification des documents essentiels
    Write-Host "`n4. Vérification documentation essentielle..."
    $essentialDocs = @(
        "docs/requirements/powershell_encoding.md",
        "docs/Components/CoreArchitecture.md",
        "docs/GIT_COMMIT_CONVENTION.md"
    )

    foreach ($doc in $essentialDocs) {
        $docPath = Join-Path $WorkspacePath $doc
        if (-not (Test-Path $docPath)) {
            Write-Warning "Document manquant: $doc"
        }
        else {
            Get-Content $docPath | Out-Null
        }
    }
    Write-Host "✅ Documentation de référence consultée"

    # 5. Validation de la session
    Write-Host "`n5. Validation de la session..."
    $sessionFiles = Get-ChildItem -Path $sessionsPath -Recurse -File | 
                   Where-Object { $_.Name -match $today }
    foreach ($file in $sessionFiles) {
        & "$WorkspacePath\tools\workflow\scripts\Test-SessionMarkdownFormat.ps1" -Path $file.FullName
    }

    Write-Host "`n==================================================="
    Write-Host "     INITIALISATION TERMINÉE AVEC SUCCÈS            "
    Write-Host "==================================================="
    Write-Host "⚠️ Contexte requis pour continuer"
}

# Exécution de la fonction
Initialize-CursorSession 