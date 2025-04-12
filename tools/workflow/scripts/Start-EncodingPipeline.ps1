# Pipeline de validation d'encodage
# Référence: chat_051 (2024-04-11 17:00)
# Source: chat_050 (Pipeline validation)

[CmdletBinding()]
param(
    [switch]$Fix
)

# Importer le module de validation
try {
    Import-Module (Join-Path $PSScriptRoot "modules/ApexWSLBridge.psm1") -Force -ErrorAction Stop
}
catch {
    Write-Error "Erreur lors du chargement du module : $_"
    exit 1
}

# Exécuter les tests d'encodage
Write-Host "🔍 Validation de l'encodage des fichiers..."
try {
    $results = Test-FileEncoding -ProjectRoot "."

    if ($null -eq $results) {
        Write-Error "Erreur : Résultats de validation invalides"
        exit 1
    }

    if ($results.Error) {
        Write-Error "Erreur lors de la validation : $($results.Error)"
        exit 1
    }

    if ($results.HasErrors) {
        Write-Host "`n❌ Problèmes d'encodage détectés:" -ForegroundColor Red
        foreach ($file in $results.InvalidFiles) {
            Write-Host "   - $($file.Path) : $($file.Encoding)" -ForegroundColor Yellow
        }

        if ($Fix) {
            Write-Host "`n🔧 Correction automatique des encodages..."
            $hasNonBOMErrors = $false
            $hasFixErrors = $false

            foreach ($file in $results.InvalidFiles) {
                if ($file.Encoding -eq "UTF-8 with BOM") {
                    try {
                        $content = Get-Content $file.Path -Raw -ErrorAction Stop
                        $utf8NoBOM = New-Object System.Text.UTF8Encoding $false
                        [System.IO.File]::WriteAllText($file.Path, $content, $utf8NoBOM)
                        Write-Host "   ✅ $($file.Path)" -ForegroundColor Green
                    }
                    catch {
                        Write-Error "   ❌ Erreur lors de la correction de $($file.Path): $_"
                        $hasFixErrors = $true
                    }
                }
                else {
                    Write-Warning "   ⚠️ $($file.Path) : Correction impossible ($($file.Encoding))"
                    $hasNonBOMErrors = $true
                }
            }

            Write-Host "`n✨ Corrections appliquées"
            if ($hasFixErrors -or $hasNonBOMErrors) {
                Write-Warning "Certains fichiers n'ont pas pu être corrigés"
                exit 1
            }
            exit 0
        }
        else {
            Write-Host "`n💡 Pour corriger automatiquement, utilisez: Start-EncodingPipeline.ps1 -Fix"
            exit 1
        }
    }
    else {
        Write-Host "✅ Tous les fichiers sont en UTF-8 sans BOM"
        exit 0
    }
}
catch {
    Write-Error "Erreur inattendue : $_"
    exit 1
} 