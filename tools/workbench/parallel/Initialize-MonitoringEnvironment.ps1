# =============================================================================
# Initialisation de l'environnement de monitoring optimisé
# =============================================================================

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'

$rootPath = (Get-Item $PSScriptRoot).Parent.Parent.Parent.FullName

# Création des répertoires requis
$paths = @(
    (Join-Path $rootPath "logs/performance"),
    (Join-Path $rootPath "logs/performance/archive"),
    (Join-Path $rootPath "monitoring"),
    (Join-Path $rootPath "monitoring/reports"),
    (Join-Path $rootPath "monitoring/data")
)

# Création des répertoires
foreach ($path in $paths) {
    if (-not (Test-Path $path)) {
        Write-Verbose "📁 Création du répertoire: $path"
        New-Item -ItemType Directory -Path $path -Force | Out-Null
    }
}

# Test des permissions
foreach ($path in $paths) {
    try {
        $testFile = Join-Path $path "test.tmp"
        [System.IO.File]::WriteAllText($testFile, "Test")
        Remove-Item $testFile -Force
        Write-Verbose "✅ Permissions validées pour: $path"
    }
    catch {
        throw "❌ Erreur de permissions sur $path : $_"
    }
}

# Nettoyage des anciens processus de monitoring
Get-Process | Where-Object { 
    $_.Name -like "*monitor*" -or 
    $_.Name -like "*watch*" -or 
    $_.Name -like "*perf*" 
} | Stop-Process -Force -ErrorAction SilentlyContinue

Write-Verbose "✨ Environnement de monitoring initialisé avec succès"
exit 0