# Script de gestion des logs de session
# Ce script peut être utilisé indépendamment du processus de commit
# Référence: chat_038 (2024-04-11 16:30 - Correction encodage)
# Source: chat_002 (2024-04-09 10:15 - Règles encodage)

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$Title,
    
    [Parameter(Mandatory = $false)]
    [string]$Test,
    
    [Parameter(Mandatory = $false)]
    [string]$Module,
    
    [Parameter(Mandatory = $false)]
    [string]$CommitMessage,
    
    [Parameter(Mandatory = $false)]
    [string]$Summary,
    
    [Parameter(Mandatory = $false)]
    [string]$Duration
)

# Trouver le module ApexWSLBridge de façon flexible
$modulePath = $null
$possiblePaths = @(
    ".\tools\workflow\modules\ApexWSLBridge",
    "..\modules\ApexWSLBridge",
    "..\..\modules\ApexWSLBridge"
)

foreach ($path in $possiblePaths) {
    if (Test-Path $path) {
        $modulePath = $path
        break
    }
}

if (-not $modulePath) {
    Write-Error "Module ApexWSLBridge introuvable"
    exit 1
}

Import-Module $modulePath -Force

# Obtenir le chemin du fichier de log
$logFile = Get-SessionLogPath

if (-not $logFile) {
    Write-Error "Impossible de déterminer le fichier de log"
    exit 1
}

# Si titre vide, demander à l'utilisateur
if (-not $Title) {
    $Title = Read-Host "Entrez le titre de la session"
}

# Créer le fichier s'il n'existe pas
if (-not (Test-Path $logFile)) {
    $template = @"
# 🔧 Session de travail - {{DATE}}

## ✨ Objectifs

## 📝 Tests effectués

## 🛠️ Modules utilisés

## 💬 Messages de commit

## 📊 Résumé

---

"@
    $template = $template -replace '{{DATE}}', (Get-Date -Format "yyyy-MM-dd")
    Set-Content -Path $logFile -Value $template -Encoding UTF8
}

# Lire le contenu actuel
$content = Get-Content -Path $logFile -Raw

# Mettre à jour le titre si fourni
if ($Title) {
    $content = $content -replace '# 🔧 Session de travail .*', "# 🔧 $Title"
}

# Écrire le fichier
Set-Content -Path $logFile -Value $content -Encoding UTF8

# Mettre à jour selon les éléments fournis
if ($Test) {
    $content = $content -replace "(## ✨ Objectifs\n\n)", "$1- [ ] $test`n"
}

if ($Module) {
    $content = $content -replace "(## 🛠️ Modules utilisés\n\n)", "$1- $Module`n"
}

if ($CommitMessage) {
    $commitText = $CommitMessage -split "`n" | ForEach-Object { "- $_" } | Join-String -Separator "`n"
    $content = $content -replace "(## 💬 Messages de commit\n\n).*?(---)", "$1$commitText`n`n---"
}

if ($Summary) {
    $content = $content -replace "(## 📊 Résumé\n\n).*?($)", "$1$Summary`n"
}

# Écrire le contenu mis à jour
Set-Content -Path $logFile -Value $content -Encoding UTF8

# Mettre à jour le résumé si fourni
if ($Duration) {
    $content = Get-Content -Path $logFile -Raw
    $content = $content -replace "(🔧 Session de travail .* {{DATE}}\n)", "$1\n**Durée de la session**: $Duration\n"
    
    # Écrire le contenu mis à jour
    Set-Content -Path $logFile -Value $content -Encoding UTF8
}

# Retourner le chemin du fichier pour utilisation dans d'autres scripts
return $logFile