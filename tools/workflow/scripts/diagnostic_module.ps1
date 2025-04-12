########################################################################
# Diagnostic et tests pour le module ApexWSLBridge
# Ce script analyse les probla"...e""mes d'inta"...e""gration entre les scripts et le module
########################################################################

# Configuration des couleurs
$infoColor = "Cyan"
$successColor = "Green"
$errorColor = "Red"
$sectionColor = "Magenta"
$highlightColor = "Yellow"

function Write-Section {
    param([string]$Title)
    Write-Host "`n=== $Title ===`n" -ForegroundColor $sectionColor
}

# Affichage des informations de base
Write-Section "INFORMATIONS SYSTa""eee""ME"
Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor $infoColor
Write-Host "Utilisateur: $env:USERNAME" -ForegroundColor $infoColor
Write-Host "Ordinateur: $env:COMPUTERNAME" -ForegroundColor $infoColor
Write-Host "Ra"...e""pertoire courant: $(Get-Location)" -ForegroundColor $infoColor

# Chemins de ra"...e""fa"...e""rence
$projectRoot = "D:\Dev\Apex_VBA_FRAMEWORK"
$toolsRoot = Join-Path -Path $projectRoot -ChildPath "tools"
$expectedModulePath = Join-Path -Path $toolsRoot -ChildPath "powershell\ApexWSLBridge.psm1"
$commitScriptPath = Join-Path -Path $projectRoot -ChildPath "tools\workflow\scripts\commit_with_context.ps1"

# Analyse des chemins du module
Write-Section "ANALYSE DES CHEMINS"
Write-Host "Chemins de ra"...e""fa"...e""rence:" -ForegroundColor $infoColor
Write-Host "1. Racine du projet: $projectRoot" -ForegroundColor $highlightColor
Write-Host "2. Ra"...e""pertoire tools: $toolsRoot" -ForegroundColor $highlightColor
Write-Host "3. Emplacement attendu du module: $expectedModulePath" -ForegroundColor $highlightColor
Write-Host "4. Script de commit: $commitScriptPath" -ForegroundColor $highlightColor

# Va"...e""rification de l'existence des fichiers
Write-Host "`nVa"...e""rifications des fichiers:" -ForegroundColor $infoColor
if (Test-Path $projectRoot) {
    Write-Host ""a""eee...aa""a""a""e"...eaaa"" Ra"...e""pertoire du projet trouva"...e""" -ForegroundColor $successColor
} else {
    Write-Host ""a""eee...aa""a""a""e"...eeee"" Ra"...e""pertoire du projet introuvable!" -ForegroundColor $errorColor
}

if (Test-Path $toolsRoot) {
    Write-Host ""a""eee...aa""a""a""e"...eaaa"" Ra"...e""pertoire tools trouva"...e""" -ForegroundColor $successColor
} else {
    Write-Host ""a""eee...aa""a""a""e"...eeee"" Ra"...e""pertoire tools introuvable!" -ForegroundColor $errorColor
}

if (Test-Path $expectedModulePath) {
    Write-Host ""a""eee...aa""a""a""e"...eaaa"" Module ApexWSLBridge trouva"...e"" a"...e"" l'emplacement attendu" -ForegroundColor $successColor
} else {
    Write-Host ""a""eee...aa""a""a""e"...eeee"" Module ApexWSLBridge introuvable a"...e"" l'emplacement attendu!" -ForegroundColor $errorColor
}

if (Test-Path $commitScriptPath) {
    Write-Host ""a""eee...aa""a""a""e"...eaaa"" Script de commit trouva"...e""" -ForegroundColor $successColor
} else {
    Write-Host ""a""eee...aa""a""a""e"...eeee"" Script de commit introuvable!" -ForegroundColor $errorColor
}

# Recherche du module dans le ra"...e""pertoire tools
Write-Section "RECHERCHE DU MODULE"
$foundModules = Get-ChildItem -Path $toolsRoot -Recurse -Filter "ApexWSLBridge.psm1" -ErrorAction SilentlyContinue

if ($foundModules.Count -gt 0) {
    Write-Host "Modules ApexWSLBridge trouva"...e""s ($($foundModules.Count)):" -ForegroundColor $successColor
    foreach ($module in $foundModules) {
        Write-Host "- $($module.FullName)" -ForegroundColor $successColor
    }
    
    # Utiliser le premier module trouva"...e"" pour les tests
    $moduleToTest = $foundModules[0].FullName
} else {
    Write-Host "Aucun module ApexWSLBridge.psm1 trouva"...e"" dans le ra"...e""pertoire tools!" -ForegroundColor $errorColor
    exit
}

# Analyse du script de commit
Write-Section "ANALYSE DU SCRIPT DE COMMIT"
if (Test-Path $commitScriptPath) {
    $commitScript = Get-Content -Path $commitScriptPath -Raw
    
    # Recherche de la section d'importation du module
    $importPattern = "Import-Module.*ApexWSLBridge\.psm1"
    $importMatch = [regex]::Match($commitScript, $importPattern)
    
    if ($importMatch.Success) {
        Write-Host "Import du module trouva"...e"" dans le script de commit:" -ForegroundColor $successColor
        $lineContext = ($commitScript -split "`n")[$importMatch.Index..($importMatch.Index + 10)]
        foreach ($line in $lineContext) {
            if ($line -match $importPattern) {
                Write-Host $line -ForegroundColor $highlightColor
            } else {
                Write-Host $line
            }
        }
        
        # Extraction du chemin relatif utilisa"...e""
        $pathPattern = 'Join-Path.*-ChildPath "(.*ApexWSLBridge\.psm1)"'
        $pathMatch = [regex]::Match($commitScript, $pathPattern)
        
        if ($pathMatch.Success) {
            $relativePath = $pathMatch.Groups[1].Value
            Write-Host "`nChemin relatif utilisa"...e"": $relativePath" -ForegroundColor $highlightColor
            
            # Simulation du calcul de chemin dans le script
            $scriptScriptRoot = Split-Path -Parent $commitScriptPath
            $scriptParent = Split-Path -Parent $scriptScriptRoot
            $calculatedPath = Join-Path -Path $scriptParent -ChildPath ($relativePath -replace '\\\\', '\')
            
            Write-Host "Chemin calcula"...e"" par le script: $calculatedPath" -ForegroundColor $highlightColor
            
            if (Test-Path $calculatedPath) {
                Write-Host ""a""eee...aa""a""a""e"...eaaa"" Le chemin calcula"...e"" par le script existe" -ForegroundColor $successColor
            } else {
                Write-Host ""a""eee...aa""a""a""e"...eeee"" Le chemin calcula"...e"" par le script N'EXISTE PAS!" -ForegroundColor $errorColor
                
                # Suggestion de correction
                $correctPath = "..\..\powershell\ApexWSLBridge.psm1"
                $suggestedPath = Join-Path -Path $scriptParent -ChildPath $correctPath
                
                if (Test-Path $suggestedPath) {
                    Write-Host "`nSuggestion de correction:" -ForegroundColor $successColor
                    Write-Host "Remplacer '$relativePath' par '$correctPath'" -ForegroundColor $highlightColor
                }
            }
        }
    } else {
        Write-Host "Import du module non trouva"...e"" dans le script de commit!" -ForegroundColor $errorColor
    }
} else {
    Write-Host "Script de commit introuvable!" -ForegroundColor $errorColor
}

# Test de chargement du module
Write-Section "TEST DE CHARGEMENT DU MODULE"
try {
    # Suppression du module s'il est da"...e""ja"...e"" charga"...e""
    if (Get-Module ApexWSLBridge) {
        Remove-Module ApexWSLBridge -Force
    }
    
    # Importation du module
    Import-Module $moduleToTest -Force
    Write-Host ""a""eee...aa""a""a""e"...eaaa"" Module charga"...e"" avec succesaa"a"aaa"'a""a"...e""s!" -ForegroundColor $successColor
    
    # Va"...e""rification des fonctions exporta"...e""es
    $commands = Get-Command -Module ApexWSLBridge
    Write-Host "Fonctions disponibles ($($commands.Count)):" -ForegroundColor $infoColor
    foreach ($cmd in $commands) {
        Write-Host "- $($cmd.Name)" -ForegroundColor $highlightColor
    }
} catch {
    Write-Host ""a""eee...aa""a""a""e"...eeee"" Erreur lors du chargement du module: $_" -ForegroundColor $errorColor
}

# Correction du probla"...e""me
Write-Section "SOLUTION PROPOSa""e"...ee""E"
Write-Host "1. Probla"...e""me identifia"...e"":" -ForegroundColor $infoColor
Write-Host "   Le script 'commit_with_context.ps1' tente de charger le module ApexWSLBridge" -ForegroundColor $highlightColor
Write-Host "   mais le chemin relatif utilisa"...e"" est incorrect ou le module n'existe pas" -ForegroundColor $highlightColor
Write-Host "   a"...e"" l'emplacement attendu par le script." -ForegroundColor $highlightColor

Write-Host "`n2. Solutions possibles:" -ForegroundColor $infoColor
Write-Host "   a) Corriger le chemin dans le script commit_with_context.ps1" -ForegroundColor $highlightColor
Write-Host "   b) Placer une copie du module a"...e"" l'emplacement attendu par le script" -ForegroundColor $highlightColor
Write-Host "   c) Modifier le script pour qu'il trouve dynamiquement le module" -ForegroundColor $highlightColor

$correctScriptContent = @'
# Script de commit avec contexte pour APEX VBA Framework
# Ce script permet d'enrichir les commits Git avec des informations contextuelles

# Trouver le module ApexWSLBridge de faa"...e""on flexible
$modulePath = $null
$possiblePaths = @(
    # Chemin relatif standard
    (Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath "..\..\powershell\ApexWSLBridge.psm1"),
    # Chemin absolu direct
    "D:\Dev\Apex_VBA_FRAMEWORK\tools\powershell\ApexWSLBridge.psm1",
    # Recherche dans tools
    (Get-ChildItem -Path "D:\Dev\Apex_VBA_FRAMEWORK\tools" -Recurse -Filter "ApexWSLBridge.psm1" -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName)
)

# Essayer chaque chemin possible
foreach ($path in $possiblePaths) {
    if ($path -and (Test-Path $path)) {
        $modulePath = $path
        break
    }
}

# Importer le module si trouva"...e""
if ($modulePath) {
    Import-Module $modulePath -Force
    $wslBridgeAvailable = $true
    Write-Host "Module ApexWSLBridge charga"...e"" depuis: $modulePath" -ForegroundColor Green
} else {
    $wslBridgeAvailable = $false
    Write-Host "Module ApexWSLBridge non trouva"...e"". Les commandes WSL pourraient a"...e""tre moins fiables." -ForegroundColor Yellow
}
'@

Write-Host "`n3. Code de correction sugga"...e""ra"...e"":" -ForegroundColor $infoColor
Write-Host $correctScriptContent -ForegroundColor $highlightColor

Write-Section "CONCLUSION"
Write-Host "Le diagnostic est termineaa"a"aaa"'a""a"...e""." -ForegroundColor $infoColor
Write-Host "Appliquez la solution sugga"...e""ra"...e""e pour ra"...e""soudre le probla"...e""me d'inta"...e""gration" -ForegroundColor $infoColor
Write-Host "entre le script de commit et le module ApexWSLBridge." -ForegroundColor $infoColor 