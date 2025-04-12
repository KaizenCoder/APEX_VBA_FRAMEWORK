[CmdletBinding()]
param (
    [switch]$NoBackup,
    [switch]$Force
)

# Fonctions de journalisation
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage
    Add-Content -Path "$PSScriptRoot\integration.log" -Value $logMessage
}

# Vérification de l'installation de VS Code/Cursor
function Test-VSCodeInstallation {
    try {
        $cursorPath = Get-Command cursor -ErrorAction SilentlyContinue
        $vscodePath = Get-Command code -ErrorAction SilentlyContinue

        if ($cursorPath) {
            Write-Log "Cursor est installé à : $($cursorPath.Source)" -Level Info
            return $true
        }
        elseif ($vscodePath) {
            Write-Log "VS Code est installé à : $($vscodePath.Source)" -Level Info
            return $true
        }
        else {
            Write-Log "Ni Cursor ni VS Code n'est installé ou n'est dans le PATH" -Level Error
            return $false
        }
    }
    catch {
        Write-Log "Erreur lors de la vérification de l'installation : $_" -Level Error
        return $false
    }
}

# Installation des extensions requises avec vérification
function Install-RequiredExtensions {
    $extensions = @{
        'PowerShell'             = 'ms-vscode.powershell'
        'Python'                 = 'ms-python.python'
        'Python Debugger'        = 'ms-python.debugpy'
        'Python Language Server' = 'ms-python.vscode-pylance'
    }

    foreach ($ext in $extensions.GetEnumerator()) {
        try {
            $installed = (cursor --list-extensions 2>$null) -contains $ext.Value -or (code --list-extensions 2>$null) -contains $ext.Value
            if (-not $installed) {
                Write-Log "Installation de l'extension : $($ext.Key) ($($ext.Value))"
                if (Get-Command cursor -ErrorAction SilentlyContinue) {
                    & cursor --install-extension $ext.Value --force
                }
                else {
                    & code --install-extension $ext.Value --force
                }
                Start-Sleep -Seconds 2
            }
            else {
                Write-Log "L'extension $($ext.Key) est déjà installée" -Level Info
            }
        }
        catch {
            Write-Log "Erreur lors de l'installation de l'extension $($ext.Key) : $_" -Level Error
        }
    }
}

# Sauvegarde des configurations avec vérification
function Backup-Configurations {
    if (-not $NoBackup) {
        $backupPath = "$PSScriptRoot\backup"
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $backupFolder = "$backupPath\vscode_$timestamp"

        try {
            # Création du dossier de sauvegarde avec timestamp
            if (-not (Test-Path $backupPath)) {
                New-Item -Path $backupPath -ItemType Directory -Force | Out-Null
            }
            New-Item -Path $backupFolder -ItemType Directory -Force | Out-Null

            # Sauvegarde des fichiers VS Code s'ils existent
            $vscodePaths = @(
                "$env:USERPROFILE\.vscode",
                "$env:APPDATA\Code\User"
            )

            foreach ($path in $vscodePaths) {
                if (Test-Path $path) {
                    $destinationFolder = Join-Path $backupFolder (Split-Path $path -Leaf)
                    Copy-Item -Path "$path\*" -Destination $destinationFolder -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Log "Sauvegarde créée pour : $path" -Level Info
                }
            }
        }
        catch {
            Write-Log "Erreur lors de la sauvegarde : $_" -Level Error
        }
    }
}

# Configuration et vérification des variables d'environnement
function Set-IntegrationEnvironment {
    $workspaceRoot = (Get-Location).Path
    
    # Suppression des anciennes variables si elles existent
    [Environment]::SetEnvironmentVariable("CURSOR_WORKSPACE", $null, "User")
    [Environment]::SetEnvironmentVariable("VSCODE_WORKSPACE", $null, "User")
    
    # Définition des nouvelles variables
    [Environment]::SetEnvironmentVariable("CURSOR_WORKSPACE", $workspaceRoot, "User")
    [Environment]::SetEnvironmentVariable("VSCODE_WORKSPACE", $workspaceRoot, "User")
    
    # Vérification
    $cursorWs = [Environment]::GetEnvironmentVariable("CURSOR_WORKSPACE", "User")
    $vscodeWs = [Environment]::GetEnvironmentVariable("VSCODE_WORKSPACE", "User")
    
    if ($cursorWs -eq $workspaceRoot -and $vscodeWs -eq $workspaceRoot) {
        Write-Log "Variables d'environnement configurées avec succès" -Level Info
        return $true
    }
    else {
        Write-Log "Échec de la configuration des variables d'environnement" -Level Error
        return $false
    }
}

# Initialisation de l'environnement
function Initialize-Environment {
    $workspaceRoot = (Get-Item $PSScriptRoot).Parent.Parent.FullName
    
    # Création des dossiers nécessaires
    $paths = @(
        "$workspaceRoot\.vscode",
        "$workspaceRoot\.cursor-rules.d",
        "$workspaceRoot\logs"
    )

    foreach ($path in $paths) {
        if (-not (Test-Path $path)) {
            New-Item -Path $path -ItemType Directory -Force | Out-Null
            Write-Log "Création du dossier : $path"
        }
    }

    # Configuration de l'environnement
    if (-not (Set-IntegrationEnvironment)) {
        throw "Échec de la configuration de l'environnement"
    }
}

# Test de l'intégration
function Test-Integration {
    $testResults = @{
        VSCodeInstalled      = Test-VSCodeInstallation
        WorkspaceExists      = Test-Path ((Get-Item $PSScriptRoot).Parent.Parent.FullName)
        EnvironmentVariables = (
            [Environment]::GetEnvironmentVariable("CURSOR_WORKSPACE", "User") -ne $null -and
            [Environment]::GetEnvironmentVariable("VSCODE_WORKSPACE", "User") -ne $null
        )
    }

    $failed = $testResults.Values | Where-Object { -not $_ } | Measure-Object | Select-Object -ExpandProperty Count
    if ($failed -eq 0) {
        Write-Log "Tests d'intégration réussis" -Level Info
        return $true
    }
    else {
        Write-Log "Certains tests d'intégration ont échoué :" -Level Warning
        $testResults.GetEnumerator() | ForEach-Object {
            if (-not $_.Value) {
                Write-Log "- Échec : $($_.Key)" -Level Warning
            }
        }
        return $false
    }
}

# Vérification de l'alignement VS Code et Cursor
function Test-VSCodeCursorAlignment {
    Write-Log "Vérification de l'alignement VS Code/Cursor..." -Level Info
    $alignmentResults = @{}

    # Vérification des chemins d'installation
    $cursorPath = (Get-Command cursor -ErrorAction SilentlyContinue).Source
    $vscodePath = (Get-Command code -ErrorAction SilentlyContinue).Source
    
    # Vérification des extensions
    try {
        $cursorExts = @(cursor --list-extensions 2>$null)
        $vscodeExts = @(code --list-extensions 2>$null)
        
        $alignmentResults.Add("ExtensionsAlignment", @{
                "Communes"         = @($cursorExts | Where-Object { $vscodeExts -contains $_ })
                "UniquementCursor" = @($cursorExts | Where-Object { $vscodeExts -notcontains $_ })
                "UniquementVSCode" = @($vscodeExts | Where-Object { $cursorExts -notcontains $_ })
            })
    }
    catch {
        Write-Log "Erreur lors de la vérification des extensions : $_" -Level Error
    }

    # Vérification des variables d'environnement
    $cursorWs = [Environment]::GetEnvironmentVariable("CURSOR_WORKSPACE", "User")
    $vscodeWs = [Environment]::GetEnvironmentVariable("VSCODE_WORKSPACE", "User")
    $alignmentResults.Add("WorkspaceAlignment", $cursorWs -eq $vscodeWs)

    # Vérification des fichiers de configuration
    $configPaths = @{
        "VSCode" = "$env:APPDATA\Code\User\settings.json"
        "Cursor" = "$env:APPDATA\Cursor\User\settings.json"
    }

    foreach ($config in $configPaths.GetEnumerator()) {
        if (Test-Path $config.Value) {
            try {
                $content = Get-Content $config.Value -Raw | ConvertFrom-Json
                $alignmentResults.Add("$($config.Key)Config", $content)
            }
            catch {
                Write-Log "Erreur lors de la lecture de $($config.Key) settings : $_" -Level Error
            }
        }
    }

    # Affichage des résultats
    Write-Log "Résultats de l'alignement :" -Level Info
    Write-Log "- Workspace : $(if ($alignmentResults.WorkspaceAlignment) { 'Aligné' } else { 'Non aligné' })" -Level Info
    
    if ($alignmentResults.ExtensionsAlignment) {
        Write-Log "- Extensions communes : $($alignmentResults.ExtensionsAlignment.Communes.Count)" -Level Info
        if ($alignmentResults.ExtensionsAlignment.UniquementCursor.Count -gt 0) {
            Write-Log "- Extensions uniquement dans Cursor : $($alignmentResults.ExtensionsAlignment.UniquementCursor -join ', ')" -Level Warning
        }
        if ($alignmentResults.ExtensionsAlignment.UniquementVSCode.Count -gt 0) {
            Write-Log "- Extensions uniquement dans VS Code : $($alignmentResults.ExtensionsAlignment.UniquementVSCode -join ', ')" -Level Warning
        }
    }

    return $alignmentResults
}

# Vérification des règles Cursor dans VS Code
function Test-CursorRulesApplication {
    Write-Log "Vérification de l'application des règles Cursor..." -Level Info
    
    # Vérification du dossier .cursor-rules.d
    $workspaceRoot = (Get-Location).Path
    $cursorRulesPath = Join-Path $workspaceRoot ".cursor-rules.d"
    $vscodePath = Join-Path $workspaceRoot ".vscode"
    
    if (-not (Test-Path $cursorRulesPath)) {
        Write-Log "Dossier .cursor-rules.d non trouvé" -Level Warning
        return $false
    }

    # Vérification des fichiers de configuration
    $configFiles = @{
        "settings.json" = @{
            "cursor" = Join-Path $cursorRulesPath "settings.json"
            "vscode" = Join-Path $vscodePath "settings.json"
        }
        "launch.json"   = @{
            "cursor" = Join-Path $cursorRulesPath "launch.json"
            "vscode" = Join-Path $vscodePath "launch.json"
        }
        "tasks.json"    = @{
            "cursor" = Join-Path $cursorRulesPath "tasks.json"
            "vscode" = Join-Path $vscodePath "tasks.json"
        }
    }

    $results = @{}

    foreach ($file in $configFiles.GetEnumerator()) {
        $cursorFile = $file.Value.cursor
        $vscodeFile = $file.Value.vscode
        
        if (Test-Path $cursorFile) {
            try {
                $cursorContent = Get-Content $cursorFile -Raw | ConvertFrom-Json
                
                if (Test-Path $vscodeFile) {
                    $vscodeContent = Get-Content $vscodeFile -Raw | ConvertFrom-Json
                    
                    # Comparaison des configurations
                    $results[$file.Key] = @{
                        "Exists"     = $true
                        "InSync"     = (ConvertTo-Json $cursorContent -Depth 10) -eq (ConvertTo-Json $vscodeContent -Depth 10)
                        "CursorPath" = $cursorFile
                        "VSCodePath" = $vscodeFile
                    }
                    
                    Write-Log "- $($file.Key): $(if ($results[$file.Key].InSync) { 'Synchronisé' } else { 'Différences détectées' })" -Level $(if ($results[$file.Key].InSync) { 'Info' } else { 'Warning' })
                }
                else {
                    Write-Log "- $($file.Key): Fichier VS Code manquant" -Level Warning
                    $results[$file.Key] = @{
                        "Exists"     = $false
                        "InSync"     = $false
                        "CursorPath" = $cursorFile
                        "VSCodePath" = $null
                    }
                }
            }
            catch {
                Write-Log "Erreur lors de la lecture de $($file.Key): $_" -Level Error
                $results[$file.Key] = @{
                    "Exists" = $true
                    "InSync" = $false
                    "Error"  = $_.Exception.Message
                }
            }
        }
    }

    # Vérification des extensions requises
    if (Test-Path (Join-Path $cursorRulesPath "extensions.json")) {
        try {
            $cursorExtensions = Get-Content (Join-Path $cursorRulesPath "extensions.json") -Raw | ConvertFrom-Json
            $vscodeExtensions = code --list-extensions
            
            $missingExtensions = @($cursorExtensions.recommendations | Where-Object { $vscodeExtensions -notcontains $_ })
            if ($missingExtensions.Count -gt 0) {
                Write-Log "Extensions recommandées manquantes dans VS Code :" -Level Warning
                foreach ($ext in $missingExtensions) {
                    Write-Log "  - $ext" -Level Warning
                }
            }
            else {
                Write-Log "Toutes les extensions recommandées sont installées" -Level Info
            }
        }
        catch {
            Write-Log "Erreur lors de la vérification des extensions : $_" -Level Error
        }
    }

    return $results
}

# Vérification des logs VS Code et hooks Cursor
function Test-CursorHooksExecution {
    Write-Log "Vérification des logs VS Code et hooks Cursor..." -Level Info
    
    $logPaths = @(
        "$env:APPDATA\Code\logs\main.log",
        "$env:APPDATA\Code\logs\renderer.log",
        "$env:APPDATA\Code\logs\exthost.log"
    )

    $hookPatterns = @(
        "Initialize Cursor Rules",
        "Register-CursorHooks",
        ".cursor-rules.d"
    )

    $results = @{
        HooksFound = $false
        LogEntries = @()
    }

    foreach ($logPath in $logPaths) {
        if (Test-Path $logPath) {
            try {
                $logContent = Get-Content $logPath -Tail 100
                foreach ($pattern in $hookPatterns) {
                    $matches = $logContent | Select-String -Pattern $pattern
                    if ($matches) {
                        $results.HooksFound = $true
                        $results.LogEntries += $matches
                    }
                }
            }
            catch {
                Write-Log "Erreur lors de la lecture du log $logPath : $_" -Level Error
            }
        }
    }

    if ($results.HooksFound) {
        Write-Log "Hooks Cursor trouvés dans les logs VS Code" -Level Info
        foreach ($entry in $results.LogEntries) {
            Write-Log "  - $entry" -Level Info
        }
    }
    else {
        Write-Log "Aucun hook Cursor trouvé dans les logs récents" -Level Warning
    }

    return $results
}

# Test d'application des règles en temps réel
function Test-CursorRulesRealTime {
    Write-Log "Test d'application des règles en temps réel..." -Level Info
    
    $testFile = "test_cursor_rules.ps1"
    $testContent = @'
# Test d'application des règles Cursor
$testVariable = "Test"  # Ligne non formatée correctement
function Test-Function {
param($param1)  # Paramètre non formaté correctement
Write-Host "Test"  # Instruction non formatée correctement
}
'@

    try {
        # Création du fichier de test
        Set-Content -Path $testFile -Value $testContent
        Write-Log "Fichier de test créé : $testFile" -Level Info

        # Attente de l'application des règles
        Start-Sleep -Seconds 2

        # Vérification des modifications automatiques
        $modifiedContent = Get-Content $testFile -Raw
        $changes = @{
            "Formatage" = $modifiedContent -ne $testContent
            "Original"  = $testContent
            "Modified"  = $modifiedContent
        }

        if ($changes.Formatage) {
            Write-Log "Les règles de formatage ont été appliquées" -Level Info
        }
        else {
            Write-Log "Aucune modification automatique détectée" -Level Warning
        }

        return $changes
    }
    catch {
        Write-Log "Erreur lors du test en temps réel : $_" -Level Error
    }
    finally {
        if (Test-Path $testFile) {
            Remove-Item $testFile -Force
        }
    }
}

# Amélioration de la détection des règles
function Test-CursorRulesDetection {
    Write-Log "Vérification de la détection des règles..." -Level Info
    
    $workspaceRoot = (Get-Location).Path
    $results = @{
        ConfigurationFiles = @{}
        VSCodeSettings     = @{}
        RulesEnabled       = $false
    }

    # Vérification des fichiers de configuration Cursor
    $configFiles = @(
        ".cursor-rules.d/settings.json",
        ".cursor-rules.d/extensions.json",
        ".cursor-rules.d/launch.json",
        ".cursor-rules.d/tasks.json"
    )

    foreach ($file in $configFiles) {
        $fullPath = Join-Path $workspaceRoot $file
        if (Test-Path $fullPath) {
            try {
                $content = Get-Content $fullPath -Raw | ConvertFrom-Json
                $results.ConfigurationFiles[$file] = @{
                    "Exists"  = $true
                    "Content" = $content
                }
            }
            catch {
                Write-Log "Erreur lors de la lecture de $file : $_" -Level Error
            }
        }
    }

    # Vérification des paramètres VS Code
    $vscodeSettingsPath = Join-Path $workspaceRoot ".vscode/settings.json"
    if (Test-Path $vscodeSettingsPath) {
        try {
            $vsCodeSettings = Get-Content $vscodeSettingsPath -Raw | ConvertFrom-Json
            $results.VSCodeSettings = @{
                "cursor.rules.enabled"        = $vsCodeSettings."cursor.rules.enabled"
                "cursor.rules.validateOnSave" = $vsCodeSettings."cursor.rules.validateOnSave"
                "cursor.rules.validateOnType" = $vsCodeSettings."cursor.rules.validateOnType"
            }
            $results.RulesEnabled = $vsCodeSettings."cursor.rules.enabled" -eq $true
        }
        catch {
            Write-Log "Erreur lors de la lecture des paramètres VS Code : $_" -Level Error
        }
    }

    # Affichage des résultats
    Write-Log "État de la détection des règles :" -Level Info
    Write-Log "- Règles Cursor activées : $($results.RulesEnabled)" -Level Info
    Write-Log "- Fichiers de configuration trouvés : $($results.ConfigurationFiles.Keys.Count)" -Level Info
    
    return $results
}

# Programme principal
try {
    Write-Log "Démarrage de l'initialisation de l'intégration"

    if (-not $Force) {
        $confirmation = Read-Host "Voulez-vous procéder à l'initialisation ? (O/N)"
        if ($confirmation -ne "O") {
            Write-Log "Initialisation annulée par l'utilisateur" -Level Warning
            exit
        }
    }

    Backup-Configurations
    Initialize-Environment
    Install-RequiredExtensions

    if (Test-Integration) {
        Write-Log "Initialisation terminée avec succès"
        # Vérification de l'alignement après l'initialisation
        $alignmentResults = Test-VSCodeCursorAlignment
        # Vérification de l'application des règles Cursor
        $rulesResults = Test-CursorRulesApplication
        
        # Tests supplémentaires
        $hooksResults = Test-CursorHooksExecution
        $realTimeResults = Test-CursorRulesRealTime
        $detectionResults = Test-CursorRulesDetection
    }
    else {
        Write-Log "L'initialisation a réussi mais certains tests ont échoué" -Level Warning
    }
}
catch {
    Write-Log "Erreur lors de l'initialisation : $_" -Level Error
    throw $_
}
finally {
    Write-Log "Fin du processus d'initialisation"
}
