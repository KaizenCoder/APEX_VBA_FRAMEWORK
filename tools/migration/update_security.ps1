# Script de migration des composants de sécurité APEX Framework
# Version: 1.0
# Date: 2024-04-11

# Configuration
$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"

# Chemins
$rootPath = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$backupPath = Join-Path $rootPath "backup"
$componentsPath = Join-Path $rootPath "apex-metier"
$logsPath = Join-Path $rootPath "logs"

# Création des dossiers nécessaires
function EnsurePaths {
    @($backupPath, $logsPath) | ForEach-Object {
        if (-not (Test-Path $_)) {
            New-Item -ItemType Directory -Path $_ | Out-Null
            Write-Verbose "Créé le dossier: $_"
        }
    }
}

# Journalisation
function Write-Log {
    param([string]$Message)
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    
    Write-Verbose $logMessage
    Add-Content -Path (Join-Path $logsPath "security_migration.log") -Value $logMessage
}

# Sauvegarde des fichiers existants
function Backup-Components {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupDir = Join-Path $backupPath "security_$timestamp"
    
    Write-Log "Début de la sauvegarde..."
    
    # Créer le dossier de backup
    New-Item -ItemType Directory -Path $backupDir | Out-Null
    
    # Copier les fichiers
    $filesToBackup = @(
        "security\clsSecurityManager.cls",
        "security\clsAES256.cls",
        "database\connection\clsConnectionPool.cls",
        "monitoring\clsMetricsCollector.cls"
    )
    
    foreach ($file in $filesToBackup) {
        $sourcePath = Join-Path $componentsPath $file
        if (Test-Path $sourcePath) {
            $targetDir = Join-Path $backupDir (Split-Path -Parent $file)
            
            # Créer la structure de dossiers
            if (-not (Test-Path $targetDir)) {
                New-Item -ItemType Directory -Path $targetDir | Out-Null
            }
            
            # Copier le fichier
            Copy-Item -Path $sourcePath -Destination $targetDir
            Write-Log "Sauvegardé: $file"
        }
    }
    
    Write-Log "Sauvegarde terminée: $backupDir"
    return $backupDir
}

# Validation des composants
function Test-Components {
    Write-Log "Validation des composants..."
    
    $testResults = @()
    
    # Vérifier les fichiers requis
    $requiredFiles = @(
        "security\clsSecurityManager.cls",
        "security\clsAES256.cls",
        "database\connection\clsConnectionPool.cls",
        "monitoring\clsMetricsCollector.cls"
    )
    
    foreach ($file in $requiredFiles) {
        $filePath = Join-Path $componentsPath $file
        $result = @{
            File         = $file
            Exists       = Test-Path $filePath
            Size         = if (Test-Path $filePath) { (Get-Item $filePath).Length } else { 0 }
            LastModified = if (Test-Path $filePath) { (Get-Item $filePath).LastWriteTime } else { $null }
        }
        $testResults += $result
        
        if (-not $result.Exists) {
            Write-Log "ERREUR: Fichier manquant: $file"
            return $false
        }
    }
    
    # Vérifier les dépendances
    $dependencies = @(
        "bcrypt.dll",
        "Rubberduck.dll"
    )
    
    foreach ($dll in $dependencies) {
        if (-not (Test-Path (Join-Path $env:SystemRoot "System32\$dll"))) {
            Write-Log "ERREUR: Dépendance manquante: $dll"
            return $false
        }
    }
    
    Write-Log "Validation réussie"
    return $true
}

# Migration des composants
function Update-Components {
    param([string]$BackupPath)
    
    Write-Log "Début de la mise à jour..."
    
    try {
        # Mettre à jour les composants
        $componentsToUpdate = @(
            @{
                Source  = "security\clsSecurityManager.cls"
                Target  = "security"
                Version = "2.0"
            },
            @{
                Source  = "security\clsAES256.cls"
                Target  = "security"
                Version = "1.0"
            },
            @{
                Source  = "database\connection\clsConnectionPool.cls"
                Target  = "database\connection"
                Version = "2.0"
            },
            @{
                Source  = "monitoring\clsMetricsCollector.cls"
                Target  = "monitoring"
                Version = "1.0"
            }
        )
        
        foreach ($component in $componentsToUpdate) {
            $targetPath = Join-Path $componentsPath $component.Target
            
            # Créer le dossier cible si nécessaire
            if (-not (Test-Path $targetPath)) {
                New-Item -ItemType Directory -Path $targetPath | Out-Null
            }
            
            # Copier et mettre à jour le composant
            $sourcePath = Join-Path $BackupPath $component.Source
            $targetFile = Join-Path $targetPath (Split-Path -Leaf $component.Source)
            
            if (Test-Path $sourcePath) {
                Copy-Item -Path $sourcePath -Destination $targetFile -Force
                Write-Log "Mis à jour: $($component.Source) -> v$($component.Version)"
            }
            else {
                Write-Log "ATTENTION: Source non trouvée: $($component.Source)"
            }
        }
        
        Write-Log "Mise à jour terminée avec succès"
        return $true
    }
    catch {
        Write-Log "ERREUR pendant la mise à jour: $($_.Exception.Message)"
        return $false
    }
}

# Fonction principale
function Main {
    Write-Log "Début du processus de migration"
    
    try {
        # Créer les dossiers nécessaires
        EnsurePaths
        
        # Sauvegarder les composants existants
        $backupDir = Backup-Components
        if (-not $backupDir) {
            throw "Échec de la sauvegarde"
        }
        
        # Valider les composants
        if (-not (Test-Components)) {
            throw "Échec de la validation"
        }
        
        # Mettre à jour les composants
        if (-not (Update-Components -BackupPath $backupDir)) {
            throw "Échec de la mise à jour"
        }
        
        Write-Log "Migration terminée avec succès"
    }
    catch {
        Write-Log "ERREUR CRITIQUE: $($_.Exception.Message)"
        Write-Log "Restauration de la sauvegarde..."
        
        if ($backupDir -and (Test-Path $backupDir)) {
            Update-Components -BackupPath $backupDir | Out-Null
            Write-Log "Restauration terminée"
        }
        
        throw
    }
}

# Exécution
try {
    Main
}
catch {
    Write-Error $_.Exception.Message
    exit 1
} 