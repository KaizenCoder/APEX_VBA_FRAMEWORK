<#
.SYNOPSIS
    Module de monitoring de performance pour APEX Framework
.DESCRIPTION
    Fournit des fonctions pour enregistrer et suivre les performances du framework APEX dans VSCode
.NOTES
    Auteur: APEX Framework Team
    Date: 2025-04-14
    Version: 1.0
#>

# Configuration du module
$script:LogsPath = Join-Path -Path "D:\Dev\Apex_VBA_FRAMEWORK" -ChildPath "logs\performance"
$script:DefaultLogFile = Join-Path -Path $script:LogsPath -ChildPath "vscode_performance.log"
$script:HistoryFile = Join-Path -Path $script:LogsPath -ChildPath "performance_history.json"
$script:ConfigFile = Join-Path -Path "D:\Dev\Apex_VBA_FRAMEWORK" -ChildPath "config\monitoring\performance_config.json"

# Créer les répertoires nécessaires s'ils n'existent pas
function Initialize-LogEnvironment {
    if (-not (Test-Path -Path $script:LogsPath)) {
        New-Item -Path $script:LogsPath -ItemType Directory -Force | Out-Null
        Write-Verbose "Répertoire de logs créé: $($script:LogsPath)"
    }
    
    # Vérifier si le fichier de configuration existe, sinon en créer un par défaut
    if (-not (Test-Path -Path $script:ConfigFile)) {
        $configDir = Split-Path -Parent $script:ConfigFile
        if (-not (Test-Path -Path $configDir)) {
            New-Item -Path $configDir -ItemType Directory -Force | Out-Null
        }
        
        $defaultConfig = @{
            LogLevel          = "INFO"
            MaxLogSizeMB      = 10
            EnableRotation    = $true
            KeepLogDays       = 7
            MonitoringEnabled = $true
            RefreshRateMs     = 5000
        } | ConvertTo-Json -Depth 3
        
        Set-Content -Path $script:ConfigFile -Value $defaultConfig -Encoding UTF8
        Write-Verbose "Fichier de configuration créé: $($script:ConfigFile)"
    }
}

<#
.SYNOPSIS
    Écrit une entrée de log de performance
.DESCRIPTION
    Enregistre une entrée de log avec niveau, message et métriques optionnelles
.PARAMETER Level
    Niveau de log (INFO, WARNING, ERROR, DEBUG)
.PARAMETER Message
    Message à enregistrer
.PARAMETER Module
    Nom du module ou composant (défaut: VSCode)
.PARAMETER Metrics
    Table de hachage contenant des métriques à enregistrer au format JSON
.PARAMETER LogFilePath
    Chemin du fichier de log (optionnel, utilise le chemin par défaut si non spécifié)
.EXAMPLE
    Write-PerformanceLog -Level "INFO" -Message "Test de performance terminé" -Module "TestRunner"
.EXAMPLE
    Write-PerformanceLog -Level "WARNING" -Message "Performance dégradée" -Metrics @{CPU=45; Memory=128MB; ResponseTime=230ms}
#>
function Write-PerformanceLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("INFO", "WARNING", "ERROR", "DEBUG")]
        [string]$Level,
        
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [string]$Module = "VSCode",
        
        [Parameter(Mandatory = $false)]
        [hashtable]$Metrics,
        
        [Parameter(Mandatory = $false)]
        [string]$LogFilePath = $script:DefaultLogFile
    )
    
    # S'assurer que l'environnement est initialisé
    Initialize-LogEnvironment
    
    # Créer le répertoire des logs s'il n'existe pas
    $logDir = Split-Path -Parent $LogFilePath
    if (-not (Test-Path -Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
    
    # Format de l'entrée de log
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Ajouter les métriques au format JSON si présentes
    $metricsJson = ""
    if ($Metrics -and $Metrics.Count -gt 0) {
        $metricsJson = " " + ($Metrics | ConvertTo-Json -Compress)
    }
    
    $logEntry = "[$timestamp][$Level][$Module] $Message$metricsJson"
    
    # Écrire dans le fichier de log
    Add-Content -Path $LogFilePath -Value $logEntry -Encoding UTF8
    
    # Afficher également dans la console selon le niveau
    if ($Level -eq "ERROR") {
        Write-Host $logEntry -ForegroundColor Red
    }
    elseif ($Level -eq "WARNING") {
        Write-Host $logEntry -ForegroundColor Yellow
    }
    elseif ($Level -eq "INFO") {
        Write-Verbose $logEntry
    }
    elseif ($Level -eq "DEBUG") {
        Write-Debug $logEntry
    }
    
    # Ajouter à l'historique pour l'analyse des tendances
    if ($Metrics -and $Level -ne "DEBUG") {
        Add-PerformanceHistory -Level $Level -Module $Module -Metrics $Metrics
    }
}

<#
.SYNOPSIS
    Ajoute une entrée dans l'historique des performances
.DESCRIPTION
    Enregistre des métriques de performance dans un fichier JSON pour analyse ultérieure
.PARAMETER Level
    Niveau de log (INFO, WARNING, ERROR)
.PARAMETER Module
    Nom du module ou composant
.PARAMETER Metrics
    Table de hachage contenant des métriques à enregistrer
#>
function Add-PerformanceHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Level,
        
        [Parameter(Mandatory = $true)]
        [string]$Module,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$Metrics
    )
    
    # Charger l'historique existant ou créer un nouveau
    $history = @()
    if (Test-Path -Path $script:HistoryFile) {
        try {
            $history = Get-Content -Path $script:HistoryFile -Raw | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            Write-Warning "Impossible de charger l'historique. Un nouveau fichier sera créé."
        }
    }
    
    # Limiter la taille de l'historique (conserver les 1000 dernières entrées)
    if ($history.Count -gt 1000) {
        $history = $history | Select-Object -Skip ($history.Count - 1000)
    }
    
    # Ajouter la nouvelle entrée
    $entry = [PSCustomObject]@{
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Level     = $Level
        Module    = $Module
        Metrics   = $Metrics
    }
    
    $history += $entry
    
    # Sauvegarder l'historique
    $history | ConvertTo-Json | Set-Content -Path $script:HistoryFile -Encoding UTF8
}

<#
.SYNOPSIS
    Effectue une rotation des fichiers de logs
.DESCRIPTION
    Vérifie la taille des fichiers de logs et crée une archive si nécessaire
#>
function Invoke-LogRotation {
    [CmdletBinding()]
    param()
    
    # Charger la configuration
    $config = Get-Content -Path $script:ConfigFile -Raw | ConvertFrom-Json
    
    if (-not $config.EnableRotation) {
        return
    }
    
    $logFile = $script:DefaultLogFile
    if (-not (Test-Path -Path $logFile)) {
        return
    }
    
    $logFileInfo = Get-Item -Path $logFile
    $maxSizeBytes = $config.MaxLogSizeMB * 1MB
    
    if ($logFileInfo.Length -gt $maxSizeBytes) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $archivePath = Join-Path -Path $script:LogsPath -ChildPath "archive"
        
        if (-not (Test-Path -Path $archivePath)) {
            New-Item -Path $archivePath -ItemType Directory -Force | Out-Null
        }
        
        $archiveFile = Join-Path -Path $archivePath -ChildPath "vscode_performance_$timestamp.log"
        
        Move-Item -Path $logFile -Destination $archiveFile
        Write-Verbose "Rotation du log effectuée: $archiveFile"
        
        # Supprimer les archives trop anciennes
        $oldLogs = Get-ChildItem -Path $archivePath -Filter "*.log" | 
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$config.KeepLogDays) }
        
        foreach ($oldLog in $oldLogs) {
            Remove-Item -Path $oldLog.FullName -Force
            Write-Verbose "Ancien log supprimé: $($oldLog.FullName)"
        }
    }
}

<#
.SYNOPSIS
    Récupère les statistiques de performance actuelles
.DESCRIPTION
    Collecte les métriques système actuelles pour le monitoring
.OUTPUTS
    Une hashtable contenant les métriques système
#>
function Get-SystemPerformanceMetrics {
    [CmdletBinding()]
    param()
    
    $metrics = @{}
    
    # Obtenir l'utilisation CPU
    $cpuLoad = (Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples.CookedValue
    $metrics.Add("CPU", [math]::Round($cpuLoad, 2))
    
    # Obtenir l'utilisation mémoire
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $memoryUsed = $os.TotalVisibleMemorySize - $os.FreePhysicalMemory
    $metrics.Add("MemoryMB", [math]::Round($memoryUsed / 1024, 2))
    $metrics.Add("MemoryPct", [math]::Round($memoryUsed * 100 / $os.TotalVisibleMemorySize, 2))
    
    # Obtenir les processus VSCode
    $vsCodeProcesses = Get-Process | Where-Object { $_.Name -like "*code*" }
    if ($vsCodeProcesses) {
        $vscodeMemory = ($vsCodeProcesses | Measure-Object -Property WorkingSet -Sum).Sum / 1MB
        $metrics.Add("VSCodeMemoryMB", [math]::Round($vscodeMemory, 2))
    }
    
    return $metrics
}

# Exporter les fonctions publiques du module
Export-ModuleMember -Function Write-PerformanceLog, Invoke-LogRotation, Get-SystemPerformanceMetrics