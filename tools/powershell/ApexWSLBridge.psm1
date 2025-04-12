# ApexWSLBridge.psm1
# Module PowerShell pour faciliter l'interaction avec WSL dans le projet APEX VBA Framework
# Ce module résout les problèmes d'interaction entre PowerShell et WSL

# Force l'encodage UTF-8 pour l'affichage
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

# Variables globales du module
$script:WorkspacePath = "D:\Dev\Apex_VBA_FRAMEWORK"
$script:LogPath = Join-Path $script:WorkspacePath "logs\powershell_wsl.log"
$script:DefaultDistro = "Ubuntu-22.04"

#region Configuration

# Fonction pour initialiser l'environnement de logs
function Initialize-ApexWSLBridge {
    param (
        [string]$WorkspacePath = $script:WorkspacePath,
        [string]$LogPath = $script:LogPath,
        [string]$Distribution = $script:DefaultDistro
    )
    
    # Mise à jour des chemins globaux
    $script:WorkspacePath = $WorkspacePath
    $script:LogPath = $LogPath
    $script:DefaultDistro = $Distribution
    
    # Création du répertoire de logs s'il n'existe pas
    $logsDir = Split-Path -Parent $script:LogPath
    if (-not (Test-Path $logsDir)) {
        New-Item -Path $logsDir -ItemType Directory -Force | Out-Null
        Write-Host "Répertoire de logs créé: $logsDir" -ForegroundColor Green
    }
    
    # Vérification de l'environnement WSL
    $wslAvailable = Test-WSLEnvironment -Silent
    if (-not $wslAvailable) {
        Write-Host "ATTENTION: L'environnement WSL n'est pas correctement configuré!" -ForegroundColor Red
    } else {
        Write-Host "Environnement WSL détecté avec la distribution $script:DefaultDistro" -ForegroundColor Green
    }
    
    # Initialisation du fichier de log
    Write-ApexLog "Module ApexWSLBridge initialisé pour $script:WorkspacePath" "INIT"
    Write-ApexLog "Distribution WSL: $script:DefaultDistro" "INIT"
    
    return $wslAvailable
}

#endregion

#region Logging

function Write-ApexLog {
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Message,
        
        [Parameter(Mandatory=$false, Position=1)]
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG", "PERF", "INIT")]
        [string]$Level = "INFO",
        
        [Parameter(Mandatory=$false)]
        [switch]$Console
    )
    
    # Création du répertoire de logs si nécessaire
    $logsDir = Split-Path -Parent $script:LogPath
    if (-not (Test-Path $logsDir)) {
        New-Item -Path $logsDir -ItemType Directory -Force | Out-Null
    }
    
    # Format du timestamp
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] [$Level] $Message"
    
    # Ecriture dans le fichier de log
    try {
        Add-Content -Path $script:LogPath -Value $logLine -ErrorAction Stop
    } catch {
        Write-Warning "Impossible d'écrire dans le fichier de log $script:LogPath : $_"
    }
    
    # Affichage console si demandé
    if ($Console) {
        $color = switch ($Level) {
            "ERROR" { "Red" }
            "WARN"  { "Yellow" }
            "INFO"  { "White" }
            "DEBUG" { "Gray" }
            "PERF"  { "Cyan" }
            "INIT"  { "Green" }
            default { "White" }
        }
        Write-Host $logLine -ForegroundColor $color
    }
}

#endregion

#region WSL Commands

function Invoke-WSLCommand {
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Command,
        
        [Parameter(Mandatory=$false)]
        [string]$Distribution = $script:DefaultDistro,
        
        [Parameter(Mandatory=$false)]
        [switch]$LogOutput,
        
        [Parameter(Mandatory=$false)]
        [switch]$UseTempFile
    )
    
    Write-ApexLog "Exécution commande WSL: $Command" "DEBUG"
    
    if ($UseTempFile) {
        # Utilisation d'un fichier temporaire pour contourner les problèmes d'interaction
        $tempOutFile = [System.IO.Path]::GetTempFileName()
        $wslCommand = "$Command > `$(wslpath -u '$tempOutFile') 2>&1"
        
        try {
            & wsl --distribution $Distribution -- bash -c $wslCommand 2>&1 | Out-Null
            $output = Get-Content -Path $tempOutFile -Raw
            Remove-Item -Path $tempOutFile -Force
        } catch {
            Write-ApexLog "Erreur lors de l'exécution de la commande WSL: $_" "ERROR" -Console
            if (Test-Path $tempOutFile) {
                Remove-Item -Path $tempOutFile -Force
            }
            return $null
        }
    } else {
        try {
            $output = & wsl --distribution $Distribution -- bash -c $Command 2>&1
        } catch {
            Write-ApexLog "Erreur lors de l'exécution de la commande WSL: $_" "ERROR" -Console
            return $null
        }
    }
    
    if ($LogOutput) {
        Write-ApexLog "Sortie de la commande: $output" "DEBUG"
    }
    
    return $output
}

function Invoke-WSLCommandWithRetry {
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Command,
        
        [Parameter(Mandatory=$false)]
        [string]$Distribution = $script:DefaultDistro,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxRetries = 3,
        
        [Parameter(Mandatory=$false)]
        [int]$DelaySeconds = 2,
        
        [Parameter(Mandatory=$false)]
        [switch]$UseTempFile
    )
    
    $retry = 0
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    
    do {
        $result = Invoke-WSLCommand -Command $Command -Distribution $Distribution -UseTempFile:$UseTempFile
        
        if ($null -ne $result) {
            $sw.Stop()
            Write-ApexLog "Commande exécutée avec succès après $($retry + 1) tentative(s) en $($sw.ElapsedMilliseconds) ms" "PERF"
            return $result
        }
        
        Write-ApexLog "Nouvelle tentative d'exécution de la commande (essai $($retry + 1)/$MaxRetries)..." "WARN" -Console
        $waitTime = $DelaySeconds * [math]::Pow(2, $retry)
        Start-Sleep -Seconds $waitTime
        $retry++
        
    } while ($retry -lt $MaxRetries)
    
    $sw.Stop()
    Write-ApexLog "Échec de la commande après $MaxRetries tentatives en $($sw.ElapsedMilliseconds) ms" "ERROR" -Console
    return $null
}

function Invoke-WSLCommandWithInput {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Command,
        
        [Parameter(Mandatory=$true)]
        [string]$Input,
        
        [Parameter(Mandatory=$false)]
        [string]$Distribution = $script:DefaultDistro
    )
    
    # Création d'un fichier temporaire pour l'entrée
    $tempInFile = [System.IO.Path]::GetTempFileName()
    $tempOutFile = [System.IO.Path]::GetTempFileName()
    
    try {
        # Écriture de l'entrée dans le fichier temporaire
        $Input | Out-File -FilePath $tempInFile -Encoding utf8
        
        # Conversion des chemins Windows en chemins WSL
        $wslInPath = Invoke-WSLCommand -Command "wslpath -u '$tempInFile'" -Distribution $Distribution
        $wslOutPath = Invoke-WSLCommand -Command "wslpath -u '$tempOutFile'" -Distribution $Distribution
        
        # Exécution de la commande avec redirection
        $wslCmd = "cat $wslInPath | $Command > $wslOutPath 2>&1"
        Invoke-WSLCommand -Command $wslCmd -Distribution $Distribution | Out-Null
        
        # Lecture du résultat
        $output = Get-Content -Path $tempOutFile -Raw
        
        return $output
    }
    catch {
        Write-ApexLog "Erreur lors de l'exécution de la commande avec entrée: $_" "ERROR" -Console
        return $null
    }
    finally {
        # Nettoyage
        if (Test-Path $tempInFile) { Remove-Item -Path $tempInFile -Force }
        if (Test-Path $tempOutFile) { Remove-Item -Path $tempOutFile -Force }
    }
}

function Measure-WSLCommand {
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Command,
        
        [Parameter(Mandatory=$false)]
        [string]$Distribution = $script:DefaultDistro,
        
        [Parameter(Mandatory=$false)]
        [switch]$UseTempFile
    )
    
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $result = Invoke-WSLCommand -Command $Command -Distribution $Distribution -UseTempFile:$UseTempFile
    $sw.Stop()
    
    Write-ApexLog "Commande '$Command' exécutée en $($sw.ElapsedMilliseconds) ms" "PERF"
    return @{
        Result = $result
        ElapsedMs = $sw.ElapsedMilliseconds
    }
}

#endregion

#region Environment Tests

function Test-WSLEnvironment {
    param(
        [Parameter(Mandatory=$false)]
        [string]$Distribution = $script:DefaultDistro,
        
        [Parameter(Mandatory=$false)]
        [switch]$Silent
    )
    
    if (-not $Silent) {
        Write-ApexLog "Test de l'environnement WSL..." "INFO" -Console
    }
    
    try {
        # Vérification que WSL est installé
        $wslInfo = Get-Command wsl -ErrorAction Stop
        
        # Vérification directe que la distribution existe/fonctionne
        try {
            $test = & wsl --distribution $Distribution -- bash -c "echo 'WSL_TEST_OK'" 2>$null
            if ($test -eq "WSL_TEST_OK") {
                if (-not $Silent) {
                    Write-ApexLog "Environnement WSL fonctionnel" "INFO" -Console
                }
                return $true
            }
        } catch {
            # Gérer les erreurs silencieusement ici
        }
        
        # Si le test direct a échoué, essayer de vérifier la liste
        $distroList = wsl --list
        
        # Pour le débogage
        if (-not $Silent) {
            Write-ApexLog "Distributions WSL trouvées: $distroList" "DEBUG" -Console
        }
        
        # Vérifier si la distribution apparaît dans la liste (peu importe le format)
        if (-not ($distroList -match [regex]::Escape($Distribution))) {
            if (-not $Silent) {
                Write-ApexLog "La distribution $Distribution n'est pas installée!" "ERROR" -Console
            }
            return $false
        }
        
        if (-not $Silent) {
            Write-ApexLog "Impossible d'exécuter des commandes dans WSL, mais la distribution existe" "WARN" -Console
        }
        
        return $false
    }
    catch {
        if (-not $Silent) {
            Write-ApexLog "Erreur lors du test WSL: $_" "ERROR" -Console
        }
        return $false
    }
}

function Get-WSLMountStatus {
    param(
        [Parameter(Mandatory=$false)]
        [string]$Drive = "d",
        
        [Parameter(Mandatory=$false)]
        [string]$Distribution = $script:DefaultDistro
    )
    
    try {
        $mounts = Invoke-WSLCommand -Command "mount | grep '/mnt/$Drive'" -Distribution $Distribution
        
        if ([string]::IsNullOrEmpty($mounts)) {
            Write-ApexLog "Aucun montage trouvé pour /mnt/$Drive" "WARN" -Console
            return $null
        }
        
        # Analyse des options de montage
        $options = [regex]::Match($mounts, ".*\((.*)\)").Groups[1].Value
        
        return @{
            MountPoint = "/mnt/$Drive"
            RawInfo = $mounts
            Options = $options
            HasMetadata = $options -match "metadata"
        }
    }
    catch {
        Write-ApexLog "Erreur lors de la vérification des montages WSL: $_" "ERROR" -Console
        return $null
    }
}

#endregion

#region Advanced Functions

function Run-SessionWithWSL {
    param(
        [Parameter(Mandatory=$true)]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory=$false)]
        [string]$SessionName = "WSLSession_$(Get-Date -Format 'yyyyMMdd_HHmmss')",
        
        [Parameter(Mandatory=$false)]
        [string]$Distribution = $script:DefaultDistro
    )
    
    Write-ApexLog "Démarrage de la session WSL: $SessionName" "INFO" -Console
    $sessionLogFile = Join-Path (Split-Path -Parent $script:LogPath) "$SessionName.log"
    
    try {
        # Vérification de l'environnement avant exécution
        if (-not (Test-WSLEnvironment -Distribution $Distribution)) {
            Write-ApexLog "Impossible de démarrer la session WSL: environnement non disponible" "ERROR" -Console
            return $false
        }
        
        # Enregistrement du début de session
        "--- Session WSL: $SessionName - Début: $(Get-Date) ---" | Out-File -FilePath $sessionLogFile
        
        # Exécution du bloc de script
        & $ScriptBlock
        
        # Enregistrement de la fin de session
        "--- Session WSL: $SessionName - Fin: $(Get-Date) - Succès ---" | Out-File -FilePath $sessionLogFile -Append
        Write-ApexLog "Session WSL $SessionName terminée avec succès" "INFO" -Console
        
        return $true
    }
    catch {
        # Enregistrement de l'erreur
        "--- Session WSL: $SessionName - Fin: $(Get-Date) - ERREUR: $_ ---" | Out-File -FilePath $sessionLogFile -Append
        Write-ApexLog "Erreur dans la session WSL $SessionName : $_" "ERROR" -Console
        
        return $false
    }
}

function Start-WSLBatchFromFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$false)]
        [string]$Distribution = $script:DefaultDistro,
        
        [Parameter(Mandatory=$false)]
        [switch]$ContinueOnError
    )
    
    if (-not (Test-Path $FilePath)) {
        Write-ApexLog "Fichier de commandes introuvable: $FilePath" "ERROR" -Console
        return $false
    }
    
    Write-ApexLog "Exécution du lot de commandes depuis $FilePath" "INFO" -Console
    
    # Lecture du fichier de commandes
    $commands = Get-Content -Path $FilePath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.StartsWith('#') }
    
    $success = $true
    $results = @()
    
    foreach ($cmd in $commands) {
        Write-ApexLog "Exécution: $cmd" "INFO" -Console
        
        $result = Invoke-WSLCommand -Command $cmd -Distribution $Distribution -UseTempFile
        
        if ($null -eq $result -and -not $ContinueOnError) {
            Write-ApexLog "Commande échouée, arrêt du lot: $cmd" "ERROR" -Console
            $success = $false
            break
        }
        
        $results += $result
    }
    
    return @{
        Success = $success
        Results = $results
    }
}

function Start-InteractiveWSLSession {
    param(
        [Parameter(Mandatory=$false)]
        [string]$InitCommand,
        
        [Parameter(Mandatory=$false)]
        [string]$WorkingDirectory = "/mnt/d/Dev/Apex_VBA_FRAMEWORK",
        
        [Parameter(Mandatory=$false)]
        [string]$Distribution = $script:DefaultDistro
    )
    
    Write-ApexLog "Démarrage d'une session WSL interactive" "INFO" -Console
    
    $command = "cd $WorkingDirectory"
    
    if (-not [string]::IsNullOrEmpty($InitCommand)) {
        $command += " && $InitCommand"
    }
    
    & wsl --distribution $Distribution -- bash -c "$command; exec bash"
}

function Invoke-PowerShellWithWSL {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ScriptPath,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$Parameters = @{},
        
        [Parameter(Mandatory=$false)]
        [string]$Distribution = $script:DefaultDistro
    )
    
    if (-not (Test-Path $ScriptPath)) {
        Write-ApexLog "Script PowerShell introuvable: $ScriptPath" "ERROR" -Console
        return $null
    }
    
    Write-ApexLog "Exécution du script PowerShell avec contexte WSL: $ScriptPath" "INFO" -Console
    
    # Préparation du script avec importation automatique du module
    $modulePath = $PSCommandPath  # Chemin du module actuel
    $scriptContent = @"
# Importation automatique du module ApexWSLBridge
Import-Module "$modulePath" -Force

# Script original avec paramètres
& "$ScriptPath" $(
    ($Parameters.GetEnumerator() | ForEach-Object {
        "-$($_.Key) $($_.Value)"
    }) -join " "
)
"@
    
    $tempScriptPath = [System.IO.Path]::GetTempFileName() + ".ps1"
    $scriptContent | Out-File -FilePath $tempScriptPath -Encoding utf8
    
    try {
        # Exécution du script temporaire
        $result = & $tempScriptPath
        return $result
    }
    catch {
        Write-ApexLog "Erreur lors de l'exécution du script avec contexte WSL: $_" "ERROR" -Console
        return $null
    }
    finally {
        # Nettoyage
        if (Test-Path $tempScriptPath) {
            Remove-Item -Path $tempScriptPath -Force
        }
    }
}

#endregion

# Initialisation automatique du module
Initialize-ApexWSLBridge | Out-Null

# Export des fonctions
Export-ModuleMember -Function Initialize-ApexWSLBridge
Export-ModuleMember -Function Write-ApexLog
Export-ModuleMember -Function Invoke-WSLCommand
Export-ModuleMember -Function Invoke-WSLCommandWithRetry
Export-ModuleMember -Function Invoke-WSLCommandWithInput
Export-ModuleMember -Function Measure-WSLCommand
Export-ModuleMember -Function Test-WSLEnvironment
Export-ModuleMember -Function Get-WSLMountStatus
Export-ModuleMember -Function Run-SessionWithWSL
Export-ModuleMember -Function Start-WSLBatchFromFile
Export-ModuleMember -Function Start-InteractiveWSLSession
Export-ModuleMember -Function Invoke-PowerShellWithWSL 