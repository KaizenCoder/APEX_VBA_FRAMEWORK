# ApexWSLBridge

Module PowerShell pour faciliter l'interaction sécurisée entre PowerShell et WSL dans le cadre du projet APEX VBA Framework.

## Introduction

Ce module résout les problèmes de communication entre PowerShell et WSL, en particulier lorsque les commandes nécessitent une interaction ou lorsqu'elles sont bloquées en raison de limitations du terminal.

## Fonctionnalités principales

- ✅ **Exécution robuste de commandes WSL** - Gère les erreurs et les interruptions
- ✅ **Système de logs intégré** - Traçabilité complète des opérations
- ✅ **Retry automatique** - Réessaie les commandes qui échouent avec backoff exponentiel
- ✅ **Entrée/sortie fiable** - Contourne les problèmes de flux avec des fichiers temporaires
- ✅ **Sessions interactives** - Permet de lancer des sessions WSL configurées
- ✅ **Exécution de batch** - Lance plusieurs commandes à partir d'un fichier
- ✅ **Mesure de performance** - Évalue la durée d'exécution des commandes

## Installation

1. Copiez le module dans `D:\Dev\Apex_VBA_FRAMEWORK\tools\powershell\`
2. Importez le module dans vos scripts PowerShell:

```powershell
Import-Module "D:\Dev\Apex_VBA_FRAMEWORK\tools\powershell\ApexWSLBridge.psm1" -Force
```

## Utilisation de base

```powershell
# Importer le module
Import-Module ".\tools\powershell\ApexWSLBridge.psm1" -Force

# Exécuter une commande WSL simple
$result = Invoke-WSLCommand -Command "ls -la /mnt/d/Dev/Apex_VBA_FRAMEWORK"

# Exécuter une commande avec retry en cas d'échec
$result = Invoke-WSLCommandWithRetry -Command "git status" -MaxRetries 3

# Mesurer la performance d'une commande
$perf = Measure-WSLCommand -Command "find /mnt/d/Dev/Apex_VBA_FRAMEWORK -name '*.ps1'"
Write-Host "Temps d'exécution: $($perf.ElapsedMs) ms"
```

## Utilisation avancée

### Exécution fiable avec fichier temporaire

Idéal pour les commandes complexes qui peuvent être bloquées:

```powershell
$result = Invoke-WSLCommand -Command "git log --oneline | head -n 10" -UseTempFile
```

### Commande avec entrée standard

Pour les commandes qui attendent une entrée:

```powershell
$input = "Contenu à analyser"
$result = Invoke-WSLCommandWithInput -Command "wc -w" -Input $input
```

### Exécution de commandes en batch

Exécuter plusieurs commandes à partir d'un fichier:

```powershell
$batchFile = "D:\Dev\Apex_VBA_FRAMEWORK\tools\workflow\scripts\git_commands.txt"
$result = Start-WSLBatchFromFile -FilePath $batchFile
```

### Session WSL interactive préparée

Lance une session WSL avec contexte préconfiguré:

```powershell
Start-InteractiveWSLSession -WorkingDirectory "/mnt/d/Dev/Apex_VBA_FRAMEWORK" -InitCommand "git status"
```

## Fonctions disponibles

| Fonction | Description |
|----------|-------------|
| `Initialize-ApexWSLBridge` | Initialise le module avec les paramètres spécifiés |
| `Write-ApexLog` | Écrit dans le journal avec niveau de log |
| `Invoke-WSLCommand` | Exécute une commande WSL |
| `Invoke-WSLCommandWithRetry` | Exécute une commande avec retry en cas d'échec |
| `Invoke-WSLCommandWithInput` | Exécute une commande avec entrée standard |
| `Measure-WSLCommand` | Mesure le temps d'exécution d'une commande |
| `Test-WSLEnvironment` | Vérifie que l'environnement WSL est disponible |
| `Get-WSLMountStatus` | Vérifie le statut de montage d'un disque |
| `Run-SessionWithWSL` | Exécute un bloc de script avec contexte WSL |
| `Start-WSLBatchFromFile` | Exécute un lot de commandes depuis un fichier |
| `Start-InteractiveWSLSession` | Lance une session WSL interactive |
| `Invoke-PowerShellWithWSL` | Exécute un script PowerShell avec contexte WSL |

## Tests

Un script de test est fourni pour vérifier que toutes les fonctionnalités fonctionnent correctement:

```powershell
.\tools\powershell\test_apex_wslbridge.ps1
```

## Logs

Les logs sont enregistrés dans `D:\Dev\Apex_VBA_FRAMEWORK\logs\powershell_wsl.log`.

## Exemples d'intégration

### Dans un workflow Git

```powershell
Import-Module ".\tools\powershell\ApexWSLBridge.psm1" -Force

# Validation du montage WSL
$mount = Get-WSLMountStatus -Drive "d"
if ($mount.HasMetadata) {
    # Exécution d'un hook Git
    $result = Invoke-WSLCommand -Command "bash /mnt/d/Dev/Apex_VBA_FRAMEWORK/tools/workflow/git-hooks/pre-commit" -UseTempFile
    Write-Host $result
}
```

### Dans un script de build

```powershell
Import-Module ".\tools\powershell\ApexWSLBridge.psm1" -Force

# Création d'un bloc de script
$buildScript = {
    # Test de l'environnement
    if (Test-WSLEnvironment) {
        # Exécution des commandes de build
        $buildCmd = "cd /mnt/d/Dev/Apex_VBA_FRAMEWORK && python build.py"
        $result = Invoke-WSLCommandWithRetry -Command $buildCmd -MaxRetries 3
        
        # Traitement des résultats
        if ($result -match "BUILD SUCCESS") {
            Write-Host "Build réussi!" -ForegroundColor Green
        } else {
            Write-Host "Build échoué!" -ForegroundColor Red
        }
    }
}

# Exécution du bloc de script avec journalisation
Run-SessionWithWSL -ScriptBlock $buildScript -SessionName "BuildProcess"
```

## Dépannage

Si vous rencontrez des problèmes:

1. Vérifiez les logs (`D:\Dev\Apex_VBA_FRAMEWORK\logs\powershell_wsl.log`)
2. Exécutez le script de test pour diagnostiquer
3. Assurez-vous que votre configuration WSL est correcte (voir `docs/WSL_SETUP_GUIDE.md`)

## Compatibilité

- PowerShell 5.1+ (Windows PowerShell)
- PowerShell Core 7.0+
- WSL1 et WSL2
- Ubuntu 20.04/22.04 LTS

## Contributeurs

- Équipe APEX VBA Framework 