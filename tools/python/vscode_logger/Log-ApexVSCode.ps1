
<#
.SYNOPSIS
    Script de journalisation pour APEX Framework dans VSCode.
.DESCRIPTION
    Permet de journaliser les interactions avec GitHub Copilot dans VSCode et de créer des fichiers de session.
.PARAMETER Command
    La commande à exécuter: 'log', 'create-session', 'archive-session' ou 'list-sessions'.
.PARAMETER Prompt
    Le prompt envoyé à GitHub Copilot.
.PARAMETER Response
    La réponse de GitHub Copilot.
.PARAMETER Agent
    Le nom de l'agent (par défaut: "GitHub Copilot").
.PARAMETER Note
    L'évaluation de la réponse (par défaut: "+").
.PARAMETER SessionId
    L'identifiant de la session (généré automatiquement si non spécifié).
.PARAMETER Description
    La description de la session (pour les commandes create-session et archive-session).
.PARAMETER ShowAll
    Pour la commande list-sessions, indique s'il faut afficher les sessions archivées.
.EXAMPLE
    .\Log-ApexVSCode.ps1 -Command log -Prompt "Comment implémenter ILogger?" -Response "Voici comment..."
.EXAMPLE
    .\Log-ApexVSCode.ps1 -Command create-session -Description "Session de développement de l'interface ILoggerBase"
.EXAMPLE
    .\Log-ApexVSCode.ps1 -Command archive-session -SessionId "20250413-1530" -Description "Implémentation terminée avec succès"
.EXAMPLE
    .\Log-ApexVSCode.ps1 -Command list-sessions -ShowAll
#>
param (
    [Parameter(Mandatory=$true)]
    [ValidateSet("log", "create-session", "archive-session", "list-sessions")]
    [string]$Command,
    
    [Parameter(Mandatory=$false)]
    [string]$Prompt,
    
    [Parameter(Mandatory=$false)]
    [string]$Response,
    
    [Parameter(Mandatory=$false)]
    [string]$Agent = "GitHub Copilot",
    
    [Parameter(Mandatory=$false)]
    [string]$Note = "+",
    
    [Parameter(Mandatory=$false)]
    [string]$SessionId,
    
    [Parameter(Mandatory=$false)]
    [string]$Description,
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowAll
)

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonScript = Join-Path $scriptPath "apex_vscode_autolog.py"

# Vérifier si Python est installé
try {
    python --version | Out-Null
}
catch {
    Write-Host "[⚠️] Python n'est pas installé ou n'est pas dans le PATH. Installation requise." -ForegroundColor Red
    exit 1
}

# Exécuter la commande appropriée
switch ($Command) {
    "log" {
        if (-not $Prompt -or -not $Response) {
            Write-Host "Pour la commande 'log', les paramètres Prompt et Response sont obligatoires." -ForegroundColor Red
            exit 1
        }
        
        $args = @("log", $Prompt, $Response, $Agent, $Note)
        if ($SessionId) {
            $args += $SessionId
        }
        
        & python $pythonScript $args
    }
    "create-session" {
        $args = @("create-session")
        if ($Description) {
            $args += $Description
        }
        if ($SessionId) {
            $args += $SessionId
        }
        
        & python $pythonScript $args
    }
    "archive-session" {
        if (-not $SessionId) {
            Write-Host "Pour la commande 'archive-session', le paramètre SessionId est obligatoire." -ForegroundColor Red
            exit 1
        }
        
        $args = @("archive-session", $SessionId)
        if ($Description) {
            $args += $Description
        }
        
        & python $pythonScript $args
    }
    "list-sessions" {
        $args = @("list-sessions")
        if ($ShowAll) {
            $args += "--all"
        }
        
        & python $pythonScript $args
    }
}
