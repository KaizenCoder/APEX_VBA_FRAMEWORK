
<#
.SYNOPSIS
    Script de journalisation pour APEX Framework.
.DESCRIPTION
    Permet de journaliser les interactions avec Cursor/Claude et de créer des fichiers de session.
.PARAMETER Command
    La commande à exécuter: 'log', 'create-session', 'archive-session' ou 'list-sessions'.
.PARAMETER Prompt
    Le prompt envoyé à Cursor/Claude.
.PARAMETER Response
    La réponse de Cursor/Claude.
.PARAMETER Agent
    Le nom de l'agent (par défaut: "Claude 3.7 Sonnet").
.PARAMETER Note
    L'évaluation de la réponse (par défaut: "+").
.PARAMETER SessionId
    L'identifiant de la session (généré automatiquement si non spécifié).
.PARAMETER Description
    La description de la session (pour les commandes create-session et archive-session).
.PARAMETER ShowAll
    Pour la commande list-sessions, indique s'il faut afficher les sessions archivées.
.EXAMPLE
    .\Log-ApexCursor.ps1 -Command log -Prompt "Comment implémenter ILogger?" -Response "Voici comment..."
.EXAMPLE
    .\Log-ApexCursor.ps1 -Command create-session -Description "Session de développement de l'interface ILoggerBase"
.EXAMPLE
    .\Log-ApexCursor.ps1 -Command archive-session -SessionId "20240411-1530" -Description "Implémentation terminée avec succès"
.EXAMPLE
    .\Log-ApexCursor.ps1 -Command list-sessions -ShowAll
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
    [string]$Agent = "Claude 3.7 Sonnet",
    
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
$pythonScript = Join-Path $scriptPath "apex_cursor_autolog.py"

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
