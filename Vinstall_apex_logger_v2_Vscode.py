#!/usr/bin/env python3
# -----------------------------------------------------------------------------
# Script: Vinstall_apex_logger_v2_Vscode.py
# Description: Installateur du système de journalisation APEX Framework optimisé
#              pour l'environnement de développement VSCode
# Author: APEX Framework Team
# Date: 2025-04-13
# Version: 2.1
# -----------------------------------------------------------------------------

import os
import sys
from pathlib import Path
from datetime import datetime
import shutil
import json

# --- CONFIGURATION ADAPTÉE À L'ENVIRONNEMENT APEX FRAMEWORK ---
# Détection automatique du répertoire racine du projet
def find_project_root():
    """Trouve la racine du projet en se basant sur la présence des dossiers apex-core ou .git"""
    current = Path.cwd()
    
    # Remonte jusqu'à 5 niveaux pour trouver la racine du projet
    for _ in range(5):
        if (current / "apex-core").exists() or (current / ".git").exists():
            return current
        parent = current.parent
        if parent == current:  # Atteint la racine du système
            break
        current = parent
    
    # Si non trouvé, utilise le répertoire courant
    return Path.cwd()

# Configuration des chemins APEX Framework adaptés pour VSCode
ROOT = find_project_root()
LOGS_DIR = ROOT / "logs"
TOOLS_DIR = ROOT / "tools" / "python" / "vscode_logger"
CONFIG_DIR = ROOT / "config" / "vscode"
TEMPLATES_DIR = ROOT / "tools" / "templates" / "vscode"
HISTORY_DIR = ROOT / "logs" / "history" / "vscode_sessions"

# Fichiers de journalisation adaptés pour VSCode
JOURNAL_FILE = LOGS_DIR / "apex-vscode-journal.md"
SESSION_TEMPLATE = TEMPLATES_DIR / "session_template.md"
AUTOLOG_SCRIPT = TOOLS_DIR / "apex_vscode_autolog.py"
POWERSHELL_WRAPPER = TOOLS_DIR / "Log-ApexVSCode.ps1"
CONFIG_FILE = CONFIG_DIR / "vscode_logger_config.json"
HISTORY_INDEX = HISTORY_DIR / "sessions_index.json"

# Vérification de l'encodage UTF-8
def ensure_utf8_no_bom():
    """Vérifie que le script est encodé en UTF-8 sans BOM"""
    with open(__file__, 'rb') as f:
        content = f.read()
        if content.startswith(b'\xef\xbb\xbf'):
            print("[⚠️] AVERTISSEMENT: Ce fichier est encodé avec BOM. Conversion nécessaire.")
            with open(__file__, 'w', encoding='utf-8') as f_out:
                f_out.write(content.decode('utf-8-sig'))
            print("[✓] Conversion en UTF-8 sans BOM effectuée.")
            return False
    return True

# Création des répertoires nécessaires
for directory in [LOGS_DIR, TOOLS_DIR, CONFIG_DIR, TEMPLATES_DIR, HISTORY_DIR]:
    directory.mkdir(parents=True, exist_ok=True)

# Contenu du script d'autologging
AUTOLOG_CONTENT = r'''#!/usr/bin/env python3
# -----------------------------------------------------------------------------
# Script: apex_vscode_autolog.py
# Description: Script de journalisation automatique pour APEX Framework (VSCode)
# Author: APEX Framework Team
# Date: 2025-04-13
# Version: 2.1
# -----------------------------------------------------------------------------

import os
import json
from datetime import datetime
import sys
from pathlib import Path
import shutil
import re

# Trouver la racine du projet
def find_project_root():
    current = Path(__file__).resolve().parent
    for _ in range(5):
        if (current / "../../apex-core").exists() or (current / "../../.git").exists():
            return current.parent.parent
        parent = current.parent
        if parent == current:
            break
        current = parent
    return Path.cwd()

ROOT = find_project_root()
LOG_PATH = ROOT / "logs" / "apex-vscode-journal.md"
CONFIG_PATH = ROOT / "config" / "vscode" / "vscode_logger_config.json"
HISTORY_DIR = ROOT / "logs" / "history" / "vscode_sessions"
HISTORY_INDEX = HISTORY_DIR / "sessions_index.json"

# Créer le répertoire d'historisation s'il n'existe pas
HISTORY_DIR.mkdir(parents=True, exist_ok=True)

# Charger la configuration
def load_config():
    try:
        if CONFIG_PATH.exists():
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception as e:
        print(f"[⚠️] Erreur de chargement de config: {e}")
    
    # Configuration par défaut
    return {
        "log_format": "markdown",
        "max_prompt_length": 300,
        "max_response_length": 500,
        "include_timestamps": True,
        "include_session_id": True,
        "session_id_format": "%Y%m%d-%H%M",
        "journal_file": str(LOG_PATH),
        "history_enabled": True,
        "history_directory": str(HISTORY_DIR)
    }

config = load_config()

def generate_session_id():
    """Génère un ID de session basé sur l'horodatage"""
    if config.get("include_session_id", True):
        format_string = config.get("session_id_format", "%Y%m%d-%H%M")
        return datetime.now().strftime(format_string)
    return ""

def log_vscode_interaction(prompt, response, agent="GitHub Copilot", note="+", session_id=None):
    """Enregistre une interaction avec VSCode/GitHub Copilot dans le journal APEX"""
    if session_id is None:
        session_id = generate_session_id()
    
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    max_prompt = config.get("max_prompt_length", 300)
    max_response = config.get("max_response_length", 500)
    
    entry = f"""---

### {now} - Session {session_id}

**Contexte :** APEX Framework
**Prompt :** {prompt.strip()[:max_prompt]}{"..." if len(prompt) > max_prompt else ""}
**Agent IA :** {agent}
**Avis :** {note}
**Réponse (extrait) :** {response.strip()[:max_response]}{"..." if len(response) > max_response else ""}
"""
    
    log_file = Path(config.get("journal_file", str(LOG_PATH)))
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(entry)
    
    print(f"[✓] Interaction ajoutée au journal: {log_file}")
    
    # Ajouter au fichier de session s'il existe
    session_file = ROOT / "logs" / f"vscode_session_{session_id}.md"
    if session_file.exists():
        with open(session_file, "a", encoding="utf-8") as f:
            f.write(f"\n### {now}\n\n**Prompt:** {prompt.strip()[:max_prompt]}\n\n**Réponse:** {response.strip()[:1000]}\n\n---\n")
        print(f"[✓] Interaction ajoutée au fichier de session: {session_file}")
    
    return True

def create_session_file(session_id=None, description=""):
    """Crée un fichier de session dans le dossier des journaux"""
    if session_id is None:
        session_id = generate_session_id()
    
    template_path = ROOT / "tools" / "templates" / "vscode" / "session_template.md"
    if template_path.exists():
        with open(template_path, "r", encoding="utf-8") as f:
            template = f.read()
    else:
        template = """# Session VSCode: {session_id}
Date: {date}
Description: {description}

## Contexte
- Framework APEX
- Phase de développement

## Objectifs
- [ ] Objectif 1
- [ ] Objectif 2
- [ ] Objectif 3

## Résultats
À compléter...

## Journal des interactions
"""
    
    now = datetime.now()
    session_file = ROOT / "logs" / f"vscode_session_{session_id}.md"
    
    content = template.format(
        session_id=session_id,
        date=now.strftime("%Y-%m-%d %H:%M"),
        description=description
    )
    
    with open(session_file, "w", encoding="utf-8") as f:
        f.write(content)
    
    print(f"[✓] Fichier de session créé: {session_file}")
    
    # Ajouter l'entrée dans l'index des sessions
    add_to_sessions_index(session_id, description, now.strftime("%Y-%m-%d %H:%M"))
    
    return session_file

def add_to_sessions_index(session_id, description, timestamp):
    """Ajoute une entrée à l'index des sessions"""
    index_file = HISTORY_INDEX
    
    # Charger l'index existant ou créer un nouvel index
    if index_file.exists():
        try:
            with open(index_file, "r", encoding="utf-8") as f:
                index = json.load(f)
        except:
            index = {"sessions": []}
    else:
        index = {"sessions": []}
    
    # Ajouter la nouvelle session
    index["sessions"].append({
        "id": session_id,
        "description": description,
        "created_at": timestamp,
        "archived": False,
        "archive_path": None
    })
    
    # Enregistrer l'index mis à jour
    with open(index_file, "w", encoding="utf-8") as f:
        json.dump(index, f, indent=2)
    
    return True

def archive_session(session_id, archive_description=None):
    """Archive une session terminée"""
    # Vérifier si l'historisation est activée
    if not config.get("history_enabled", True):
        print("[⚠️] Historisation désactivée dans la configuration.")
        return False
    
    session_file = ROOT / "logs" / f"vscode_session_{session_id}.md"
    
    # Vérifier que le fichier de session existe
    if not session_file.exists():
        print(f"[⚠️] Fichier de session introuvable: {session_file}")
        return False
    
    # Lire le contenu du fichier pour extraction des métadonnées
    with open(session_file, "r", encoding="utf-8") as f:
        content = f.read()
    
    # Extraire la description et la date
    description_match = re.search(r"Description: (.*?)\n", content)
    description = description_match.group(1) if description_match else "Sans description"
    
    date_match = re.search(r"Date: (.*?)\n", content)
    date = date_match.group(1) if date_match else datetime.now().strftime("%Y-%m-%d %H:%M")
    
    # Créer un sous-dossier par année-mois
    date_obj = datetime.strptime(date, "%Y-%m-%d %H:%M")
    month_dir = HISTORY_DIR / f"{date_obj.year}-{date_obj.month:02d}"
    month_dir.mkdir(exist_ok=True)
    
    # Ajouter un résumé de fin à la session
    with open(session_file, "a", encoding="utf-8") as f:
        f.write(f"\n## Résumé de fin de session\n\n")
        if archive_description:
            f.write(f"{archive_description}\n\n")
        f.write(f"Session archivée le {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")
    
    # Copier le fichier vers l'archive
    archive_path = month_dir / f"session_{session_id}.md"
    shutil.copy2(session_file, archive_path)
    
    # Mettre à jour l'index des sessions
    update_session_index(session_id, archive_path)
    
    print(f"[✓] Session {session_id} archivée avec succès: {archive_path}")
    return True

def update_session_index(session_id, archive_path):
    """Met à jour l'index des sessions après archivage"""
    index_file = HISTORY_INDEX
    
    # Charger l'index
    if index_file.exists():
        try:
            with open(index_file, "r", encoding="utf-8") as f:
                index = json.load(f)
        except:
            index = {"sessions": []}
    else:
        index = {"sessions": []}
    
    # Mettre à jour l'entrée de la session
    found = False
    for session in index["sessions"]:
        if session["id"] == session_id:
            session["archived"] = True
            session["archive_path"] = str(archive_path)
            session["archived_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            found = True
            break
    
    # Si la session n'est pas dans l'index, l'ajouter
    if not found:
        index["sessions"].append({
            "id": session_id,
            "description": "Session archivée",
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "archived": True,
            "archive_path": str(archive_path),
            "archived_at": datetime.now().strftime("%Y-%m-%d %H:%M")
        })
    
    # Enregistrer l'index mis à jour
    with open(index_file, "w", encoding="utf-8") as f:
        json.dump(index, f, indent=2)
    
    return True

def list_sessions(show_archived=False):
    """Liste toutes les sessions existantes"""
    index_file = HISTORY_INDEX
    
    if not index_file.exists():
        print("[ℹ️] Aucune session enregistrée.")
        return []
    
    try:
        with open(index_file, "r", encoding="utf-8") as f:
            index = json.load(f)
    except:
        print("[⚠️] Erreur de lecture de l'index des sessions.")
        return []
    
    sessions = index.get("sessions", [])
    
    if not sessions:
        print("[ℹ️] Aucune session enregistrée.")
        return []
    
    # Filtrer les sessions archivées si demandé
    if not show_archived:
        sessions = [s for s in sessions if not s.get("archived", False)]
    
    # Afficher les sessions
    print(f"\n{'=' * 80}")
    print(f"{'ID':12} | {'DATE':16} | {'STATUT':10} | DESCRIPTION")
    print(f"{'-' * 12}-+-{'-' * 16}-+-{'-' * 10}-+-{'-' * 40}")
    
    for session in sessions:
        status = "Archivée" if session.get("archived", False) else "Active"
        print(f"{session['id']:12} | {session['created_at']:16} | {status:10} | {session['description'][:40]}")
    
    print(f"{'=' * 80}\n")
    return sessions

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: apex_vscode_autolog.py <command> [args...]")
        print("Commands:")
        print("  log '<prompt>' '<response>' [agent] [note] [session_id]")
        print("  create-session [description] [session_id]")
        print("  archive-session <session_id> [description]")
        print("  list-sessions [--all]")
        sys.exit(1)
    
    command = sys.argv[1]
    
    if command == "log":
        if len(sys.argv) < 4:
            print("Usage: apex_vscode_autolog.py log '<prompt>' '<response>' [agent] [note] [session_id]")
            sys.exit(1)
        
        prompt = sys.argv[2]
        response = sys.argv[3]
        agent = sys.argv[4] if len(sys.argv) > 4 else "GitHub Copilot" 
        note = sys.argv[5] if len(sys.argv) > 5 else "+"
        session_id = sys.argv[6] if len(sys.argv) > 6 else None
        
        log_vscode_interaction(prompt, response, agent, note, session_id)
    
    elif command == "create-session":
        description = sys.argv[2] if len(sys.argv) > 2 else ""
        session_id = sys.argv[3] if len(sys.argv) > 3 else None
        
        create_session_file(session_id, description)
    
    elif command == "archive-session":
        if len(sys.argv) < 3:
            print("Usage: apex_vscode_autolog.py archive-session <session_id> [description]")
            sys.exit(1)
            
        session_id = sys.argv[2]
        description = sys.argv[3] if len(sys.argv) > 3 else None
        
        archive_session(session_id, description)
    
    elif command == "list-sessions":
        show_all = len(sys.argv) > 2 and sys.argv[2] == "--all"
        list_sessions(show_all)
    
    else:
        print(f"Commande inconnue: {command}")
        sys.exit(1)
'''

# Contenu du wrapper PowerShell
POWERSHELL_CONTENT = r'''
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
'''

# Contenu du fichier de configuration
CONFIG_CONTENT = r'''{
    "log_format": "markdown",
    "max_prompt_length": 300,
    "max_response_length": 500,
    "include_timestamps": true,
    "include_session_id": true,
    "session_id_format": "%Y%m%d-%H%M",
    "journal_file": "__LOG_PATH__",
    "history_enabled": true,
    "history_directory": "__HISTORY_PATH__",
    "editor": "vscode"
}'''

# Contenu du template de session
SESSION_TEMPLATE_CONTENT = r'''# Session VSCode: {session_id}
Date: {date}
Description: {description}

## Contexte
- Framework APEX
- Phase de développement

## Objectifs
- [ ] Objectif 1
- [ ] Objectif 2
- [ ] Objectif 3

## Résultats
À compléter...

## Journal des interactions
'''

# Écriture des fichiers
def write_file(path, content):
    """Écrit le contenu dans un fichier avec encodage UTF-8"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)
    return path

# Installation des composants
def install_components():
    """Installe tous les composants du logger"""
    # Écriture des fichiers avec gestion des erreurs
    try:
        # Création du journal s'il n'existe pas
        if not JOURNAL_FILE.exists():
            journal_init = f"# Journal des interactions Cursor/Claude - APEX Framework\n\n### {datetime.now().strftime('%Y-%m-%d')} – Initialisation\n\n*Journal initialisé pour le projet APEX Framework.*\n"
            write_file(JOURNAL_FILE, journal_init)
        
        # Configuration
        config_content = CONFIG_CONTENT.replace("__LOG_PATH__", str(JOURNAL_FILE).replace("\\", "/"))
        config_content = config_content.replace("__HISTORY_PATH__", str(HISTORY_DIR).replace("\\", "/"))
        write_file(CONFIG_FILE, config_content)
        
        # Création de l'index des sessions s'il n'existe pas
        if not HISTORY_INDEX.exists():
            sessions_index = {"sessions": []}
            with open(HISTORY_INDEX, "w", encoding="utf-8") as f:
                json.dump(sessions_index, f, indent=2)
        
        # Scripts
        write_file(AUTOLOG_SCRIPT, AUTOLOG_CONTENT)
        write_file(POWERSHELL_WRAPPER, POWERSHELL_CONTENT)
        write_file(SESSION_TEMPLATE, SESSION_TEMPLATE_CONTENT)
        
        # Rendre les scripts exécutables sur Unix
        if os.name != "nt":  # Si pas Windows
            os.chmod(AUTOLOG_SCRIPT, 0o755)
        
        return True
    except Exception as e:
        print(f"[⚠️] Erreur lors de l'installation: {e}")
        return False

# Vérification de l'environnement Python
def check_environment():
    """Vérifie l'environnement Python"""
    python_version = sys.version_info
    if python_version.major < 3 or (python_version.major == 3 and python_version.minor < 6):
        print(f"[⚠️] Version Python incompatible: {python_version.major}.{python_version.minor}")
        print("Python 3.6 ou supérieur est requis.")
        return False
    return True

# Fonction principale
def main():
    """Fonction principale d'installation"""
    print("\n=== Installation du Logger APEX pour VSCode avec Historisation ===\n")
    
    # Vérification de l'encodage
    ensure_utf8_no_bom()
    
    # Vérification de l'environnement
    if not check_environment():
        return False
    
    # Installation des composants
    if install_components():
        print("\n[✓] Logger APEX pour VSCode installé avec succès!")
        print(f"  - Racine du projet:    {ROOT}")
        print(f"  - Script Python:       {AUTOLOG_SCRIPT}")
        print(f"  - Script PowerShell:   {POWERSHELL_WRAPPER}")
        print(f"  - Journal:             {JOURNAL_FILE}")
        print(f"  - Répertoire archives: {HISTORY_DIR}")
        print(f"  - Index des sessions:  {HISTORY_INDEX}")
        print(f"  - Configuration:       {CONFIG_FILE}")
        print(f"  - Template:            {SESSION_TEMPLATE}")
        
        print("\nPour utiliser le logger:")
        print("  - PowerShell: .\\Log-ApexVSCode.ps1 -Command log -Prompt 'Question' -Response 'Réponse'")
        print("  - Python:     python apex_vscode_autolog.py log 'Question' 'Réponse'")
        
        print("\nPour créer une nouvelle session:")
        print("  - PowerShell: .\\Log-ApexVSCode.ps1 -Command create-session -Description 'Description de la session'")
        print("  - Python:     python apex_vscode_autolog.py create-session 'Description de la session'")
        
        print("\nPour archiver une session terminée:")
        print("  - PowerShell: .\\Log-ApexVSCode.ps1 -Command archive-session -SessionId '20250413-1530' -Description 'Résumé de la session'")
        print("  - Python:     python apex_vscode_autolog.py archive-session '20250413-1530' 'Résumé de la session'")
        
        print("\nPour lister les sessions:")
        print("  - PowerShell: .\\Log-ApexVSCode.ps1 -Command list-sessions [-ShowAll]")
        print("  - Python:     python apex_vscode_autolog.py list-sessions [--all]")
        
        return True
    else:
        print("\n[❌] L'installation a échoué. Voir les erreurs ci-dessus.")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)