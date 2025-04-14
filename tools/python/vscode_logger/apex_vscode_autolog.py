#!/usr/bin/env python3
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
