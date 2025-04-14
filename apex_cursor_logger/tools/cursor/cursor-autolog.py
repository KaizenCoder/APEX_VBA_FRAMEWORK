#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import json
import locale
from datetime import datetime
from pathlib import Path
try:
    import emoji
except ImportError:
    print("Installation de la bibliothèque emoji...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "emoji"])
    import emoji

def setup_encoding():
    """Configure l'encodage pour l'environnement"""
    if sys.platform == 'win32':
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    return 'utf-8'

def sanitize_text(text):
    """Nettoie et valide le texte pour l'encodage UTF-8"""
    try:
        # Décomposition des émojis composés
        text = emoji.demojize(text)
        # Conversion en émojis simples
        text = emoji.emojize(text, language='alias')
        # Encodage/décodage pour validation
        return text.encode('utf-8').decode('utf-8')
    except UnicodeError as e:
        print(f"⚠️ Attention : Problème avec certains caractères ({str(e)})")
        # Remplacement des caractères problématiques
        return ''.join(char if ord(char) < 0x10000 else '?' for char in text)

def format_markdown_text(text):
    """Formate le texte pour Markdown en échappant les caractères spéciaux"""
    markdown_chars = ['*', '_', '`', '#', '[', ']', '(', ')', '<', '>', '|']
    for char in markdown_chars:
        text = text.replace(char, '\\' + char)
    return text

def log_cursor_interaction(prompt, response, agent, note):
    try:
        # Configuration de l'encodage
        encoding = setup_encoding()
        
        # Nettoyage et formatage des entrées
        prompt = format_markdown_text(sanitize_text(prompt))
        response = format_markdown_text(sanitize_text(response))
        agent = sanitize_text(agent)
        note = format_markdown_text(sanitize_text(note))
        
        # Chargement de la configuration
        config_path = Path(__file__).parent.parent.parent / "logger_config.json"
        with open(config_path, encoding=encoding) as f:
            config = json.load(f)
            
        # Construction du chemin du journal
        journal_path = Path(__file__).parent.parent.parent / "logs" / "cursor-journal.md"
        
        # Création du dossier logs si nécessaire
        journal_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Formatage de l'entrée
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entry = f"\n### {timestamp} - {agent}\n\n"
        entry += f"**Prompt**: {prompt}\n\n"
        entry += f"**Réponse**: {response}\n\n"
        entry += f"**Note**: {note}\n"
        entry += "---\n"
        
        # Vérification de l'encodage du fichier existant
        if journal_path.exists():
            with open(journal_path, 'rb') as f:
                content = f.read()
                if content.startswith(b'\xef\xbb\xbf'):  # BOM UTF-8
                    print("⚠️ Attention : BOM UTF-8 détecté, conversion en UTF-8 sans BOM")
                    content = content[3:]
                    with open(journal_path, 'wb') as f:
                        f.write(content)
        
        # Écriture dans le journal
        with open(journal_path, 'a', encoding=encoding) as f:
            f.write(entry)
            
        print("✅ Journalisation réussie")
        return 0
            
    except Exception as e:
        print(f"❌ Erreur : {str(e)}", file=sys.stderr)
        return 1

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: cursor-autolog.py <prompt> <response> [agent] [note]", file=sys.stderr)
        sys.exit(1)
        
    prompt = sys.argv[1]
    response = sys.argv[2]
    agent = sys.argv[3] if len(sys.argv) > 3 else "Claude-3"
    note = sys.argv[4] if len(sys.argv) > 4 else "+"
    
    sys.exit(log_cursor_interaction(prompt, response, agent, note))
