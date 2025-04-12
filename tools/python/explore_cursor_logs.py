import os
import sqlite3
import json
import hashlib
from pathlib import Path
from datetime import datetime
from contextlib import contextmanager
from typing import Dict, List, Optional
import textwrap
import re

class CursorDBError(Exception):
    """Erreur personnalisée pour la gestion de la base Cursor."""
    pass

@contextmanager
def safe_db_connection(db_path):
    """Gère la connexion à la base de données de manière sécurisée."""
    if not os.path.exists(db_path):
        raise CursorDBError(f"Base de données non trouvée: {db_path}")
        
    try:
        # Vérifier si le fichier est accessible en lecture
        with open(db_path, 'rb') as f:
            # Calculer un hash pour détecter les modifications
            file_hash = hashlib.md5(f.read()).hexdigest()
            
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        
        # Vérifier la structure de la base
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = [row[0] for row in cursor.fetchall()]
        
        if 'ItemTable' not in tables:
            raise CursorDBError("Structure de base invalide: table ItemTable manquante")
            
        yield conn
        
        # Vérifier si le fichier a été modifié pendant l'utilisation
        with open(db_path, 'rb') as f:
            new_hash = hashlib.md5(f.read()).hexdigest()
            if new_hash != file_hash:
                print("⚠️ Attention: La base a été modifiée pendant la lecture")
                
    except sqlite3.Error as e:
        raise CursorDBError(f"Erreur SQLite: {e}")
    finally:
        if 'conn' in locals():
            conn.close()

def get_cursor_storage_path():
    """Retourne le chemin du dossier de stockage Cursor."""
    appdata = os.getenv('APPDATA')
    if not appdata:
        raise CursorDBError("Variable d'environnement APPDATA non trouvée")
    return os.path.join(appdata, 'Cursor', 'User', 'workspaceStorage')

def find_workspace_db(workspace_path):
    """Trouve le fichier state.vscdb dans un workspace."""
    try:
        db_path = os.path.join(workspace_path, 'state.vscdb')
        if os.path.exists(db_path):
            # Vérifier les permissions
            if not os.access(db_path, os.R_OK):
                raise CursorDBError(f"Pas de permission de lecture: {db_path}")
            return db_path
        return None
    except Exception as e:
        raise CursorDBError(f"Erreur lors de la recherche de la base: {e}")

def analyze_message_role(text: str, key: str) -> str:
    """Analyse le contenu du message pour déterminer son rôle."""
    text = text.strip()
    
    # Vérifier les préfixes explicites
    if text.startswith(('Human:', 'User:')):
        return 'user'
    if text.startswith('Assistant:'):
        return 'assistant'
    
    # Analyser le contenu pour déduire le rôle
    user_indicators = [
        'PS C:\\',  # Commandes PowerShell
        '>>',       # Suite de commande
        '@https://', # Liens partagés
        '?',        # Questions
        'attends',  # Commandes à l'assistant
        'stop',
        'non',
        'oui',
        'ok'
    ]
    
    assistant_indicators = [
        'Je vais',
        'Voici',
        'J\'ai',
        'Pour',
        'Le script',
        'Les modifications',
        'Maintenant',
        'Analysons'
    ]
    
    # Vérifier les indicateurs
    for indicator in user_indicators:
        if text.startswith(indicator):
            return 'user'
            
    for indicator in assistant_indicators:
        if text.startswith(indicator):
            return 'assistant'
    
    # Analyser la structure du message
    if len(text.split('\n')) > 3:  # Messages longs sont souvent de l'assistant
        return 'assistant'
    if len(text) < 20:  # Messages courts sont souvent de l'utilisateur
        return 'user'
    
    return 'unknown'

def extract_conversations(db_path, limit=100, offset=0):
    """Extrait les conversations de la base de données avec pagination."""
    try:
        with safe_db_connection(db_path) as conn:
            cursor = conn.cursor()
            
            # Récupérer les conversations avec pagination
            cursor.execute("""
                SELECT key, value 
                FROM ItemTable 
                WHERE key IN (
                    'aiService.prompts', 
                    'workbench.panel.aichat.view.aichat.chatdata',
                    'workbench.panel.aichat.view.aichat'
                )
                ORDER BY key DESC
                LIMIT ? OFFSET ?
            """, (limit, offset))
            
            conversations = []
            for row in cursor:
                try:
                    data = json.loads(row['value'])
                    
                    # Traiter les différents formats de données
                    if isinstance(data, list):
                        for item in data:
                            if isinstance(item, dict):
                                text = item.get('text', '').strip()
                                role = item.get('role', 'unknown')
                                
                                # Si le rôle n'est pas explicite, l'analyser
                                if role == 'unknown':
                                    role = analyze_message_role(text, row['key'])
                                
                                conv = {
                                    'key': row['key'],
                                    'timestamp': item.get('timestamp', 
                                                item.get('created_at', 
                                                item.get('date', 'N/A'))),
                                    'text': text,
                                    'role': role,
                                    'model': item.get('model', 
                                             item.get('assistant_model', 
                                             item.get('ai_model', 'unknown')))
                                }
                                
                                # Nettoyer les données sensibles
                                if 'api_key' in conv:
                                    conv['api_key'] = '[MASQUÉ]'
                                    
                                conversations.append(conv)
                    
                    elif isinstance(data, dict):
                        # Traiter les conversations au format dict
                        messages = data.get('messages', [])
                        if isinstance(messages, list):
                            for msg in messages:
                                if isinstance(msg, dict):
                                    text = msg.get('content', msg.get('text', '')).strip()
                                    role = msg.get('role', 'unknown')
                                    
                                    # Si le rôle n'est pas explicite, l'analyser
                                    if role == 'unknown':
                                        role = analyze_message_role(text, row['key'])
                                    
                                    conv = {
                                        'key': row['key'],
                                        'timestamp': msg.get('timestamp', 'N/A'),
                                        'text': text,
                                        'role': role,
                                        'model': msg.get('model', 'unknown')
                                    }
                                    conversations.append(conv)
                                    
                except json.JSONDecodeError:
                    print(f"⚠️ Erreur de décodage JSON pour la clé: {row['key']}")
            
            # Trier les conversations par timestamp si possible
            conversations.sort(key=lambda x: x.get('timestamp', 'N/A'), reverse=True)
            
            return conversations, len(conversations)
            
    except CursorDBError as e:
        print(f"❌ {e}")
        return [], 0

def format_conversation(conv):
    """Formate une conversation pour l'affichage."""
    if 'timestamp' in conv:
        return f"""
  🕒 Timestamp: {conv['timestamp']}
  🤖 Modèle: {conv['model']}
  👤 Rôle: {conv['role']}
  💬 Texte: {conv['text'][:200]}...
"""
    else:
        return f"""
  🔑 Clé: {conv['key']}
  📝 Type: {conv['data']}
"""

def format_chat_for_human(conversations: List[Dict]) -> str:
    """Formate une conversation pour une lecture humaine."""
    output = []
    output.append("\n" + "="*80)
    output.append("📝 CONVERSATION")
    output.append("="*80 + "\n")

    for idx, conv in enumerate(conversations, 1):
        # En-tête du message
        output.append(f"Message #{idx}")
        output.append("-" * 40)
        
        # Rôle avec émoji approprié
        role_display = {
            'user': '👤 UTILISATEUR',
            'assistant': '🤖 ASSISTANT',
            'system': '⚙️ SYSTÈME',
            'unknown': '❓ INCONNU'
        }.get(conv.get('role', 'unknown'), '❓ INCONNU')
        
        # Formater le timestamp
        timestamp = conv.get('timestamp', 'N/A')
        if timestamp != 'N/A':
            try:
                if isinstance(timestamp, (int, float)):
                    timestamp = datetime.fromtimestamp(timestamp/1000 if timestamp > 1e10 else timestamp)
                elif isinstance(timestamp, str):
                    timestamp = datetime.fromisoformat(timestamp.replace('Z', '+00:00'))
                timestamp = timestamp.strftime('%Y-%m-%d %H:%M:%S')
            except:
                pass
        
        output.append(f"{role_display} - {timestamp}")
        
        # Contenu du message
        text = conv.get('text', '').strip()
        if text:
            # Formater le texte avec indentation et largeur fixe
            wrapped_text = textwrap.fill(text, width=70, initial_indent='    ', subsequent_indent='    ')
            output.append(f"\n{wrapped_text}\n")
        
        # Modèle utilisé (si disponible)
        model = conv.get('model')
        if model and model != 'unknown':
            output.append(f"    Modèle: {model}")
        
        output.append("")  # Ligne vide entre les messages

    return "\n".join(output)

def generate_conversation_summary(conversations: List[Dict]) -> str:
    """Génère un résumé structuré d'une conversation."""
    if not conversations:
        return "Aucune conversation trouvée."
    
    # Compter les messages par rôle
    role_counts = {'user': 0, 'assistant': 0, 'system': 0, 'unknown': 0}
    for conv in conversations:
        role = conv.get('role', 'unknown')
        role_counts[role] += 1
    
    # Extraire les modèles utilisés
    models = set()
    for conv in conversations:
        model = conv.get('model')
        if model and model != 'unknown':
            models.add(model)
    
    # Déterminer la période de la conversation
    timestamps = []
    for conv in conversations:
        timestamp = conv.get('timestamp', None)
        if timestamp and timestamp != 'N/A':
            try:
                if isinstance(timestamp, (int, float)):
                    dt = datetime.fromtimestamp(timestamp/1000 if timestamp > 1e10 else timestamp)
                    timestamps.append(dt)
                elif isinstance(timestamp, str):
                    dt = datetime.fromisoformat(timestamp.replace('Z', '+00:00'))
                    timestamps.append(dt)
            except:
                pass
    
    # Déterminer la période
    start_date = min(timestamps) if timestamps else None
    end_date = max(timestamps) if timestamps else None
    
    # Extraire les sujets principaux (mots-clés)
    all_text = " ".join([conv.get('text', '') for conv in conversations if conv.get('text')])
    words = re.findall(r'\b\w{4,}\b', all_text.lower())
    stop_words = {'dans', 'avec', 'pour', 'cette', 'votre', 'vous', 'nous', 'mais', 'sont', 'comme'}
    meaningful_words = [w for w in words if w not in stop_words]
    
    # Compter les mots et trouver les plus fréquents
    from collections import Counter
    word_counts = Counter(meaningful_words)
    common_words = [word for word, count in word_counts.most_common(5)]
    
    # Construire le résumé
    summary = []
    summary.append("\n" + "="*80)
    summary.append("📊 RÉSUMÉ DE CONVERSATION")
    summary.append("="*80 + "\n")
    
    summary.append(f"🔢 Nombre total de messages: {len(conversations)}")
    summary.append(f"👤 Messages de l'utilisateur: {role_counts['user']}")
    summary.append(f"🤖 Messages de l'assistant: {role_counts['assistant']}")
    
    if models:
        summary.append(f"🧠 Modèles utilisés: {', '.join(models)}")
    
    if start_date and end_date:
        summary.append(f"📅 Période: {start_date.strftime('%Y-%m-%d %H:%M')} à {end_date.strftime('%Y-%m-%d %H:%M')}")
    
    if common_words:
        summary.append(f"🔍 Mots-clés fréquents: {', '.join(common_words)}")
    
    # Extraire les premiers échanges (début de conversation)
    if conversations:
        first_user_msg = next((conv for conv in conversations if conv.get('role') == 'user'), None)
        if first_user_msg:
            text = first_user_msg.get('text', '').strip()
            if text:
                summary.append("\n📝 Premier message utilisateur:")
                summary.append(f"    {text[:150]}..." if len(text) > 150 else f"    {text}")
    
    # Résumé des actions/questions
    summary.append("\n🔄 Résumé des interactions:")
    
    # Détection de modèles de questions/actions
    questions = [conv for conv in conversations if conv.get('role') == 'user' and '?' in conv.get('text', '')]
    code_requests = [conv for conv in conversations if conv.get('role') == 'user' and any(kw in conv.get('text', '').lower() for kw in ['code', 'script', 'fonction', 'créer', 'modifier'])]
    
    if questions:
        summary.append(f"    • {len(questions)} questions posées")
    if code_requests:
        summary.append(f"    • {len(code_requests)} demandes liées au code")
    
    return "\n".join(summary)

def main():
    """Fonction principale."""
    try:
        storage_path = get_cursor_storage_path()
        print(f"🔍 Recherche dans: {storage_path}")
        
        if not os.path.exists(storage_path):
            raise CursorDBError("Dossier de stockage Cursor non trouvé!")
        
        # Parcourir tous les workspaces
        for workspace in os.listdir(storage_path):
            workspace_path = os.path.join(storage_path, workspace)
            if not os.path.isdir(workspace_path):
                continue
                
            print(f"\n📁 Workspace: {workspace}")
            
            try:
                db_path = find_workspace_db(workspace_path)
                if not db_path:
                    print("  ⚠️ Pas de base de données trouvée")
                    continue
                    
                print(f"  📊 Base de données: {db_path}")
                conversations, total = extract_conversations(db_path, limit=50)  # Augmenté à 50 pour un meilleur résumé
                
                if not conversations:
                    print("  ⚠️ Pas de conversations trouvées")
                    continue
                    
                print(f"  💬 {len(conversations)}/{total} conversations trouvées")
                
                # Générer et afficher le résumé
                summary = generate_conversation_summary(conversations)
                print(summary)
                
                # Afficher aussi le détail des messages
                print("\n📜 Détail des messages:")
                print(format_chat_for_human(conversations))  # Afficher tous les messages trouvés
                    
            except CursorDBError as e:
                print(f"  ❌ Erreur pour {workspace}: {e}")
                continue
                
    except CursorDBError as e:
        print(f"❌ Erreur critique: {e}")
        return 1
    except Exception as e:
        print(f"❌ Erreur inattendue: {e}")
        return 1
        
    return 0

if __name__ == "__main__":
    exit(main()) 