import sqlite3
import os
import json
from datetime import datetime

DB_PATH = r"C:\Users\Pape\AppData\Roaming\Cursor\User\workspaceStorage\6cf1141d0fb451733055df2e0dee7b7b\state.vscdb"

def format_json(value):
    """Formate une valeur JSON pour l'affichage."""
    try:
        if isinstance(value, bytes):
            value_str = value.decode('utf-8')
        else:
            value_str = str(value)
        
        if value_str.strip().startswith('{') or value_str.strip().startswith('['):
            json_data = json.loads(value_str)
            return json.dumps(json_data, indent=2)
        return value_str
    except Exception as e:
        return f"[Erreur de formatage: {e}] {value}"

def explore_prompts():
    """Explorer les prompts et conversations stockés dans la base de données."""
    if not os.path.exists(DB_PATH):
        print(f"❌ Base de données non trouvée: {DB_PATH}")
        return

    try:
        # Connexion en lecture seule
        conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
        cursor = conn.cursor()
        
        # Clés intéressantes à explorer
        interesting_keys = [
            "aiService.prompts",
            "aiService.generations",
            "workbench.panel.aichat.view.aichat.chatdata",
            "anysphere.cursor-retrieval"
        ]
        
        for key in interesting_keys:
            print(f"\n🔍 Recherche de la clé: {key}")
            cursor.execute("SELECT value FROM ItemTable WHERE key = ?;", (key,))
            row = cursor.fetchone()
            
            if row:
                print(f"✅ Clé trouvée!")
                value_json = format_json(row[0])
                
                # Analyser le JSON pour comprendre sa structure
                try:
                    if isinstance(row[0], bytes):
                        value_str = row[0].decode('utf-8')
                    else:
                        value_str = str(row[0])
                    
                    data = json.loads(value_str)
                    
                    if isinstance(data, list):
                        print(f"📊 Structure: Liste contenant {len(data)} éléments")
                        if len(data) > 0:
                            print(f"📌 Premier élément de type: {type(data[0]).__name__}")
                            
                            # Explorer le premier élément s'il s'agit d'un dictionnaire
                            if isinstance(data[0], dict):
                                print(f"🔑 Clés disponibles dans le premier élément: {', '.join(data[0].keys())}")
                                
                                # Vérifier s'il s'agit d'un message (chercher des indices)
                                if 'text' in data[0] or 'content' in data[0] or 'prompt' in data[0]:
                                    print("\n💬 Exemples de messages:")
                                    for i, item in enumerate(data[:5], 1):  # Limiter à 5 exemples
                                        # Déterminer le rôle du message (utilisateur ou IA)
                                        role = item.get('role', 'inconnu')
                                        if role == 'inconnu':
                                            # Tenter de déduire le rôle
                                            if 'user' in str(item).lower():
                                                role = 'utilisateur'
                                            elif 'assistant' in str(item).lower() or 'ai' in str(item).lower():
                                                role = 'assistant'
                                        
                                        # Récupérer le contenu du message
                                        text = item.get('text', item.get('content', item.get('prompt', '[Contenu non trouvé]')))
                                        
                                        # Récupérer le timestamp si disponible
                                        timestamp = item.get('timestamp', item.get('createdAt', item.get('date', None)))
                                        if timestamp and isinstance(timestamp, (int, float)):
                                            # Convertir timestamp (ms ou s) en datetime
                                            timestamp = datetime.fromtimestamp(timestamp/1000 if timestamp > 1e10 else timestamp)
                                            timestamp_str = timestamp.strftime('%Y-%m-%d %H:%M:%S')
                                        else:
                                            timestamp_str = 'Inconnu'
                                        
                                        print(f"\n📝 Message #{i}:")
                                        print(f"👤 Rôle: {role}")
                                        print(f"🕒 Timestamp: {timestamp_str}")
                                        print(f"💬 Contenu: {text[:200]}...")
                    
                    elif isinstance(data, dict):
                        print(f"📊 Structure: Dictionnaire contenant {len(data.keys())} clés")
                        print(f"🔑 Clés disponibles: {', '.join(data.keys())}")
                        
                        # Vérifier s'il y a une clé messages ou conversations
                        if 'messages' in data and isinstance(data['messages'], list):
                            messages = data['messages']
                            print(f"\n💬 Liste de messages trouvée ({len(messages)} messages)")
                            for i, msg in enumerate(messages[:5], 1):  # Limiter à 5 exemples
                                print(f"\n📝 Message #{i}:")
                                for key, value in msg.items():
                                    if isinstance(value, str) and len(value) > 100:
                                        print(f"  - {key}: {value[:100]}...")
                                    else:
                                        print(f"  - {key}: {value}")
                        
                        elif 'conversations' in data and isinstance(data['conversations'], list):
                            conversations = data['conversations']
                            print(f"\n💬 Liste de conversations trouvée ({len(conversations)} conversations)")
                            for i, conv in enumerate(conversations[:3], 1):  # Limiter à 3 exemples
                                print(f"\n📝 Conversation #{i}:")
                                for key, value in conv.items():
                                    if isinstance(value, str) and len(value) > 100:
                                        print(f"  - {key}: {value[:100]}...")
                                    else:
                                        print(f"  - {key}: {value}")
                    
                    print(f"\n📋 Aperçu des données (limité à 500 caractères):")
                    print(value_json[:500] + "..." if len(value_json) > 500 else value_json)
                    
                except json.JSONDecodeError:
                    print("❌ Impossible de décoder les données JSON")
                except Exception as e:
                    print(f"❌ Erreur lors de l'analyse: {e}")
            else:
                print(f"❌ Clé non trouvée!")
                
        # Recherche générique pour trouver d'autres clés potentiellement intéressantes
        print("\n🔎 Recherche d'autres clés potentiellement intéressantes...")
        patterns = ["%chat%", "%ai%"]
        for pattern in patterns:
            cursor.execute("SELECT key FROM ItemTable WHERE key LIKE ? AND key NOT IN (SELECT value FROM (SELECT value FROM ItemTable LIMIT 0));", (pattern,))
            keys = cursor.fetchall()
            if keys:
                print(f"\n📑 Clés contenant '{pattern[1:-1]}':")
                for key in keys:
                    print(f"  - {key[0]}")
                    
    except sqlite3.Error as e:
        print(f"❌ Erreur SQLite: {e}")
    except Exception as e:
        print(f"❌ Erreur générale: {e}")
    finally:
        if 'conn' in locals():
            conn.close()

if __name__ == "__main__":
    print(f"🔍 Exploration des prompts et conversations dans: {DB_PATH}")
    explore_prompts() 