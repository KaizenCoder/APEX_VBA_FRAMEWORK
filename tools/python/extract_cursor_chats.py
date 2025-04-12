import sqlite3
import json
from datetime import datetime
from pathlib import Path
import os
import re
from collections import Counter

def get_cursor_db_path():
    """Obtient le chemin de la base de données Cursor."""
    base_path = Path(os.environ["APPDATA"]) / "Cursor" / "User" / "workspaceStorage"
    
    if not base_path.exists():
        print(f"Chemin non trouvé : {base_path}")
        return None
        
    # Chercher le fichier state.vscdb le plus récent
    db_files = list(base_path.rglob("state.vscdb"))
    if not db_files:
        print("Aucun fichier state.vscdb trouvé")
        return None
        
    # Trier par date de modification
    db_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    db_path = str(db_files[0])
    print(f"Base de données trouvée : {db_path}")
    return db_path

def extract_conversations(db_path):
    """Extrait les conversations de la base de données."""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Requêtes pour les différents types de données de conversation
    queries = [
        ("SELECT value FROM ItemTable WHERE key = 'aiService.prompts'", "prompts"),
        ("SELECT value FROM ItemTable WHERE key = 'aiService.generations'", "generations"),
        ("SELECT value FROM ItemTable WHERE key = 'workbench.panel.aichat.view.aichat.chatdata'", "chat_data")
    ]
    
    conversations = {}
    for query, key in queries:
        cursor.execute(query)
        result = cursor.fetchone()
        if result:
            try:
                conversations[key] = json.loads(result[0])
            except json.JSONDecodeError:
                print(f"Erreur de décodage JSON pour {key}")
                conversations[key] = None
    
    conn.close()
    return conversations

def extract_keywords(text, min_length=4):
    """Extrait les mots-clés significatifs d'un texte."""
    words = re.findall(r'\b\w+\b', text.lower())
    stopwords = {'dans', 'avec', 'pour', 'cette', 'mais', 'les', 'des', 'est', 'sont'}
    keywords = [w for w in words if len(w) >= min_length and w not in stopwords]
    return Counter(keywords).most_common(5)

def detect_theme(text):
    """Détecte le thème d'une conversation basé sur des mots-clés."""
    themes = {
        "Installation": ["install", "setup", "configuration", "pip", "npm", "node", "python"],
        "Documentation": ["doc", "readme", "markdown", "documentation"],
        "Tests": ["test", "unittest", "validation", "vérification"],
        "Développement": ["code", "développement", "feature", "fonction"],
        "CI/CD": ["ci", "cd", "pipeline", "automation", "github"],
        "Erreurs": ["error", "erreur", "bug", "problème", "issue"],
        "Configuration": ["config", "setting", "paramètre", "variable"]
    }
    
    text = text.lower()
    theme_scores = {theme: 0 for theme in themes}
    
    for theme, keywords in themes.items():
        for keyword in keywords:
            if keyword.lower() in text:
                theme_scores[theme] += 1
                
    if not any(theme_scores.values()):
        return "Divers"
        
    return max(theme_scores.items(), key=lambda x: x[1])[0]

def format_code_block(text):
    """Formate un bloc de code avec la syntaxe appropriée."""
    if any(cmd in text.lower() for cmd in ["ps ", "npm", "pip", "python", "> "]):
        return f"```powershell\n{text}\n```"
    elif ".py" in text or "import" in text:
        return f"```python\n{text}\n```"
    elif ".json" in text or "{" in text:
        return f"```json\n{text}\n```"
    else:
        return f"```\n{text}\n```"

def format_conversation(conversations):
    """Formate les conversations en Markdown avec une meilleure structure."""
    output = "# Conversations Cursor\n\n"
    output += f"Exporté le : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
    
    # Initialiser les conversations par thème
    themed_conversations = {}
    theme_stats = {}
    total_messages = 0
    
    # Traiter les prompts et générations ensemble
    if conversations.get("prompts") and conversations.get("generations"):
        prompts = conversations["prompts"]
        generations = conversations.get("generations", [])
        
        for i, prompt in enumerate(prompts):
            if not isinstance(prompt, dict):
                continue
                
            text = prompt.get("text", "").strip()
            if not text:
                continue
                
            total_messages += 1
            
            # Détecter le thème
            theme = detect_theme(text)
            if theme not in themed_conversations:
                themed_conversations[theme] = []
                theme_stats[theme] = {
                    "messages": 0,
                    "first_timestamp": None,
                    "last_timestamp": None,
                    "keywords": Counter()
                }
            
            # Mettre à jour les statistiques
            theme_stats[theme]["messages"] += 1
            timestamp = datetime.fromtimestamp(prompt.get("unixMs", 0)/1000) if "unixMs" in prompt else None
            if timestamp:
                if not theme_stats[theme]["first_timestamp"] or timestamp < theme_stats[theme]["first_timestamp"]:
                    theme_stats[theme]["first_timestamp"] = timestamp
                if not theme_stats[theme]["last_timestamp"] or timestamp > theme_stats[theme]["last_timestamp"]:
                    theme_stats[theme]["last_timestamp"] = timestamp
            
            # Extraire les mots-clés
            keywords = extract_keywords(text)
            for word, count in keywords:
                theme_stats[theme]["keywords"][word] += count
            
            # Formater le message de l'utilisateur
            msg = f"**User**: {text}\n\n"
            
            # Ajouter la réponse correspondante si disponible
            if i < len(generations):
                gen = generations[i]
                if isinstance(gen, dict):
                    response = gen.get("text", "").strip()
                    if response:
                        # Détecter et formater les blocs de code
                        if any(cmd in response.lower() for cmd in ["ps ", "npm", "pip", "python", "> "]):
                            response = format_code_block(response)
                        msg += f"**Assistant**: {response}\n\n"
                        total_messages += 1
                        theme_stats[theme]["messages"] += 1
            
            themed_conversations[theme].append(msg)
    
    # Ajouter les statistiques globales
    output += "## 📊 Statistiques\n\n"
    output += f"- Messages totaux : {total_messages}\n"
    output += "- Répartition par thème :\n"
    for theme, stats in theme_stats.items():
        output += f"  - {theme} : {stats['messages']} messages\n"
    output += "\n"
    
    # Générer la table des matières
    output += "## 📑 Table des matières\n\n"
    for theme in themed_conversations:
        output += f"- [{theme}](#{theme.lower()})\n"
    output += "\n"
    
    # Ajouter les conversations par thème
    for theme, messages in themed_conversations.items():
        output += f"## {theme}\n\n"
        
        # Ajouter les métadonnées de la section
        stats = theme_stats[theme]
        output += "### 📌 Métadonnées\n\n"
        output += f"- Messages : {stats['messages']}\n"
        if stats['first_timestamp']:
            output += f"- Premier message : {stats['first_timestamp'].strftime('%Y-%m-%d %H:%M:%S')}\n"
            output += f"- Dernier message : {stats['last_timestamp'].strftime('%Y-%m-%d %H:%M:%S')}\n"
        output += "- Mots-clés principaux : " + ", ".join(f"`{k}`" for k, _ in stats['keywords'].most_common(5)) + "\n\n"
        
        output += "### 💬 Conversations\n\n"
        output += "".join(messages)
        output += "---\n\n"
    
    return output

def main():
    try:
        # Créer le dossier d'export s'il n'existe pas
        export_dir = Path("D:/Dev/Apex_VBA_FRAMEWORK/cursor_exports")
        export_dir.mkdir(exist_ok=True)
        
        # Obtenir le chemin de la base de données
        db_path = get_cursor_db_path()
        print(f"Base de données trouvée : {db_path}")
        
        # Extraire les conversations
        conversations = extract_conversations(db_path)
        print("Conversations extraites avec succès")
        
        # Formater et sauvegarder
        formatted_content = format_conversation(conversations)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = export_dir / f"cursor_chat_{timestamp}.md"
        
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(formatted_content)
        
        print(f"Conversations exportées dans : {output_file}")
        
        # Sauvegarder aussi les données brutes en JSON
        json_file = export_dir / f"cursor_chat_raw_{timestamp}.json"
        with open(json_file, "w", encoding="utf-8") as f:
            json.dump(conversations, f, ensure_ascii=False, indent=2)
        
        print(f"Données brutes sauvegardées dans : {json_file}")
        
    except Exception as e:
        print(f"Erreur : {str(e)}")

if __name__ == "__main__":
    main() 