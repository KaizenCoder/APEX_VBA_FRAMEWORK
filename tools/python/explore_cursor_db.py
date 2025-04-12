import sqlite3
import os
import json

DB_PATH = r"C:\Users\Pape\AppData\Roaming\Cursor\User\workspaceStorage\6cf1141d0fb451733055df2e0dee7b7b\state.vscdb"

def explore_db_structure():
    """Explorer la structure de la base de données SQLite."""
    if not os.path.exists(DB_PATH):
        print(f"❌ Base de données non trouvée: {DB_PATH}")
        return

    try:
        # Connexion en lecture seule
        conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
        cursor = conn.cursor()
        
        # Récupérer la liste des tables
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()
        
        print(f"📊 Tables dans la base de données:")
        for i, table in enumerate(tables, 1):
            print(f"  {i}. {table[0]}")
        
        # Explorer la structure des tables
        for table in tables:
            table_name = table[0]
            print(f"\n📋 Structure de la table '{table_name}':")
            cursor.execute(f"PRAGMA table_info({table_name});")
            columns = cursor.fetchall()
            for col in columns:
                print(f"  - {col[1]} ({col[2]})")
            
            # Compter le nombre d'enregistrements
            cursor.execute(f"SELECT COUNT(*) FROM {table_name};")
            count = cursor.fetchone()[0]
            print(f"  👉 Total enregistrements: {count}")
            
            # Échantillon de données (limité aux 3 premiers)
            if count > 0:
                print(f"\n📝 Échantillon de données (3 premiers enregistrements):")
                cursor.execute(f"SELECT * FROM {table_name} LIMIT 3;")
                rows = cursor.fetchall()
                for i, row in enumerate(rows, 1):
                    print(f"\n  📄 Enregistrement #{i}:")
                    for j, col in enumerate(columns):
                        col_name = col[1]
                        value = row[j]
                        
                        if col_name.lower() in ["key", "id", "name"]:
                            print(f"    - {col_name}: {value}")
                        elif "json" in col_name.lower() or col_name.lower() in ["value", "data"]:
                            # Tenter de décoder JSON
                            try:
                                if isinstance(value, bytes):
                                    value_str = value.decode('utf-8')
                                else:
                                    value_str = str(value)
                                
                                if value_str.strip().startswith('{') or value_str.strip().startswith('['):
                                    try:
                                        json_data = json.loads(value_str)
                                        json_preview = json.dumps(json_data, indent=2)[:200]
                                        print(f"    - {col_name}: {json_preview}...")
                                    except json.JSONDecodeError:
                                        print(f"    - {col_name}: [Données JSON invalides] {value_str[:100]}...")
                                else:
                                    print(f"    - {col_name}: {value_str[:100]}...")
                            except Exception as e:
                                print(f"    - {col_name}: [Erreur de décodage: {e}] {value}")
                        else:
                            # Limiter l'affichage des valeurs trop longues
                            if isinstance(value, str) and len(value) > 100:
                                print(f"    - {col_name}: {value[:100]}...")
                            else:
                                print(f"    - {col_name}: {value}")
        
        # Rechercher des clés spécifiques liées aux conversations Cursor
        print("\n🔍 Recherche des clés liées aux conversations Cursor:")
        search_keys = [
            "%chat%", 
            "%cursor%", 
            "%prompt%", 
            "%message%",
            "%ai%",
            "%conversation%"
        ]
        
        for pattern in search_keys:
            cursor.execute("SELECT COUNT(*) FROM ItemTable WHERE key LIKE ?;", (pattern,))
            count = cursor.fetchone()[0]
            if count > 0:
                print(f"  ✅ Trouvé {count} entrées avec clé contenant '{pattern[1:-1]}'")
                # Afficher quelques exemples
                cursor.execute("SELECT key FROM ItemTable WHERE key LIKE ? LIMIT 5;", (pattern,))
                keys = cursor.fetchall()
                for key in keys:
                    print(f"    - {key[0]}")
            else:
                print(f"  ❌ Aucune entrée avec clé contenant '{pattern[1:-1]}'")
                
    except sqlite3.Error as e:
        print(f"❌ Erreur SQLite: {e}")
    except Exception as e:
        print(f"❌ Erreur générale: {e}")
    finally:
        if 'conn' in locals():
            conn.close()

if __name__ == "__main__":
    print(f"🔍 Exploration de la base de données: {DB_PATH}")
    explore_db_structure() 