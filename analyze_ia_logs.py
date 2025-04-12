def process_cursor_data(data):
    processed = []
    
    if isinstance(data, list):
        print(f"📝 Traitement d'une liste de {len(data)} éléments")
        for i, item in enumerate(data, 1):
            if isinstance(item, dict):
                entry = {
                    'timestamp': datetime.now().isoformat(),  # Pas de timestamp dans les données
                    'runner': 'text',  # Type par défaut
                    'file': 'unknown',
                    'prompt': item.get('text', ''),
                    'type': item.get('commandType', 'other')
                }
                processed.append(entry)
    
    return processed

def load_prompt_logs(log_dir):
    logs = []
    cursor_dir = os.path.expandvars(log_dir)
    print(f"📂 Recherche des logs dans: {cursor_dir}")
    
    if not os.path.exists(cursor_dir):
        print(f"❌ Répertoire non trouvé: {cursor_dir}")
        return logs
        
    subdirs = [d for d in os.listdir(cursor_dir) if os.path.isdir(os.path.join(cursor_dir, d))]
    print(f"📂 Sous-répertoires trouvés: {len(subdirs)}")
    
    for subdir in subdirs:
        db_path = os.path.join(cursor_dir, subdir, "state.vscdb")
        if os.path.exists(db_path):
            print(f"🔍 Vérification de {db_path}")
            try:
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                
                # Afficher les tables disponibles
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
                tables = cursor.fetchall()
                print(f"📊 Tables dans la base de données: {tables}")
                
                # Afficher la structure de ItemTable
                cursor.execute("PRAGMA table_info(ItemTable);")
                structure = cursor.fetchall()
                print(f"📊 Structure de ItemTable: {structure}")
                
                # Rechercher les entrées pertinentes
                cursor.execute("SELECT key, value FROM ItemTable WHERE key IN ('aiService.prompts', 'workbench.panel.aichat.view.aichat.chatdata');")
                entries = cursor.fetchall()
                print(f"📊 Entrées trouvées dans {db_path}: {len(entries)}")
                
                for key, value in entries:
                    print(f"🔑 Clé: {key}")
                    try:
                        data = json.loads(value)
                        print(f"📝 Type de données: {type(data)}")
                        if isinstance(data, list):
                            print(f"📝 Nombre d'éléments: {len(data)}")
                            if data:
                                print(f"📝 Type du premier élément: {type(data[0])}")
                                if isinstance(data[0], dict):
                                    print(f"📝 Clés du premier élément: {list(data[0].keys())}")
                        processed_entries = process_cursor_data(data)
                        logs.extend(processed_entries)
                    except json.JSONDecodeError as e:
                        print(f"❌ Erreur de décodage JSON pour la clé {key}: {e}")
                    except Exception as e:
                        print(f"❌ Erreur lors du traitement des données pour la clé {key}: {e}")
                
                conn.close()
            except sqlite3.Error as e:
                print(f"❌ Erreur SQLite pour {db_path}: {e}")
            except Exception as e:
                print(f"❌ Erreur générale pour {db_path}: {e}")
    
    return logs 