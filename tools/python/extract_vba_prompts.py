import sqlite3
import os
import json
from datetime import datetime
import re
import pandas as pd

DB_PATH = r"C:\Users\Pape\AppData\Roaming\Cursor\User\workspaceStorage\6cf1141d0fb451733055df2e0dee7b7b\state.vscdb"
OUTPUT_FILE = "vba_prompts_analysis.md"

# Mots-clés pour filtrer les prompts pertinents
VBA_KEYWORDS = ['vba', 'excel', 'access', 'sub', 'function', 'range', 'cells', 'workbook', 'worksheet', 'recordset', 'module', 'macro']
APEX_KEYWORDS = ['apex', 'framework', 'apex framework', 'test', 'python', 'cursor', 'powershell']
ERROR_KEYWORDS = ['error', 'bug', 'fail', "doesn't work", 'ne fonctionne pas', 'erreur', 'debug', 'fix', 'problème', 'issue']

# Combiner tous les mots-clés pour le filtrage
RELEVANT_KEYWORDS = set(VBA_KEYWORDS + APEX_KEYWORDS)

def is_relevant_prompt(text):
    """Vérifie si un prompt est pertinent pour notre analyse VBA/Apex."""
    text_lower = text.lower()
    return any(keyword in text_lower for keyword in RELEVANT_KEYWORDS)

def extract_vba_prompts():
    """Extrait les prompts liés à VBA/Apex depuis la base de données."""
    if not os.path.exists(DB_PATH):
        print(f"❌ Base de données non trouvée: {DB_PATH}")
        return None

    try:
        # Connexion en lecture seule
        conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
        cursor = conn.cursor()
        
        # Récupérer les prompts depuis aiService.prompts
        print("🔍 Extraction des prompts depuis aiService.prompts...")
        cursor.execute("SELECT value FROM ItemTable WHERE key = 'aiService.prompts';")
        row = cursor.fetchone()
        
        if not row:
            print("❌ Aucun prompt trouvé!")
            return None
            
        # Décoder les données JSON
        if isinstance(row[0], bytes):
            value_str = row[0].decode('utf-8')
        else:
            value_str = str(row[0])
        
        data = json.loads(value_str)
        
        # Vérifier que nous avons une liste de prompts
        if not isinstance(data, list):
            print(f"❌ Format de données inattendu: {type(data)}")
            return None
            
        print(f"✅ {len(data)} prompts trouvés.")
        
        # Récupérer les générations pour les timestamps
        print("🔍 Récupération des métadonnées de génération...")
        cursor.execute("SELECT value FROM ItemTable WHERE key = 'aiService.generations';")
        gen_row = cursor.fetchone()
        
        generations = {}
        if gen_row:
            if isinstance(gen_row[0], bytes):
                gen_str = gen_row[0].decode('utf-8')
            else:
                gen_str = str(gen_row[0])
            
            gen_data = json.loads(gen_str)
            if isinstance(gen_data, list):
                for gen in gen_data:
                    if 'generationUUID' in gen and 'unixMs' in gen:
                        generations[gen['generationUUID']] = {
                            'timestamp': gen['unixMs'],
                            'type': gen.get('type', 'unknown'),
                            'description': gen.get('textDescription', '')
                        }
        
        # Extraire les prompts pertinents
        vba_prompts = []
        for i, prompt in enumerate(data):
            if not isinstance(prompt, dict) or 'text' not in prompt:
                continue
                
            text = prompt.get('text', '')
            if is_relevant_prompt(text):
                # Créer un objet pour le prompt
                prompt_obj = {
                    'text': text,
                    'index': i,
                    'command_type': prompt.get('commandType', 'unknown'),
                    'uuid': prompt.get('generationUUID', ''),
                    'timestamp': None,
                    'timestamp_str': 'Inconnu'
                }
                
                # Ajouter le timestamp si disponible
                if prompt_obj['uuid'] and prompt_obj['uuid'] in generations:
                    gen_info = generations[prompt_obj['uuid']]
                    prompt_obj['timestamp'] = gen_info['timestamp']
                    # Convertir le timestamp en chaîne lisible
                    if isinstance(prompt_obj['timestamp'], (int, float)):
                        dt = datetime.fromtimestamp(prompt_obj['timestamp']/1000)
                        prompt_obj['timestamp_str'] = dt.strftime('%Y-%m-%d %H:%M:%S')
                
                vba_prompts.append(prompt_obj)
        
        print(f"✅ {len(vba_prompts)} prompts pertinents pour VBA/Apex extraits.")
        
        return vba_prompts
        
    except sqlite3.Error as e:
        print(f"❌ Erreur SQLite: {e}")
        return None
    except Exception as e:
        print(f"❌ Erreur générale: {e}")
        return None
    finally:
        if 'conn' in locals():
            conn.close()

def analyze_prompts(prompts):
    """Analyse les prompts pour extraire des informations utiles."""
    if not prompts:
        return None
        
    # Convertir en DataFrame pour faciliter l'analyse
    df = pd.DataFrame(prompts)
    
    # Trier par timestamp si disponible
    if 'timestamp' in df and df['timestamp'].notna().any():
        df = df.sort_values('timestamp', ascending=False)
    
    # Détecter les problèmes mentionnés
    problems = []
    for _, row in df.iterrows():
        text = row['text'].lower()
        
        # Vérifier si le prompt mentionne un problème ou une erreur
        if any(keyword in text for keyword in ERROR_KEYWORDS):
            # Extraire un contexte autour du problème
            for err_keyword in ERROR_KEYWORDS:
                if err_keyword in text:
                    # Trouver la phrase contenant le mot-clé d'erreur
                    sentences = re.split(r'[.!?]', row['text'])
                    for sentence in sentences:
                        if err_keyword in sentence.lower():
                            # Ajouter le problème avec son contexte
                            problems.append({
                                'text': sentence.strip(),
                                'keyword': err_keyword,
                                'timestamp': row['timestamp_str']
                            })
                            break
    
    # Compter les mots-clés VBA et Apex mentionnés
    keyword_counts = {keyword: 0 for keyword in RELEVANT_KEYWORDS}
    for _, row in df.iterrows():
        text = row['text'].lower()
        for keyword in RELEVANT_KEYWORDS:
            if keyword in text:
                keyword_counts[keyword] += 1
    
    # Trier les mots-clés par fréquence
    sorted_keywords = sorted(keyword_counts.items(), key=lambda x: x[1], reverse=True)
    
    return {
        'prompts': df,
        'problems': problems,
        'keyword_counts': sorted_keywords
    }

def generate_report(analysis):
    """Génère un rapport Markdown des résultats de l'analyse."""
    if not analysis:
        return False
        
    prompts_df = analysis['prompts']
    problems = analysis['problems']
    keyword_counts = analysis['keyword_counts']
    
    content = """# Analyse des Prompts VBA/Apex Framework

## Résumé

"""
    content += f"* **Nombre total de prompts pertinents:** {len(prompts_df)}\n"
    content += f"* **Période d'analyse:** {prompts_df['timestamp_str'].min() if not prompts_df.empty and 'timestamp_str' in prompts_df else 'Inconnue'} à {prompts_df['timestamp_str'].max() if not prompts_df.empty and 'timestamp_str' in prompts_df else 'Inconnue'}\n\n"

    content += """## Mots-clés les plus fréquents

"""
    for keyword, count in keyword_counts:
        if count > 0:
            content += f"* **{keyword}**: {count} mentions\n"
    
    content += """
## Problèmes détectés

Les problèmes suivants ont été identifiés dans les conversations:

"""
    if problems:
        for i, problem in enumerate(problems, 1):
            content += f"{i}. **Problème ({problem['timestamp']})**: \"{problem['text']}\"\n"
    else:
        content += "Aucun problème spécifique n'a été détecté.\n"
    
    content += """
## Suggestions de tests

Basé sur les problèmes détectés, les tests suivants sont recommandés:

"""
    if problems:
        unique_tests = set()
        for problem in problems:
            # Créer une suggestion de test basée sur le problème
            test_text = f"Tester le cas où {problem['text'].lower().replace(problem['keyword'], '')}"
            # Nettoyer et normaliser le texte du test
            test_text = re.sub(r'\s+', ' ', test_text).strip()
            if test_text and len(test_text) > 15:  # Ignorer les tests trop courts
                unique_tests.add(test_text)
        
        for i, test in enumerate(unique_tests, 1):
            content += f"{i}. {test}\n"
    else:
        content += "Aucune suggestion de test spécifique basée sur des problèmes détectés.\n"
    
    content += """
## Extraits de prompts pertinents

Voici quelques exemples de prompts pertinents trouvés:

"""
    # Afficher quelques exemples de prompts (les plus récents)
    for i, (_, row) in enumerate(prompts_df.head(5).iterrows(), 1):
        content += f"### Prompt {i} ({row['timestamp_str']})\n\n"
        content += f"```\n{row['text']}\n```\n\n"
    
    try:
        with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
            f.write(content)
        return True
    except Exception as e:
        print(f"❌ Erreur lors de la génération du rapport: {e}")
        return False

if __name__ == "__main__":
    print(f"🔍 Extraction et analyse des prompts VBA/Apex depuis: {DB_PATH}")
    prompts = extract_vba_prompts()
    
    if prompts:
        print("📊 Analyse des prompts...")
        analysis_results = analyze_prompts(prompts)
        
        if analysis_results:
            print(f"📝 Génération du rapport dans: {OUTPUT_FILE}")
            if generate_report(analysis_results):
                print(f"✅ Rapport généré avec succès!")
            else:
                print("❌ Échec de la génération du rapport.")
        else:
            print("❌ Échec de l'analyse des prompts.")
    else:
        print("❌ Échec de l'extraction des prompts.") 