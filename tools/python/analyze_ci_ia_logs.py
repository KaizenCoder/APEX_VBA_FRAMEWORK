import sqlite3
import os
import json
import re
import pandas as pd
from datetime import datetime
import argparse
import hashlib
import shutil
from pathlib import Path
import matplotlib.pyplot as plt
import seaborn as sns

# Constantes
VBA_KEYWORDS = ['vba', 'excel', 'access', 'sub', 'function', 'range', 'cells', 'workbook', 'worksheet', 'recordset', 'module', 'macro']
APEX_KEYWORDS = ['apex', 'framework', 'apex framework', 'test', 'ci', 'cd', 'pipeline']
ERROR_KEYWORDS = ['error', 'bug', 'fail', "doesn't work", 'ne fonctionne pas', 'erreur', 'debug', 'fix', 'problème', 'issue']
TEST_KEYWORDS = ['test', 'testing', 'unit test', 'validation', 'ci', 'automated test']
CI_KEYWORDS = ['ci', 'cd', 'pipeline', 'build', 'deploy', 'integration', 'continuous', 'automation', 'workflow']

class CursorLogAnalyzer:
    """Analyseur des logs de conversation IA pour Apex Framework CI/CD."""
    
    def __init__(self, db_path=None, output_dir="reports", search_paths=None):
        self.db_path = db_path
        self.output_dir = output_dir
        self.prompts = []
        self.analysis_results = {}
        
        # Si aucun chemin de base de données n'est fourni, rechercher aux emplacements standard
        if not self.db_path:
            self.db_path = self.find_cursor_db(search_paths)
            
        # Créer le répertoire de sortie s'il n'existe pas
        os.makedirs(self.output_dir, exist_ok=True)
    
    def find_cursor_db(self, search_paths=None):
        """Trouver la base de données Cursor aux emplacements standard."""
        if not search_paths:
            search_paths = []
            
            # Ajout des emplacements standard
            appdata = os.getenv('APPDATA', '')
            if appdata:
                # Chercher dans tous les workspaces
                cursor_dir = os.path.join(appdata, 'Cursor', 'User', 'workspaceStorage')
                if os.path.exists(cursor_dir):
                    for workspace in os.listdir(cursor_dir):
                        workspace_path = os.path.join(cursor_dir, workspace)
                        if os.path.isdir(workspace_path):
                            db_path = os.path.join(workspace_path, 'state.vscdb')
                            if os.path.exists(db_path):
                                search_paths.append(db_path)
        
        # Si des chemins ont été trouvés, retourner le premier valide
        for path in search_paths:
            if os.path.exists(path):
                print(f"✅ Base de données trouvée: {path}")
                return path
                
        print("❌ Aucune base de données Cursor trouvée!")
        return None
    
    def extract_prompts(self):
        """Extraire les prompts de la base de données SQLite."""
        if not self.db_path or not os.path.exists(self.db_path):
            print(f"❌ Base de données non trouvée: {self.db_path}")
            return False
            
        # Créer une copie temporaire de la base pour éviter les problèmes de verrouillage
        temp_db = os.path.join(self.output_dir, "temp_cursor_db.sqlite")
        try:
            shutil.copy2(self.db_path, temp_db)
            print(f"✅ Copie temporaire créée: {temp_db}")
        except Exception as e:
            print(f"⚠️ Impossible de créer une copie temporaire: {e}")
            temp_db = self.db_path
            
        try:
            # Connexion en lecture seule
            uri = f"file:{temp_db}?mode=ro" if temp_db == self.db_path else f"file:{temp_db}"
            conn = sqlite3.connect(uri, uri=True)
            cursor = conn.cursor()
            
            # Récupérer les prompts depuis aiService.prompts
            print("🔍 Extraction des prompts depuis aiService.prompts...")
            cursor.execute("SELECT value FROM ItemTable WHERE key = 'aiService.prompts';")
            row = cursor.fetchone()
            
            if not row:
                print("❌ Aucun prompt trouvé dans aiService.prompts!")
                return False
                
            # Décoder les données JSON
            if isinstance(row[0], bytes):
                value_str = row[0].decode('utf-8')
            else:
                value_str = str(row[0])
            
            prompts_data = json.loads(value_str)
            
            # Vérifier que nous avons une liste de prompts
            if not isinstance(prompts_data, list):
                print(f"❌ Format inattendu dans aiService.prompts: {type(prompts_data)}")
                return False
                
            print(f"✅ {len(prompts_data)} prompts trouvés.")
            
            # Récupérer les métadonnées des générations
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
            
            # Construire la liste des prompts avec contexte
            prompts = []
            
            for i, prompt in enumerate(prompts_data):
                if not isinstance(prompt, dict) or 'text' not in prompt:
                    continue
                    
                # Créer l'objet de prompt
                prompt_obj = {
                    'text': prompt.get('text', ''),
                    'index': i,
                    'command_type': prompt.get('commandType', 'unknown'),
                    'uuid': prompt.get('generationUUID', ''),
                    'timestamp': None,
                    'timestamp_str': 'Inconnu',
                    'type': self.categorize_prompt(prompt.get('text', '')),
                    'has_ci_keywords': any(keyword in prompt.get('text', '').lower() for keyword in CI_KEYWORDS)
                }
                
                # Ajouter le timestamp si disponible
                if prompt_obj['uuid'] and prompt_obj['uuid'] in generations:
                    gen_info = generations[prompt_obj['uuid']]
                    prompt_obj['timestamp'] = gen_info['timestamp']
                    # Convertir le timestamp en chaîne lisible
                    if isinstance(prompt_obj['timestamp'], (int, float)):
                        dt = datetime.fromtimestamp(prompt_obj['timestamp']/1000)
                        prompt_obj['timestamp_str'] = dt.strftime('%Y-%m-%d %H:%M:%S')
                
                # Déterminer le rôle (utilisateur ou IA)
                # Dans Cursor, commandType=4 semble être utilisé pour les messages
                prompt_obj['role'] = 'user' if prompt.get('commandType') == 4 else 'system'
                
                prompts.append(prompt_obj)
            
            self.prompts = prompts
            print(f"✅ {len(prompts)} prompts traités.")
            return True
            
        except sqlite3.Error as e:
            print(f"❌ Erreur SQLite: {e}")
            return False
        except json.JSONDecodeError as e:
            print(f"❌ Erreur décodage JSON: {e}")
            return False
        except Exception as e:
            print(f"❌ Erreur générale: {e}")
            return False
        finally:
            if 'conn' in locals():
                conn.close()
            
            # Supprimer la copie temporaire si elle existe
            if temp_db != self.db_path and os.path.exists(temp_db):
                try:
                    os.remove(temp_db)
                except:
                    pass
    
    def categorize_prompt(self, text):
        """Catégoriser le prompt en fonction de son contenu."""
        text_lower = text.lower()
        
        # Détection des types de prompts
        if any(keyword in text_lower for keyword in TEST_KEYWORDS):
            return 'test'
        elif any(keyword in text_lower for keyword in CI_KEYWORDS):
            return 'ci_cd'
        elif any(keyword in text_lower for keyword in ERROR_KEYWORDS):
            return 'debug'
        elif any(keyword in text_lower for keyword in VBA_KEYWORDS):
            return 'vba'
        elif any(keyword in text_lower for keyword in APEX_KEYWORDS):
            return 'apex'
        else:
            return 'other'
    
    def analyze_prompts(self):
        """Analyser les prompts pour extraire des informations utiles pour la CI."""
        if not self.prompts:
            print("❌ Aucun prompt à analyser!")
            return False
            
        # Convertir en DataFrame pour faciliter l'analyse
        df = pd.DataFrame(self.prompts)
        
        # Trier par timestamp si disponible
        if 'timestamp' in df and df['timestamp'].notna().any():
            df = df.sort_values('timestamp', ascending=True)
        
        # 1. Analyser la distribution des types de prompts
        prompt_types = df['type'].value_counts().to_dict()
        
        # 2. Détecter les problèmes mentionnés (focus sur CI/CD)
        problems = []
        for _, row in df.iterrows():
            text = row['text'].lower()
            
            # Vérifier si le prompt mentionne un problème lié à CI/CD
            if any(keyword in text for keyword in ERROR_KEYWORDS) and (
               row['type'] in ['ci_cd', 'test'] or row['has_ci_keywords']):
                # Extraire un contexte autour du problème
                for err_keyword in ERROR_KEYWORDS:
                    if err_keyword in text:
                        # Trouver la phrase contenant le mot-clé d'erreur
                        sentences = re.split(r'[.!?]', row['text'])
                        for sentence in sentences:
                            if err_keyword in sentence.lower():
                                # Ajouter le problème avec contexte
                                problems.append({
                                    'text': sentence.strip(),
                                    'keyword': err_keyword,
                                    'timestamp': row['timestamp_str'],
                                    'type': row['type']
                                })
                                break
        
        # 3. Analyser la chronologie des interactions CI/CD
        ci_timeline = df[df['has_ci_keywords']].copy()
        if not ci_timeline.empty and 'timestamp' in ci_timeline and ci_timeline['timestamp'].notna().any():
            ci_timeline['date'] = pd.to_datetime(ci_timeline['timestamp'], unit='ms')
            ci_timeline = ci_timeline.set_index('date')
            ci_timeline = ci_timeline.sort_index()
        
        # 4. Détecter les tests et pipelines mentionnés
        tests_mentioned = []
        for _, row in df.iterrows():
            if row['type'] == 'test' or any(keyword in row['text'].lower() for keyword in TEST_KEYWORDS):
                # Extraire les mentions de tests
                test_matches = re.findall(r'test\s+([a-zA-Z0-9_]+)', row['text'], re.IGNORECASE)
                for match in test_matches:
                    tests_mentioned.append({
                        'name': match,
                        'context': row['text'][:100] + '...',
                        'timestamp': row['timestamp_str']
                    })
        
        # 5. Compter les mots-clés CI/CD mentionnés
        ci_keywords_count = {keyword: 0 for keyword in CI_KEYWORDS}
        for _, row in df.iterrows():
            text = row['text'].lower()
            for keyword in CI_KEYWORDS:
                if keyword in text:
                    ci_keywords_count[keyword] += 1
        
        # Trier les mots-clés par fréquence
        ci_keywords_count = dict(sorted(ci_keywords_count.items(), key=lambda x: x[1], reverse=True))
        
        # Stocker les résultats
        self.analysis_results = {
            'prompt_types': prompt_types,
            'problems': problems,
            'ci_timeline': ci_timeline,
            'tests_mentioned': tests_mentioned,
            'ci_keywords_count': ci_keywords_count,
            'total_prompts': len(df),
            'ci_related_prompts': df['has_ci_keywords'].sum(),
            'prompts_df': df
        }
        
        print(f"✅ Analyse terminée. {len(problems)} problèmes détectés.")
        return True
    
    def generate_visualizations(self):
        """Générer des visualisations pour l'analyse."""
        if not self.analysis_results:
            return False
            
        viz_dir = os.path.join(self.output_dir, 'ia_usage')
        os.makedirs(viz_dir, exist_ok=True)
        
        # 1. Distribution des types de prompts (pie chart)
        if self.analysis_results['prompt_types']:
            plt.figure(figsize=(10, 6))
            plt.pie(
                self.analysis_results['prompt_types'].values(), 
                labels=self.analysis_results['prompt_types'].keys(),
                autopct='%1.1f%%',
                startangle=90
            )
            plt.title('Distribution des Types de Prompts')
            plt.axis('equal')
            plt.tight_layout()
            plt.savefig(os.path.join(viz_dir, 'prompt_types_distribution.png'))
            plt.close()
        
        # 2. Chronologie des interactions CI/CD
        if not self.analysis_results['ci_timeline'].empty and 'type' in self.analysis_results['ci_timeline'].columns:
            # Compter par jour et type
            try:
                daily_counts = self.analysis_results['ci_timeline'].resample('D')['type'].value_counts().unstack().fillna(0)
                plt.figure(figsize=(12, 6))
                daily_counts.plot(kind='bar', stacked=True)
                plt.title('Interactions CI/CD par Jour')
                plt.xlabel('Date')
                plt.ylabel('Nombre de Prompts')
                plt.tight_layout()
                plt.savefig(os.path.join(viz_dir, 'ci_interactions_timeline.png'))
                plt.close()
            except:
                pass
        
        # 3. Mots-clés CI/CD les plus fréquents
        plt.figure(figsize=(10, 6))
        keywords = list(self.analysis_results['ci_keywords_count'].keys())
        counts = list(self.analysis_results['ci_keywords_count'].values())
        # Filtrer les mots-clés avec des occurrences
        keywords = [k for i, k in enumerate(keywords) if counts[i] > 0]
        counts = [c for c in counts if c > 0]
        
        if keywords and counts:
            plt.bar(keywords, counts)
            plt.title('Mots-clés CI/CD les Plus Fréquents')
            plt.xlabel('Mot-clé')
            plt.ylabel('Nombre d\'occurrences')
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            plt.savefig(os.path.join(viz_dir, 'ci_keywords_frequency.png'))
            plt.close()
        
        return True
    
    def generate_ci_report(self, format='markdown'):
        """Générer un rapport d'analyse orienté CI."""
        if not self.analysis_results:
            print("❌ Aucun résultat d'analyse à inclure dans le rapport!")
            return False
            
        if format.lower() == 'markdown':
            return self.generate_markdown_report()
        elif format.lower() == 'html':
            return self.generate_html_report()
        else:
            print(f"⚠️ Format non supporté: {format}")
            return self.generate_markdown_report()
    
    def generate_markdown_report(self):
        """Générer un rapport en format Markdown."""
        report_file = os.path.join(self.output_dir, 'ia_usage_report.md')
        
        try:
            with open(report_file, 'w', encoding='utf-8') as f:
                f.write(f"""# Rapport d'Analyse des Logs IA pour CI/CD Apex Framework

## Résumé

* **Total de prompts analysés:** {self.analysis_results['total_prompts']}
* **Prompts liés à CI/CD:** {self.analysis_results['ci_related_prompts']} ({self.analysis_results['ci_related_prompts']/self.analysis_results['total_prompts']*100:.1f}%)
* **Problèmes détectés:** {len(self.analysis_results['problems'])}
* **Tests mentionnés:** {len(self.analysis_results['tests_mentioned'])}

## Distribution des Types de Prompts

""")
                # Distribution des types
                for prompt_type, count in self.analysis_results['prompt_types'].items():
                    percentage = count / self.analysis_results['total_prompts'] * 100
                    f.write(f"* **{prompt_type}:** {count} ({percentage:.1f}%)\n")
                
                f.write("""
## Problèmes CI/CD Détectés

Les problèmes suivants ont été identifiés dans les conversations liées à CI/CD:

""")
                # Problèmes détectés
                if self.analysis_results['problems']:
                    for i, problem in enumerate(self.analysis_results['problems'], 1):
                        f.write(f"{i}. **Problème ({problem['timestamp']}, Type: {problem['type']})**: \"{problem['text']}\"\n")
                else:
                    f.write("Aucun problème spécifique à CI/CD n'a été détecté.\n")
                
                f.write("""
## Tests Mentionnés

Les tests suivants ont été mentionnés dans les conversations:

""")
                # Tests mentionnés
                if self.analysis_results['tests_mentioned']:
                    for i, test in enumerate(self.analysis_results['tests_mentioned'], 1):
                        f.write(f"{i}. **{test['name']}** ({test['timestamp']}): \"{test['context']}\"\n")
                else:
                    f.write("Aucun test spécifique n'a été mentionné explicitement.\n")
                
                f.write("""
## Mots-clés CI/CD les Plus Fréquents

""")
                # Mots-clés CI/CD
                for keyword, count in self.analysis_results['ci_keywords_count'].items():
                    if count > 0:
                        f.write(f"* **{keyword}**: {count} mentions\n")
                
                f.write("""
## Suggestions pour l'Amélioration de CI/CD

Basé sur l'analyse des conversations, les améliorations suivantes sont recommandées:

""")
                # Générer des suggestions basées sur les problèmes
                suggestions = self.generate_ci_suggestions()
                if suggestions:
                    for i, suggestion in enumerate(suggestions, 1):
                        f.write(f"{i}. {suggestion}\n")
                else:
                    f.write("Aucune suggestion automatique n'a pu être générée à partir des données disponibles.\n")
                
                f.write("""
## Conclusion

Ce rapport a été généré automatiquement à partir de l'analyse des logs d'IA pour soutenir le développement CI/CD d'Apex Framework.

""")
                
                # Ajouter des liens vers les visualisations
                viz_dir = os.path.join('ia_usage')
                f.write("""
## Visualisations

""")
                for viz_file in ['prompt_types_distribution.png', 'ci_interactions_timeline.png', 'ci_keywords_frequency.png']:
                    file_path = os.path.join(viz_dir, viz_file)
                    if os.path.exists(os.path.join(self.output_dir, file_path)):
                        f.write(f"![{viz_file}]({file_path})\n\n")
            
            print(f"✅ Rapport généré: {report_file}")
            return True
                
        except Exception as e:
            print(f"❌ Erreur lors de la génération du rapport: {e}")
            return False
    
    def generate_html_report(self):
        """Générer un rapport en format HTML."""
        # Simplification: utiliser pandas pour générer un rapport HTML basique
        report_file = os.path.join(self.output_dir, 'ia_usage_report.html')
        
        try:
            # Créer un HTML de base
            html_content = f"""<!DOCTYPE html>
<html>
<head>
    <title>Rapport d'Analyse IA pour CI/CD</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        h1, h2, h3 {{ color: #333; }}
        .container {{ max-width: 1200px; margin: 0 auto; }}
        .chart {{ margin: 20px 0; }}
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; }}
        th {{ background-color: #f2f2f2; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>Rapport d'Analyse des Logs IA pour CI/CD Apex Framework</h1>
        
        <h2>Résumé</h2>
        <ul>
            <li><strong>Total de prompts analysés:</strong> {self.analysis_results['total_prompts']}</li>
            <li><strong>Prompts liés à CI/CD:</strong> {self.analysis_results['ci_related_prompts']} ({self.analysis_results['ci_related_prompts']/self.analysis_results['total_prompts']*100:.1f}%)</li>
            <li><strong>Problèmes détectés:</strong> {len(self.analysis_results['problems'])}</li>
            <li><strong>Tests mentionnés:</strong> {len(self.analysis_results['tests_mentioned'])}</li>
        </ul>
        
        <h2>Visualisations</h2>
"""
            
            # Ajouter des liens vers les visualisations
            viz_dir = 'ia_usage'
            for viz_file in ['prompt_types_distribution.png', 'ci_interactions_timeline.png', 'ci_keywords_frequency.png']:
                file_path = os.path.join(viz_dir, viz_file)
                if os.path.exists(os.path.join(self.output_dir, file_path)):
                    html_content += f"""
        <div class="chart">
            <h3>{viz_file.replace('_', ' ').replace('.png', '').title()}</h3>
            <img src="{file_path}" alt="{viz_file}" style="max-width: 100%;">
        </div>
"""
            
            # Ajouter des tableaux pour les problèmes et tests
            if self.analysis_results['problems']:
                html_content += """
        <h2>Problèmes CI/CD Détectés</h2>
        <table>
            <tr>
                <th>#</th>
                <th>Timestamp</th>
                <th>Type</th>
                <th>Description</th>
            </tr>
"""
                for i, problem in enumerate(self.analysis_results['problems'], 1):
                    html_content += f"""
            <tr>
                <td>{i}</td>
                <td>{problem['timestamp']}</td>
                <td>{problem['type']}</td>
                <td>{problem['text']}</td>
            </tr>"""
                
                html_content += """
        </table>
"""
            
            # Fermer le HTML
            html_content += """
    </div>
</body>
</html>
"""
            
            with open(report_file, 'w', encoding='utf-8') as f:
                f.write(html_content)
                
            print(f"✅ Rapport HTML généré: {report_file}")
            return True
            
        except Exception as e:
            print(f"❌ Erreur lors de la génération du rapport HTML: {e}")
            return False
    
    def generate_ci_suggestions(self):
        """Générer des suggestions pour l'amélioration de CI/CD basées sur l'analyse."""
        if not self.analysis_results:
            return []
            
        suggestions = []
        
        # 1. Suggestions basées sur les problèmes détectés
        if self.analysis_results['problems']:
            # Regrouper les problèmes par type
            problems_by_type = {}
            for problem in self.analysis_results['problems']:
                problem_type = problem['type']
                if problem_type not in problems_by_type:
                    problems_by_type[problem_type] = []
                problems_by_type[problem_type].append(problem)
            
            # Générer des suggestions spécifiques par type
            if 'ci_cd' in problems_by_type and len(problems_by_type['ci_cd']) > 0:
                suggestions.append("Améliorer les scripts de pipeline CI/CD pour résoudre les problèmes récurrents.")
                
            if 'test' in problems_by_type and len(problems_by_type['test']) > 0:
                suggestions.append("Renforcer la suite de tests automatisés pour couvrir les cas problématiques identifiés.")
        
        # 2. Suggestions basées sur l'utilisation
        if self.analysis_results['ci_related_prompts'] / self.analysis_results['total_prompts'] < 0.1:
            suggestions.append("Augmenter l'utilisation de CI/CD dans le processus de développement (moins de 10% actuellement).")
        
        # 3. Suggestions basées sur les tests
        if not self.analysis_results['tests_mentioned']:
            suggestions.append("Implémenter une stratégie de test automatisé intégrée à la CI/CD.")
        elif len(self.analysis_results['tests_mentioned']) < 5:
            suggestions.append("Étendre la couverture des tests automatisés pour couvrir plus de fonctionnalités.")
        
        # Ajouter des suggestions génériques si nécessaire
        if len(suggestions) < 3:
            suggestions.extend([
                "Mettre en place une validation automatique des prompts IA dans les pipelines de développement.",
                "Implémenter des métriques de qualité pour évaluer les interactions avec l'IA.",
                "Considérer l'intégration d'une étape d'analyse des logs IA dans le pipeline CI/CD."
            ])
        
        return suggestions[:5]  # Limiter à 5 suggestions

def main():
    parser = argparse.ArgumentParser(description='Analyser les logs de conversation IA pour CI/CD.')
    parser.add_argument('--db-path', help='Chemin vers la base de données SQLite de Cursor.')
    parser.add_argument('--output-dir', default='reports', help='Répertoire de sortie pour les rapports.')
    parser.add_argument('--format', choices=['markdown', 'html'], default='markdown', help='Format du rapport généré.')
    args = parser.parse_args()
    
    # Initialiser l'analyseur
    analyzer = CursorLogAnalyzer(db_path=args.db_path, output_dir=args.output_dir)
    
    # Extraire et analyser les prompts
    if analyzer.extract_prompts():
        if analyzer.analyze_prompts():
            # Générer des visualisations
            analyzer.generate_visualizations()
            
            # Générer le rapport
            if analyzer.generate_ci_report(format=args.format):
                print(f"✅ Analyse terminée avec succès! Rapport généré dans le répertoire: {args.output_dir}")
                return 0
    
    print("❌ L'analyse a échoué.")
    return 1

if __name__ == "__main__":
    exit(main()) 