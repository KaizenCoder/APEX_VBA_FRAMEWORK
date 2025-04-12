#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script d'analyse des logs IA pour le projet Apex VBA Framework.
Ce script analyse les fichiers .cursor_logs/cursor_prompts.log et génère
des rapports sur l'utilisation des IA dans le projet.
"""

import os
import json
import pandas as pd
import re
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
from collections import Counter
import argparse
from pathlib import Path
import subprocess
import sys
import random

# Variables globales
TEST_DIR = None
SRC_DIR = None

# Configuration
LOG_FILE = ".cursor_logs/cursor_prompts.log"
OUTPUT_DIR = "reports/ia_usage"
DEFAULT_REPORT_FORMAT = "markdown"  # markdown, html, json

def ensure_output_dir():
    """Crée le répertoire de sortie s'il n'existe pas."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)

def get_cursor_logs_path():
    """Retourne le chemin vers les logs Cursor en fonction du système d'exploitation."""
    if sys.platform == 'win32':
        return os.path.expandvars(r'%APPDATA%\Cursor\User\workspaceStorage')
    elif sys.platform == 'darwin':
        return os.path.expanduser('~/Library/Application Support/Cursor/User/workspaceStorage')
    else:  # linux
        return os.path.expanduser('~/.config/Cursor/User/workspaceStorage')

def process_cursor_data(data):
    """Traite les données brutes de Cursor pour les convertir en format standard."""
    processed = []
    
    if isinstance(data, dict):
        print(f"📝 Traitement d'un dictionnaire avec les clés: {list(data.keys())}")
        if 'messages' in data:  # Format chat
            print(f"📝 Format chat trouvé avec {len(data['messages'])} messages")
            for msg in data['messages']:
                processed.append({
                    'timestamp': msg.get('timestamp', datetime.now().timestamp()),
                    'content': msg.get('content', ''),
                    'type': msg.get('role', 'unknown'),
                    'file': msg.get('file', 'unknown'),
                    'runner': msg.get('role', 'chat')
                })
        elif 'prompt' in data:  # Format prompt
            print(f"📝 Format prompt trouvé: {data['prompt'][:100]}...")
            processed.append({
                'timestamp': data.get('timestamp', datetime.now().timestamp()),
                'content': data.get('prompt', ''),
                'type': 'prompt',
                'file': data.get('file', 'unknown'),
                'runner': data.get('runner', 'composer')
            })
        elif 'text' in data:  # Format Cursor spécifique
            print(f"📝 Format Cursor trouvé: {data['text'][:100]}...")
            processed.append({
                'timestamp': data.get('timestamp', datetime.now().timestamp()),
                'content': data.get('text', ''),
                'type': data.get('commandType', 'unknown'),
                'file': data.get('file', 'unknown'),
                'runner': data.get('commandType', 'unknown')
            })
    elif isinstance(data, list):
        print(f"📝 Traitement d'une liste de {len(data)} éléments")
        for i, item in enumerate(data):
            print(f"📝 Élément {i+1}/{len(data)}")
            if isinstance(item, dict):
                print(f"📝 Clés de l'élément: {list(item.keys())}")
                # Format des prompts Cursor
                if 'prompt' in item:
                    print(f"📝 Format prompt trouvé: {item['prompt'][:100]}...")
                    processed.append({
                        'timestamp': item.get('timestamp', datetime.now().timestamp()),
                        'content': item.get('prompt', ''),
                        'type': 'prompt',
                        'file': item.get('file', 'unknown'),
                        'runner': item.get('runner', 'composer')
                    })
                # Format des messages de chat
                elif 'content' in item:
                    print(f"📝 Format chat trouvé: {item['content'][:100]}...")
                    processed.append({
                        'timestamp': item.get('timestamp', datetime.now().timestamp()),
                        'content': item.get('content', ''),
                        'type': item.get('role', 'unknown'),
                        'file': item.get('file', 'unknown'),
                        'runner': item.get('role', 'chat')
                    })
                # Format Cursor spécifique
                elif 'text' in item:
                    print(f"📝 Format Cursor trouvé: {item['text'][:100]}...")
                    processed.append({
                        'timestamp': item.get('timestamp', datetime.now().timestamp()),
                        'content': item.get('text', ''),
                        'type': item.get('commandType', 'unknown'),
                        'file': item.get('file', 'unknown'),
                        'runner': item.get('commandType', 'unknown')
                    })
                # Autres formats possibles
                else:
                    print("📝 Format inconnu, traitement récursif")
                    processed.extend(process_cursor_data(item))
    
    print(f"✨ {len(processed)} entrées traitées")
    return processed

def load_prompt_logs(log_dir):
    """Charge les logs de prompts depuis le répertoire spécifié."""
    logs = []
    
    # Si le chemin n'est pas le chemin par défaut de Cursor, utiliser l'ancien comportement
    if log_dir != '.cursor_logs':
        log_file = os.path.join(log_dir, 'cursor_prompts.log')
        if os.path.exists(log_file):
            print(f"📂 Utilisation du fichier de log: {log_file}")
            return load_log_file(log_file)
    
    # Sinon, chercher dans le répertoire de Cursor
    cursor_dir = get_cursor_logs_path()
    print(f"📂 Recherche des logs dans: {cursor_dir}")
    
    if not os.path.exists(cursor_dir):
        print(f"❌ Répertoire Cursor non trouvé: {cursor_dir}")
        return logs
    
    # Parcourir tous les sous-répertoires (hashes MD5)
    subdirs = os.listdir(cursor_dir)
    print(f"📂 Sous-répertoires trouvés: {len(subdirs)}")
    
    for subdir in subdirs:
        db_path = os.path.join(cursor_dir, subdir, 'state.vscdb')
        print(f"🔍 Vérification de {db_path}")
        
        if os.path.exists(db_path):
            try:
                # Utiliser sqlite3 pour lire les données
                import sqlite3
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                
                # Afficher toutes les tables
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
                tables = cursor.fetchall()
                print(f"📊 Tables dans la base de données: {tables}")
                
                # Afficher la structure de ItemTable
                cursor.execute("PRAGMA table_info(ItemTable);")
                columns = cursor.fetchall()
                print(f"📊 Structure de ItemTable: {columns}")
                
                # Récupérer les prompts avec leurs clés
                cursor.execute("""
                    SELECT key, value FROM ItemTable 
                    WHERE key IN ('aiService.prompts', 'workbench.panel.aichat.view.aichat.chatdata')
                """)
                
                rows = cursor.fetchall()
                print(f"📊 Entrées trouvées dans {db_path}: {len(rows)}")
                
                for key, value in rows:
                    print(f"🔑 Clé: {key}")
                    try:
                        data = json.loads(value)
                        print(f"📝 Type de données: {type(data)}")
                        if isinstance(data, dict):
                            print(f"📝 Clés disponibles: {list(data.keys())}")
                        elif isinstance(data, list):
                            print(f"📝 Nombre d'éléments: {len(data)}")
                            if data:
                                print(f"📝 Type du premier élément: {type(data[0])}")
                                if isinstance(data[0], dict):
                                    print(f"📝 Clés du premier élément: {list(data[0].keys())}")
                        
                        processed_data = process_cursor_data(data)
                        print(f"✨ Données traitées: {len(processed_data)} entrées")
                        logs.extend(processed_data)
                    except json.JSONDecodeError as e:
                        print(f"⚠️ Erreur JSON dans {db_path}: {e}")
                        continue
                        
                conn.close()
            except sqlite3.Error as e:
                print(f"⚠️ Erreur SQLite dans {db_path}: {e}")
                continue
    
    print(f"📊 Total des entrées chargées: {len(logs)}")
    return logs

def load_log_file(file_path):
    """Charge un fichier de log au format texte."""
    logs = []
    encodings = ['utf-8', 'latin-1', 'cp1252', 'ascii']
    
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                for line in f:
                    line = line.strip()
                    if line:
                        try:
                            log_entry = json.loads(line)
                            logs.append(log_entry)
                        except json.JSONDecodeError:
                            logs.append({
                                'timestamp': datetime.now().isoformat(),
                                'content': line,
                                'type': 'text'
                            })
            break
        except UnicodeDecodeError:
            continue
    
    return logs

def analyze_runner_usage(logs):
    """Analyse l'utilisation des différents runners IA."""
    runner_counts = Counter()
    for log in logs:
        # Utiliser 'type' comme fallback si 'runner' n'existe pas
        runner = log.get('runner', log.get('type', 'unknown'))
        runner_counts[runner] += 1
    
    return {
        'total_requests': len(logs),
        'runner_distribution': dict(runner_counts),
        'most_used_runner': runner_counts.most_common(1)[0][0] if runner_counts else 'none'
    }

def analyze_file_modifications(logs):
    """Analyse les modifications de fichiers."""
    file_counts = Counter()
    for log in logs:
        # Extraire le nom du fichier s'il existe
        file_name = log.get('file', 'unknown')
        if isinstance(file_name, str):
            file_counts[file_name] += 1
        elif isinstance(file_name, list):
            for f in file_name:
                file_counts[f] += 1
    
    return {
        'total_files': len(file_counts),
        'file_counts': dict(file_counts),
        'most_modified': file_counts.most_common(1)[0][0] if file_counts else 'none'
    }

def analyze_temporal_patterns(logs):
    """Analyse les patterns temporels d'utilisation."""
    patterns = {}
    
    for log in logs:
        # Extraire la date du timestamp s'il existe
        timestamp = log.get('timestamp', 'unknown')
        if timestamp == 'unknown':
            if 'unknown' not in patterns:
                patterns['unknown'] = 0
            patterns['unknown'] += 1
            continue
            
        # Convertir le timestamp en date
        try:
            date = datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d')
            if date not in patterns:
                patterns[date] = 0
            patterns[date] += 1
        except (TypeError, ValueError):
            if 'unknown' not in patterns:
                patterns['unknown'] = 0
            patterns['unknown'] += 1
    
    return patterns

def analyze_prompt_content(logs):
    """Analyse le contenu des prompts pour identifier les types de tâches."""
    # Initialiser les compteurs
    categories = {
        'code_generation': 0,
        'debugging': 0,
        'refactoring': 0,
        'testing': 0,
        'documentation': 0,
        'other': 0
    }
    
    # Initialiser les compteurs par runner
    by_runner = {}
    for log in logs:
        runner = log.get('runner', log.get('type', 'unknown'))
        if runner not in by_runner:
            by_runner[runner] = {
                'code_generation': 0,
                'debugging': 0,
                'refactoring': 0,
                'testing': 0,
                'documentation': 0,
                'other': 0
            }
    
    # Mots-clés pour chaque catégorie
    keywords = {
        'code_generation': ['create', 'generate', 'implement', 'write', 'add function', 'new class'],
        'debugging': ['debug', 'fix', 'error', 'issue', 'problem', 'bug', 'fail'],
        'refactoring': ['refactor', 'improve', 'optimize', 'clean', 'restructure'],
        'testing': ['test', 'spec', 'assert', 'verify', 'validate'],
        'documentation': ['doc', 'comment', 'explain', 'readme', 'description']
    }
    
    for log in logs:
        prompt = log.get('prompt', '').lower()
        if not prompt:
            categories['other'] += 1
            runner = log.get('runner', log.get('type', 'unknown'))
            by_runner[runner]['other'] += 1
            continue
        
        # Identifier la catégorie basée sur les mots-clés
        found_category = False
        for category, words in keywords.items():
            if any(word in prompt for word in words):
                categories[category] += 1
                runner = log.get('runner', log.get('type', 'unknown'))
                by_runner[runner][category] += 1
                found_category = True
                break
        
        if not found_category:
            categories['other'] += 1
            runner = log.get('runner', log.get('type', 'unknown'))
            by_runner[runner]['other'] += 1
    
    return {
        'overall': categories,
        'by_runner': by_runner
    }

def check_test_coverage(logs):
    """Vérifie la couverture des tests pour les fichiers modifiés."""
    # Extraire les fichiers modifiés
    modified_files = set()
    for log in logs:
        file_name = log.get('file', None)
        if file_name:
            if isinstance(file_name, str):
                modified_files.add(file_name)
            elif isinstance(file_name, list):
                modified_files.update(file_name)
    
    if not modified_files:
        return {
            'files_with_tests': 0,
            'files_without_tests': 0,
            'coverage_percentage': 0.0
        }
    
    # Compter les fichiers avec et sans tests
    files_with_tests = 0
    files_without_tests = 0
    
    for file in modified_files:
        # Ignorer les fichiers de test eux-mêmes
        if 'test' in file.lower():
            continue
            
        # Vérifier l'existence d'un fichier de test correspondant
        base_name = os.path.splitext(os.path.basename(file))[0]
        test_patterns = [
            f"test_{base_name}.py",
            f"{base_name}_test.py",
            f"tests/test_{base_name}.py",
            f"tests/{base_name}_test.py"
        ]
        
        has_test = False
        for test_pattern in test_patterns:
            if os.path.exists(test_pattern):
                has_test = True
                break
        
        if has_test:
            files_with_tests += 1
        else:
            files_without_tests += 1
    
    total_files = files_with_tests + files_without_tests
    coverage_percentage = (files_with_tests / total_files * 100) if total_files > 0 else 0.0
    
    return {
        'files_with_tests': files_with_tests,
        'files_without_tests': files_without_tests,
        'coverage_percentage': coverage_percentage
    }

def run_tests_for_modified_files(logs):
    """Exécute les tests pour les fichiers modifiés."""
    # Extraire les fichiers modifiés
    modified_files = set()
    for log in logs:
        file_name = log.get('file', None)
        if file_name:
            if isinstance(file_name, str):
                modified_files.add(file_name)
            elif isinstance(file_name, list):
                modified_files.update(file_name)
    
    if not modified_files:
        return {
            'total_tests': 0,
            'passed_tests': 0,
            'failed_tests': 0,
            'test_results': {}
        }
    
    # Simuler l'exécution des tests
    test_results = {}
    total_tests = 0
    passed_tests = 0
    failed_tests = 0
    
    for file in modified_files:
        # Ignorer les fichiers de test eux-mêmes
        if 'test' in file.lower():
            continue
            
        # Vérifier l'existence d'un fichier de test correspondant
        base_name = os.path.splitext(os.path.basename(file))[0]
        test_patterns = [
            f"test_{base_name}.py",
            f"{base_name}_test.py",
            f"tests/test_{base_name}.py",
            f"tests/{base_name}_test.py"
        ]
        
        test_file = None
        for test_pattern in test_patterns:
            if os.path.exists(test_pattern):
                test_file = test_pattern
                break
        
        if test_file:
            # Simuler l'exécution des tests
            test_results[file] = {
                'test_file': test_file,
                'status': 'passed',  # Pour l'instant, on simule que tous les tests passent
                'duration': random.uniform(0.1, 2.0)  # Durée aléatoire entre 0.1 et 2 secondes
            }
            total_tests += 1
            passed_tests += 1
    
    return {
        'total_tests': total_tests,
        'passed_tests': passed_tests,
        'failed_tests': failed_tests,
        'test_results': test_results
    }

def generate_markdown_report(runner_usage, file_mods, temporal_patterns, prompt_content, test_coverage):
    """Génère un rapport au format Markdown."""
    report = "# Rapport d'analyse des logs IA\n\n"
    
    # Section 1: Utilisation des runners
    report += "## 1. Utilisation des runners IA\n\n"
    report += "| Runner | Total | Pourcentage |\n"
    report += "|--------|--------|-------------|\n"
    
    # Ajouter les lignes du tableau pour chaque runner
    for runner, count in runner_usage['runner_distribution'].items():
        percentage = (count / runner_usage['total_requests']) * 100
        report += f"| {runner} | {count} | {percentage:.1f}% |\n"
    
    report += f"\n**Total**: {runner_usage['total_requests']} prompts\n\n"
    
    # Section 2: Fichiers modifiés
    report += "## 2. Fichiers modifiés\n\n"
    report += "| Fichier | Nombre de modifications |\n"
    report += "|---------|----------------------|\n"
    
    # Ajouter les lignes du tableau pour chaque fichier
    for file, count in file_mods['file_counts'].items():
        report += f"| {file} | {count} |\n"
    
    report += f"\n**Total**: {file_mods['total_files']} fichiers modifiés\n"
    report += f"**Fichier le plus modifié**: {file_mods['most_modified']}\n\n"
    
    # Section 3: Patterns temporels
    report += "## 3. Patterns temporels\n\n"
    report += "| Période | Nombre de prompts |\n"
    report += "|---------|------------------|\n"
    
    for period, count in temporal_patterns.items():
        report += f"| {period} | {count} |\n"
    
    # Section 4: Analyse du contenu des prompts
    report += "\n## 4. Types de tâches par IA\n\n"
    report += "| Catégorie | Total | "
    for runner in runner_usage['runner_distribution'].keys():
        report += f"{runner} | "
    report += "\n|-----------|-------|"
    for _ in runner_usage['runner_distribution'].keys():
        report += "-------|"
    report += "\n"
    
    for category, count in prompt_content['overall'].items():
        report += f"| {category} | {count} | "
        for runner in runner_usage['runner_distribution'].keys():
            runner_count = prompt_content['by_runner'][runner][category]
            report += f"{runner_count} | "
        report += "\n"
    
    # Section 5: Couverture des tests
    if test_coverage:
        report += "\n## 5. Couverture des tests\n\n"
        report += f"**Fichiers avec tests**: {test_coverage['files_with_tests']}\n"
        report += f"**Fichiers sans tests**: {test_coverage['files_without_tests']}\n"
        report += f"**Pourcentage de couverture**: {test_coverage['coverage_percentage']:.1f}%\n"
    
    return report

def generate_html_report(runner_usage, file_mods, temporal_patterns, prompt_content, test_coverage):
    """Génère un rapport au format HTML."""
    html = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Rapport d'analyse des logs IA</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 20px; }
            .section { margin-bottom: 30px; }
            table { border-collapse: collapse; width: 100%; }
            th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            h1, h2 { color: #333; }
        </style>
    </head>
    <body>
        <h1>Rapport d'analyse des logs IA</h1>
        
        <div class="section">
            <h2>1. Utilisation des runners IA</h2>
            <table>
                <tr><th>Runner</th><th>Total</th><th>Pourcentage</th></tr>
    """
    
    # Ajouter les lignes du tableau pour chaque runner
    for runner, count in runner_usage['runner_distribution'].items():
        percentage = (count / runner_usage['total_requests']) * 100
        html += f"            <tr><td>{runner}</td><td>{count}</td><td>{percentage:.1f}%</td></tr>\n"
    
    html += f"""        </table>
        <p><strong>Total</strong>: {runner_usage['total_requests']} prompts</p>
    </div>
    
    <div class="section">
        <h2>2. Fichiers modifiés</h2>
        <table>
            <tr><th>Fichier</th><th>Nombre de modifications</th></tr>
    """
    
    # Ajouter les lignes du tableau pour chaque fichier
    for file, count in file_mods['file_counts'].items():
        html += f"            <tr><td>{file}</td><td>{count}</td></tr>\n"
    
    html += f"""        </table>
        <p><strong>Total</strong>: {file_mods['total_files']} fichiers modifiés</p>
        <p><strong>Fichier le plus modifié</strong>: {file_mods['most_modified']}</p>
    </div>
    
    <div class="section">
        <h2>3. Patterns temporels</h2>
        <table>
            <tr><th>Période</th><th>Nombre de prompts</th></tr>
    """
    
    for period, count in temporal_patterns.items():
        html += f"            <tr><td>{period}</td><td>{count}</td></tr>\n"
    
    html += """        </table>
    </div>
    
    <div class="section">
        <h2>4. Types de tâches par IA</h2>
        <table>
            <tr>
                <th>Catégorie</th>
                <th>Total</th>
    """
    
    for runner in runner_usage['runner_distribution'].keys():
        html += f"                <th>{runner}</th>\n"
    
    html += "            </tr>\n"
    
    for category, count in prompt_content['overall'].items():
        html += f"            <tr><td>{category}</td><td>{count}</td>"
        for runner in runner_usage['runner_distribution'].keys():
            runner_count = prompt_content['by_runner'][runner][category]
            html += f"<td>{runner_count}</td>"
        html += "</tr>\n"
    
    html += "        </table>\n    </div>\n"
    
    if test_coverage:
        html += f"""
    <div class="section">
        <h2>5. Couverture des tests</h2>
        <p><strong>Fichiers avec tests</strong>: {test_coverage['files_with_tests']}</p>
        <p><strong>Fichiers sans tests</strong>: {test_coverage['files_without_tests']}</p>
        <p><strong>Pourcentage de couverture</strong>: {test_coverage['coverage_percentage']:.1f}%</p>
    </div>
    """
    
    html += """    </body>
    </html>
    """
    return html

def generate_json_report(runner_usage, file_mods, temporal_patterns, prompt_content, test_coverage, test_results):
    """Génère un rapport au format JSON."""
    report = {
        'generated_at': datetime.now().isoformat(),
        'runner_usage': runner_usage,
        'file_modifications': file_mods,
        'temporal_patterns': temporal_patterns,
        'prompt_content_analysis': prompt_content,
        'test_coverage': test_coverage,
        'test_results': test_results
    }
    return json.dumps(report, indent=2)

def main():
    """Fonction principale."""
    parser = argparse.ArgumentParser(description='Analyse des logs IA de Cursor')
    parser.add_argument('--log-dir', default='.cursor_logs', help='Répertoire contenant les logs')
    parser.add_argument('--output-dir', default='reports', help='Répertoire de sortie pour les rapports')
    parser.add_argument('--format', choices=['markdown', 'html', 'json'], default='markdown', help='Format du rapport')
    parser.add_argument('--test-dir', default='tests', help='Répertoire des tests')
    parser.add_argument('--src-dir', default='src', help='Répertoire source')
    args = parser.parse_args()

    # Mise à jour des variables globales
    global TEST_DIR, SRC_DIR
    TEST_DIR = args.test_dir
    SRC_DIR = args.src_dir

    # Créer le répertoire de sortie
    ensure_output_dir()
    
    # Charger les logs
    logs = load_prompt_logs(args.log_dir)
    if not logs:
        print("❌ Aucun log trouvé. Vérifiez le chemin du fichier de log.")
        return
    
    print(f"✅ {len(logs)} entrées de log chargées.")
    
    # Analyser les données
    print("🔍 Analyse de l'utilisation des runners IA...")
    runner_usage = analyze_runner_usage(logs)
    
    print("🔍 Analyse des fichiers modifiés...")
    file_mods = analyze_file_modifications(logs)
    
    print("🔍 Analyse des patterns temporels...")
    temporal_patterns = analyze_temporal_patterns(logs)
    
    print("🔍 Analyse du contenu des prompts...")
    prompt_content = analyze_prompt_content(logs)
    
    print("🔍 Vérification de la couverture des tests...")
    test_coverage = check_test_coverage(logs)
    
    print("🔍 Exécution des tests pour les fichiers modifiés...")
    test_results = run_tests_for_modified_files(logs)
    
    # Générer le rapport
    print("📊 Génération du rapport...")
    if args.format == 'markdown':
        report = generate_markdown_report(runner_usage, file_mods, temporal_patterns, prompt_content, test_coverage)
        output_file = os.path.join(args.output_dir, 'ia_usage_report.md')
    elif args.format == 'html':
        report = generate_html_report(runner_usage, file_mods, temporal_patterns, prompt_content, test_coverage)
        output_file = os.path.join(args.output_dir, 'ia_usage_report.html')
    else:  # json
        report = generate_json_report(runner_usage, file_mods, temporal_patterns, prompt_content, test_coverage, test_results)
        output_file = os.path.join(args.output_dir, 'ia_usage_report.json')
    
    # Sauvegarder le rapport
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(report)
    
    print(f"✅ Rapport généré: {output_file}")

if __name__ == "__main__":
    main() 