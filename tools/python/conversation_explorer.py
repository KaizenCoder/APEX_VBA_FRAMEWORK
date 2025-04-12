"""
Conversation Explorer - Interface de consultation des logs IA pour Apex Framework
--------------------------------------------------------------------------------
Interface web qui permet de consulter facilement les conversations IA
avec indexation multi-critères (temporel, agent, sujet, etc.)
"""

import os
import re
import json
import sqlite3
import argparse
from datetime import datetime, timedelta
from pathlib import Path
import pandas as pd
import numpy as np
from collections import Counter, defaultdict
import hashlib
import logging

# Modules web
from flask import Flask, render_template, request, jsonify, abort, send_file, redirect, url_for
import plotly
import plotly.express as px
import plotly.graph_objects as go

# Analyse de sentiment et classification
from textblob import TextBlob
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize

# Configuration du logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("conversation_explorer.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("ConversationExplorer")

# Constantes
VBA_KEYWORDS = ['vba', 'excel', 'access', 'sub', 'function', 'range', 'cells', 'workbook', 'worksheet', 'recordset', 'module', 'macro']
APEX_KEYWORDS = ['apex', 'framework', 'apex framework', 'test', 'ci', 'cd', 'pipeline']
ERROR_KEYWORDS = ['error', 'bug', 'fail', "doesn't work", 'ne fonctionne pas', 'erreur', 'debug', 'fix', 'problème', 'issue']
TEST_KEYWORDS = ['test', 'testing', 'unit test', 'validation', 'ci', 'automated test']
CI_KEYWORDS = ['ci', 'cd', 'pipeline', 'build', 'deploy', 'integration', 'continuous', 'automation', 'workflow']

# Classes pour la gestion des données
class ConversationStore:
    """Gère le stockage et l'indexation des conversations."""
    
    def __init__(self, db_path='conversations.db'):
        """Initialise le stockage de conversations avec SQLite."""
        self.db_path = db_path
        self.ensure_db_exists()
        
    def ensure_db_exists(self):
        """Crée la base de données et les tables si elles n'existent pas."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Table des conversations
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS conversations (
            id TEXT PRIMARY KEY,
            title TEXT,
            timestamp INTEGER,
            date_str TEXT,
            agent TEXT,
            role TEXT,
            num_messages INTEGER,
            summary TEXT,
            sentiment REAL,
            word_count INTEGER,
            has_error INTEGER,
            has_code INTEGER
        )
        ''')
        
        # Table des messages
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS messages (
            id TEXT PRIMARY KEY,
            conversation_id TEXT,
            content TEXT,
            role TEXT,
            timestamp INTEGER,
            date_str TEXT,
            order_index INTEGER,
            sentiment REAL,
            word_count INTEGER,
            FOREIGN KEY (conversation_id) REFERENCES conversations(id)
        )
        ''')
        
        # Table des tags/thématiques
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS tags (
            conversation_id TEXT,
            tag TEXT,
            score REAL,
            FOREIGN KEY (conversation_id) REFERENCES conversations(id),
            PRIMARY KEY (conversation_id, tag)
        )
        ''')
        
        # Table pour stocker les métriques agrégées
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS metrics (
            key TEXT PRIMARY KEY,
            value REAL,
            description TEXT
        )
        ''')
        
        conn.commit()
        conn.close()
    
    def add_conversation(self, conversation_data):
        """Ajoute une conversation et ses messages dans la base de données."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Vérifier si la conversation existe déjà
        cursor.execute("SELECT id FROM conversations WHERE id = ?", (conversation_data['id'],))
        if cursor.fetchone():
            conn.close()
            return False  # Conversation déjà importée
        
        # Insérer la conversation
        cursor.execute('''
        INSERT INTO conversations (id, title, timestamp, date_str, agent, role, num_messages, summary, sentiment, word_count, has_error, has_code)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            conversation_data['id'],
            conversation_data.get('title', 'Sans titre'),
            conversation_data.get('timestamp', 0),
            conversation_data.get('date_str', ''),
            conversation_data.get('agent', 'unknown'),
            conversation_data.get('role', 'unknown'),
            len(conversation_data.get('messages', [])),
            conversation_data.get('summary', ''),
            conversation_data.get('sentiment', 0.0),
            conversation_data.get('word_count', 0),
            1 if conversation_data.get('has_error', False) else 0,
            1 if conversation_data.get('has_code', False) else 0
        ))
        
        # Insérer les messages
        for i, message in enumerate(conversation_data.get('messages', [])):
            message_id = hashlib.md5(f"{conversation_data['id']}_{i}".encode()).hexdigest()
            cursor.execute('''
            INSERT INTO messages (id, conversation_id, content, role, timestamp, date_str, order_index, sentiment, word_count)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                message_id,
                conversation_data['id'],
                message.get('content', ''),
                message.get('role', 'unknown'),
                message.get('timestamp', 0),
                message.get('date_str', ''),
                i,
                message.get('sentiment', 0.0),
                message.get('word_count', 0)
            ))
        
        # Insérer les tags
        for tag, score in conversation_data.get('tags', {}).items():
            cursor.execute('''
            INSERT INTO tags (conversation_id, tag, score)
            VALUES (?, ?, ?)
            ''', (conversation_data['id'], tag, score))
        
        conn.commit()
        conn.close()
        return True
    
    def get_conversation(self, conversation_id):
        """Récupère une conversation complète avec ses messages."""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        # Récupérer la conversation
        cursor.execute("SELECT * FROM conversations WHERE id = ?", (conversation_id,))
        conversation = cursor.fetchone()
        
        if not conversation:
            conn.close()
            return None
        
        # Convertir en dictionnaire
        conversation_dict = dict(conversation)
        
        # Récupérer les messages associés
        cursor.execute("SELECT * FROM messages WHERE conversation_id = ? ORDER BY order_index", (conversation_id,))
        messages = [dict(row) for row in cursor.fetchall()]
        conversation_dict['messages'] = messages
        
        # Récupérer les tags associés
        cursor.execute("SELECT tag, score FROM tags WHERE conversation_id = ?", (conversation_id,))
        tags = {row['tag']: row['score'] for row in cursor.fetchall()}
        conversation_dict['tags'] = tags
        
        conn.close()
        return conversation_dict
    
    def search_conversations(self, filters=None, sort_by='timestamp', sort_order='desc', page=1, per_page=20):
        """Recherche des conversations selon des filtres spécifiés."""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        query = "SELECT * FROM conversations"
        params = []
        
        # Construire la clause WHERE basée sur les filtres
        if filters:
            where_clauses = []
            
            if 'text' in filters:
                where_clauses.append("(title LIKE ? OR summary LIKE ?)")
                search_text = f"%{filters['text']}%"
                params.extend([search_text, search_text])
            
            if 'date_from' in filters:
                where_clauses.append("timestamp >= ?")
                params.append(filters['date_from'])
            
            if 'date_to' in filters:
                where_clauses.append("timestamp <= ?")
                params.append(filters['date_to'])
            
            if 'agent' in filters:
                where_clauses.append("agent = ?")
                params.append(filters['agent'])
            
            if 'has_error' in filters:
                where_clauses.append("has_error = ?")
                params.append(1 if filters['has_error'] else 0)
            
            if 'has_code' in filters:
                where_clauses.append("has_code = ?")
                params.append(1 if filters['has_code'] else 0)
            
            if 'tag' in filters:
                query = "SELECT c.* FROM conversations c JOIN tags t ON c.id = t.conversation_id"
                where_clauses.append("t.tag = ?")
                params.append(filters['tag'])
            
            if where_clauses:
                query += " WHERE " + " AND ".join(where_clauses)
        
        # Compter le nombre total de résultats pour la pagination
        count_query = query.replace("SELECT *", "SELECT COUNT(*)")
        cursor.execute(count_query, params)
        total_count = cursor.fetchone()[0]
        
        # Ajouter le tri et la pagination
        query += f" ORDER BY {sort_by} {sort_order}"
        query += f" LIMIT {per_page} OFFSET {(page - 1) * per_page}"
        
        cursor.execute(query, params)
        results = [dict(row) for row in cursor.fetchall()]
        
        conn.close()
        
        return {
            'conversations': results,
            'total': total_count,
            'page': page,
            'per_page': per_page,
            'pages': (total_count + per_page - 1) // per_page
        }
    
    def get_stats(self):
        """Récupère des statistiques sur les conversations stockées."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        stats = {}
        
        # Nombre total de conversations
        cursor.execute("SELECT COUNT(*) FROM conversations")
        stats['total_conversations'] = cursor.fetchone()[0]
        
        # Nombre total de messages
        cursor.execute("SELECT COUNT(*) FROM messages")
        stats['total_messages'] = cursor.fetchone()[0]
        
        # Répartition par agent
        cursor.execute("SELECT agent, COUNT(*) FROM conversations GROUP BY agent")
        stats['agents'] = dict(cursor.fetchall())
        
        # Répartition par tag
        cursor.execute("SELECT tag, COUNT(*) FROM tags GROUP BY tag")
        stats['tags'] = dict(cursor.fetchall())
        
        # Conversations par mois
        cursor.execute("""
        SELECT 
            strftime('%Y-%m', date_str) as month, 
            COUNT(*) 
        FROM conversations 
        GROUP BY month 
        ORDER BY month
        """)
        stats['by_month'] = dict(cursor.fetchall())
        
        # Sentiment moyen
        cursor.execute("SELECT AVG(sentiment) FROM conversations")
        stats['avg_sentiment'] = cursor.fetchone()[0]
        
        conn.close()
        return stats
    
    def get_tags(self):
        """Récupère la liste des tags disponibles."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("SELECT DISTINCT tag FROM tags")
        tags = [row[0] for row in cursor.fetchall()]
        
        conn.close()
        return tags
    
    def get_agents(self):
        """Récupère la liste des agents disponibles."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("SELECT DISTINCT agent FROM conversations")
        agents = [row[0] for row in cursor.fetchall()]
        
        conn.close()
        return agents

# Classe pour l'analyse des données
class ConversationAnalyzer:
    """Analyse des conversations pour extraire des métadonnées."""
    
    @staticmethod
    def analyze_conversation(messages):
        """Analyse une conversation complète et extrait des métadonnées."""
        if not messages:
            return {}
        
        # Jointure de tous les messages pour analyse globale
        all_text = " ".join([msg.get('content', '') for msg in messages])
        
        # Analyse de sentiment
        blob = TextBlob(all_text)
        sentiment = blob.sentiment.polarity
        
        # Détection de code
        has_code = bool(re.search(r'```[\s\S]*?```', all_text) or 
                        re.search(r'<code[\s\S]*?</code>', all_text))
        
        # Détection d'erreurs
        has_error = any(keyword in all_text.lower() for keyword in ERROR_KEYWORDS)
        
        # Classification thématique
        tags = {}
        for category, keywords in {
            'vba': VBA_KEYWORDS,
            'apex': APEX_KEYWORDS,
            'error': ERROR_KEYWORDS,
            'test': TEST_KEYWORDS,
            'ci_cd': CI_KEYWORDS
        }.items():
            # Calcul du score pour chaque catégorie
            score = sum(1 for keyword in keywords if keyword in all_text.lower())
            if score > 0:
                tags[category] = min(1.0, score / len(keywords) * 2)  # Normaliser entre 0 et 1
        
        # Générer un titre à partir du premier message utilisateur
        title = "Sans titre"
        for msg in messages:
            if msg.get('role') == 'user':
                text = msg.get('content', '')
                # Extraire la première phrase ou les 50 premiers caractères
                title = text.split('\n')[0].strip()
                title = (title[:47] + '...') if len(title) > 50 else title
                break
        
        # Générer un résumé
        summary = ConversationAnalyzer.generate_summary(all_text)
        
        return {
            'title': title,
            'sentiment': sentiment,
            'has_code': has_code,
            'has_error': has_error,
            'tags': tags,
            'summary': summary,
            'word_count': len(all_text.split())
        }
    
    @staticmethod
    def generate_summary(text, max_length=150):
        """Génère un résumé concis du texte."""
        # Si le texte est court, le retourner directement
        if len(text) <= max_length:
            return text
        
        # Sinon, extraire les premières phrases
        sentences = re.split(r'[.!?]', text)
        summary = ""
        
        for sentence in sentences:
            if len(summary) + len(sentence) < max_length:
                summary += sentence + ". "
            else:
                break
        
        return summary.strip()

# Application Flask
app = Flask(__name__)
app.config['DATABASE_PATH'] = os.path.join(os.getcwd(), 'conversations.db')
app.config['TEMPLATES_AUTO_RELOAD'] = True
conversation_store = ConversationStore()

@app.route('/')
def index():
    """Page d'accueil avec tableau de bord et statistiques."""
    stats = conversation_store.get_stats()
    return render_template('index.html', stats=stats)

@app.route('/conversations')
def list_conversations():
    """Liste des conversations avec filtres et pagination."""
    # Récupérer les paramètres de recherche
    text = request.args.get('text', '')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    agent = request.args.get('agent', '')
    tag = request.args.get('tag', '')
    has_error = request.args.get('has_error') == 'true'
    has_code = request.args.get('has_code') == 'true'
    
    # Paramètres de tri et pagination
    sort_by = request.args.get('sort_by', 'timestamp')
    sort_order = request.args.get('sort_order', 'desc')
    page = int(request.args.get('page', 1))
    per_page = int(request.args.get('per_page', 20))
    
    # Construire les filtres
    filters = {}
    if text:
        filters['text'] = text
    if date_from:
        filters['date_from'] = int(datetime.strptime(date_from, '%Y-%m-%d').timestamp())
    if date_to:
        # Ajouter un jour pour inclure toute la journée
        end_date = datetime.strptime(date_to, '%Y-%m-%d') + timedelta(days=1)
        filters['date_to'] = int(end_date.timestamp())
    if agent:
        filters['agent'] = agent
    if tag:
        filters['tag'] = tag
    if has_error:
        filters['has_error'] = has_error
    if has_code:
        filters['has_code'] = has_code
    
    # Récupérer les conversations
    result = conversation_store.search_conversations(
        filters=filters,
        sort_by=sort_by,
        sort_order=sort_order,
        page=page,
        per_page=per_page
    )
    
    # Récupérer les listes de tags et agents pour les filtres
    tags = conversation_store.get_tags()
    agents = conversation_store.get_agents()
    
    return render_template(
        'conversations.html',
        conversations=result['conversations'],
        total=result['total'],
        page=result['page'],
        pages=result['pages'],
        per_page=result['per_page'],
        tags=tags,
        agents=agents,
        filters=filters,
        sort_by=sort_by,
        sort_order=sort_order
    )

@app.route('/conversation/<conversation_id>')
def view_conversation(conversation_id):
    """Affichage détaillé d'une conversation."""
    conversation = conversation_store.get_conversation(conversation_id)
    
    if not conversation:
        abort(404)
    
    return render_template('conversation.html', conversation=conversation)

@app.route('/api/conversations')
def api_conversations():
    """API pour récupérer les conversations (format JSON)."""
    # Mêmes paramètres que pour la liste web
    text = request.args.get('text', '')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    agent = request.args.get('agent', '')
    tag = request.args.get('tag', '')
    has_error = request.args.get('has_error') == 'true'
    has_code = request.args.get('has_code') == 'true'
    
    sort_by = request.args.get('sort_by', 'timestamp')
    sort_order = request.args.get('sort_order', 'desc')
    page = int(request.args.get('page', 1))
    per_page = int(request.args.get('per_page', 20))
    
    filters = {}
    if text:
        filters['text'] = text
    if date_from:
        filters['date_from'] = int(datetime.strptime(date_from, '%Y-%m-%d').timestamp())
    if date_to:
        end_date = datetime.strptime(date_to, '%Y-%m-%d') + timedelta(days=1)
        filters['date_to'] = int(end_date.timestamp())
    if agent:
        filters['agent'] = agent
    if tag:
        filters['tag'] = tag
    if has_error:
        filters['has_error'] = has_error
    if has_code:
        filters['has_code'] = has_code
    
    result = conversation_store.search_conversations(
        filters=filters,
        sort_by=sort_by,
        sort_order=sort_order,
        page=page,
        per_page=per_page
    )
    
    return jsonify(result)

@app.route('/api/conversation/<conversation_id>')
def api_conversation(conversation_id):
    """API pour récupérer une conversation spécifique (format JSON)."""
    conversation = conversation_store.get_conversation(conversation_id)
    
    if not conversation:
        abort(404)
    
    return jsonify(conversation)

@app.route('/api/stats')
def api_stats():
    """API pour récupérer les statistiques (format JSON)."""
    stats = conversation_store.get_stats()
    return jsonify(stats)

@app.route('/api/tags')
def api_tags():
    """API pour récupérer la liste des tags (format JSON)."""
    tags = conversation_store.get_tags()
    return jsonify(tags)

@app.route('/api/agents')
def api_agents():
    """API pour récupérer la liste des agents (format JSON)."""
    agents = conversation_store.get_agents()
    return jsonify(agents)

@app.route('/visualize')
def visualize():
    """Page avec visualisations interactives des données."""
    stats = conversation_store.get_stats()
    
    # Créer les graphiques avec Plotly
    
    # 1. Evolution temporelle des conversations
    months = list(stats.get('by_month', {}).keys())
    counts = list(stats.get('by_month', {}).values())
    
    if months:
        time_fig = px.line(
            x=months,
            y=counts,
            title='Évolution du nombre de conversations',
            labels={'x': 'Mois', 'y': 'Nombre de conversations'}
        )
        time_chart = json.dumps(time_fig, cls=plotly.utils.PlotlyJSONEncoder)
    else:
        time_chart = None
    
    # 2. Répartition des agents
    if stats.get('agents'):
        agent_fig = px.pie(
            values=list(stats['agents'].values()),
            names=list(stats['agents'].keys()),
            title='Répartition par agent'
        )
        agent_chart = json.dumps(agent_fig, cls=plotly.utils.PlotlyJSONEncoder)
    else:
        agent_chart = None
    
    # 3. Répartition des tags
    if stats.get('tags'):
        tag_fig = px.bar(
            x=list(stats['tags'].keys()),
            y=list(stats['tags'].values()),
            title='Distribution des sujets',
            labels={'x': 'Sujet', 'y': 'Nombre de conversations'}
        )
        tag_chart = json.dumps(tag_fig, cls=plotly.utils.PlotlyJSONEncoder)
    else:
        tag_chart = None
    
    return render_template(
        'visualize.html',
        stats=stats,
        time_chart=time_chart,
        agent_chart=agent_chart,
        tag_chart=tag_chart
    )

@app.route('/import', methods=['GET', 'POST'])
def import_data():
    """Interface pour importer des données de conversations."""
    if request.method == 'POST':
        # Vérifier si un fichier a été soumis
        if 'file' not in request.files:
            return render_template('import.html', error='Aucun fichier sélectionné')
        
        file = request.files['file']
        
        if file.filename == '':
            return render_template('import.html', error='Aucun fichier sélectionné')
        
        if file and file.filename.endswith('.json'):
            try:
                data = json.load(file)
                
                imported = 0
                for conv in data:
                    # Analyser et enrichir les données
                    analysis = ConversationAnalyzer.analyze_conversation(conv.get('messages', []))
                    conv.update(analysis)
                    
                    # Générer un ID s'il n'existe pas
                    if 'id' not in conv:
                        # Créer un hash basé sur le contenu
                        content = json.dumps(conv.get('messages', []))
                        conv['id'] = hashlib.md5(content.encode()).hexdigest()
                    
                    # Ajouter à la base de données
                    if conversation_store.add_conversation(conv):
                        imported += 1
                
                return render_template('import.html', success=f'{imported} conversations importées avec succès')
            except Exception as e:
                return render_template('import.html', error=f'Erreur lors de l\'importation: {str(e)}')
        else:
            return render_template('import.html', error='Format de fichier non supporté. Veuillez importer un fichier JSON')
    
    return render_template('import.html')

@app.route('/export')
def export_data():
    """Exporte les données au format JSON."""
    # Récupérer les paramètres de filtre (les mêmes que pour la recherche)
    text = request.args.get('text', '')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    agent = request.args.get('agent', '')
    tag = request.args.get('tag', '')
    has_error = request.args.get('has_error') == 'true'
    has_code = request.args.get('has_code') == 'true'
    
    filters = {}
    if text:
        filters['text'] = text
    if date_from:
        filters['date_from'] = int(datetime.strptime(date_from, '%Y-%m-%d').timestamp())
    if date_to:
        end_date = datetime.strptime(date_to, '%Y-%m-%d') + timedelta(days=1)
        filters['date_to'] = int(end_date.timestamp())
    if agent:
        filters['agent'] = agent
    if tag:
        filters['tag'] = tag
    if has_error:
        filters['has_error'] = has_error
    if has_code:
        filters['has_code'] = has_code
    
    # Récupérer toutes les conversations correspondant aux filtres (sans pagination)
    result = conversation_store.search_conversations(
        filters=filters,
        per_page=10000  # Grand nombre pour récupérer toutes les conversations
    )
    
    # Récupérer les détails complets de chaque conversation
    conversations = []
    for conv_summary in result['conversations']:
        conv = conversation_store.get_conversation(conv_summary['id'])
        if conv:
            conversations.append(conv)
    
    # Créer un fichier temporaire
    export_file = f"conversation_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    export_path = os.path.join(os.getcwd(), export_file)
    
    with open(export_path, 'w', encoding='utf-8') as f:
        json.dump(conversations, f, ensure_ascii=False, indent=2)
    
    return send_file(export_path, as_attachment=True, download_name=export_file)

@app.route('/compare', methods=['GET', 'POST'])
def compare_conversations():
    """Compare two or more conversations side by side."""
    if request.method == 'POST':
        # Récupérer les IDs des conversations à comparer
        conversation_ids = request.form.getlist('conversation_ids')
        
        if not conversation_ids or len(conversation_ids) < 2:
            return render_template('compare.html', error='Veuillez sélectionner au moins deux conversations')
        
        # Récupérer les conversations complètes
        conversations = []
        for conv_id in conversation_ids:
            conv = conversation_store.get_conversation(conv_id)
            if conv:
                conversations.append(conv)
            else:
                return render_template('compare.html', error=f'Conversation {conv_id} non trouvée')
        
        return render_template('compare_view.html', conversations=conversations)
    
    # Afficher le formulaire de sélection
    # Récupérer les 100 dernières conversations pour le sélecteur
    result = conversation_store.search_conversations(per_page=100)
    
    return render_template('compare.html', conversations=result['conversations'])

def main():
    """Point d'entrée principal pour exécuter l'application."""
    parser = argparse.ArgumentParser(description="Interface d'exploration des conversations IA.")
    parser.add_argument('--db-path', default='conversations.db', help='Chemin vers la base de données SQLite')
    parser.add_argument('--port', type=int, default=5000, help='Port du serveur web')
    parser.add_argument('--host', default='127.0.0.1', help='Hôte du serveur web')
    parser.add_argument('--debug', action='store_true', help='Mode debug')
    
    args = parser.parse_args()
    
    # Initialiser la base de données
    global conversation_store
    conversation_store = ConversationStore(db_path=args.db_path)
    
    # Lancer l'application Flask
    app.run(host=args.host, port=args.port, debug=args.debug)

if __name__ == '__main__':
    main() 