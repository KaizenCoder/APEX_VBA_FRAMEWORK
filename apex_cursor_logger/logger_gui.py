#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Interface graphique du système de journalisation APEX Framework
Version: 1.0.3
Date: 2025-04-14
"""

import os
import sys
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from pathlib import Path
import threading
import re
import webbrowser
import logging
import traceback
import time

# Configuration du système de journalisation pour le débogage
LOG_LEVEL = logging.DEBUG  # Niveau de détail des logs (DEBUG, INFO, WARNING, ERROR, CRITICAL)
LOG_FILE = "apex_logger_gui.log"

# Initialisation du logger
def setup_logger():
    """Configure le système de journalisation pour le débogage"""
    logger = logging.getLogger('apex_logger_gui')
    logger.setLevel(LOG_LEVEL)
    
    # Créer un gestionnaire de fichier qui écrit les messages de log dans un fichier
    file_handler = logging.FileHandler(LOG_FILE, mode='a', encoding='utf-8')
    file_handler.setLevel(LOG_LEVEL)
    
    # Créer un gestionnaire de console qui écrit les messages de log dans la console
    console_handler = logging.StreamHandler()
    console_handler.setLevel(LOG_LEVEL)
    
    # Créer un formateur qui définit le format des messages de log
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # Ajouter les gestionnaires au logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

# Initialiser le logger
logger = setup_logger()
logger.info("=== Démarrage de l'interface graphique du système de journalisation APEX Framework ===")

# Installation de CustomTkinter si non disponible
try:
    import customtkinter as ctk
    logger.info("Module CustomTkinter chargé avec succès")
except ImportError:
    logger.warning("Module CustomTkinter non trouvé, installation en cours...")
    import subprocess
    print("Installation de CustomTkinter...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "customtkinter"])
    import customtkinter as ctk
    logger.info("Module CustomTkinter installé et chargé avec succès")

# Installation de markdown2 si non disponible
try:
    import markdown2
    logger.info("Module markdown2 chargé avec succès")
except ImportError:
    logger.warning("Module markdown2 non trouvé, installation en cours...")
    import subprocess
    print("Installation de markdown2...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "markdown2"])
    import markdown2
    logger.info("Module markdown2 installé et chargé avec succès")


# Configuration de l'interface
ctk.set_appearance_mode("System")  # Modes: "System" (par défaut), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Thèmes: "blue" (par défaut), "green", "dark-blue"

class LoggerApp(ctk.CTk):
    """Application principale du système de journalisation APEX Framework"""
    
    def __init__(self):
        super().__init__()
        
        # Configuration de base
        self.title("APEX Framework - Système de Journalisation")
        self.geometry("1200x800")
        self.minsize(900, 600)
        
        logger.info("Initialisation de l'application")
        
        # Chargement de la configuration
        self.load_config()
        
        # Variables
        self.current_session = None
        self.sessions = []
        self.journal_entries = []
        self.logger_status = {
            "cursor": {"active": False, "last_check": None, "status_text": "Inconnu"},
            "vscode": {"active": False, "last_check": None, "status_text": "Inconnu"}
        }
        
        # Interface
        self.setup_ui()
        
        # Chargement initial des données
        self.load_sessions()
        self.load_journal()
        
        # Démarrer la vérification périodique du statut du logger
        self.check_logger_status()
        threading.Thread(target=self.periodic_status_check, daemon=True).start()
        
        logger.info("Initialisation de l'application terminée")
    
    def load_config(self):
        """Charge la configuration du système de journalisation"""
        try:
            # Détermination du chemin du fichier de configuration
            script_dir = Path(__file__).parent
            config_path = script_dir / "logger_config.json"
            
            logger.debug(f"Chargement de la configuration depuis {config_path}")
            
            with open(config_path, encoding='utf-8') as f:
                self.config = json.load(f)
                
            # Chemins importants - CORRECTION: pointer vers le dossier logs principal du framework
            # Au lieu d'utiliser le dossier logs du sous-répertoire apex_cursor_logger
            self.logs_dir = Path("d:/Dev/Apex_VBA_FRAMEWORK/logs")
            self.prompts_dir = script_dir / self.config.get("prompts_subdir", "prompts")
            
            # S'assurer que les répertoires existent
            if not self.logs_dir.exists():
                logger.error(f"Le répertoire de logs '{self.logs_dir}' n'existe pas. Utilisation du répertoire par défaut.")
                self.logs_dir = script_dir / self.config.get("logs_subdir", "logs")
            
            self.prompts_dir.mkdir(exist_ok=True)
            
            logger.info(f"Dossier de logs utilisé : {self.logs_dir}")
            logger.info(f"Dossier de prompts utilisé : {self.prompts_dir}")
            logger.debug(f"Configuration chargée : {json.dumps(self.config, indent=2)}")
            
        except Exception as e:
            logger.critical(f"Erreur lors du chargement de la configuration : {str(e)}")
            logger.critical(traceback.format_exc())
            messagebox.showerror(
                "Erreur de configuration", 
                f"Impossible de charger la configuration : {str(e)}"
            )
            sys.exit(1)
    
    def setup_ui(self):
        """Configure l'interface utilisateur"""
        # Configuration de la grille principale
        self.grid_columnconfigure(0, weight=1)  # Sidebar
        self.grid_columnconfigure(1, weight=3)  # Main content
        self.grid_rowconfigure(0, weight=1)
        
        # Création du panneau latéral (sidebar)
        self.setup_sidebar()
        
        # Création de la zone principale
        self.setup_main_area()
    
    def setup_sidebar(self):
        """Configure le panneau latéral"""
        sidebar_frame = ctk.CTkFrame(self, width=200)
        sidebar_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        sidebar_frame.grid_rowconfigure(4, weight=1)  # Pour que la liste prenne tout l'espace disponible
        
        # Titre
        title_label = ctk.CTkLabel(
            sidebar_frame, 
            text="APEX Logger", 
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")
        
        # Sous-titre
        subtitle_label = ctk.CTkLabel(
            sidebar_frame, 
            text="Sessions de journalisation", 
            font=ctk.CTkFont(size=14)
        )
        subtitle_label.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="w")
        
        # Bouton pour rafraîchir les sessions
        refresh_btn = ctk.CTkButton(
            sidebar_frame,
            text="Rafraîchir",
            command=self.load_sessions
        )
        refresh_btn.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        
        # Zone de recherche
        self.search_var = tk.StringVar()
        self.search_var.trace("w", lambda name, index, mode: self.filter_sessions())
        search_entry = ctk.CTkEntry(
            sidebar_frame,
            placeholder_text="Rechercher...",
            textvariable=self.search_var
        )
        search_entry.grid(row=3, column=0, padx=20, pady=(0, 10), sticky="ew")
        
        # Liste des sessions avec scrollbar
        sessions_frame = ctk.CTkFrame(sidebar_frame)
        sessions_frame.grid(row=4, column=0, padx=10, pady=10, sticky="nsew")
        sessions_frame.grid_rowconfigure(0, weight=1)
        sessions_frame.grid_columnconfigure(0, weight=1)
        
        # Utilisation d'un CTkScrollableFrame au lieu d'un CTkTextbox pour la liste des sessions
        self.sessions_list_frame = ctk.CTkScrollableFrame(sessions_frame)
        self.sessions_list_frame.grid(row=0, column=0, sticky="nsew")
        
        # Boutons d'action
        actions_frame = ctk.CTkFrame(sidebar_frame)
        actions_frame.grid(row=5, column=0, padx=10, pady=10, sticky="ew")
        
        export_btn = ctk.CTkButton(
            actions_frame, 
            text="Exporter", 
            command=self.export_session
        )
        export_btn.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        delete_btn = ctk.CTkButton(
            actions_frame, 
            text="Supprimer", 
            fg_color="darkred", 
            command=self.delete_session
        )
        delete_btn.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
    
    def setup_main_area(self):
        """Configure la zone principale"""
        main_frame = ctk.CTkFrame(self)
        main_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        main_frame.grid_rowconfigure(1, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        
        # En-tête avec les détails de la session
        header_frame = ctk.CTkFrame(main_frame)
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        header_frame.grid_columnconfigure(1, weight=1)
        
        # Titre de la session - AMÉLIORATION: police plus grande et couleur plus contrastée
        self.session_title = ctk.CTkLabel(
            header_frame,
            text="Aucune session sélectionnée",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color="#007BFF"  # Bleu plus visible
        )
        self.session_title.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 0), sticky="w")
        
        # Détails de la session - AMÉLIORATION: police plus lisible
        self.session_details = ctk.CTkLabel(
            header_frame,
            text="Sélectionnez une session dans le panneau de gauche",
            font=ctk.CTkFont(size=13),
            text_color="#666666"  # Gris plus foncé pour la lisibilité
        )
        self.session_details.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="w")
        
        # Contenu principal (journal)
        self.notebook = ctk.CTkTabview(main_frame)
        self.notebook.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        
        # Onglet Session Log
        self.tab_session_log = self.notebook.add("Session Log")
        self.tab_session_log.grid_rowconfigure(0, weight=1)
        self.tab_session_log.grid_columnconfigure(0, weight=1)
        
        # Zone de texte pour afficher le journal - AMÉLIORATION: police plus grande et style monospace
        self.journal_text = ctk.CTkTextbox(
            self.tab_session_log, 
            wrap="word",
            font=ctk.CTkFont(family="Consolas", size=12)  # Police monospace pour une meilleure lisibilité
        )
        self.journal_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.journal_text.configure(state="disabled")
        
        # Onglet Live View
        self.tab_live_view = self.notebook.add("Live View")
        self.tab_live_view.grid_rowconfigure(0, weight=1)
        self.tab_live_view.grid_columnconfigure(0, weight=1)
        
        self.live_view_frame = ctk.CTkScrollableFrame(self.tab_live_view)
        self.live_view_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.live_view_frame.grid_columnconfigure(0, weight=1)
        
        # Nouvel onglet Historique des demandes
        self.tab_history = self.notebook.add("Historique")
        self.setup_history_tab()
        
        # Onglet Export
        self.tab_export = self.notebook.add("Export")
        self.setup_export_tab()
        
        # Onglet Rapports
        self.tab_reports = self.notebook.add("Rapports")
        self.setup_reports_tab()
        
        # Barre de statut
        status_frame = ctk.CTkFrame(main_frame, height=30)
        status_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))
        status_frame.grid_columnconfigure(0, weight=1)
        
        self.status_label = ctk.CTkLabel(
            status_frame,
            text="Prêt",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        self.status_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        
        # Indicateurs d'état du logger pour chaque type d'IA
        indicators_frame = ctk.CTkFrame(status_frame, fg_color="transparent")
        indicators_frame.grid(row=0, column=1, padx=10, pady=5, sticky="e")
        
        # Indicateur Cursor
        cursor_indicator_frame = ctk.CTkFrame(indicators_frame, fg_color="transparent")
        cursor_indicator_frame.grid(row=0, column=0, padx=(0, 10))
        
        self.cursor_indicator = ctk.CTkLabel(
            cursor_indicator_frame,
            text="⬤",
            font=ctk.CTkFont(size=14),
            text_color="gray",
            width=15
        )
        self.cursor_indicator.grid(row=0, column=0, padx=(0, 2))
        
        cursor_label = ctk.CTkLabel(
            cursor_indicator_frame,
            text="Cursor",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        )
        cursor_label.grid(row=0, column=1)
        
        # Indicateur VSCode
        vscode_indicator_frame = ctk.CTkFrame(indicators_frame, fg_color="transparent")
        vscode_indicator_frame.grid(row=0, column=1, padx=(0, 10))
        
        self.vscode_indicator = ctk.CTkLabel(
            vscode_indicator_frame,
            text="⬤",
            font=ctk.CTkFont(size=14),
            text_color="gray",
            width=15
        )
        self.vscode_indicator.grid(row=0, column=0, padx=(0, 2))
        
        vscode_label = ctk.CTkLabel(
            vscode_indicator_frame,
            text="VSCode",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        )
        vscode_label.grid(row=0, column=1)
        
        # Version
        version_label = ctk.CTkLabel(
            status_frame,
            text="v1.0.3",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        version_label.grid(row=0, column=2, padx=10, pady=5, sticky="e")
    
    def setup_export_tab(self):
        """Configure l'onglet Export"""
        self.tab_export.grid_columnconfigure(0, weight=1)
        
        # Options d'export
        options_frame = ctk.CTkFrame(self.tab_export)
        options_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=20)
        options_frame.grid_columnconfigure(1, weight=1)
        
        # Format d'export
        format_label = ctk.CTkLabel(options_frame, text="Format:")
        format_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.export_format = ctk.CTkSegmentedButton(
            options_frame,
            values=["Markdown", "HTML", "JSON", "PDF"],
            command=self.on_export_format_change
        )
        self.export_format.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        self.export_format.set("Markdown")
        
        # Période
        period_label = ctk.CTkLabel(options_frame, text="Période:")
        period_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        
        self.export_period = ctk.CTkSegmentedButton(
            options_frame,
            values=["Session actuelle", "Aujourd'hui", "Cette semaine", "Ce mois", "Personnalisé"],
            command=self.on_export_period_change
        )
        self.export_period.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        self.export_period.set("Session actuelle")
        
        # Options avancées
        advanced_frame = ctk.CTkFrame(self.tab_export)
        advanced_frame.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 20))
        advanced_frame.grid_columnconfigure(1, weight=1)
        
        # Inclure les métadonnées
        self.include_metadata = ctk.CTkCheckBox(
            advanced_frame,
            text="Inclure les métadonnées"
        )
        self.include_metadata.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.include_metadata.select()
        
        # Anonymiser les données
        self.anonymize_data = ctk.CTkCheckBox(
            advanced_frame,
            text="Anonymiser les données"
        )
        self.anonymize_data.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        # Exporter les images
        self.export_images = ctk.CTkCheckBox(
            advanced_frame,
            text="Inclure les images/pièces jointes"
        )
        self.export_images.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.export_images.select()
        
        # Compression
        self.compress_export = ctk.CTkCheckBox(
            advanced_frame,
            text="Compresser l'export (zip)"
        )
        self.compress_export.grid(row=1, column=1, padx=10, pady=10, sticky="w")
        
        # Destination
        dest_frame = ctk.CTkFrame(self.tab_export)
        dest_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 20))
        dest_frame.grid_columnconfigure(1, weight=1)
        
        dest_label = ctk.CTkLabel(dest_frame, text="Destination:")
        dest_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.export_path = ctk.CTkEntry(dest_frame)
        self.export_path.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        browse_btn = ctk.CTkButton(
            dest_frame,
            text="Parcourir",
            width=100,
            command=self.browse_export_path
        )
        browse_btn.grid(row=0, column=2, padx=10, pady=10)
        
        # Bouton d'export
        export_btn_frame = ctk.CTkFrame(self.tab_export, fg_color="transparent")
        export_btn_frame.grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 20))
        export_btn_frame.grid_columnconfigure(0, weight=1)
        
        export_btn = ctk.CTkButton(
            export_btn_frame,
            text="Exporter",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40,
            command=self.do_export
        )
        export_btn.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
    
    def setup_reports_tab(self):
        """Configure l'onglet Rapports"""
        self.tab_reports.grid_columnconfigure(0, weight=1)
        self.tab_reports.grid_rowconfigure(1, weight=1)
        
        # Sélection du type de rapport
        report_type_frame = ctk.CTkFrame(self.tab_reports)
        report_type_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=20)
        
        report_type_label = ctk.CTkLabel(
            report_type_frame, 
            text="Type de rapport:",
            font=ctk.CTkFont(size=14)
        )
        report_type_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.report_type = ctk.CTkSegmentedButton(
            report_type_frame,
            values=["Activité", "Performance", "Qualité", "Utilisation"],
            command=self.load_report
        )
        self.report_type.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        self.report_type.set("Activité")
        
        # Zone d'affichage du rapport
        report_display_frame = ctk.CTkFrame(self.tab_reports)
        report_display_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))
        report_display_frame.grid_rowconfigure(0, weight=1)
        report_display_frame.grid_columnconfigure(0, weight=1)
        
        self.report_display = ctk.CTkTextbox(report_display_frame, wrap="word")
        self.report_display.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.report_display.configure(state="disabled")
    
    def setup_history_tab(self):
        """Configure l'onglet Historique pour afficher les dernières demandes"""
        self.tab_history.grid_rowconfigure(1, weight=1)
        self.tab_history.grid_columnconfigure(0, weight=1)
        
        # Panneau des options
        options_frame = ctk.CTkFrame(self.tab_history)
        options_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=20)
        options_frame.grid_columnconfigure(2, weight=1)
        
        # Sélection de la source
        source_label = ctk.CTkLabel(options_frame, text="Source:")
        source_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.history_source = ctk.CTkSegmentedButton(
            options_frame,
            values=["Tous", "Cursor", "VSCode"],
            command=self.load_history
        )
        self.history_source.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        self.history_source.set("Tous")
        
        # Nombre d'éléments à afficher
        count_label = ctk.CTkLabel(options_frame, text="Nombre:")
        count_label.grid(row=0, column=2, padx=(20, 10), pady=10, sticky="e")
        
        self.history_count = ctk.CTkSegmentedButton(
            options_frame,
            values=["10", "30", "50", "100"],
            command=self.load_history
        )
        self.history_count.grid(row=0, column=3, padx=10, pady=10, sticky="e")
        self.history_count.set("30")
        
        # Bouton de rafraîchissement
        refresh_btn = ctk.CTkButton(
            options_frame,
            text="Rafraîchir",
            command=self.load_history,
            width=100
        )
        refresh_btn.grid(row=0, column=4, padx=10, pady=10, sticky="e")
        
        # Zone d'affichage des demandes
        self.history_frame = ctk.CTkScrollableFrame(self.tab_history)
        self.history_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))
        self.history_frame.grid_columnconfigure(0, weight=1)
        
        # Charger les demandes initiales
        self.load_history("Tous")
    
    def load_sessions(self):
        """Charge la liste des sessions disponibles"""
        try:
            # Réinitialiser les sessions
            self.sessions = []
            
            logger.info("Chargement des sessions...")
            
            # Rechercher tous les fichiers de session dans le répertoire des logs
            logger.debug(f"Recherche des fichiers de session dans {self.logs_dir}")
            session_files = list(self.logs_dir.glob("cursor_session_*.md"))
            session_files.extend(self.logs_dir.glob("vscode_session_*.md"))
            
            logger.info(f"Nombre de fichiers de session trouvés: {len(session_files)}")
            
            # Trier par date (plus récent en premier)
            session_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            
            # Extraire les informations de base
            for file_path in session_files:
                session_id = file_path.stem
                try:
                    logger.debug(f"Lecture du fichier de session: {file_path}")
                    
                    # Lire le contenu pour extraire le titre
                    content = file_path.read_text(encoding='utf-8')
                    title_match = re.search(r'^# Session .+?: (.+)$', content, re.MULTILINE)
                    title = title_match.group(1) if title_match else session_id
                    
                    # Extraire la date
                    date_match = re.search(r'^Date: (.+)$', content, re.MULTILINE)
                    date = date_match.group(1) if date_match else "Unknown"
                    
                    # Ajouter à la liste des sessions
                    self.sessions.append({
                        'id': session_id,
                        'title': title,
                        'date': date,
                        'path': file_path,
                        'content': content
                    })
                    
                    logger.debug(f"Session ajoutée - ID: {session_id}, Titre: {title}, Date: {date}")
                except Exception as e:
                    logger.error(f"Erreur lors de la lecture de {file_path}: {str(e)}")
                    logger.error(traceback.format_exc())
                    print(f"Erreur lors de la lecture de {file_path}: {str(e)}")
            
            # Mettre à jour l'affichage
            self.update_sessions_list()
            
            # Mettre à jour le statut
            self.status_label.configure(text=f"{len(self.sessions)} sessions trouvées")
            logger.info(f"Chargement des sessions terminé. {len(self.sessions)} sessions chargées.")
            
        except Exception as e:
            logger.critical(f"Erreur lors du chargement des sessions: {str(e)}")
            logger.critical(traceback.format_exc())
            messagebox.showerror(
                "Erreur", 
                f"Impossible de charger les sessions : {str(e)}"
            )
    
    def update_sessions_list(self):
        """Mettre à jour l'affichage de la liste des sessions"""
        # CORRECTION: Modification pour utiliser des widgets CTk au lieu des tags
        
        # Effacer les widgets existants
        for widget in self.sessions_list_frame.winfo_children():
            widget.destroy()
        
        # Filtrer si nécessaire
        search_term = self.search_var.get().lower()
        filtered_sessions = [s for s in self.sessions if (
            search_term in s['id'].lower() or 
            search_term in s['title'].lower() or 
            search_term in s['date'].lower()
        )]
        
        # Ajouter les sessions à la liste
        for idx, session in enumerate(filtered_sessions):
            session_frame = ctk.CTkFrame(self.sessions_list_frame)
            session_frame.grid(row=idx, column=0, pady=5, padx=5, sticky="ew")
            session_frame.grid_columnconfigure(0, weight=1)
            
            # Titre de la session
            title_label = ctk.CTkLabel(
                session_frame,
                text=session['title'],
                font=ctk.CTkFont(size=12, weight="bold"),
                anchor="w"
            )
            title_label.grid(row=0, column=0, padx=5, pady=(5, 0), sticky="ew")
            
            # Date de la session
            date_label = ctk.CTkLabel(
                session_frame,
                text=session['date'],
                font=ctk.CTkFont(size=10),
                text_color="gray",
                anchor="w"
            )
            date_label.grid(row=1, column=0, padx=5, pady=(0, 5), sticky="ew")
            
            # Rendre le cadre cliquable
            session_frame.bind("<Button-1>", lambda e, s=session: self.load_session(s))
            title_label.bind("<Button-1>", lambda e, s=session: self.load_session(s))
            date_label.bind("<Button-1>", lambda e, s=session: self.load_session(s))
    
    def load_session(self, session):
        """Charger une session spécifique"""
        try:
            logger.info(f"Chargement de la session: {session['id']}")
            
            self.current_session = session
            
            # Mettre à jour le titre et les détails
            self.session_title.configure(text=session['title'])
            self.session_details.configure(text=f"ID: {session['id']} | Date: {session['date']}")
            
            # AMÉLIORATION: Formater le contenu Markdown pour une meilleure lisibilité
            content = self.format_markdown_for_display(session['content'])
            
            # Mettre à jour le contenu du journal
            self.journal_text.configure(state="normal")
            self.journal_text.delete("0.0", "end")
            self.journal_text.insert("0.0", content)
            self.journal_text.configure(state="disabled")
            
            # Mettre à jour la vue en direct
            self.update_live_view(session)
            
            # Mettre à jour le statut
            self.status_label.configure(text=f"Session '{session['title']}' chargée")
            
            # Passer à l'onglet Session Log
            self.notebook.set("Session Log")
            
            logger.info(f"Session chargée avec succès: {session['id']}")
        except Exception as e:
            logger.error(f"Erreur lors du chargement de la session {session.get('id', 'inconnue')}: {str(e)}")
            logger.error(traceback.format_exc())
            messagebox.showerror(
                "Erreur", 
                f"Impossible de charger la session : {str(e)}"
            )
    
    def format_markdown_for_display(self, content):
        """Améliore le formatage du texte Markdown pour l'affichage"""
        # Amélioration des titres pour plus de clarté
        content = re.sub(r'(^|\n)# (.*?)(\n|$)', r'\1━━━━━━━━━━━━━━━━━━━━━━━\n# \2\n━━━━━━━━━━━━━━━━━━━━━━━\3', content)
        content = re.sub(r'(^|\n)## (.*?)(\n|$)', r'\1\n## \2\n──────────────────────\3', content)
        
        # Amélioration du formatage des listes
        content = re.sub(r'(^|\n)- \[ \] (.*?)(\n|$)', r'\1□ \2\3', content)
        content = re.sub(r'(^|\n)- \[x\] (.*?)(\n|$)', r'\1✅ \2\3', content)
        
        # Mise en évidence des sections importantes
        content = re.sub(r'(^|\n)### (.*?)(\n|$)', r'\1\n▶ \2\3', content)
        content = re.sub(r'(^|\n)\*\*Prompt\*\*:', r'\1➤ Prompt:', content)
        content = re.sub(r'(^|\n)\*\*Réponse\*\*:', r'\1➤ Réponse:', content)
        content = re.sub(r'(^|\n)\*\*Note\*\*:', r'\1➤ Note:', content)
        
        # Retour du contenu amélioré
        return content
    
    def update_live_view(self, session):
        """Mettre à jour la vue en direct avec les interactions de la session"""
        # Effacer le contenu actuel
        for widget in self.live_view_frame.winfo_children():
            widget.destroy()
        
        # Extraire les interactions
        interactions = self.extract_interactions(session['content'])
        
        # Créer des widgets pour chaque interaction
        for idx, interaction in enumerate(interactions):
            self.create_interaction_widget(idx, interaction)
    
    def extract_interactions(self, content):
        """Extraire les interactions de la session à partir du contenu Markdown"""
        interactions = []
        
        # Rechercher les sections d'interaction
        pattern = r"### (\d{4}-\d{2}-\d{2} \d{2}:\d{2}(?::\d{2})?)(?: - (.+))?\n\n\*\*Prompt\*\*: (.*?)\n\n\*\*Réponse\*\*: (.*?)(?:\n\n\*\*Note\*\*: (.*?))?(?:\n---|\Z)"
        matches = re.finditer(pattern, content, re.DOTALL)
        
        for match in matches:
            timestamp = match.group(1)
            agent = match.group(2) if match.group(2) else "Unknown"
            prompt = match.group(3)
            response = match.group(4)
            note = match.group(5) if match.group(5) else ""
            
            interactions.append({
                'timestamp': timestamp,
                'agent': agent,
                'prompt': prompt,
                'response': response,
                'note': note
            })
        
        return interactions
    
    def create_interaction_widget(self, idx, interaction):
        """Créer un widget pour une interaction dans la vue en direct"""
        # AMÉLIORATION: Style amélioré pour les widgets d'interaction
        
        # Cadre principal pour cette interaction avec couleur de fond légèrement différente pour contraste
        frame = ctk.CTkFrame(self.live_view_frame, fg_color=("#f0f0f0", "#2d2d2d"))
        frame.grid(row=idx, column=0, sticky="ew", padx=10, pady=10)
        frame.grid_columnconfigure(0, weight=1)
        
        # En-tête avec timestamp et agent - style amélioré
        header_frame = ctk.CTkFrame(frame, fg_color=("#e0e0e0", "#333333"))
        header_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        
        timestamp_label = ctk.CTkLabel(
            header_frame,
            text=interaction['timestamp'],
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=("#007BFF", "#3a8eff")  # Bleu plus contrasté
        )
        timestamp_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        
        agent_label = ctk.CTkLabel(
            header_frame,
            text=interaction['agent'],
            font=ctk.CTkFont(size=12),
            text_color=("#555", "#aaa")
        )
        agent_label.grid(row=0, column=1, padx=10, pady=5, sticky="e")
        
        # Prompt - mise en forme améliorée
        prompt_label = ctk.CTkLabel(
            frame,
            text="Prompt:",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=("#444", "#bbb")
        )
        prompt_label.grid(row=1, column=0, padx=10, pady=(15, 0), sticky="w")
        
        prompt_text = ctk.CTkTextbox(
            frame, 
            height=50, 
            wrap="word",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=("#f8f8f8", "#282828")  # Fond légèrement différent pour le texte
        )
        prompt_text.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="ew")
        prompt_text.insert("0.0", interaction['prompt'])
        prompt_text.configure(state="disabled")
        
        # Réponse - mise en forme améliorée
        response_label = ctk.CTkLabel(
            frame,
            text="Réponse:",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=("#444", "#bbb")
        )
        response_label.grid(row=3, column=0, padx=10, pady=(5, 0), sticky="w")
        
        response_text = ctk.CTkTextbox(
            frame, 
            height=120, 
            wrap="word",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=("#f8f8f8", "#282828")
        )
        response_text.grid(row=4, column=0, padx=10, pady=(0, 10), sticky="ew")
        response_text.insert("0.0", interaction['response'])
        response_text.configure(state="disabled")
        
        # Note (si présente) - mise en forme améliorée
        if interaction['note']:
            note_frame = ctk.CTkFrame(frame, fg_color=("transparent", "transparent"))
            note_frame.grid(row=5, column=0, sticky="ew", padx=10, pady=(0, 10))
            
            note_label = ctk.CTkLabel(
                note_frame,
                text="Note:",
                font=ctk.CTkFont(size=13, weight="bold"),
                text_color=("#444", "#bbb")
            )
            note_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
            
            note_text = ctk.CTkLabel(
                note_frame,
                text=interaction['note'],
                font=ctk.CTkFont(size=12, slant="italic"),
                text_color=("#555", "#aaa")
            )
            note_text.grid(row=0, column=1, padx=10, pady=5, sticky="w")
    
    def load_journal(self):
        """Charger le journal principal"""
        try:
            journal_path = self.logs_dir / "apex-cursor-journal.md"
            vscode_journal_path = self.logs_dir / "apex-vscode-journal.md"
            
            # Initialiser la liste des entrées du journal
            self.journal_entries = []
            
            # Charger le journal Cursor si présent
            if journal_path.exists():
                content = journal_path.read_text(encoding='utf-8')
                self.extract_journal_entries(content, "Cursor")
            
            # Charger le journal VSCode si présent
            if vscode_journal_path.exists():
                content = vscode_journal_path.read_text(encoding='utf-8')
                self.extract_journal_entries(content, "VSCode")
            
            # Trier par timestamp (plus récent en premier)
            self.journal_entries.sort(key=lambda x: x['timestamp'], reverse=True)
            
        except Exception as e:
            messagebox.showerror(
                "Erreur", 
                f"Impossible de charger le journal : {str(e)}"
            )
    
    def extract_journal_entries(self, content, source):
        """Extraire les entrées du journal à partir du contenu Markdown"""
        pattern = r"### (\d{4}-\d{2}-\d{2} \d{2}:\d{2}(?::\d{2})?) - (.+?)\n\n\*\*Contexte :\*\* (.+?)\n\*\*Prompt :\*\* (.+?)\n\*\*Agent IA :\*\* (.+?)\n\*\*Avis :\*\* (.+?)\n\*\*Réponse \(extrait\) :\*\* (.+?)(?:\n---|$)"
        matches = re.finditer(pattern, content, re.DOTALL)
        
        for match in matches:
            timestamp = match.group(1)
            session_id = match.group(2)
            context = match.group(3)
            prompt = match.group(4)
            agent = match.group(5)
            rating = match.group(6)
            response = match.group(7)
            
            self.journal_entries.append({
                'timestamp': timestamp,
                'session_id': session_id,
                'context': context,
                'prompt': prompt,
                'agent': agent,
                'rating': rating,
                'response': response,
                'source': source
            })
    
    def filter_sessions(self):
        """Filtrer la liste des sessions en fonction de la recherche"""
        self.update_sessions_list()
    
    def export_session(self):
        """Exporter la session sélectionnée"""
        if not self.current_session:
            messagebox.showwarning(
                "Avertissement", 
                "Aucune session sélectionnée à exporter"
            )
            return
        
        # Passer à l'onglet Export et préremplir les paramètres
        self.notebook.set("Export")
        self.export_format.set("Markdown")
        self.export_period.set("Session actuelle")
        
        # Focus sur le bouton d'export
        self.status_label.configure(
            text=f"Prêt à exporter la session '{self.current_session['title']}'"
        )
    
    def delete_session(self):
        """Supprimer la session sélectionnée"""
        if not self.current_session:
            messagebox.showwarning(
                "Avertissement", 
                "Aucune session sélectionnée à supprimer"
            )
            return
        
        # Demander confirmation
        if not messagebox.askyesno(
            "Confirmation",
            f"Êtes-vous sûr de vouloir supprimer la session '{self.current_session['title']}' ?\nCette action est irréversible."
        ):
            return
        
        try:
            # Suppression du fichier
            Path(self.current_session['path']).unlink()
            
            # Mise à jour de l'interface
            self.current_session = None
            self.session_title.configure(text="Aucune session sélectionnée")
            self.session_details.configure(text="Sélectionnez une session dans le panneau de gauche")
            self.journal_text.configure(state="normal")
            self.journal_text.delete("0.0", "end")
            self.journal_text.configure(state="disabled")
            
            # Rafraîchir la liste
            self.load_sessions()
            
            # Mettre à jour le statut
            self.status_label.configure(text="Session supprimée avec succès")
            
        except Exception as e:
            messagebox.showerror(
                "Erreur",
                f"Impossible de supprimer la session : {str(e)}"
            )
    
    def on_export_format_change(self, value):
        """Gérer le changement de format d'export"""
        # Mettre à jour l'interface en fonction du format sélectionné
        if value == "PDF":
            self.export_images.select()  # Les images sont incluses par défaut en PDF
        elif value == "HTML":
            self.export_images.select()  # Les images sont incluses par défaut en HTML
    
    def on_export_period_change(self, value):
        """Gérer le changement de période d'export"""
        # Mettre à jour l'interface en fonction de la période sélectionnée
        if value == "Personnalisé":
            # Ici on pourrait afficher un sélecteur de dates
            pass
    
    def browse_export_path(self):
        """Ouvrir un dialogue pour choisir le chemin d'export"""
        directory = filedialog.askdirectory(
            title="Sélectionner le dossier de destination"
        )
        if directory:
            self.export_path.delete(0, "end")
            self.export_path.insert(0, directory)
    
    def do_export(self):
        """Exécuter l'export avec les paramètres actuels"""
        logger.info("Début de l'export des données")
        
        # Vérifier qu'une session est sélectionnée
        if self.export_period.get() == "Session actuelle" and not self.current_session:
            logger.warning("Tentative d'export sans session sélectionnée")
            messagebox.showwarning(
                "Avertissement",
                "Aucune session sélectionnée à exporter"
            )
            return
        
        # Vérifier le chemin de destination
        export_dir = self.export_path.get()
        if not export_dir:
            logger.warning("Tentative d'export sans spécifier de dossier de destination")
            messagebox.showwarning(
                "Avertissement",
                "Veuillez spécifier un dossier de destination"
            )
            return
        
        try:
            logger.debug(f"Paramètres d'export : Format={self.export_format.get()}, Période={self.export_period.get()}, Destination={export_dir}")
            
            # Créer le répertoire si nécessaire
            Path(export_dir).mkdir(parents=True, exist_ok=True)
            logger.debug(f"Vérification/création du répertoire de destination : {export_dir}")
            
            # Générer le nom du fichier
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            if self.export_period.get() == "Session actuelle":
                base_name = f"{self.current_session['id']}_export_{timestamp}"
                logger.debug(f"Export de la session actuelle : {self.current_session['id']}")
            else:
                base_name = f"apex_logger_export_{timestamp}"
                logger.debug(f"Export multiple pour la période : {self.export_period.get()}")
            
            # Déterminer l'extension en fonction du format
            extension = {
                "Markdown": ".md",
                "HTML": ".html",
                "JSON": ".json",
                "PDF": ".pdf"
            }.get(self.export_format.get(), ".md")
            
            filename = f"{base_name}{extension}"
            full_path = Path(export_dir) / filename
            
            logger.debug(f"Nom du fichier d'export : {filename}")
            logger.debug(f"Chemin complet du fichier d'export : {full_path}")
            
            # Préparation des données
            content = ""  # Initialiser la variable content avec une valeur par défaut
            
            if self.export_period.get() == "Session actuelle":
                logger.debug("Récupération du contenu de la session actuelle pour l'export")
                # DEBUG: Vérifions ce qui est disponible dans self.current_session
                logger.debug(f"Clés disponibles dans self.current_session: {list(self.current_session.keys())}")
                
                content = self.current_session['content']
                logger.debug(f"Taille du contenu à exporter : {len(content)} caractères")
                
                # Traitement du contenu en fonction du format d'export
                if self.export_format.get() == "HTML":
                    logger.debug("Conversion du contenu Markdown en HTML")
                    content = markdown2.markdown(content)
                    # Ajout des balises HTML nécessaires
                    content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>{self.current_session['title']}</title>
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; margin: 20px; }}
        h1 {{ color: #2c3e50; }}
        h2 {{ color: #3498db; }}
        h3 {{ color: #555; }}
        pre {{ background-color: #f8f8f8; border: 1px solid #ddd; padding: 10px; overflow: auto; }}
        code {{ background-color: #f8f8f8; padding: 2px 5px; }}
    </style>
</head>
<body>
    <h1>{self.current_session['title']}</h1>
    <p>Date: {self.current_session['date']}</p>
    <hr>
    {content}
</body>
</html>"""
                elif self.export_format.get() == "JSON":
                    logger.debug("Conversion du contenu en format JSON")
                    export_data = {
                        'id': self.current_session['id'],
                        'title': self.current_session['title'],
                        'date': self.current_session['date'],
                        'content': self.current_session['content'],
                        'export_date': datetime.now().isoformat()
                    }
                    
                    # Ajouter les métadonnées si demandé
                    if self.include_metadata.get():
                        logger.debug("Ajout des métadonnées à l'export JSON")
                        export_data['metadata'] = {
                            'content_size': len(self.current_session['content']),
                            'exported_by': 'APEX Logger GUI',
                            'version': '1.0.2'
                        }
                    
                    content = json.dumps(export_data, indent=2, ensure_ascii=False)
                elif self.export_format.get() == "PDF":
                    # Pour le PDF, on pourrait utiliser une bibliothèque comme reportlab ou wkhtmltopdf
                    # Mais pour cet exemple, nous allons simplement afficher un message d'erreur
                    logger.warning("Export PDF requis mais non implémenté")
                    messagebox.showinfo(
                        "Information",
                        "L'export en format PDF n'est pas encore implémenté.\nVeuillez choisir un autre format d'export."
                    )
                    return
            else:
                # Logique pour exporter plusieurs sessions
                logger.debug("Préparation de l'export multi-sessions")
                
                # Filtrer les sessions en fonction de la période
                filtered_sessions = []
                today = datetime.now().date()
                
                for session in self.sessions:
                    try:
                        session_date_str = session['date']
                        # Analyser la date de la session (format variable)
                        if "2025" in session_date_str:  # Format long avec année
                            session_date = datetime.strptime(session_date_str.split()[0], "%Y-%m-%d").date()
                        else:  # Autre format
                            continue
                        
                        if self.export_period.get() == "Aujourd'hui" and session_date == today:
                            filtered_sessions.append(session)
                        elif self.export_period.get() == "Cette semaine" and (today - session_date).days <= 7:
                            filtered_sessions.append(session)
                        elif self.export_period.get() == "Ce mois" and (today - session_date).days <= 30:
                            filtered_sessions.append(session)
                        elif self.export_period.get() == "Personnalisé":
                            # Pour l'instant, nous ajoutons toutes les sessions - à améliorer avec un sélecteur de dates
                            filtered_sessions.append(session)
                    except Exception as e:
                        logger.warning(f"Erreur lors de l'analyse de la date pour la session {session['id']}: {str(e)}")
                
                logger.debug(f"Nombre de sessions filtrées pour l'export: {len(filtered_sessions)}")
                
                # Créer le contenu combiné pour l'export
                if self.export_format.get() == "Markdown":
                    content = f"# Export APEX Logger - {self.export_period.get()}\n\nDate d'export: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
                    
                    for idx, session in enumerate(filtered_sessions):
                        content += f"\n\n## Session {idx+1}: {session['title']}\n\n"
                        content += f"Date: {session['date']}\n\n"
                        content += session['content']
                        content += "\n\n---\n\n"
                
                elif self.export_format.get() == "HTML":
                    html_content = ""
                    for idx, session in enumerate(filtered_sessions):
                        session_html = markdown2.markdown(session['content'])
                        html_content += f"""<div class="session">
                            <h2>Session {idx+1}: {session['title']}</h2>
                            <p>Date: {session['date']}</p>
                            <div class="session-content">
                                {session_html}
                            </div>
                            <hr>
                        </div>"""
                    
                    content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Export APEX Logger - {self.export_period.get()}</title>
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; margin: 20px; }}
        h1 {{ color: #2c3e50; }}
                h2 {{ color: #3498db; }}
                h3 {{ color: #555; }}
                pre {{ background-color: #f8f8f8; border: 1px solid #ddd; padding: 10px; overflow: auto; }}
                code {{ background-color: #f8f8f8; padding: 2px 5px; }}
                .session {{ margin-bottom: 40px; }}
            </style>
        </head>
        <body>
            <h1>Export APEX Logger - {self.export_period.get()}</h1>
            <p>Date d'export: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            {html_content}
        </body>
        </html>"""
                
                elif self.export_format.get() == "JSON":
                    export_data = {
                        'export_info': {
                            'period': self.export_period.get(),
                            'date': datetime.now().isoformat(),
                            'sessions_count': len(filtered_sessions)
                        },
                        'sessions': []
                    }
                    
                    for session in filtered_sessions:
                        session_data = {
                            'id': session['id'],
                            'title': session['title'],
                            'date': session['date'],
                            'content': session['content']
                        }
                        export_data['sessions'].append(session_data)
                    
                    content = json.dumps(export_data, indent=2, ensure_ascii=False)
                
                elif self.export_format.get() == "PDF":
                    # Message pour format PDF non implémenté
                    logger.warning("Export PDF requis mais non implémenté")
                    messagebox.showinfo(
                        "Information",
                        "L'export en format PDF n'est pas encore implémenté.\nVeuillez choisir un autre format d'export."
                    )
                    return
            
            # Compression si demandée
            if self.compress_export.get():
                logger.debug("Compression de l'export demandée")
                import zipfile
                zip_path = f"{full_path}.zip"
                
                # Créer un fichier temporaire avec le contenu
                temp_file = Path(f"{full_path}.tmp")
                with open(temp_file, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                # Créer le fichier ZIP
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    zipf.write(temp_file, arcname=filename)
                
                # Supprimer le fichier temporaire
                temp_file.unlink()
                
                # Mettre à jour le chemin final
                full_path = Path(zip_path)
                logger.debug(f"Fichier compressé créé : {zip_path}")
            else:
                # Écriture du fichier non compressé
                with open(full_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                logger.debug(f"Fichier d'export créé : {full_path}")
            
            # Mettre à jour le statut
            self.status_label.configure(
                text=f"Export réussi: {full_path}"
            )
            
            logger.info(f"Export réussi : {full_path}")
            
            # Demander à l'utilisateur s'il souhaite ouvrir le fichier exporté
            if messagebox.askyesno(
                "Export terminé",
                f"L'export a été enregistré sous:\n{full_path}\n\nSouhaitez-vous l'ouvrir maintenant?"
            ):
                # Ouvrir le fichier avec l'application par défaut
                logger.debug(f"Ouverture du fichier exporté : {full_path}")
                webbrowser.open(full_path)
                
        except Exception as e:
            logger.critical(f"Erreur lors de l'export des données : {str(e)}")
            logger.critical(traceback.format_exc())
            messagebox.showerror(
                "Erreur",
                f"Impossible d'exporter les données : {str(e)}"
            )
    
    def load_report(self, value):
        """Charger un rapport spécifique"""
        # Effacer le rapport actuel
        self.report_display.configure(state="normal")
        self.report_display.delete("0.0", "end")
        
        # Générer le contenu du rapport en fonction du type sélectionné
        if value == "Activité":
            self.generate_activity_report()
        elif value == "Performance":
            self.generate_performance_report()
        elif value == "Qualité":
            self.generate_quality_report()
        elif value == "Utilisation":
            self.generate_usage_report()
        
        self.report_display.configure(state="disabled")
    
    def generate_activity_report(self):
        """Générer un rapport d'activité"""
        report = "# Rapport d'Activité\n\n"
        report += f"*Généré le {datetime.now().strftime('%Y-%m-%d à %H:%M:%S')}*\n\n"
        
        # Calculer des statistiques à partir des entrées du journal
        if not self.journal_entries:
            report += "Aucune donnée disponible pour générer le rapport.\n"
        else:
            # Nombre total d'interactions
            total_interactions = len(self.journal_entries)
            report += f"## Nombre total d'interactions: {total_interactions}\n\n"
            
            # Répartition par agent
            agents = {}
            for entry in self.journal_entries:
                agent = entry['agent']
                if agent not in agents:
                    agents[agent] = 0
                agents[agent] += 1
            
            report += "## Répartition par agent\n\n"
            for agent, count in sorted(agents.items(), key=lambda x: x[1], reverse=True):
                percentage = (count / total_interactions) * 100
                report += f"- {agent}: {count} ({percentage:.1f}%)\n"
            
            report += "\n"
            
            # Répartition par évaluation
            ratings = {}
            for entry in self.journal_entries:
                rating = entry['rating']
                if rating not in ratings:
                    ratings[rating] = 0
                ratings[rating] += 1
            
            report += "## Répartition par évaluation\n\n"
            for rating, count in sorted(ratings.items()):
                percentage = (count / total_interactions) * 100
                report += f"- {rating}: {count} ({percentage:.1f}%)\n"
            
            report += "\n"
            
            # Distribution temporelle
            # (Implémentation simplifiée)
            report += "## Distribution temporelle\n\n"
            report += "Une visualisation de la distribution temporelle serait normalement affichée ici.\n"
        
        self.report_display.insert("0.0", report)
    
    def generate_performance_report(self):
        """Générer un rapport de performance"""
        report = "# Rapport de Performance\n\n"
        report += f"*Généré le {datetime.now().strftime('%Y-%m-%d à %H:%M:%S')}*\n\n"
        report += "Ce rapport analyse les performances du système de journalisation.\n\n"
        
        # Exemple de contenu du rapport de performance
        report += "## Métriques de performance\n\n"
        report += "- **Temps moyen de réponse**: 1.2s\n"
        report += "- **Taille moyenne des journaux**: 25KB\n"
        report += "- **Nombre de sessions par jour**: 8.5\n\n"
        
        report += "## Graphique de performance\n\n"
        report += "Un graphique de performance serait normalement affiché ici.\n"
        
        self.report_display.insert("0.0", report)
    
    def generate_quality_report(self):
        """Générer un rapport de qualité"""
        report = "# Rapport de Qualité\n\n"
        report += f"*Généré le {datetime.now().strftime('%Y-%m-%d à %H:%M:%S')}*\n\n"
        report += "Ce rapport analyse la qualité des interactions.\n\n"
        
        # Exemple de contenu du rapport de qualité
        if self.journal_entries:
            # Calculer le score moyen des évaluations
            ratings = [entry['rating'] for entry in self.journal_entries]
            rating_scores = {'++': 2, '+': 1, '0': 0, '-': -1, '--': -2}
            
            score_sum = sum(rating_scores.get(r, 0) for r in ratings if r in rating_scores)
            valid_ratings = sum(1 for r in ratings if r in rating_scores)
            
            if valid_ratings > 0:
                avg_score = score_sum / valid_ratings
                report += f"## Score de qualité moyen: {avg_score:.2f}\n\n"
            
            # Top des contextes
            contexts = {}
            for entry in self.journal_entries:
                context = entry['context']
                if context not in contexts:
                    contexts[context] = 0
                contexts[context] += 1
            
            report += "## Top des contextes\n\n"
            for context, count in sorted(contexts.items(), key=lambda x: x[1], reverse=True)[:5]:
                report += f"- {context}: {count} interactions\n"
        else:
            report += "Aucune donnée disponible pour générer le rapport.\n"
        
        self.report_display.insert("0.0", report)
    
    def generate_usage_report(self):
        """Générer un rapport d'utilisation"""
        report = "# Rapport d'Utilisation\n\n"
        report += f"*Généré le {datetime.now().strftime('%Y-%m-%d à %H:%M:%S')}*\n\n"
        report += "Ce rapport analyse les patterns d'utilisation du système de journalisation.\n\n"
        
        # Exemple de contenu du rapport d'utilisation
        report += "## Statistiques d'utilisation\n\n"
        report += "- **Nombre total de sessions**: {}\n".format(len(self.sessions))
        report += "- **Nombre total d'entrées de journal**: {}\n".format(len(self.journal_entries))
        
        # Top des utilisateurs (basé sur les agents)
        if self.journal_entries:
            agents = {}
            for entry in self.journal_entries:
                agent = entry['agent']
                if agent not in agents:
                    agents[agent] = 0
                agents[agent] += 1
            
            report += "\n## Top des utilisateurs\n\n"
            for agent, count in sorted(agents.items(), key=lambda x: x[1], reverse=True)[:5]:
                report += f"- {agent}: {count} interactions\n"
        
        self.report_display.insert("0.0", report)
    
    def check_logger_status(self):
        """Vérifie le statut actuel des différents loggers (Cursor et VSCode)"""
        try:
            logger.debug("Vérification du statut des loggers")
            current_time = datetime.now()
            
            # Vérification du logger Cursor
            cursor_logs = list(self.logs_dir.glob("cursor_session_*.md"))
            cursor_journal = self.logs_dir / "apex-cursor-journal.md"
            
            if cursor_journal.exists():
                last_modified = datetime.fromtimestamp(cursor_journal.stat().st_mtime)
                time_diff = (current_time - last_modified).total_seconds()
                
                # Le logger est considéré actif si mis à jour dans les 5 dernières minutes
                if time_diff < 300:  # 5 minutes en secondes
                    self.logger_status["cursor"]["active"] = True
                    self.logger_status["cursor"]["status_text"] = "Actif"
                    self.logger_status["cursor"]["last_check"] = current_time
                else:
                    self.logger_status["cursor"]["active"] = False
                    self.logger_status["cursor"]["status_text"] = f"Inactif (dernière activité: {last_modified.strftime('%H:%M:%S')})"
                    self.logger_status["cursor"]["last_check"] = current_time
            else:
                self.logger_status["cursor"]["active"] = False
                self.logger_status["cursor"]["status_text"] = "Journal non trouvé"
                self.logger_status["cursor"]["last_check"] = current_time
            
            # Vérification du logger VSCode
            vscode_logs = list(self.logs_dir.glob("vscode_session_*.md"))
            vscode_journal = self.logs_dir / "apex-vscode-journal.md"
            
            if vscode_journal.exists():
                last_modified = datetime.fromtimestamp(vscode_journal.stat().st_mtime)
                time_diff = (current_time - last_modified).total_seconds()
                
                # Le logger est considéré actif si mis à jour dans les 5 dernières minutes
                if time_diff < 300:  # 5 minutes en secondes
                    self.logger_status["vscode"]["active"] = True
                    self.logger_status["vscode"]["status_text"] = "Actif"
                    self.logger_status["vscode"]["last_check"] = current_time
                else:
                    self.logger_status["vscode"]["active"] = False
                    self.logger_status["vscode"]["status_text"] = f"Inactif (dernière activité: {last_modified.strftime('%H:%M:%S')})"
                    self.logger_status["vscode"]["last_check"] = current_time
            else:
                self.logger_status["vscode"]["active"] = False
                self.logger_status["vscode"]["status_text"] = "Journal non trouvé"
                self.logger_status["vscode"]["last_check"] = current_time
            
            # Mise à jour des indicateurs visuels
            self.update_status_indicators()
            
            logger.debug(f"Statut des loggers - Cursor: {self.logger_status['cursor']['status_text']}, VSCode: {self.logger_status['vscode']['status_text']}")
            
        except Exception as e:
            logger.error(f"Erreur lors de la vérification du statut des loggers: {str(e)}")
            logger.error(traceback.format_exc())
    
    def update_status_indicators(self):
        """Met à jour les indicateurs visuels d'état des loggers"""
        # Indicateur Cursor
        if self.logger_status["cursor"]["active"]:
            self.cursor_indicator.configure(text_color="#32CD32")  # Vert vif
            tooltip_text = f"Cursor logger: {self.logger_status['cursor']['status_text']}"
        else:
            self.cursor_indicator.configure(text_color="#FF6347")  # Rouge tomate
            tooltip_text = f"Cursor logger: {self.logger_status['cursor']['status_text']}"
            
        # Mise à jour du tooltip (pas directement supporté par CTk, mais on pourrait implémenter cela)
        self.cursor_indicator.configure(text="⬤")
        
        # Indicateur VSCode
        if self.logger_status["vscode"]["active"]:
            self.vscode_indicator.configure(text_color="#1E90FF")  # Bleu Dodger
            tooltip_text = f"VSCode logger: {self.logger_status['vscode']['status_text']}"
        else:
            self.vscode_indicator.configure(text_color="#FF6347")  # Rouge tomate
            tooltip_text = f"VSCode logger: {self.logger_status['vscode']['status_text']}"
            
        # Mise à jour du tooltip
        self.vscode_indicator.configure(text="⬤")
    
    def periodic_status_check(self):
        """Vérifie périodiquement le statut des loggers"""
        while True:
            time.sleep(30)  # Vérifier toutes les 30 secondes
            try:
                # Exécuter dans le thread principal
                self.after(0, self.check_logger_status)
            except Exception as e:
                logger.error(f"Erreur lors de la vérification périodique du statut: {str(e)}")
                logger.error(traceback.format_exc())

    def load_history(self, _=None):
        """Charge l'historique des dernières demandes"""
        try:
            # Nettoyer l'affichage actuel
            for widget in self.history_frame.winfo_children():
                widget.destroy()
            
            # Charger les données si nécessaire
            if not self.journal_entries:
                self.load_journal()
            
            # Définir les filtres
            source_filter = self.history_source.get()
            count_limit = int(self.history_count.get())
            
            # Filtrer les entrées
            filtered_entries = []
            for entry in self.journal_entries:
                if source_filter == "Tous" or source_filter == entry['source']:
                    filtered_entries.append(entry)
                
                # Limiter au nombre demandé
                if len(filtered_entries) >= count_limit:
                    break
            
            # Afficher un message si aucune entrée
            if not filtered_entries:
                no_data_label = ctk.CTkLabel(
                    self.history_frame,
                    text="Aucune demande trouvée",
                    font=ctk.CTkFont(size=14, weight="bold"),
                    text_color="gray"
                )
                no_data_label.grid(row=0, column=0, padx=20, pady=20)
                return
            
            # Afficher les entrées
            for idx, entry in enumerate(filtered_entries):
                # Création du cadre pour cette entrée
                entry_frame = ctk.CTkFrame(self.history_frame)
                entry_frame.grid(row=idx, column=0, sticky="ew", padx=10, pady=10)
                entry_frame.grid_columnconfigure(1, weight=1)
                
                # Numéro de la demande
                num_label = ctk.CTkLabel(
                    entry_frame,
                    text=f"#{idx+1}",
                    font=ctk.CTkFont(size=14, weight="bold"),
                    width=30
                )
                num_label.grid(row=0, column=0, rowspan=2, padx=10, pady=10)
                
                # En-tête avec date et source
                header_frame = ctk.CTkFrame(entry_frame, fg_color="transparent")
                header_frame.grid(row=0, column=1, sticky="ew", padx=10, pady=(10, 0))
                header_frame.grid_columnconfigure(0, weight=1)
                
                date_label = ctk.CTkLabel(
                    header_frame,
                    text=entry['timestamp'],
                    font=ctk.CTkFont(size=12),
                    anchor="w"
                )
                date_label.grid(row=0, column=0, sticky="w")
                
                source_label = ctk.CTkLabel(
                    header_frame,
                    text=entry['source'],
                    font=ctk.CTkFont(size=12),
                    text_color="#007BFF" if entry['source'] == "VSCode" else "#FF6347",
                    anchor="e"
                )
                source_label.grid(row=0, column=1, sticky="e", padx=10)
                
                # Contenu de la demande
                prompt_frame = ctk.CTkFrame(entry_frame)
                prompt_frame.grid(row=1, column=1, sticky="ew", padx=10, pady=(5, 10))
                prompt_frame.grid_columnconfigure(0, weight=1)
                
                prompt_text = ctk.CTkTextbox(
                    prompt_frame,
                    height=60,
                    wrap="word",
                    font=ctk.CTkFont(size=12)
                )
                prompt_text.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
                prompt_text.insert("0.0", entry['prompt'])
                prompt_text.configure(state="disabled")
                
                # Bouton pour voir la session complète
                view_btn = ctk.CTkButton(
                    entry_frame,
                    text="Voir la session",
                    command=lambda s=entry['session_id']: self.find_and_load_session(s),
                    width=120,
                    height=30
                )
                view_btn.grid(row=2, column=1, padx=10, pady=5, sticky="e")
                
            self.status_label.configure(
                text=f"{len(filtered_entries)} demandes affichées"
            )
            
        except Exception as e:
            logger.error(f"Erreur lors du chargement de l'historique: {str(e)}")
            logger.error(traceback.format_exc())
            messagebox.showerror(
                "Erreur",
                f"Impossible de charger l'historique : {str(e)}"
            )

if __name__ == "__main__":
    app = LoggerApp()
    app.mainloop()