import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from pathlib import Path
import json
from typing import Dict, Any
import logging
from datetime import datetime
import subprocess
import sys

# Configuration du logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('orchestrator.log'),
        logging.StreamHandler()
    ]
)

class ModernFrame(ttk.Frame):
    def __init__(self, master, title: str, **kwargs):
        super().__init__(master, **kwargs)
        
        # Style du cadre
        self.configure(padding="10")
        
        # Titre avec style moderne
        title_label = ttk.Label(self, 
                              text=title,
                              style='Title.TLabel',
                              anchor='center')
        title_label.pack(fill='x', pady=(0, 10))

class ApexOrchestratorGUI:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("APEX Orchestrator")
        self.window.geometry("1024x768")
        
        # État de l'application
        self.current_project = None
        self.projects = {}
        self.is_modified = False
        self.vba_modules = []
        self.test_results = {}
        
        # Chargement de l'icône
        icon_path = Path("assets/icon.ico")
        if icon_path.exists():
            self.window.iconbitmap(icon_path)
        
        # Configuration des styles
        self.setup_styles()
        
        # Menu principal
        self.create_menu()
        
        # Layout principal
        self.create_layout()
        
        # Statut
        self.create_status_bar()
        
        # Chargement initial
        self.load_config()
        
        # Gestionnaires d'événements
        self.bind_events()
        
        # Protocole de fermeture
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_styles(self):
        """Configuration des styles."""
        style = ttk.Style()
        
        # Couleurs APEX
        style.configure('.',
                       background='#2B2D42',
                       foreground='#EDF2F4')
        
        # Style des titres
        style.configure('Title.TLabel',
                       font=('Segoe UI', 12, 'bold'),
                       foreground='#EF233C',
                       background='#2B2D42',
                       padding=5)
        
        # Style des boutons
        style.configure('Action.TButton',
                       font=('Segoe UI', 10),
                       background='#EF233C',
                       foreground='#EDF2F4',
                       padding=5)
        style.map('Action.TButton',
                 background=[('active', '#D90429')],
                 foreground=[('active', '#FFFFFF')])
        
        # Style des cadres
        style.configure('Card.TFrame',
                       background='#8D99AE',
                       relief='raised',
                       borderwidth=1)
        
        # Style des arbres
        style.configure('Treeview',
                       background='#2B2D42',
                       foreground='#EDF2F4',
                       fieldbackground='#2B2D42',
                       font=('Segoe UI', 9))
        style.configure('Treeview.Heading',
                       background='#8D99AE',
                       foreground='#2B2D42',
                       font=('Segoe UI', 10, 'bold'))
        style.map('Treeview',
                 background=[('selected', '#EF233C')],
                 foreground=[('selected', '#FFFFFF')])
        
        # Style des onglets
        style.configure('TNotebook.Tab',
                       background='#8D99AE',
                       foreground='#2B2D42',
                       padding=[10, 2],
                       font=('Segoe UI', 9))
        style.map('TNotebook.Tab',
                 background=[('selected', '#EF233C')],
                 foreground=[('selected', '#FFFFFF')])
        
        # Style de la barre de statut
        style.configure('Status.TLabel',
                       background='#2B2D42',
                       foreground='#8D99AE',
                       font=('Segoe UI', 9))
        
        # Style des menus
        self.window.option_add('*Menu.font', ('Segoe UI', 9))
        self.window.option_add('*Menu.background', '#2B2D42')
        self.window.option_add('*Menu.foreground', '#EDF2F4')
        self.window.option_add('*Menu.selectColor', '#EF233C')

    def create_menu(self):
        """Création du menu principal."""
        menubar = tk.Menu(self.window)
        
        # Menu Fichier
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Nouveau projet", command=self.new_project)
        file_menu.add_command(label="Ouvrir...", command=self.open_project)
        file_menu.add_command(label="Enregistrer", command=self.save_project)
        file_menu.add_separator()
        file_menu.add_command(label="Quitter", command=self.on_close)
        menubar.add_cascade(label="Fichier", menu=file_menu)
        
        # Menu APEX
        apex_menu = tk.Menu(menubar, tearoff=0)
        apex_menu.add_command(label="Scanner les modules", command=self.scan_vba_modules)
        apex_menu.add_command(label="Exécuter les tests", command=self.run_tests)
        apex_menu.add_command(label="Générer la documentation", command=self.generate_docs)
        apex_menu.add_separator()
        apex_menu.add_command(label="Vérifier l'encodage", command=self.check_encoding)
        menubar.add_cascade(label="APEX", menu=apex_menu)
        
        # Menu Outils
        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Configuration", command=self.show_config)
        tools_menu.add_command(label="Logs", command=self.show_logs)
        tools_menu.add_separator()
        tools_menu.add_command(label="Analyser le code", command=self.analyze_code)
        menubar.add_cascade(label="Outils", menu=tools_menu)
        
        # Menu Aide
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Documentation", command=self.show_docs)
        help_menu.add_command(label="À propos", command=self.show_about)
        menubar.add_cascade(label="Aide", menu=help_menu)
        
        self.window.config(menu=menubar)

    def create_layout(self):
        """Création du layout principal."""
        # Container principal
        self.main_container = ttk.PanedWindow(self.window, orient='horizontal')
        self.main_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Panneau de gauche (navigation)
        left_panel = ModernFrame(self.main_container, "Navigation")
        self.main_container.add(left_panel, weight=1)
        
        # Barre d'outils de navigation
        nav_toolbar = ttk.Frame(left_panel)
        nav_toolbar.pack(fill='x', pady=(0, 5))
        
        ttk.Button(nav_toolbar,
                  text="+ Nouveau",
                  style='Action.TButton',
                  command=self.new_project).pack(side='left', padx=2)
        ttk.Button(nav_toolbar,
                  text="Rafraîchir",
                  style='Action.TButton',
                  command=self.update_projects_tree).pack(side='left', padx=2)
        
        # Liste des projets avec style
        self.projects_tree = ttk.Treeview(left_panel,
                                        columns=('status',),
                                        show='tree headings',
                                        style='Treeview')
        self.projects_tree.heading('status', text='Statut')
        self.projects_tree.pack(fill='both', expand=True)
        
        # Panneau central (contenu)
        center_panel = ModernFrame(self.main_container, "Contenu")
        self.main_container.add(center_panel, weight=3)
        
        # Notebook pour les onglets
        self.notebook = ttk.Notebook(center_panel, style='TNotebook')
        self.notebook.pack(fill='both', expand=True)
        
        # Onglets
        self.dashboard_tab = ModernFrame(self.notebook, "Tableau de bord")
        self.create_dashboard()
        self.notebook.add(self.dashboard_tab, text="Tableau de bord")
        
        self.tests_tab = ModernFrame(self.notebook, "Tests")
        self.create_tests_view()
        self.notebook.add(self.tests_tab, text="Tests")
        
        self.modules_tab = ModernFrame(self.notebook, "Modules")
        self.create_modules_view()
        self.notebook.add(self.modules_tab, text="Modules")
        
        self.logs_tab = ModernFrame(self.notebook, "Logs")
        self.create_logs_view()
        self.notebook.add(self.logs_tab, text="Logs")
        
        # Panneau de droite (détails)
        right_panel = ModernFrame(self.main_container, "Détails")
        self.main_container.add(right_panel, weight=1)
        
        # Barre d'outils des propriétés
        props_toolbar = ttk.Frame(right_panel)
        props_toolbar.pack(fill='x', pady=(0, 5))
        
        ttk.Button(props_toolbar,
                  text="Éditer",
                  style='Action.TButton',
                  command=self.edit_property).pack(side='left', padx=2)
        ttk.Button(props_toolbar,
                  text="Supprimer",
                  style='Action.TButton',
                  command=self.delete_property).pack(side='left', padx=2)
        
        # Zone de propriétés avec style
        self.props_tree = ttk.Treeview(right_panel,
                                     columns=('value',),
                                     show='tree headings',
                                     style='Treeview')
        self.props_tree.heading('value', text='Valeur')
        self.props_tree.pack(fill='both', expand=True)

    def create_dashboard(self):
        """Création du tableau de bord."""
        # Statistiques
        stats_frame = ttk.LabelFrame(self.dashboard_tab, text="Statistiques", padding=10)
        stats_frame.pack(fill='x', pady=5)
        
        self.stats_labels = {
            'modules': ttk.Label(stats_frame, text="Modules: 0"),
            'tests': ttk.Label(stats_frame, text="Tests: 0"),
            'coverage': ttk.Label(stats_frame, text="Couverture: 0%"),
            'quality': ttk.Label(stats_frame, text="Qualité: N/A")
        }
        
        for label in self.stats_labels.values():
            label.pack(side='left', padx=10)
        
        # Dernières actions
        actions_frame = ttk.LabelFrame(self.dashboard_tab, text="Dernières actions", padding=10)
        actions_frame.pack(fill='both', expand=True, pady=5)
        
        self.actions_list = ttk.Treeview(actions_frame,
                                       columns=('date', 'type', 'status'),
                                       show='headings')
        self.actions_list.heading('date', text='Date')
        self.actions_list.heading('type', text='Type')
        self.actions_list.heading('status', text='Statut')
        self.actions_list.pack(fill='both', expand=True)

    def create_tests_view(self):
        """Création de la vue des tests."""
        # Barre d'outils
        toolbar = ttk.Frame(self.tests_tab)
        toolbar.pack(fill='x', pady=5)
        
        ttk.Button(toolbar,
                  text="Exécuter tous",
                  style='Action.TButton',
                  command=self.run_tests).pack(side='left', padx=2)
        ttk.Button(toolbar,
                  text="Exécuter sélection",
                  style='Action.TButton',
                  command=self.run_selected_tests).pack(side='left', padx=2)
        
        # Liste des tests
        self.tests_tree = ttk.Treeview(self.tests_tab,
                                     columns=('module', 'status', 'duration'),
                                     show='headings')
        self.tests_tree.heading('module', text='Module')
        self.tests_tree.heading('status', text='Statut')
        self.tests_tree.heading('duration', text='Durée')
        self.tests_tree.pack(fill='both', expand=True)

    def create_modules_view(self):
        """Création de la vue des modules."""
        # Barre d'outils
        toolbar = ttk.Frame(self.modules_tab)
        toolbar.pack(fill='x', pady=5)
        
        ttk.Button(toolbar,
                  text="Scanner",
                  style='Action.TButton',
                  command=self.scan_vba_modules).pack(side='left', padx=2)
        ttk.Button(toolbar,
                  text="Analyser",
                  style='Action.TButton',
                  command=self.analyze_code).pack(side='left', padx=2)
        
        # Liste des modules
        self.modules_tree = ttk.Treeview(self.modules_tab,
                                       columns=('type', 'lines', 'quality'),
                                       show='headings')
        self.modules_tree.heading('type', text='Type')
        self.modules_tree.heading('lines', text='Lignes')
        self.modules_tree.heading('quality', text='Qualité')
        self.modules_tree.pack(fill='both', expand=True)

    def create_logs_view(self):
        """Création de la vue des logs."""
        # Zone de texte pour les logs
        self.logs_text = tk.Text(self.logs_tab, wrap='none')
        self.logs_text.pack(fill='both', expand=True)
        
        # Scrollbars
        x_scroll = ttk.Scrollbar(self.logs_tab, orient='horizontal', command=self.logs_text.xview)
        y_scroll = ttk.Scrollbar(self.logs_tab, orient='vertical', command=self.logs_text.yview)
        self.logs_text.configure(xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)
        
        x_scroll.pack(side='bottom', fill='x')
        y_scroll.pack(side='right', fill='y')

    def create_status_bar(self):
        """Création de la barre de statut."""
        self.status_frame = ttk.Frame(self.window, style='Card.TFrame')
        self.status_frame.pack(fill='x', side='bottom')
        
        self.status_label = ttk.Label(self.status_frame,
                                    text="Prêt",
                                    style='Status.TLabel',
                                    anchor='w',
                                    padding=5)
        self.status_label.pack(side='left')
        
        self.version_label = ttk.Label(self.status_frame,
                                     text="v1.0.0",
                                     style='Status.TLabel',
                                     anchor='e',
                                     padding=5)
        self.version_label.pack(side='right')

    def bind_events(self):
        """Association des gestionnaires d'événements."""
        self.projects_tree.bind('<<TreeviewSelect>>', self.on_project_select)
        self.notebook.bind('<<NotebookTabChanged>>', self.on_tab_change)
        self.props_tree.bind('<<TreeviewSelect>>', self.on_property_select)

    def load_config(self):
        """Chargement de la configuration."""
        try:
            with open('config/orchestrator_config.json', 'r') as f:
                config = json.load(f)
                self.projects = config.get('projects', {})
                self.update_projects_tree()
                logging.info("Configuration chargée avec succès")
        except Exception as e:
            logging.error(f"Erreur lors du chargement de la configuration: {e}")
            messagebox.showerror("Erreur", 
                               "Impossible de charger la configuration")

    def save_config(self):
        """Sauvegarde de la configuration."""
        try:
            config = {
                'projects': self.projects,
                'last_update': datetime.now().isoformat()
            }
            with open('config/orchestrator_config.json', 'w') as f:
                json.dump(config, f, indent=4)
            self.is_modified = False
            self.update_status("Configuration sauvegardée")
            logging.info("Configuration sauvegardée avec succès")
        except Exception as e:
            logging.error(f"Erreur lors de la sauvegarde: {e}")
            messagebox.showerror("Erreur", 
                               "Impossible de sauvegarder la configuration")

    def update_projects_tree(self):
        """Mise à jour de l'arbre des projets."""
        self.projects_tree.delete(*self.projects_tree.get_children())
        for project_id, project in self.projects.items():
            self.projects_tree.insert('', 'end', 
                                   text=project['name'],
                                   values=(project['status'],),
                                   tags=(project['status'],))

    def update_properties(self, item_id):
        """Mise à jour des propriétés."""
        self.props_tree.delete(*self.props_tree.get_children())
        if item_id in self.projects:
            project = self.projects[item_id]
            for key, value in project.items():
                self.props_tree.insert('', 'end',
                                    text=key,
                                    values=(str(value),))

    def update_status(self, message: str):
        """Mise à jour du message de statut."""
        self.status_label.config(text=message)
        logging.info(message)

    # Gestionnaires d'événements
    def on_project_select(self, event):
        """Gestion de la sélection d'un projet."""
        selection = self.projects_tree.selection()
        if selection:
            self.current_project = selection[0]
            self.update_properties(self.current_project)
            self.update_status(f"Projet sélectionné: {self.current_project}")

    def on_tab_change(self, event):
        """Gestion du changement d'onglet."""
        tab = self.notebook.select()
        tab_text = self.notebook.tab(tab, "text")
        self.update_status(f"Onglet actif: {tab_text}")

    def on_property_select(self, event):
        """Gestion de la sélection d'une propriété."""
        selection = self.props_tree.selection()
        if selection:
            item = self.props_tree.item(selection[0])
            self.update_status(f"Propriété: {item['text']} = {item['values'][0]}")

    def on_close(self):
        """Gestion de la fermeture de l'application."""
        if self.is_modified:
            if messagebox.askyesno("Sauvegarder",
                                 "Des modifications non sauvegardées existent. Sauvegarder?"):
                self.save_config()
        self.window.destroy()

    # Actions du menu
    def new_project(self):
        """Création d'un nouveau projet."""
        name = simpledialog.askstring("Nouveau projet", 
                                    "Nom du projet:")
        if name:
            project_id = str(len(self.projects) + 1)
            self.projects[project_id] = {
                'name': name,
                'status': 'new',
                'created_at': datetime.now().isoformat()
            }
            self.is_modified = True
            self.update_projects_tree()
            self.update_status(f"Nouveau projet créé: {name}")

    def open_project(self):
        """Ouverture d'un projet existant."""
        filename = filedialog.askopenfilename(
            title="Ouvrir un projet",
            filetypes=[("Fichiers JSON", "*.json")]
        )
        if filename:
            try:
                with open(filename, 'r') as f:
                    project = json.load(f)
                    project_id = str(len(self.projects) + 1)
                    self.projects[project_id] = project
                    self.is_modified = True
                    self.update_projects_tree()
                    self.update_status(f"Projet ouvert: {project['name']}")
            except Exception as e:
                logging.error(f"Erreur lors de l'ouverture du projet: {e}")
                messagebox.showerror("Erreur",
                                   "Impossible d'ouvrir le projet")

    def save_project(self):
        """Sauvegarde du projet courant."""
        if self.current_project:
            try:
                project = self.projects[self.current_project]
                filename = filedialog.asksaveasfilename(
                    title="Enregistrer le projet",
                    defaultextension=".json",
                    filetypes=[("Fichiers JSON", "*.json")]
                )
                if filename:
                    with open(filename, 'w') as f:
                        json.dump(project, f, indent=4)
                    self.update_status(f"Projet sauvegardé: {project['name']}")
            except Exception as e:
                logging.error(f"Erreur lors de la sauvegarde du projet: {e}")
                messagebox.showerror("Erreur",
                                   "Impossible de sauvegarder le projet")

    def show_config(self):
        """Affichage de la configuration."""
        config_window = tk.Toplevel(self.window)
        config_window.title("Configuration")
        config_window.geometry("600x400")
        
        text = tk.Text(config_window)
        text.pack(fill='both', expand=True)
        
        try:
            with open('config/orchestrator_config.json', 'r') as f:
                config = json.load(f)
                text.insert('1.0', json.dumps(config, indent=4))
        except Exception as e:
            text.insert('1.0', f"Erreur: {str(e)}")

    def show_logs(self):
        """Affichage des logs."""
        logs_window = tk.Toplevel(self.window)
        logs_window.title("Logs")
        logs_window.geometry("800x600")
        
        text = tk.Text(logs_window)
        text.pack(fill='both', expand=True)
        
        try:
            with open('orchestrator.log', 'r') as f:
                text.insert('1.0', f.read())
        except Exception as e:
            text.insert('1.0', f"Erreur: {str(e)}")

    def show_docs(self):
        """Affichage de la documentation."""
        messagebox.showinfo("Documentation",
                          "La documentation est disponible sur le wiki du projet.")

    def show_about(self):
        """Affichage des informations sur l'application."""
        messagebox.showinfo("À propos",
                          "APEX Orchestrator v1.0.0\n"
                          "© 2024 APEX Framework")

    def edit_property(self):
        """Édition d'une propriété."""
        selection = self.props_tree.selection()
        if selection and self.current_project:
            item = self.props_tree.item(selection[0])
            key = item['text']
            value = item['values'][0]
            new_value = simpledialog.askstring("Éditer la propriété",
                                             f"Nouvelle valeur pour {key}:",
                                             initialvalue=value)
            if new_value is not None:
                self.projects[self.current_project][key] = new_value
                self.is_modified = True
                self.update_properties(self.current_project)
                self.update_status(f"Propriété modifiée: {key}")

    def delete_property(self):
        """Suppression d'une propriété."""
        selection = self.props_tree.selection()
        if selection and self.current_project:
            item = self.props_tree.item(selection[0])
            key = item['text']
            if messagebox.askyesno("Supprimer la propriété",
                                 f"Supprimer la propriété {key}?"):
                del self.projects[self.current_project][key]
                self.is_modified = True
                self.update_properties(self.current_project)
                self.update_status(f"Propriété supprimée: {key}")

    # Fonctionnalités APEX
    def scan_vba_modules(self):
        """Scan des modules VBA."""
        try:
            # Simulation du scan
            self.vba_modules = [
                {'name': 'clsExcelWorkbookAccessor', 'type': 'Class', 'lines': 150},
                {'name': 'clsExcelSheetAccessor', 'type': 'Class', 'lines': 120},
                {'name': 'TestWorkbookAccessor', 'type': 'TestModule', 'lines': 200}
            ]
            
            # Mise à jour de la vue des modules
            self.modules_tree.delete(*self.modules_tree.get_children())
            for module in self.vba_modules:
                self.modules_tree.insert('', 'end',
                                      values=(module['type'],
                                             module['lines'],
                                             'A'))
            
            self.update_stats()
            self.update_status("Modules VBA scannés avec succès")
            self.log_action("Scan des modules", "success")
        except Exception as e:
            logging.error(f"Erreur lors du scan des modules: {e}")
            messagebox.showerror("Erreur", "Impossible de scanner les modules")
            self.log_action("Scan des modules", "error")

    def run_tests(self):
        """Exécution des tests."""
        try:
            # Simulation des tests
            self.test_results = {
                'TestWorkbookAccessor': {'status': 'success', 'duration': '1.2s'},
                'TestSheetAccessor': {'status': 'success', 'duration': '0.8s'},
                'TestRangeAccessor': {'status': 'failed', 'duration': '0.5s'}
            }
            
            # Mise à jour de la vue des tests
            self.tests_tree.delete(*self.tests_tree.get_children())
            for test, result in self.test_results.items():
                self.tests_tree.insert('', 'end',
                                    values=(test,
                                           result['status'],
                                           result['duration']))
            
            self.update_stats()
            self.update_status("Tests exécutés")
            self.log_action("Exécution des tests", "success")
        except Exception as e:
            logging.error(f"Erreur lors de l'exécution des tests: {e}")
            messagebox.showerror("Erreur", "Impossible d'exécuter les tests")
            self.log_action("Exécution des tests", "error")

    def run_selected_tests(self):
        """Exécution des tests sélectionnés."""
        selection = self.tests_tree.selection()
        if selection:
            try:
                for item in selection:
                    test = self.tests_tree.item(item)['values'][0]
                    # Simulation du test
                    result = {'status': 'success', 'duration': '0.5s'}
                    self.tests_tree.item(item, values=(test,
                                                     result['status'],
                                                     result['duration']))
                self.update_status("Tests sélectionnés exécutés")
                self.log_action("Exécution des tests sélectionnés", "success")
            except Exception as e:
                logging.error(f"Erreur lors de l'exécution des tests: {e}")
                messagebox.showerror("Erreur", 
                                   "Impossible d'exécuter les tests sélectionnés")
                self.log_action("Exécution des tests sélectionnés", "error")

    def generate_docs(self):
        """Génération de la documentation."""
        try:
            # Simulation de la génération
            self.update_status("Documentation générée")
            self.log_action("Génération de la documentation", "success")
            messagebox.showinfo("Documentation",
                              "Documentation générée avec succès")
        except Exception as e:
            logging.error(f"Erreur lors de la génération de la documentation: {e}")
            messagebox.showerror("Erreur",
                               "Impossible de générer la documentation")
            self.log_action("Génération de la documentation", "error")

    def check_encoding(self):
        """Vérification de l'encodage des fichiers."""
        try:
            # Simulation de la vérification
            results = {
                'total': 10,
                'utf8': 8,
                'other': 2
            }
            
            message = (f"Résultats de l'analyse:\n"
                      f"Total: {results['total']} fichiers\n"
                      f"UTF-8: {results['utf8']} fichiers\n"
                      f"Autres: {results['other']} fichiers")
            
            messagebox.showinfo("Encodage", message)
            self.update_status("Vérification de l'encodage terminée")
            self.log_action("Vérification de l'encodage", "success")
        except Exception as e:
            logging.error(f"Erreur lors de la vérification de l'encodage: {e}")
            messagebox.showerror("Erreur",
                               "Impossible de vérifier l'encodage")
            self.log_action("Vérification de l'encodage", "error")

    def analyze_code(self):
        """Analyse du code."""
        try:
            # Simulation de l'analyse
            results = {
                'complexity': 'B',
                'maintainability': 'A',
                'documentation': 'A',
                'test_coverage': '95%'
            }
            
            message = (f"Résultats de l'analyse:\n"
                      f"Complexité: {results['complexity']}\n"
                      f"Maintenabilité: {results['maintainability']}\n"
                      f"Documentation: {results['documentation']}\n"
                      f"Couverture de tests: {results['test_coverage']}")
            
            messagebox.showinfo("Analyse du code", message)
            self.update_status("Analyse du code terminée")
            self.log_action("Analyse du code", "success")
        except Exception as e:
            logging.error(f"Erreur lors de l'analyse du code: {e}")
            messagebox.showerror("Erreur",
                               "Impossible d'analyser le code")
            self.log_action("Analyse du code", "error")

    def update_stats(self):
        """Mise à jour des statistiques."""
        stats = {
            'modules': len(self.vba_modules),
            'tests': len(self.test_results),
            'coverage': '95%',
            'quality': 'A'
        }
        
        self.stats_labels['modules'].config(text=f"Modules: {stats['modules']}")
        self.stats_labels['tests'].config(text=f"Tests: {stats['tests']}")
        self.stats_labels['coverage'].config(text=f"Couverture: {stats['coverage']}")
        self.stats_labels['quality'].config(text=f"Qualité: {stats['quality']}")

    def log_action(self, action_type: str, status: str):
        """Enregistrement d'une action."""
        date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.actions_list.insert('', 0, values=(date, action_type, status))
        logging.info(f"Action: {action_type} - Statut: {status}")

def start_gui():
    """Démarre l'interface graphique."""
    app = ApexOrchestratorGUI()
    app.window.mainloop()

if __name__ == "__main__":
    start_gui() 