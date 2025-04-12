#!/usr/bin/env python3
"""
VSCode Performance Monitor
=========================

Outil de surveillance des performances de VS Code et ses extensions.
Interface graphique pour le monitoring en temps réel.

Auteur: APEX Framework Team
Version: 1.0.0
"""

import tkinter as tk
from tkinter import ttk
import psutil
import json
import sys
import os
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import threading
import queue
import logging

class VSCodeMonitor:
    def __init__(self, root):
        self.root = root
        self.root.title("VS Code Performance Monitor")
        self.root.geometry("800x600")
        
        # Configuration du logging
        logging.basicConfig(
            filename='logs/vscode_monitor.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        
        # Création des onglets
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill='both', padx=5, pady=5)
        
        # Onglet Performance
        self.perf_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.perf_frame, text='Performance')
        
        # Onglet Extensions
        self.ext_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.ext_frame, text='Extensions')
        
        # Onglet Rapports
        self.report_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.report_frame, text='Rapports')
        
        # Initialisation des composants
        self.init_performance_tab()
        self.init_extensions_tab()
        self.init_reports_tab()
        
        # Variables pour le monitoring
        self.monitoring = False
        self.data_queue = queue.Queue()
        self.vscode_process = None
        
        # Démarrage automatique du monitoring
        self.start_monitoring()

    def init_performance_tab(self):
        """Initialise l'onglet de surveillance des performances."""
        # Graphique CPU
        self.fig_cpu = Figure(figsize=(7, 3))
        self.ax_cpu = self.fig_cpu.add_subplot(111)
        self.canvas_cpu = FigureCanvasTkAgg(self.fig_cpu, self.perf_frame)
        self.canvas_cpu.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Graphique RAM
        self.fig_ram = Figure(figsize=(7, 3))
        self.ax_ram = self.fig_ram.add_subplot(111)
        self.canvas_ram = FigureCanvasTkAgg(self.fig_ram, self.perf_frame)
        self.canvas_ram.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Informations en temps réel
        self.info_frame = ttk.LabelFrame(self.perf_frame, text="Informations")
        self.info_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.cpu_label = ttk.Label(self.info_frame, text="CPU: 0%")
        self.cpu_label.pack(side=tk.LEFT, padx=5)
        
        self.ram_label = ttk.Label(self.info_frame, text="RAM: 0 MB")
        self.ram_label.pack(side=tk.LEFT, padx=5)
        
        self.ext_count_label = ttk.Label(self.info_frame, text="Extensions: 0")
        self.ext_count_label.pack(side=tk.LEFT, padx=5)

    def init_extensions_tab(self):
        """Initialise l'onglet de gestion des extensions."""
        # Liste des extensions
        self.ext_tree = ttk.Treeview(self.ext_frame, columns=('Nom', 'Version', 'RAM', 'CPU'))
        self.ext_tree.heading('Nom', text='Nom')
        self.ext_tree.heading('Version', text='Version')
        self.ext_tree.heading('RAM', text='RAM (MB)')
        self.ext_tree.heading('CPU', text='CPU (%)')
        self.ext_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Boutons de gestion
        btn_frame = ttk.Frame(self.ext_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(btn_frame, text="Rafraîchir", command=self.refresh_extensions).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Analyser Impact", command=self.analyze_extensions).pack(side=tk.LEFT, padx=5)

    def init_reports_tab(self):
        """Initialise l'onglet des rapports."""
        # Liste des rapports
        self.report_list = tk.Listbox(self.report_frame)
        self.report_list.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Boutons de gestion des rapports
        btn_frame = ttk.Frame(self.report_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(btn_frame, text="Générer Rapport", command=self.generate_report).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Exporter", command=self.export_report).pack(side=tk.LEFT, padx=5)

    def start_monitoring(self):
        """Démarre le monitoring des performances."""
        self.monitoring = True
        threading.Thread(target=self.monitor_performance, daemon=True).start()

    def monitor_performance(self):
        """Fonction de monitoring en arrière-plan."""
        while self.monitoring:
            try:
                # Recherche du processus VS Code
                for proc in psutil.process_iter(['pid', 'name', 'memory_info']):
                    if 'code' in proc.info['name'].lower():
                        self.vscode_process = proc
                        break
                
                if self.vscode_process:
                    # Collecte des données
                    cpu_percent = self.vscode_process.cpu_percent()
                    ram_mb = self.vscode_process.memory_info().rss / 1024 / 1024
                    
                    # Mise à jour de l'interface
                    self.data_queue.put({
                        'cpu': cpu_percent,
                        'ram': ram_mb,
                        'timestamp': datetime.now()
                    })
                    
                    self.root.after(0, self.update_ui)
                
                # Attente avant la prochaine mesure
                threading.Event().wait(1)
                
            except Exception as e:
                logging.error(f"Erreur de monitoring: {str(e)}")

    def update_ui(self):
        """Met à jour l'interface utilisateur avec les nouvelles données."""
        try:
            while not self.data_queue.empty():
                data = self.data_queue.get_nowait()
                
                # Mise à jour des labels
                self.cpu_label.config(text=f"CPU: {data['cpu']:.1f}%")
                self.ram_label.config(text=f"RAM: {data['ram']:.1f} MB")
                
                # Mise à jour des graphiques
                self.update_graphs(data)
                
        except Exception as e:
            logging.error(f"Erreur de mise à jour UI: {str(e)}")

    def update_graphs(self, data):
        """Met à jour les graphiques de performance."""
        # Mise à jour graphique CPU
        self.ax_cpu.clear()
        self.ax_cpu.set_title('Utilisation CPU')
        self.ax_cpu.set_ylabel('CPU %')
        self.ax_cpu.grid(True)
        self.canvas_cpu.draw()
        
        # Mise à jour graphique RAM
        self.ax_ram.clear()
        self.ax_ram.set_title('Utilisation RAM')
        self.ax_ram.set_ylabel('RAM (MB)')
        self.ax_ram.grid(True)
        self.canvas_ram.draw()

    def refresh_extensions(self):
        """Rafraîchit la liste des extensions."""
        self.ext_tree.delete(*self.ext_tree.get_children())
        
        # Lecture du fichier d'extensions VS Code
        ext_file = os.path.join(os.environ['USERPROFILE'], '.vscode', 'extensions.json')
        try:
            with open(ext_file, 'r') as f:
                extensions = json.load(f)
                for ext in extensions:
                    self.ext_tree.insert('', 'end', values=(
                        ext.get('name', ''),
                        ext.get('version', ''),
                        '-',
                        '-'
                    ))
        except Exception as e:
            logging.error(f"Erreur de lecture des extensions: {str(e)}")

    def analyze_extensions(self):
        """Analyse l'impact des extensions sur les performances."""
        pass  # À implémenter

    def generate_report(self):
        """Génère un rapport de performance."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_file = f"reports/performance_report_{timestamp}.md"
        
        try:
            os.makedirs('reports', exist_ok=True)
            with open(report_file, 'w') as f:
                f.write("# Rapport de Performance VS Code\n\n")
                f.write(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                
                # Ajout des données de performance
                if self.vscode_process:
                    f.write("## Performances Actuelles\n")
                    f.write(f"- CPU: {self.vscode_process.cpu_percent()}%\n")
                    f.write(f"- RAM: {self.vscode_process.memory_info().rss / 1024 / 1024:.1f} MB\n")
                
                # Liste des extensions
                f.write("\n## Extensions Installées\n")
                for item in self.ext_tree.get_children():
                    values = self.ext_tree.item(item)['values']
                    f.write(f"- {values[0]} (v{values[1]})\n")
            
            self.report_list.insert(tk.END, f"Rapport {timestamp}")
            
        except Exception as e:
            logging.error(f"Erreur de génération du rapport: {str(e)}")

    def export_report(self):
        """Exporte le rapport sélectionné."""
        selection = self.report_list.curselection()
        if selection:
            report_name = self.report_list.get(selection[0])
            # Implémenter l'export
            pass

if __name__ == "__main__":
    root = tk.Tk()
    app = VSCodeMonitor(root)
    root.mainloop() 