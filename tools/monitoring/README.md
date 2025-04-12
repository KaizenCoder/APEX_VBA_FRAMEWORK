# 📊 VS Code Performance Monitor

## Description
Outil de surveillance des performances de VS Code et ses extensions avec interface graphique.

## 🚀 Fonctionnalités

### 1. Monitoring en Temps Réel
- Utilisation CPU
- Consommation RAM
- Nombre d'extensions actives
- Graphiques dynamiques

### 2. Gestion des Extensions
- Liste des extensions installées
- Analyse d'impact
- Recommandations d'optimisation

### 3. Rapports
- Génération de rapports détaillés
- Export au format Markdown
- Historique des performances

## 📋 Prérequis

### Python
- Version 3.8 ou supérieure
- Packages requis :
  ```bash
  pip install -r requirements.txt
  ```

### Dépendances
```txt
tkinter
psutil
matplotlib
```

## 💻 Installation

1. **Cloner le répertoire**
   ```bash
   git clone <repository>
   cd tools/monitoring
   ```

2. **Installer les dépendances**
   ```bash
   pip install -r requirements.txt
   ```

3. **Lancer l'application**
   ```bash
   python VSCodeMonitor.py
   ```

## 🔧 Utilisation

### Démarrage
```bash
python VSCodeMonitor.py
```

### Interface
1. **Onglet Performance**
   - Graphiques en temps réel
   - Indicateurs CPU/RAM
   - Statistiques globales

2. **Onglet Extensions**
   - Liste des extensions
   - Impact sur les performances
   - Recommandations

3. **Onglet Rapports**
   - Génération de rapports
   - Historique
   - Export des données

## 📈 Fonctionnalités Détaillées

### Monitoring
- Intervalle de rafraîchissement : 1s
- Historique conservé : 1h
- Alertes configurables

### Analyse des Extensions
- Impact individuel
- Conflits potentiels
- Suggestions d'optimisation

### Rapports
- Format Markdown
- Graphiques inclus
- Recommandations automatiques

## ⚙️ Configuration

### Fichier `config.json`
```json
{
    "refresh_rate": 1,
    "history_length": 3600,
    "alert_thresholds": {
        "cpu": 80,
        "ram": 1000
    }
}
```

### Variables d'Environnement
```bash
VSCODE_MONITOR_LOG_LEVEL=INFO
VSCODE_MONITOR_CONFIG_PATH=config/custom.json
```

## 📝 Logs

### Structure
```
logs/
├── vscode_monitor.log
└── performance/
    ├── YYYYMMDD_cpu.log
    └── YYYYMMDD_ram.log
```

### Format
```
2024-04-11 10:15:30 - INFO - Démarrage du monitoring
2024-04-11 10:15:31 - INFO - CPU: 15%, RAM: 450MB
```

## 🔍 Dépannage

### Problèmes Courants
1. **Graphiques non mis à jour**
   - Vérifier le processus VS Code
   - Redémarrer l'application

2. **Extensions non détectées**
   - Vérifier le chemin VS Code
   - Actualiser manuellement

3. **Erreurs de Performance**
   - Vérifier les logs
   - Ajuster les seuils

## 🤝 Contribution

### Guidelines
1. Fork le projet
2. Créer une branche (`git checkout -b feature/AmazingFeature`)
3. Commit les changements (`git commit -m 'Add AmazingFeature'`)
4. Push vers la branche (`git push origin feature/AmazingFeature`)
5. Ouvrir une Pull Request

## 📄 Licence
Distribué sous la licence MIT. Voir `LICENSE` pour plus d'informations.

## ✨ Auteurs
- APEX Framework Team 