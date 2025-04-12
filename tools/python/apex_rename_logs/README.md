# apex-rename-logs

Outil CLI pour renommer les fichiers journaux obsolètes en ajoutant `.DEPRECATED` à leur nom.

## 📦 Installation

### Installation en mode développement

```bash
cd tools/python/apex_rename_logs/
pip install -e .
```

Cette commande installera l'outil en mode développement, ce qui permet de modifier le code source sans avoir à réinstaller.

## 🚀 Utilisation

### Commande de base

```bash
apex-rename-logs --dir "D:\chemin\vers\mon\projet" --dry-run --verbose
```

### Export des résultats au format CSV

```bash
apex-rename-logs --dir "D:\chemin\vers\mon\projet" --export-csv resultats.csv
```

### Exécution réelle (sans simulation)

```bash
apex-rename-logs --dir "D:\chemin\vers\mon\projet"
```

## ⚙️ Options

| Option | Description |
|--------|-------------|
| `--dir` | Dossier racine à analyser (défaut: répertoire courant) |
| `--dry-run` | Simule sans renommer les fichiers |
| `--log-file` | Nom du fichier log généré |
| `--export-csv` | Génère un rapport CSV des fichiers renommés |
| `--verbose`, `-v` | Active les logs détaillés (DEBUG) |

## 📋 Exemple de rapport CSV

Le rapport CSV généré contient les colonnes suivantes :

- **Fichier** : Chemin complet du fichier
- **Statut** : "Simulation", "Renommé" ou "Erreur: [message]"
- **Motif** : Motif correspondant au fichier

## 📝 Notes importantes

- Ce script doit être exécuté sous PowerShell/Windows, pas sous WSL
- Il recherche les motifs suivants :
  - `create_addin_log*.txt`
  - `fix_classes_log.txt`
  - `apex_addin_generator.log` 