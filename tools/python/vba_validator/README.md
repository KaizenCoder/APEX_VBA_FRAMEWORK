# Validateur de Code VBA pour APEX Framework

## Vue d'ensemble

Le validateur de code VBA est un outil conçu pour assurer la qualité et la conformité du code VBA au sein du projet APEX Framework. Il permet d'identifier automatiquement les problèmes de code, les violations de conventions de nommage et les pratiques non optimales.

Version actuelle : 1.0 (2025-04-13)

## Fonctionnalités

- **Vérification des conventions de nommage** : Préfixes de variables, style de casse, etc.
- **Analyse de la complexité du code** : Longueur des fonctions, nombre de paramètres, niveaux d'imbrication
- **Vérification du style de code** : Indentation, longueur des lignes, utilisation de Option Explicit
- **Détection des problèmes de performance** : Utilisation inefficace des objets Excel, utilisation de Select/Activate
- **Vérification des bonnes pratiques** : Documentation des fonctions, détection du code mort
- **Génération de rapports** : Rapports détaillés au format texte ou JSON

## Prérequis

- Python 3.6 ou supérieur
- PowerShell 5.1 ou supérieur (pour l'utilisation de l'interface PowerShell)

## Structure du dossier

```
tools/python/vba_validator/
├── vba_validator.py        # Script Python principal du validateur
├── default_config.json     # Configuration par défaut
├── Test-VBACode.ps1        # Interface PowerShell
├── README.md               # Ce fichier
└── reports/                # Dossier contenant les rapports générés (créé automatiquement)
```

## Utilisation

### Via PowerShell (recommandé)

Le script PowerShell `Test-VBACode.ps1` fournit une interface conviviale pour utiliser le validateur.

```powershell
# Valider un fichier VBA spécifique
.\Test-VBACode.ps1 -Target "chemin\vers\monmodule.bas"

# Valider tous les fichiers VBA dans un répertoire
.\Test-VBACode.ps1 -Target "chemin\vers\dossier"

# Générer un rapport dans un fichier
.\Test-VBACode.ps1 -Target "chemin\vers\dossier" -OutputFile "rapport_validation.txt"

# Utiliser une configuration personnalisée
.\Test-VBACode.ps1 -Target "chemin\vers\dossier" -Config "ma_config.json"

# Générer un rapport au format JSON
.\Test-VBACode.ps1 -Target "chemin\vers\dossier" -Format "json" -OutputFile "rapport.json"
```

### Via Python directement

Vous pouvez également utiliser directement le script Python.

```bash
# Valider un fichier VBA
python vba_validator.py "chemin/vers/monmodule.bas"

# Valider un répertoire avec options
python vba_validator.py "chemin/vers/dossier" --config "config.json" --output "rapport.txt"

# Options disponibles
python vba_validator.py --help
```

## Configuration

Le validateur peut être configuré via un fichier JSON pour ajuster les règles à vos besoins spécifiques.

### Exemple de configuration

Voir le fichier `default_config.json` pour un exemple complet de configuration. Voici un aperçu des principales sections :

```json
{
  "naming": {
    "module_prefix": {
      "standard": "mod",
      "class": "cls"
    },
    "variable_prefixes": {
      "String": "s",
      "Integer": "i"
    }
  },
  "complexity": {
    "max_function_length": 100,
    "max_line_length": 120
  },
  "rules_enabled": {
    "naming": true,
    "style": true,
    "best_practices": true
  }
}
```

## Intégration continue

Le validateur peut être intégré dans un pipeline CI/CD pour garantir la qualité du code avant les commits ou les déploiements.

### Exemple pour Git pre-commit hook

```bash
# Dans .git/hooks/pre-commit
pwsh -Command "& { cd 'chemin/vers/projet'; ./tools/python/vba_validator/Test-VBACode.ps1 -Target './src' -Verbosity 0; if ($LASTEXITCODE -ne 0) { exit 1 } }"
```

## Exemples de rapports

### Exemple de rapport texte

```
=== Rapport de validation VBA APEX Framework ===
Date: 2025-04-13 22:30

--- Statistiques ---
Fichiers analysés: 5
Lignes analysées: 1250
Problèmes détectés: 8
  - Erreurs: 1
  - Avertissements: 3
  - Informations: 4

--- Problèmes détectés ---
Fichier: D:\Project\modExample.bas
---------------------------------
❌ Ligne 1: Il manque 'Option Explicit' en début de fichier [style.option_explicit]
⚠️ Ligne 15: La fonction 'processData' est trop longue (142 lignes, max recommandé: 100) [complexity.function_length]
ℹ️ Ligne 23: La variable 'counter' ne suit pas la convention camelCase [naming.variable_case]
   Dim Counter As Integer
```

## Contribution

Les contributions pour améliorer le validateur sont les bienvenues ! N'hésitez pas à signaler des bugs ou à proposer des fonctionnalités supplémentaires.

## Licence

Ce projet est sous licence selon les termes de la licence APEX Framework.
