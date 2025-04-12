# Guide de l'Agent Documentaire APEX Framework

## Vue d'ensemble

L'Agent Documentaire est un outil automatisé qui vérifie et maintient la conformité de la documentation dans le projet APEX VBA Framework. Il analyse les fichiers de code et de documentation pour s'assurer qu'ils respectent les standards définis dans le fichier de configuration.

## Fonctionnalités

- Vérification automatique de la documentation des fichiers VBA (.cls, .bas, .frm)
- Validation des fichiers Markdown selon les modèles définis
- Correction automatique des problèmes courants
- Génération de rapports d'analyse détaillés
- Exécution manuelle ou programmée

## Prérequis

- Python 3.6 ou supérieur
- PowerShell 5.1 ou supérieur (pour l'exécution planifiée)
- Droits d'administrateur (pour la planification des tâches)

## Utilisation

### Exécution manuelle via VS Code

Plusieurs tâches VS Code sont disponibles pour exécuter l'agent documentaire :

1. **Documentation: Vérifier Conformité** - Analyse l'ensemble du projet et génère un rapport
2. **Documentation: Corriger Automatiquement** - Détecte et corrige les problèmes automatiquement
3. **Documentation: Analyser Fichier Courant** - Analyse uniquement le fichier actuellement ouvert
4. **Documentation: Pre-Commit Check** - Vérifie la conformité avant un commit Git

Pour lancer une de ces tâches :

- Ouvrir la palette de commandes VS Code (Ctrl+Shift+P)
- Taper "Tasks: Run Task"
- Sélectionner la tâche souhaitée

### Exécution via ligne de commande

L'agent peut également être exécuté directement en ligne de commande :

```bash
python doc_agent.py --target=<dossier_cible> [--fix] [--report=<chemin_rapport>] [--config=<chemin_config>]
```

Options :

- `--target` : Dossier ou fichier à analyser (par défaut: dossier courant)
- `--fix` : Tente de corriger automatiquement les problèmes détectés
- `--report` : Chemin où générer le rapport d'analyse (format Markdown)
- `--config` : Chemin vers un fichier de configuration personnalisé

### Planification d'exécutions automatiques

Un script PowerShell est fourni pour configurer une tâche planifiée qui exécutera régulièrement l'agent documentaire :

```powershell
.\schedule_doc_agent.ps1 -Interval Daily -Time 09:00 -GenerateReport
```

Options :

- `-Interval` : Fréquence d'exécution (Daily, Weekly, Monthly)
- `-Time` : Heure d'exécution (format HH:MM)
- `-GenerateReport` : Génère un rapport à chaque exécution
- `-AutoFix` : Corrige automatiquement les problèmes détectés
- `-TaskName` : Nom personnalisé pour la tâche planifiée

**Note** : Ce script nécessite des droits d'administrateur pour créer une tâche planifiée Windows.

## Configuration

Les règles et standards de documentation sont définis dans le fichier `doc_guidelines.json`. Ce fichier comprend :

- Les patterns d'en-têtes pour modules VBA
- Les conventions de nommage
- Les sections requises pour les différents types de documents
- Les extensions de fichiers à analyser
- La structure attendue des dossiers

### Structure du fichier de configuration

```json
{
    "vba_patterns": {
        "module_header": ["@Module", "@Folder", "@Description", ...],
        "method_header": ["@Description", "@Param", "@Returns"],
        "class_prefixes": ["cls"],
        "module_prefixes": ["mod"],
        "form_prefixes": ["frm"]
    },
    "markdown_patterns": {
        "required_sections": {
            "guide": ["Objectif", "Prérequis", "Utilisation", "Exemples"],
            "api": ["Description", "Interface", "Méthodes", "Exemples d'utilisation"],
            "component": ["Vue d'ensemble", "Architecture", "Dépendances", "Configuration", "Utilisation"]
        }
    },
    "file_patterns": {
        "vba": [".cls", ".bas", ".frm"],
        "markdown": [".md"],
        "config": [".json", ".ini"]
    }
}
```

## Rapports

Les rapports générés par l'agent sont au format Markdown et contiennent :

- Un résumé des fichiers analysés
- Le nombre total de problèmes détectés (erreurs et avertissements)
- Une liste détaillée des problèmes par fichier
- Des suggestions de correction

Les rapports sont enregistrés dans le dossier `reports/` du projet.

## Intégration continue

L'agent documentaire peut être intégré dans votre workflow CI/CD en exécutant la tâche "Documentation: Pre-Commit Check" avant chaque commit, garantissant ainsi que tous les nouveaux fichiers ou modifications respectent les standards documentaires du projet.

Pour configurer un hook pre-commit Git :

1. Créer un fichier `.git/hooks/pre-commit` avec le contenu suivant :

```bash
#!/bin/sh
python ./tools/workflow/documentation/doc_agent.py --target .
exit_code=$?

if [ $exit_code -ne 0 ]; then
    echo "❌ Des problèmes de documentation ont été détectés. Consultez le rapport pour plus de détails."
    exit 1
fi

exit 0
```

2. Rendre ce fichier exécutable :

```bash
chmod +x .git/hooks/pre-commit
```

## Extension et personnalisation

L'agent est conçu pour être extensible. Pour ajouter de nouvelles règles :

1. Créez une classe qui hérite de `DocumentationRule` dans `doc_agent.py`
2. Implémentez les méthodes `check()` et `fix()`
3. Ajoutez votre nouvelle règle dans la méthode `_init_rules()` de la classe `DocumentationAgent`

## Support

Pour toute question ou problème concernant l'Agent Documentaire, contactez l'équipe APEX Framework.
