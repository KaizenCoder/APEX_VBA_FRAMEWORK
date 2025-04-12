# APEX VBA FRAMEWORK

<ai_requirements>
ATTENTION IA : Prérequis et règles de base OBLIGATOIRES

1. ENCODAGE : [📝 Règles d'encodage](docs/requirements/powershell_encoding.md)
   - Validation systématique avant toute modification
   - Utilisation des outils de validation fournis
2. DOCUMENTATION : Maintenir la cohérence et la structure
3. VALIDATION : Tester toute modification avant commit
   </ai_requirements>

## ⚠️ PRÉREQUIS CRITIQUES

Avant toute utilisation ou modification :

1. [📝 Encodage PowerShell](docs/requirements/powershell_encoding.md) - **LECTURE OBLIGATOIRE**
   - Règles d'encodage strictes
   - Validation automatique
   - Correction des erreurs

![Version](https://img.shields.io/badge/Version-1.1.0-blue)
![VBA](https://img.shields.io/badge/Language-VBA-brightgreen)
![Excel](https://img.shields.io/badge/Platform-Excel-green)
![Python](https://img.shields.io/badge/Integration-Python-blue)

Framework VBA modulaire pour Excel, offrant une architecture robuste et extensible à trois couches avec intégration Python.

## Architecture

Le framework APEX est désormais organisé en trois couches distinctes :

### 1. Apex.Core

Modules techniques, transversaux et stables formant le socle du framework :

- Logging et journalisation
- Configuration et paramètres
- Utilitaires (date, texte, fichiers)
- Tests unitaires
- Interfaces techniques

### 2. Apex.Métier

Modules applicatifs et fonctionnels implémentant la logique métier :

- Traitement des recettes
- Parsing et manipulation XML
- Intégration Outlook
- Accès aux données
- ORM (Object-Relational Mapping)
- Services REST

### 3. Apex.UI

Interface utilisateur et composants d'interaction :

- Ruban personnalisé
- Formulaires
- Gestionnaires d'événements

## Intégration Python

Le framework propose une intégration avec Python via xlwings pour :

- Automatisation des tests
- Manipulation avancée des données via pandas
- Exécution de macros VBA depuis des scripts externes
- Création de rapports dynamiques

Pour installer et configurer xlwings :

```powershell
# Exécuter depuis PowerShell en tant qu'administrateur
./tools/python/install_xlwings.ps1
```

Consultez [XLWings_Integration.md](docs/Components/XLWings_Integration.md) pour plus de détails.

## Installation

Plusieurs méthodes d'installation sont disponibles :

1. **Installation de l'Add-In** : méthode recommandée utilisant le fichier complément `.xlam`
2. **Création automatisée de l'Add-In** : génération du fichier `.xlam` à partir des sources
3. **Importation manuelle des modules** : intégration directe des modules dans votre projet

Pour des instructions détaillées sur les différentes méthodes d'installation, consultez notre [guide d'installation complet](docs/Installation.md).

```powershell
# Pour créer automatiquement l'Add-In à partir des sources :
.\tools\CreateApexAddIn.ps1
```

## Onboarding

Pour une prise en main rapide du framework :

- [Guide de démarrage rapide](docs/QuickStartGuide.md) - Vue d'ensemble pour les développeurs
- [Guide d'onboarding IA](docs/AI_ONBOARDING_GUIDE.md) - Documentation spécifique pour les assistants IA
- [Guide d'architecture](docs/ARCHITECTURE.md) - Description détaillée de la structure
- [Directives de documentation](docs/DOCUMENTATION_GUIDELINES.md) - Règles pour la création et maintenance de documentation

## Documentation

- [Documentation Core](docs/CORE.md)
- [Documentation Métier](docs/METIER.md)
- [Documentation UI](docs/UI.md)
- [Guide d'onboarding IA](docs/AI_ONBOARDING_GUIDE.md)
- [Configuration Cursor](docs/CURSOR_SETUP.md) - Guide de paramétrage Cursor dans VSCode
- [Format Feedback IA](docs/FEEDBACK_FORMAT.md) - Template standard pour le feedback IA
- [Intégration XLWings](docs/Components/XLWings_Integration.md)
- [Système de tests](docs/Components/Testing.md)
- [Modules planifiés](docs/MODULES_PLANIFIES.md) - Liste des modules à développer
- [Génération augmentée d'add-in](docs/GENERATE_ADDIN_AUGMENTED.md) - Génération avec gestion automatique des modules manquants

## Fonctionnalités principales

- Framework de tests unitaires intégré
- Gestion avancée des logs
- Configuration flexible
- Traitement XML
- Intégration avec Outlook
- Interface utilisateur personnalisable via ruban
- Accès aux données et ORM
- Automatisation via Python (xlwings)

## Prérequis

- Excel 2013+
- VBA 7.0+
- Windows (pour les fonctionnalités DPAPI)
- Python 3.8+ (pour l'intégration xlwings)

## Licence

Ce projet est sous licence. Voir le fichier [LICENSE](LICENSE) pour plus d'informations.

## 🔍 Validation d'Encodage

Le framework utilise un pipeline de validation pour assurer la cohérence de l'encodage et du format des fichiers :

### Pipeline de Validation

Le script `Start-EncodingPipeline.ps1` vérifie :

- Les fichiers de session (format Markdown)
- L'historique des chats
- Les scripts PowerShell
- La documentation

### Utilisation

```powershell
# Validation simple
.\tools\workflow\scripts\Start-EncodingPipeline.ps1

# Validation avec détails
.\tools\workflow\scripts\Start-EncodingPipeline.ps1 -Verbose

# Correction automatique
.\tools\workflow\scripts\Start-EncodingPipeline.ps1 -Fix
```

### Intégration Git

Le pipeline est automatiquement exécuté avant chaque commit via un hook pre-commit.
Pour corriger les erreurs d'encodage :

1. Exécutez le pipeline avec l'option `-Fix`
2. Vérifiez les modifications
3. Ajoutez les fichiers corrigés et recommencez le commit
