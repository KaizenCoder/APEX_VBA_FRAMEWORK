# Guide d'Onboarding - APEX VBA FRAMEWORK

Ce document sert de point d'entrée pour comprendre le framework APEX VBA et naviguer efficacement dans le projet avec Cursor.

## Objectif du Projet

APEX VBA Framework est une solution d'architecture professionnelle pour les applications Excel VBA d'entreprise. Il transforme VBA d'un simple langage de script en environnement de développement structuré avec une architecture en couches, testing, ORM et plugins.

## Navigation Rapide dans le Projet

- **Pour explorer l'architecture**: Commencez par `docs/ARCHITECTURE.md`
- **Pour comprendre un composant**: Voir le dossier `docs/Components/`
- **Pour examiner le code**: Naviguez dans les dossiers `apex-core/`, `apex-metier/` et `apex-ui/`
- **Pour comprendre les tests**: Examinez `tests/` et `apex-core/testing/`

## Architecture Générale

- Framework VBA à trois couches distinctes : Core, Métier, UI
- Architecture modulaire structurée en composants indépendants
- Flux de dépendances unidirectionnel : UI → Métier → Core
- Intégration Python via xlwings pour l'automatisation et les tests

## Intégration d'Outils IA

Afin d'accélérer et d'améliorer le développement au sein d'Apex Framework, des outils d'assistance IA en ligne de commande ont été intégrés à l'environnement de développement (Cursor).

- **`vibe-tools` (basé sur @eastlondoner/cursor-tools) :** Un ensemble d'outils CLI permettant d'étendre les capacités de l'agent IA pour des tâches telles que l'analyse de code, la génération de documentation, l'interaction avec GitHub, l'analyse de vidéos techniques, etc.
  - **État actuel :** Installation fonctionnelle, exécutable via `npx vibe-tools`. La configuration finale des clés API est prévue pour le Lot 2. Voir `tools/python/CURSOR_TOOLS_STATUS.md` pour les détails d'installation.
  - **Utilisation :** Un ensemble de prompts optimisés pour les cas d'usage APEX VBA est disponible dans `APEX_VBA_CURSOR_PROMPTS.md`.

- **Analyse de conversations IA :** Des scripts Python (`explore_cursor_logs.py`, `analyze_ia_logs.py`) ont été développés pour extraire et analyser les logs de conversation de Cursor, permettant d'identifier des patterns d'usage et des points d'amélioration pour le framework (Voir `reports/cursor_analysis/ia_usage_report.md`). Un explorateur web interactif (`conversation_explorer.py`) est également prévu.

## Structure des Dossiers

```
apex-core/            # Composants fondamentaux et transversaux
  ├── interfaces/     # Interfaces et contrats
  ├── testing/        # Framework de tests
  └── utils/          # Utilitaires techniques
apex-metier/          # Modules applicatifs et métier
  ├── database/       # Accès aux données
  ├── orm/            # Object-Relational Mapping
  ├── recette/        # Comparaison de recettes
  ├── xml/            # Traitement XML
  └── outlook/        # Intégration Outlook
apex-ui/              # Interface utilisateur
  ├── ribbon/         # Ruban personnalisé
  ├── forms/          # Formulaires VBA
  └── handlers/       # Gestionnaires d'événements
config/               # Fichiers de configuration
docs/                 # Documentation complète
tests/                # Tests unitaires et d'intégration
tools/                # Scripts de déploiement
  ├── python/         # Scripts Python (xlwings)
  │   ├── install_cursor_tools.ps1  # Installation des outils
  │   └── test_xlwings.py           # Tests d'automatisation Excel
  ├── scripts/        # Scripts de déploiement
  └── workflow/       # Gestion des sessions de développement
      ├── sessions/   # Sessions de développement organisées
      ├── scripts/    # Scripts de gestion des sessions
      └── docs/       # Documentation du workflow
```

## Composants Principaux

### Core

- `clsLogger.cls` - Système de journalisation avancé
- `modConfigManager.bas` - Gestion de configuration
- `clsPluginManager.cls` - Système de plugins extensible
- `modSecurityDPAPI.bas` - Sécurité et cryptographie
- `modTestRunner.bas`, `modTestAssertions.bas` - Framework de test

### Métier

- `clsOrmBase.cls`, `IRelationalObject.cls` - Système ORM
- `clsDbAccessor.cls`, `IDbAccessorBase.cls` - Accès aux données
- `modRecipeComparer.bas` - Comparaison de recettes
- `clsXmlParser.cls` - Traitement XML

### UI

- `customUI.xml` - Configuration du ruban Excel
- `modRibbonCallbacks.bas` - Gestionnaires d'événements ruban

### Intégration Python (xlwings)

- `tools/python/test_xlwings.py` - Scripts d'automatisation Excel
- `tools/python/install_cursor_tools.ps1` - Installation des outils

## Fonctionnalités Clés

- Journalisation multi-destinations (console, fichier, feuille)
- Framework de test unitaire complet
- Système ORM avec relations (One-to-Many, Many-to-Many)
- Système de plugins pour extensions
- Configuration externalisée et paramétrable
- Gestion sécurisée des données sensibles
- Automatisation Excel via Python (xlwings)
- Tests automatisés avec xlwings

## Configuration

Configuration via fichiers INI sectionnés :

- `logger_config.ini` - Configuration du logger
- `recipe_config.ini` - Paramètres de comparaison
- `test_config.ini` - Configuration des tests
- `.cursor-tools.env` - Configuration des outils Python

## Déploiement

- Script `BuildRelease.bat` pour générer une distribution
- Guide de migration `MIGRATION_GUIDE.md` pour projets existants
- Scripts Python pour l'automatisation des tests

## Tests

Plan de test complet préparé pour :

- Tests unitaires par couche
- Tests d'intégration du workflow complet
- Validation de l'interface utilisateur
- Tests automatisés avec xlwings

## Avantages Distinctifs

- Architecture professionnelle rarement vue en VBA
- ORM relationnel complet (unique dans l'écosystème VBA)
- Framework de test intégré mature
- Extensibilité via plugins et interfaces
- Intégration Python pour l'automatisation

## Instructions pour Cursor

Pour explorer ce framework avec Cursor:

1. Commencez par examiner la documentation dans le dossier `docs/`
2. Explorez les interfaces dans `apex-core/interfaces/` pour comprendre les contrats
3. Examinez les tests dans `tests/` pour voir des exemples d'utilisation
4. Suivez les scénarios d'utilisation typiques décrits dans `docs/QuickStartGuide.md`
5. Utilisez les scripts Python dans `tools/python/` pour l'automatisation

## Utilisation de xlwings

Le framework utilise xlwings pour l'automatisation Excel et les tests :

1. **Installation** :
   - Exécuter `tools/python/install_cursor_tools.ps1` en tant qu'administrateur
   - Configurer les clés API dans `.cursor-tools.env`

2. **Tests automatisés** :
   - Utiliser `tools/python/test_xlwings.py` pour les tests d'automatisation
   - Exécuter des tests spécifiques au framework

3. **Automatisation** :
   - Créer des scripts Python pour automatiser les tâches Excel
   - Intégrer les tests dans le workflow de développement

## Questions Typiques à Explorer

1. "Comment fonctionne le système de journalisation?"
2. "Comment implémenter une entité avec relations dans le système ORM?"
3. "Comment étendre le framework avec un plugin personnalisé?"
4. "Comment configurer et exécuter les tests unitaires?"
5. "Comment déployer une application basée sur ce framework?"
6. "Comment utiliser xlwings pour automatiser les tests?"

## Workflow des Sessions de Développement

Le framework intègre un système structuré de gestion des sessions de développement :

### Objectif
- Traçabilité des développements
- Organisation chronologique des sessions
- Validation automatique du format
- Migration et archivage automatisés

### Structure des Sessions
- Organisation par année/mois : `sessions/YYYY/MM/`
- Statuts : `active/` et `completed/`
- Format : `YYYY_MM_DD_description.md`

### Outils de Gestion
- **`Manage-SessionMigration.ps1`** : Migration et organisation des sessions
- **`Test-SessionMarkdownFormat.ps1`** : Validation du format et de l'encodage
- Hooks Git pour validation automatique

### Utilisation
1. Créer une nouvelle session : Format `YYYY_MM_DD_description.md`
2. Remplir les sections obligatoires :
   - 🎯 Objectif(s)
   - 📌 Suivi des tâches
   - ✅ Clôture de session
3. Valider le format : `.\tools\workflow\scripts\Test-SessionMarkdownFormat.ps1`
4. Migrer si nécessaire : `.\tools\workflow\scripts\Manage-SessionMigration.ps1`

Pour plus de détails, consultez `tools/workflow/docs/`.
