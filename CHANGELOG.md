# CHANGELOG - Apex VBA Framework

## Version 1.1.0 (Avril 2025)

### Refonte architecturale majeure

- Migration vers une architecture à trois couches:
  - **Apex.Core**: Modules techniques, transversaux et stables
  - **Apex.Métier**: Modules applicatifs et fonctionnels
  - **Apex.UI**: Interface utilisateur et composants d'interaction

### Génération augmentée d'add-in

- **Support des modules manquants**: Introduction du script `resolve_missing.py`
  - Détection automatique des modules manquants par rapport à une liste prédéfinie
  - Génération de stubs fonctionnels pour les modules non implémentés
  - Classification des modules par priorité (Haute, Moyenne, Basse)
  - Rapports détaillés au format Markdown
- **Amélioration du générateur d'add-in**: 
  - Support des fichiers `.stub` pour génération incrémentale du framework
  - Intégration avec `resolve_missing.py` pour un workflow continu
  - Meilleure gestion des erreurs et logging amélioré
  - [Documentation complète](docs/GENERATE_ADDIN_AUGMENTED.md)

### Structure et organisation

- Restructuration complète des répertoires selon la nouvelle architecture
- Séparation claire des responsabilités entre couches
- Ajout des commentaires de migration dans chaque fichier déplacé
- Optimisation de la structure pour une meilleure maintenabilité

### Améliorations UI

- Introduction d'un ruban personnalisé (Ribbon) via `customUI.xml`
- Création du module `modRibbonCallbacks.bas` pour gérer les interactions
- Préparation pour l'ajout de formulaires utilisateur

### Documentation

- Création de guides spécifiques par couche (`CORE.md`, `METIER.md`, `UI.md`)
- Mise à jour de la documentation principale
- Plan de migration détaillé pour les projets existants

### Build et déploiement

- Mise à jour du script `BuildRelease.bat` pour la nouvelle structure
- Support de génération de documentation automatisée par couche
- Amélioration du processus de validation des releases

## Version 5.0.0 (Avril 2025)

### Ajouts

- Support multi-base de données avec nouveaux drivers:
  - `clsSqlServerDriver.cls` avec support de SCOPE_IDENTITY(), pagination OFFSET/FETCH
  - `clsMySqlDriver.cls` avec support LIMIT et adaptations spécifiques MySQL
  - `clsAccessDriver.cls` avec optimisations pour Microsoft Access
- Client REST API `clsRestApiClient.cls` pour intégration avec services externes
- Système de configuration amélioré avec `clsConfigLoader.cls` v3.0:
  - Support de formats JSON/INI
  - Implémentation du pattern Observer pour notification de changements
- Nouvelle structure de dossiers thématique pour une meilleure organisation

### Améliorations

- Refactoring complet du Query Builder avec support des identifiants sécurisés
- Optimisation des performances de l'ORM et extension de la compatibilité DB
- Amélioration du système de logging avec filtrage par catégorie
- Documentation complète du code et exemples d'utilisation
- Suite de tests étendue avec nouveaux cas

### Corrections

- Résolution des fuites de mémoire dans la gestion des connexions DB
- Correction du traitement des dates dans les requêtes paramétrées
- Gestion améliorée des erreurs dans le client REST
- Correction des problèmes de concurrence dans les transactions

## Version 4.5.0 (Janvier 2025)

### Ajouts

- Interface IDbDriver pour supporter différents moteurs de base de données
- Implémentation de clsAccessDriver comme référence
- Nouvelle classe pour les tests automatisés clsTestResult
- Fonctionnalités ORM étendues avec support de transactions

### Améliorations

- Refactoring du code de test avec introduction de clsTestSuite
- Performance améliorée pour les requêtes volumineuses
- Extensions des assertions pour tests unitaires
- Organisation de code plus modulaire

### Corrections

- Correction des problèmes d'échappement SQL
- Gestion des erreurs plus robuste pour DbAccessor
- Prévention des fuites de ressources dans les connexions

## Version 4.0.0 (Octobre 2024)

### Ajouts

- Générateur ORM automatique via modOrmGenerator
- Support des procédures stockées avec paramètres de sortie
- Nouvelles fonctionnalités du Query Builder: TOP, ORDER BY, GROUP BY
- Introduction des classes de mock pour testing

### Améliorations

- Industrialisation du logging avec rotation de fichiers
- Optimisation des performances générales
- Support complet des paramètres typés

### Corrections

- Résolution des problèmes de mémoire sur longues sessions
- Correction des erreurs de conversion de type dans l'ORM

## Version 3.5.0 (Juillet 2024)

### Ajouts

- Première version du Query Builder avec clauses WHERE avancées
- Introduction de l'ORM basique avec clsOrmBase
- Support de la configuration centralisée

### Améliorations

- Refactoring de DbAccessor pour plus de robustesse
- Meilleure gestion des transactions
- Support amélioré des types de données ADO

### Corrections

- Correction des erreurs avec les types de données NULL
- Gestion des connexions multiples améliorée

## Version 3.0.0 (Mars 2024)

### Ajouts

- Framework de logging multi-cibles (console, fichier, feuille)
- DbAccessor avec support complet des paramètres
- Base du Factory pattern pour les connexions

### Améliorations

- Support des gros volumes de données
- Meilleure gestion des ressources

### Corrections

- Corrections de stabilité diverses

## Version 2.0.0 (Novembre 2023)

### Ajouts

- Première version de DbAccessor avec fonctions de base
- Support des transactions SQL
- Module d'utilitaires de base

### Améliorations

- Ajout de la validation de paramètres
- Support basique des erreurs SQL

## Version 1.0.0 (Août 2023)

### Caractéristiques initiales

- Module de logging basique
- Connexions ADO simples
- Exécution de requêtes SQL
- Structure de base du projet
