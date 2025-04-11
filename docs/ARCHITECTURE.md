# Architecture Apex Framework v1.1

## Vue d'ensemble

Le framework Apex VBA est conçu selon une architecture à trois couches qui sépare clairement les préoccupations, améliore la maintenabilité et facilite l'évolution du code. Cette architecture s'inspire des principes SOLID et des bonnes pratiques de conception logicielle.

```
┌─────────────────────────────────────────────────────────────────┐
│                       APEX FRAMEWORK v1.1                        │
└─────────────────────────────────────────────────────────────────┘
                               │
     ┌─────────────────────────┼─────────────────────────┐
     │                         │                         │
┌────▼───────────┐     ┌───────▼──────────┐     ┌───────▼──────────┐
│   Apex.Core    │     │   Apex.Métier    │     │    Apex.UI       │
│                │     │                  │     │                  │
│ Modules        │     │ Modules          │     │ Modules          │
│ techniques     │◄────┤ applicatifs     │◄────┤ d'interface      │
│ et transversaux│     │ et métier        │     │ utilisateur      │
└────────────────┘     └──────────────────┘     └──────────────────┘
```

## Principes de conception

1. **Séparation des préoccupations (SoC)**

   - Chaque couche a une responsabilité unique et bien définie
   - Les dépendances circulent dans une seule direction (UI → Métier → Core)
   - Les modules d'une même couche peuvent interagir entre eux

2. **Abstraction et interfaces**

   - Utilisation d'interfaces pour définir des contrats clairs entre composants
   - Implémentations concrètes séparées de leurs interfaces
   - Support du polymorphisme et des patterns d'inversion de dépendances

3. **Encapsulation**

   - L'exposition des fonctionnalités se fait à travers des API clairement définies
   - Les détails d'implémentation sont masqués aux autres composants
   - Chaque module gère son propre état interne

4. **Testabilité**
   - Conception facilitant les tests unitaires et d'intégration
   - Minimisation des dépendances externes directes
   - Possibilité de mocker les dépendances

## Structure des couches

### Apex.Core

La couche Core contient les composants techniques, transversaux et stables du framework. C'est le socle sur lequel reposent les autres couches.

**Caractéristiques**:

- Faiblement couplée avec des dépendances externes minimales
- Fortement testable et stable
- Pas de dépendances vers les couches Métier ou UI
- Composants génériques réutilisables

**Responsabilités**:

- Journalisation (logging)
- Configuration et paramétrage
- Utilitaires techniques (dates, textes, fichiers)
- Framework de tests
- Abstractions et interfaces fondamentales
- Gestion de la sécurité
- Analyse IA et journalisation des interactions

### Apex.Métier

La couche Métier contient les modules applicatifs qui implémentent la logique fonctionnelle du framework.

**Caractéristiques**:

- Peut dépendre de la couche Core, mais pas de la couche UI
- Encapsule la logique métier spécifique au domaine
- Expose des services consommables par l'interface utilisateur

**Responsabilités**:

- Traitement des recettes
- Parsing et manipulation XML
- Intégration avec Outlook
- Accès aux données et ORM
- Logique d'affaires et règles métier
- Explorateur de Conversations IA (extraction et analyse)

### Apex.UI

La couche UI est responsable de l'interface utilisateur et de l'interaction avec l'utilisateur.

**Caractéristiques**:

- Peut dépendre des couches Core et Métier
- Ne contient que du code lié à la présentation et aux interactions
- Séparation de la logique de présentation et du comportement

**Responsabilités**:

- Ruban personnalisé
- Formulaires utilisateur
- Gestion des événements d'interface
- Présentation des données
- Interactions utilisateur
- Interface web de l'Explorateur de Conversations IA

## Flux de dépendances

Les dépendances entre les couches suivent un flux unidirectionnel qui va de la couche la plus externe (UI) vers la couche la plus interne (Core):

```
UI → Métier → Core
```

Cela signifie:

- La couche UI peut dépendre de la couche Métier et de la couche Core
- La couche Métier peut dépendre de la couche Core, mais pas de la couche UI
- La couche Core ne dépend d'aucune autre couche

## Composants Spécialisés

### Explorateur de Conversations IA

L'Explorateur de Conversations IA est un composant transversal qui s'étend sur les trois couches de l'architecture. Il permet d'extraire, d'analyser et de visualiser les interactions entre les développeurs et les assistants IA.

**Structure**:

- **Core** : Extraction des logs et journalisation (Python)
- **Métier** : Analyse et classification des conversations (Python)
- **UI** : Interface web Flask avec visualisations (HTML, CSS, JS)

Pour une description détaillée de ce composant, consultez :
- [Documentation principale](Components/ConversationExplorer.md)
- [Analyse des besoins](Components/ConversationExplorer_Analysis.md)
- [Analyse technique détaillée](../tools/python/ANALYSIS.md)

### Intégration Python via XLWings

L'intégration XLWings est un composant transversal qui établit un pont bidirectionnel entre Python et Excel VBA. Elle permet d'étendre les capacités du framework en offrant une automatisation avancée, une manipulation sophistiquée des données, et des tests d'intégration robustes.

**Structure**:

- **Core** : Utilitaires d'automatisation et journalisation des interactions Python-Excel
- **Métier** : Scripts d'analyse de données et pont d'appel aux fonctionnalités métier
- **UI** : Génération de rapports et tests d'interface utilisateur

Cette intégration s'inscrit parfaitement dans l'architecture en couches tout en respectant les principes de séparation des préoccupations.

Pour une description détaillée de cette intégration, consultez :
- [Documentation complète](Components/XLWings_Integration.md)
- [Script de test](../tools/python/test_xlwings.py)
- [Script d'installation](../tools/python/install_xlwings.ps1)

## Avantages de cette architecture

1. **Maintenabilité améliorée**

   - Les modifications dans une couche ont un impact limité sur les autres
   - Organisation claire facilitant la compréhension du code
   - Séparation des préoccupations permettant des interventions ciblées

2. **Évolutivité**

   - Possibilité d'ajouter de nouvelles fonctionnalités sans perturber l'existant
   - Remplacement de composants par couche sans impact global
   - Extension par couche selon les besoins

3. **Testabilité**

   - Facilité à écrire des tests unitaires par couche
   - Possibilité de mocker les dépendances
   - Tests d'intégration simplifiés

4. **Collaboration améliorée**
   - Répartition claire des responsabilités entre développeurs
   - Travail en parallèle sur différentes couches
   - Documentation par couche facilitant l'intégration de nouveaux développeurs

## Exemples d'interactions

### Exemple 1: Exécution d'une recette

1. **UI**: L'utilisateur clique sur le bouton "Run Recipe" dans le ruban
2. **UI**: Le handler `OnAction_RunRecipe` capte l'événement et journalise l'action
3. **Métier**: Le module `modRecipeComparer` est appelé pour exécuter la comparaison
4. **Core**: Le module de logging est utilisé pour tracer l'exécution
5. **Métier**: Les résultats de la comparaison sont générés
6. **UI**: Les résultats sont présentés à l'utilisateur

### Exemple 2: Configuration du framework

1. **Core**: Le module `modConfigManager` gère les paramètres de configuration
2. **Métier**: Les modules métier lisent leur configuration via le gestionnaire
3. **UI**: Les interfaces permettent de modifier les paramètres
4. **Core**: Les modifications sont persistées par le gestionnaire de configuration

### Exemple 3: Analyse des conversations IA

1. **Core**: Le module d'extraction récupère les logs depuis SQLite
2. **Métier**: L'analyseur de conversations traite et classifie les interactions
3. **Métier**: Les données sont stockées dans la base conversations.db
4. **UI**: L'interface web Flask présente les visualisations et statistiques
5. **UI**: L'utilisateur peut filtrer et explorer les conversations

### Exemple 4: Automatisation via XLWings

1. **UI**: L'utilisateur lance un traitement par lots via l'interface
2. **Métier**: Le module métier appelle un script Python via xlwings
3. **Python**: Le script traite les données et utilise pandas pour l'analyse
4. **Python**: Les résultats sont renvoyés à Excel via xlwings
5. **UI**: Les données traitées sont présentées à l'utilisateur

## Documentation détaillée

Pour plus d'informations sur chaque composant de l'architecture, consultez :

- [Documentation du Core](CORE.md)
- [Documentation du Métier](METIER.md)
- [Documentation de l'UI](UI.md)
- [Explorateur de Conversations IA](Components/ConversationExplorer.md)
- [Intégration XLWings](Components/XLWings_Integration.md)
- [Système de tests](Components/Testing.md)
- [Guide d'onboarding IA](AI_ONBOARDING_GUIDE.md)
- [Modules planifiés](MODULES_PLANIFIES.md) - Liste des modules à développer
- [Maintenance](Maintenance.md) - Gestion des fichiers obsolètes
- [Génération augmentée d'add-in](GENERATE_ADDIN_AUGMENTED.md) - Génération avec gestion automatique des modules manquants

## Conclusion

L'architecture à trois couches du framework Apex VBA offre une structure claire, maintenable et évolutive. Elle favorise les bonnes pratiques de développement, améliore la qualité du code et facilite la collaboration entre développeurs. Cette architecture constitue une base solide pour le développement continu du framework.
