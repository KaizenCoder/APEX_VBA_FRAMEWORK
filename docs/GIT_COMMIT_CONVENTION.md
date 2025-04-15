# Conventions de Commit Git - APEX VBA Framework

## Format standard
Tous les commits doivent suivre le format [Conventional Commits](https://www.conventionalcommits.org) :

```
<type>(<portée>): <description>

<corps>

<pied>
```

## Types de modifications
- **feat**: Nouvelle fonctionnalité
- **fix**: Correction de bug
- **docs**: Modification de documentation
- **refactor**: Modification du code sans changer son comportement
- **test**: Ajout ou modification de tests
- **chore**: Tâches de maintenance, build, etc.
- **perf**: Amélioration de performance
- **style**: Modifications de style (formatage, espaces, etc.)
- **ci**: Modification des fichiers de CI/CD
- **build**: Modification du système de build ou des dépendances externes

## Portée (scope)
La portée indique le module concerné:
- **core**: Composants fondamentaux (logger, config...)
- **metier**: Modules métier (recette, xml...)
  - **xml**: Module de traitement XML
  - **database**: Module d'accès aux données
  - **orm**: Module ORM
  - **security**: Module de sécurité
  - **cache**: Module de cache
  - **recette**: Module de validation métier
  - **monitoring**: Module de surveillance
  - **restapi**: API REST
  - **outlook**: Intégration Outlook
- **ui**: Composants d'interface utilisateur
- **orm**: Couche d'accès aux données
- **docs**: Documentation
- **tests**: Infrastructure de tests
- **ci**: Intégration continue
- **infra**: Infrastructure technique

## Exemples par type

### feat - Nouvelle fonctionnalité
```
feat(orm): ajouter support des transactions imbriquées

Implémentation du support des transactions imbriquées dans le module ORM
permettant la gestion de la profondeur des transactions et le rollback partiel.

Ref: #54
```

### fix - Correction de bug
```
fix(xml): correction du traitement des caractères spéciaux

Résout le problème de parsing des balises XML contenant
des caractères accentués.

Close: #27
```

### docs - Documentation
```
docs(metier): documenter les modules métier

Ajout des fichiers README.md pour tous les modules métier
avec description des fonctions, dépendances et cas d'usage.

Ref: #42
```

### refactor - Refactorisation
```
refactor(core): restructurer l'initialisation du logger

Simplifie la chaîne d'initialisation du logger pour réduire
les dépendances cycliques et améliorer la testabilité.
```

### test - Tests
```
test(modTestRegistry): implémenter tests d'intégration

Ajout des tests d'intégration complets pour le module TestRegistry
avec validation des cas limites et des scénarios d'erreur.

Ref: #63
```

### chore - Maintenance
```
chore(build): mettre à jour les scripts de génération

Optimisation des scripts PowerShell de build pour accélérer
le processus de déploiement et réduire les erreurs.
```

### perf - Performance
```
perf(cache): optimiser la gestion de la mémoire

Réduction de l'empreinte mémoire du module de cache en
implémentant une stratégie d'éviction LRU plus efficace.

Benchmark: -35% consommation mémoire
```

## Exemple de commit complet
```
feat(database): implémenter connecteur SQL Server natif

Ajoute un nouveau connecteur SQL Server utilisant l'API native
au lieu de l'interface ADODB pour améliorer les performances
et la stabilité des connexions longue durée.

- Implémente la classe `clsSqlServerNativeConnector`
- Ajoute support des connexions sécurisées avec encryption
- Résout les fuites mémoire identifiées dans #32
- Ajoute tests unitaires et documentation

Tests:
- Tests de performance montrent +45% d'amélioration
- Test de charge avec 1000 connexions simultanées validé

Ref: #32, #45
Close: #48
```

## Hooks Git
Les hooks Git vérifient automatiquement le respect de ces conventions.
Si votre message est rejeté, utilisez:
```
tools/workflow/scripts/commit_with_context.ps1
```

## Fréquence des commits
- **Un commit par intention atomique** (non groupé)
- **Un log par session** dans `/logs/sessions/`
- **Un log JSON des prompts IA** dans `/logs/prompts/`

## Documentation des prompts IA
Toute modification significative utilisant l'assistance IA doit:
1. Être documentée dans le journal de session
2. Inclure un résumé du prompt utilisé
3. Être validée par un humain avant le commit

## Lien avec les tests et la documentation
Chaque commit doit mentionner:
- Les tests associés (nouveaux ou modifiés)
- La documentation mise à jour si nécessaire
- Le statut de validation des modifications

## Ressources
- [Conventional Commits](https://www.conventionalcommits.org)
- [Angular Commit Convention](https://github.com/angular/angular/blob/master/CONTRIBUTING.md#commit)
- [Semantic Versioning](https://semver.org)

---
*Version: 2.0*  
*Dernière mise à jour: 15/04/2025* 