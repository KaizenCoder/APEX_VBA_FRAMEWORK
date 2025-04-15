# Migration du Core APEX Framework v1.6 → v1.7

Ce répertoire contient les nouvelles interfaces et implémentations pour la migration du Core du framework APEX v1.6 vers v1.7.

## Nouvelles Interfaces

### Interfaces de Base

- **IConfigManagerBase**: Interface pour la gestion de configuration standardisée
- **ILoggerBase**: Interface améliorée pour le logging avec niveaux de log
- **IErrorHandlerBase**: Interface pour la gestion centralisée des erreurs

### Interfaces de Factory

- **ILoggerFactoryBase**: Interface pour la création et gestion des loggers
- **IConfigFactoryBase**: (à implémenter) Interface pour la création des gestionnaires de configuration

## Nouvelles Implémentations

### Factories
- **LoggerFactory**: Implémentation de ILoggerFactoryBase, responsable de la création et gestion des loggers

## Principaux Changements

1. **Injection de Dépendances**: Toutes les implémentations utilisent l'injection de dépendances via les factories
2. **Gestion d'Erreurs Centralisée**: Nouveau système unifié de gestion d'erreurs
3. **Configuration Améliorée**: API de configuration standardisée
4. **Logging Amélioré**: Système de logging avec niveaux de verbosité

## Implémentations Manquantes

- Logger (implémentation de ILoggerBase)
- FileLogger (implémentation de ILoggerBase pour fichiers)
- ExcelLogger (implémentation de ILoggerBase pour Excel)
- ConfigManager (implémentation de IConfigManagerBase)
- ConfigFactory (implémentation de IConfigFactoryBase)
- ErrorHandler (implémentation de IErrorHandlerBase)

## Prochaines Étapes

1. Implémenter les classes manquantes
2. Créer des adaptateurs pour la compatibilité avec v1.6
3. Ajouter des tests unitaires
4. Documenter les nouveaux patterns d'utilisation 