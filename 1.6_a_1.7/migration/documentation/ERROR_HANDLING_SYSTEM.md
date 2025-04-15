# Système de Gestion d'Erreurs APEX Framework v1.7

## Introduction

Le système de gestion d'erreurs du framework APEX v1.7 permet une gestion centralisée, configurable et traçable des erreurs à travers l'application. Cette documentation présente l'architecture, les fonctionnalités et l'utilisation de ce nouveau système.

## Architecture

Le système de gestion d'erreurs est construit autour de ces composants clés :

```
                ┌─────────────────┐
                │  ErrorFactory   │
                └────────┬────────┘
                         │ crée
                         ▼
┌───────────────┐  ┌────────────────────┐  ┌─────────────────┐
│ ILoggerBase   ├─►│ IErrorHandlerBase  │◄─┤ IConfigManager  │
└───────────────┘  └────────────────────┘  └─────────────────┘
                           │
                   ┌───────▼───────┐
                   │  ErrorHandler │
                   └───────────────┘
```

### Interfaces

- **IErrorHandlerBase** : Interface principale définissant les méthodes et propriétés du gestionnaire d'erreurs.
- **ErrorFactory** : Factory unique pour la création et gestion des gestionnaires d'erreurs.

### Implémentations

- **ErrorHandler** : Implémentation standard du gestionnaire d'erreurs.

## Types d'Erreurs et Niveaux de Gravité

### Types d'Erreurs

Le système définit plusieurs catégories d'erreurs pour une meilleure organisation :

- `ERROR_TYPE_VALIDATION` (5000) : Erreurs de validation des entrées utilisateur
- `ERROR_TYPE_BUSINESS` (5100) : Erreurs liées aux règles métier
- `ERROR_TYPE_DATA` (5200) : Erreurs d'accès ou de manipulation de données
- `ERROR_TYPE_SECURITY` (5300) : Erreurs liées à la sécurité
- `ERROR_TYPE_CONFIGURATION` (5400) : Erreurs de configuration
- `ERROR_TYPE_SYSTEM` (5500) : Erreurs système générales

### Niveaux de Gravité

Quatre niveaux de gravité sont définis pour classifier l'importance des erreurs :

1. `ERROR_SEVERITY_CRITICAL` (1) : Erreur critique nécessitant une attention immédiate
2. `ERROR_SEVERITY_ERROR` (2) : Erreur standard empêchant une opération normale
3. `ERROR_SEVERITY_WARNING` (3) : Avertissement qui n'empêche pas l'opération
4. `ERROR_SEVERITY_INFO` (4) : Information sur un événement qui pourrait être problématique

## Fonctionnalités Clés

1. **Journalisation Intégrée** : Intégration avec le système de logging du framework
2. **Notifications** : Possibilité d'activer des notifications pour certains types d'erreurs
3. **Statistiques** : Suivi des erreurs par type et fréquence
4. **Contextualisation** : Capture de l'emplacement (module, procédure) des erreurs
5. **Centralisation** : Gestion unifiée via ErrorFactory
6. **Configuration** : Paramétrage via le système de configuration

## Utilisation

### Création d'un Gestionnaire d'Erreurs

```vba
' Création via factory (recommandé)
Dim factory As ErrorFactory
Set factory = ErrorFactory
Dim errorHandler As IErrorHandlerBase
Set errorHandler = factory.CreateErrorHandler()

' Avec dépendances explicites
Dim logger As ILoggerBase
Set logger = LoggerFactory.GetLogger("ErrorLogger")
Dim config As IConfigManagerBase
Set config = ConfigFactory.GetConfigManager("AppConfig")
Set errorHandler = factory.CreateErrorHandler(logger, config)

' Récupération depuis le cache
Set errorHandler = factory.GetErrorHandler("AppErrorHandler", logger, config)
```

### Gestion des Erreurs

```vba
' Gestion d'une erreur spécifique
errorHandler.HandleError 5, "Erreur de division par zéro", "ModuleCalcul", "Calculator", "Divide", "Tentative de division par zéro", 2

' Gestion de l'erreur en cours (dans un bloc On Error)
On Error Resume Next
x = 1 / 0  ' Génère une erreur
If Err.Number <> 0 Then
    errorHandler.HandleCurrentError "ModuleCalcul", "Calculator", "Divide"
End If
```

### Déclenchement d'Erreurs

```vba
' Déclenchement d'une erreur personnalisée
If montant <= 0 Then
    errorHandler.RaiseError 5100, "Le montant doit être positif", "ModuleFinance"
End If

' Déclenchement d'une erreur typée
If Not IsDate(dateStr) Then
    errorHandler.RaiseTypedError "Format de date invalide", "ModuleDate", ERROR_TYPE_VALIDATION
End If
```

### Journalisation d'Erreurs

```vba
' Journalisation sans gestion
errorHandler.LogError 1004, "Fichier non trouvé: rapport.xlsx", "ModuleExport", ERROR_SEVERITY_WARNING

' Journalisation de l'erreur en cours
If Err.Number <> 0 Then
    errorHandler.LogCurrentError "ModuleImport", ERROR_SEVERITY_ERROR
End If
```

### Contrôle des Notifications

```vba
' Activer les notifications pour les erreurs de validation
errorHandler.SetNotificationEnabled ERROR_TYPE_VALIDATION, True

' Désactiver les notifications pour les avertissements
errorHandler.SetNotificationEnabled ERROR_SEVERITY_WARNING, False

' Vérifier si les notifications sont activées
If errorHandler.IsNotificationEnabled(ERROR_TYPE_SECURITY) Then
    ' Les notifications sont activées pour les erreurs de sécurité
End If
```

### Utilisation des Méthodes Utilitaires

```vba
' Utilisation directe via la factory (sans créer d'instance)
ErrorFactory.HandleError 5, "Erreur de division par zéro", "ModuleCalcul"

ErrorFactory.RaiseTypedError "Accès refusé", "ModuleSecurity", ERROR_TYPE_SECURITY

' Récupération des statistiques
Dim stats As Object
Set stats = errorHandler.GetErrorStats()
Debug.Print "Nombre d'erreurs: " & errorHandler.ErrorCount
Debug.Print "Erreurs critiques: " & stats("SEVERITY_1")
```

## Configuration

Le système peut être configuré via le système de configuration du framework :

```ini
[error]
rethrowAfterHandling=true

[error.notifications]
1=true    # Activer pour les erreurs critiques
2=true    # Activer pour les erreurs standard
3=false   # Désactiver pour les avertissements
4=false   # Désactiver pour les infos
5000=true # Activer pour les erreurs de validation
```

## Intégration avec le Logging

Le système s'intègre naturellement avec le système de logging :

```vba
' Configuration avec un logger
Dim logger As ILoggerBase
Set logger = LoggerFactory.GetLogger("ErrorLog")
logger.SetLogLevel LogLevel.ERROR ' Capturer les erreurs et au-dessus

' Le gestionnaire d'erreurs utilisera le logger
Set errorHandler = ErrorFactory.CreateErrorHandler(logger)

' Les erreurs seront automatiquement journalisées
errorHandler.HandleError 5, "Erreur de calcul"
' → Génère une entrée de log via le logger
```

## Bonnes Pratiques

1. **Centralisation** : Utilisez toujours ErrorFactory pour créer et obtenir les gestionnaires
2. **Contextualisation** : Fournissez toujours le module et la procédure pour faciliter le débogage
3. **Typage** : Utilisez les types d'erreurs prédéfinis pour une meilleure organisation
4. **Gestion appropriée** : Utilisez HandleError pour les erreurs à gérer, LogError pour simplement les journaliser
5. **Intégration** : Configurez toujours un logger pour garantir la traçabilité
6. **Notifications** : Réservez les notifications aux erreurs critiques ou nécessitant une intervention

## Migration depuis v1.6

Pour migrer du système de gestion d'erreurs v1.6 vers v1.7:

```vba
' Ancien code (v1.6)
Dim oldErrorHandler As clsErrorManager
Set oldErrorHandler = New clsErrorManager
oldErrorHandler.LogError errNumber, errDescription, errSource

' Nouveau code (v1.7)
Dim newErrorHandler As IErrorHandlerBase
Set newErrorHandler = ErrorFactory.GetDefaultErrorHandler()
newErrorHandler.HandleError errNumber, errDescription, errSource
```

## Conclusion

Le nouveau système de gestion d'erreurs offre une approche robuste, flexible et centralisée pour gérer les erreurs dans les applications APEX. La séparation entre interface et implémentation permet d'étendre facilement le système, tandis que l'intégration avec les systèmes de logging et de configuration garantit une expérience cohérente. 