# Système de Configuration APEX Framework v1.7

## Introduction

Le système de configuration du framework APEX v1.7 a été repensé pour offrir une meilleure flexibilité, une gestion centralisée et un accès simplifié aux paramètres de l'application. Cette documentation présente l'architecture et l'utilisation de ce nouveau système.

## Architecture

Le système de configuration est construit autour de ces composants clés :

```
                ┌──────────────────┐
                │  ConfigFactory   │
                └─────────┬────────┘
                          │ crée
                          ▼
┌─────────────┐   ┌───────────────────┐   ┌──────────────┐
│ ILoggerBase ├──►│ IConfigManagerBase├───► ErrorHandler │
└─────────────┘   └─────────┬─────────┘   └──────────────┘
                           ┌┴──────────┐
                           │           │
                   ┌───────▼─┐   ┌─────▼─────┐
                   │ConfigMgr│   │SourceConfig│
                   └─────────┘   └───────────┘
```

### Interfaces

- **IConfigManagerBase** : Interface principale pour le gestionnaire de configuration, définissant les méthodes et propriétés communes.
- **ConfigFactory** : Factory unique pour la création et gestion des gestionnaires de configuration.

### Implémentations

- **ConfigManager** : Implémentation standard du gestionnaire de configuration.
- **TypedConfigManager** : (Prévu) Extension pour la gestion de configurations typées.

## Fonctionnalités Clés

1. **Formats multiples** : Support pour les formats INI, XML, JSON, CSV, Excel
2. **Sections** : Organisation hiérarchique des paramètres
3. **Typages** : Conversion automatique des valeurs (String, Integer, Boolean, Date, Double)
4. **Cache** : Optimisation des performances avec mise en cache
5. **Centralisation** : Gestion unifiée via ConfigFactory
6. **Injection** : Intégration avec le système de logging

## Utilisation

### Création d'un Gestionnaire de Configuration

```vba
' Création via factory (recommandé)
Dim factory As ConfigFactory
Set factory = ConfigFactory
Dim config As IConfigManagerBase
Set config = factory.CreateConfigManager("C:\config\app.ini")

' Avec options personnalisées
Dim options As Object
Set options = CreateObject("Scripting.Dictionary")
Set options("Logger") = LoggerFactory.GetLogger("ConfigLogger")
options("IgnoreErrors") = True
Set config = factory.CreateConfigManager("C:\config\app.ini", options)

' Récupération du cache
Set config = factory.GetConfigManager("AppConfig", "C:\config\app.ini")
```

### Lecture des Valeurs

```vba
' Lecture de valeurs typées
Dim serverName As String
serverName = config.GetString("database.server", "localhost")

Dim port As Long
port = config.GetInteger("database.port", 1433)

Dim isActive As Boolean
isActive = config.GetBoolean("features.reporting", False)

Dim startDate As Date
startDate = config.GetDate("report.startDate", #1/1/2025#)

' Vérification de l'existence
If config.HasKey("database.password") Then
    ' La clé existe
End If

' Accès à une section complète
Dim dbSection As Object
Set dbSection = config.GetSection("database")
' dbSection est un dictionnaire contenant toutes les clés de la section
```

### Modification des Valeurs

```vba
' Définir des valeurs
config.SetValue "app.name", "APEX Application"
config.SetValue "app.version", "1.7.0"
config.SetValue "app.isDebug", True

' Supprimer des valeurs
config.RemoveValue "temp.sessionId"

' Effacer toutes les valeurs
config.Clear
```

### Sauvegarde et Rechargement

```vba
' Sauvegarde dans le même fichier
config.Save

' Sauvegarde dans un nouveau fichier
config.Save "C:\config\app_backup.ini"

' Rechargement depuis la source
config.Reload
```

## Configuration par Section

Le système supporte une organisation hiérarchique avec des sections:

```ini
[database]
server=sqlserver01
port=1433
username=sa
password=****

[ui]
theme=dark
language=fr
showToolbar=true

[logging]
level=info
format=json
path=C:\logs\app.log
```

Accès dans le code:

```vba
' Utilisation de la notation par point
serverName = config.GetString("database.server")

' Ou via section
Dim dbConfig As Object
Set dbConfig = config.GetSection("database")
serverName = dbConfig("server")
```

## Gestion Centralisée

Le ConfigFactory permet de gérer facilement plusieurs configurations:

```vba
' Initialisation avec config par défaut
factory.Initialize "C:\config\default.ini", logger

' Configuration par défaut
Set defaultConfig = factory.GetDefaultConfig()

' Configurations spécifiques
Set appConfig = factory.GetConfigManager("AppConfig", "C:\config\app.ini")
Set userConfig = factory.GetConfigManager("UserConfig", "C:\config\user.ini")
```

## Bonnes Pratiques

1. **Accès unifié** : Utilisez toujours la factory pour créer et obtenir les gestionnaires
2. **Structure logique** : Organisez vos paramètres en sections cohérentes
3. **Valeurs par défaut** : Spécifiez toujours des valeurs par défaut pour la robustesse
4. **Validation** : Vérifiez la présence des clés critiques au démarrage
5. **Centralisation** : Évitez les valeurs codées en dur, utilisez la configuration
6. **Logging** : Utilisez le logger pour tracer les problèmes de configuration

## Migration depuis v1.6

Pour migrer du système de configuration v1.6 vers v1.7:

```vba
' Ancien code (v1.6)
Dim oldConfig As clsConfig
Set oldConfig = New clsConfig
oldConfig.LoadFromFile "config.ini"
serverName = oldConfig.GetParam("SERVER_NAME")

' Nouveau code (v1.7)
Dim newConfig As IConfigManagerBase
Set newConfig = ConfigFactory.GetConfigManager("AppConfig", "config.ini")
serverName = newConfig.GetString("database.server", "localhost")
```

Un adaptateur `LegacyConfigAdapter` est prévu pour faciliter la migration.

## Conclusion

Le nouveau système de configuration offre une flexibilité et une modularité accrues, tout en simplifiant l'accès aux paramètres de l'application. La séparation claire entre interfaces et implémentations permet d'ajouter facilement de nouveaux types de sources de configuration sans modifier le code existant. 