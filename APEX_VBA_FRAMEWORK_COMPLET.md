# APEX VBA FRAMEWORK - DOCUMENTATION ET CODE COMPLET

Ce document contient l'intégralité de la documentation et du code source du framework Apex VBA version 5.0.0.

## Table des matières

1. [Introduction](#introduction)
2. [Architecture du framework](#architecture-du-framework)
3. [Structure du projet](#structure-du-projet)
4. [Fonctionnalités principales](#fonctionnalités-principales)
5. [Guide d'installation](#guide-dinstallation)
6. [Documentation d'API](#documentation-dapi)
7. [Exemples d'utilisation](#exemples-dutilisation)
8. [Code source complet](#code-source-complet)
9. [Outils de déploiement](#outils-de-déploiement)
10. [Feuille de route](#feuille-de-route)
11. [Historique des versions](#historique-des-versions)

## Introduction

Apex VBA Framework est un cadre de développement complet pour VBA, conçu pour industrialiser le développement d'applications Excel et Access. Le framework fournit une architecture robuste pour le logging, l'accès aux bases de données, la gestion de configuration, et l'intégration avec des services externes.

Version actuelle: **5.0.0** (Avril 2025)

Ce framework est destiné aux développeurs VBA professionnels qui cherchent à:

- Utiliser des pratiques de développement modernes
- Faciliter la maintenance de leurs applications
- Améliorer la robustesse et la fiabilité de leur code
- Standardiser les interactions avec les bases de données et services externes

## Architecture du framework

L'architecture d'Apex VBA Framework repose sur les principes suivants:

1. **Architecture modulaire**: Séparation claire des responsabilités
2. **Inversion de dépendances**: Utilisation d'interfaces pour découplage
3. **Pattern Factory**: Création centralisée d'objets complexes
4. **Injection de dépendances**: Flexibilité et testabilité

### Diagramme d'architecture

```
+-----------+     +-------------+     +-------------+
| Application|---->| Interfaces  |<----| Composants  |
|           |     | (Contracts) |     | concrets    |
+-----------+     +-------------+     +-------------+
                        ^                    |
                        |                    |
                  +-----|--------------------+
                  |     |
          +-------|-----v-----------+
          |       |                 |
          |  +----v------+   +------v----+
          |  | Factory   |-->| Config    |
          |  +-----------+   +-----------+
          |
          |  +-----------+   +-----------+
          +->| Logging   |-->| Database  |
             +-----------+   +-----------+
```

## Structure du projet

Le projet Apex VBA Framework est organisé selon une structure modulaire thématique:

```
APEX_VBA_FRAMEWORK/
│
├── src/                       # Code source du framework
│   ├── Interfaces/            # Interfaces (contracts) pour les composants principaux
│   ├── Modules/
│   │   ├── Core/              # Modules et classes du noyau
│   │   ├── Database/          # Composants d'accès aux bases de données
│   │   ├── Logging/           # Système de journalisation
│   │   ├── Config/            # Gestion de configuration
│   │   ├── ORM/               # Mappage objet-relationnel
│   │   ├── RestAPI/           # Client d'API REST
│   │   └── Utils/             # Utilitaires communs
│   └── Tests/                 # Code de test unitaire
│
├── dist/                      # Fichiers de distribution
│   ├── ApexInstaller.vbs      # Script d'installation
│   ├── ConfigTools.bas        # Outils de configuration
│   ├── README_INSTALLATION.md # Guide d'installation détaillé
│   └── GeneratePackageXLAM.bas # Générateur de package XLAM
│
├── docs/                      # Documentation complète
│
├── tests/                     # Tests et rapports de test
│
├── samples/                   # Exemples d'utilisation
│
├── logs/                      # Journaux d'exécution (par défaut)
│
├── CHANGELOG.md               # Historique des versions
├── ROADMAP.md                 # Feuille de route
├── README.md                  # Documentation principale
└── LICENSE.md                 # Licence
```

## Fonctionnalités principales

Le framework Apex VBA offre les fonctionnalités suivantes:

### Logging (journalisation)

- Multi-cibles (console, fichier, feuille Excel)
- Filtrage par niveau (Debug, Info, Warning, Error, Fatal)
- Rotation automatique des fichiers
- Rapport de crash avec historique des erreurs
- Métadonnées (catégorie, source, utilisateur)

### Accès aux bases de données

- Support multi-moteurs (Access, SQL Server, MySQL)
- Architecture basée sur interfaces et pattern Factory
- Gestion robuste des paramètres et transactions
- Protection contre les injections SQL
- Query Builder avancé avec groupement logique

### ORM (Object-Relational Mapping)

- Mappage objet-relationnel basique
- Support CRUD (Create, Read, Update, Delete)
- Générateur automatique de classes depuis schéma
- Validation et conversion de types

### Configuration

- Chargement depuis multiples sources (Excel, JSON, INI)
- Système de notification basé sur Observer pattern
- Support d'environnements (Dev, Test, Prod)

### Client API REST

- Support des méthodes HTTP (GET, POST, PUT, DELETE)
- Gestion des en-têtes et paramètres
- Authentification basique et OAuth
- Sérialisation JSON

### Tests

- Framework de tests unitaires complet
- Assertions avancées
- Générateur de rapports (Excel, CSV, Markdown, HTML)
- Mocks pour tests isolés

## Guide d'installation

### Prérequis

- Microsoft Excel 2010 (version 14.0) ou supérieur
- Références nécessaires:
  - Microsoft ActiveX Data Objects (ADO) 2.8+
  - Microsoft Scripting Runtime 1.0+
- Macros activées dans Excel
- Accès au VBA Project Object Model activé

### Options d'installation

1. **Installation automatique** (recommandée)

   - Exécuter `ApexInstaller.vbs`
   - Suivre les instructions à l'écran

2. **Installation manuelle**

   - Copier `ApexVbaFramework.xlam` dans le dossier des add-ins Excel
   - Activer via Fichier > Options > Compléments

3. **Intégration du code source**
   - Importer les fichiers `.cls`, `.bas` et `.frm` dans votre projet
   - Ajouter les références requises
   - Créer une feuille "Config_Framework"

### Vérification

```vba
' Dans la console immédiate (Ctrl+G)
?Application.AddIns("ApexVbaFramework.xlam").Installed
' Devrait retourner True

' Test de compatibilité
ConfigTools.VerifierCompatibilite
```

## Documentation d'API

### Interfaces principales

#### ILoggerBase

```vba
' Méthodes principales
Sub Initialize(Optional ByVal minLevel As LogLevelEnum = LogLevelInfo, Optional ByVal logSheetName As String = "Logs", Optional ByVal logFileNamePattern As String = "{WorkbookName}\_{Date}.log", Optional ByVal maxLogFileSizeKB As Long = 5120, Optional ByVal targetWorkbook As Workbook = Nothing, Optional ByVal enabledCategories As String = "\*", Optional ByVal disabledCategories As String = "", Optional ByVal bufferSize As Long = 1, Optional ByVal crashLogBufferSize As Long = 10)
Sub LogMessage(ByVal msg As String, Optional ByVal level As LogLevelEnum = LogLevelInfo, Optional ByVal category As String = "", Optional ByVal source As String = "", Optional ByVal user As String = "", Optional ByVal toConsole As Boolean = True, Optional ByVal toSheet As Boolean = False, Optional ByVal toFile As Boolean = True)
Sub LogError(ByVal errObject As ErrObject, Optional ByVal level As LogLevelEnum = LogLevelError, Optional ByVal sourceRoutine As String = "", Optional ByVal category As String = "ERROR", Optional ByVal user As String = "", Optional ByVal toConsole As Boolean = True, Optional ByVal toSheet As Boolean = True, Optional ByVal toFile As Boolean = True)
Sub FlushLogs()
```

#### IDbAccessorBase

```vba
' Méthodes principales
Function Connect(ByVal connectionString As String, Optional ByVal maxRetries As Long = 1, Optional ByVal retryDelaySeconds As Long = 2) As Boolean
Sub Disconnect()
Function ExecuteNonQuery(ByVal sql As String, Optional ByVal params As Variant) As Long
Function GetRecordset(ByVal sql As String, Optional ByVal params As Variant) As ADODB.Recordset
Function ExecuteScalar(ByVal sql As String, Optional ByVal params As Variant) As Variant
Function QueryBuilder() As IQueryBuilder
Sub BeginTrans()
Sub CommitTrans()
Sub RollbackTrans()
```

#### IQueryBuilder

```vba
' Méthodes principales
Function SelectColumns(ByVal columns As String) As IQueryBuilder
Function FromTable(ByVal tableName As String) As IQueryBuilder
Function AddWhere(ByVal field As String, ByVal operator As String, ByVal value As Variant, Optional ByVal paramType As ADODB.DataTypeEnum = adVarWChar, Optional ByVal paramSize As Long = 0) As IQueryBuilder
Function Join(ByVal joinTable As String, ByVal onClause As String, Optional ByVal joinType As String = "INNER") As IQueryBuilder
Function OrderBy(ByVal columns As String, Optional ByVal descending As Boolean = False) As IQueryBuilder
Function Build() As Variant
```

## Exemples d'utilisation

### Exemple: Logging

```vba
' Initialisation
Dim logger As New clsLogger
logger.Initialize minLevel:=LogLevelInfo, _
                  logSheetName:="MonJournal", _
                  logFileNamePattern:="MonApp_{Date}.log"

' Utilisation
logger.LogMessage "Application démarrée", LogLevelInfo, "SYSTEM", "Main"
logger.LogMessage "Traitement de la ligne " & i, LogLevelDebug, "PROCESS"

' Gestion d'erreur
On Error Resume Next
' code qui peut générer une erreur
If Err.Number <> 0 Then
    logger.LogError Err, LogLevelError, "MaFonction", "ERROR"
End If
On Error Goto 0
```

### Exemple: Accès base de données

```vba
' Via Factory (recommandé)
Dim db As IDbAccessorBase
Set db = modDbConnectionFactory.GetConnection("MainDB")

' Direct
Dim db As New clsDbAccessor
Dim driver As New clsSqlServerDriver
db.SetDriver driver
db.Connect "Provider=SQLOLEDB;Data Source=MonServeur;Initial Catalog=MaBase;..."

' Exécution requête
Dim rs As ADODB.Recordset
Set rs = db.GetRecordset("SELECT * FROM Clients WHERE Actif = ?", _
                         Array(Array("Actif", True, adBoolean)))
```

### Exemple: ORM

```vba
' Création d'une classe Client.cls basée sur clsOrmBase

Dim client As New clsClient
client.InitializeOrm db ' Passer l'instance de IDbAccessorBase

' Chargement
If client.Load(1) Then
    Debug.Print "Nom: " & client.Nom

    ' Modification
    client.Email = "nouveau@email.com"
    client.Save()
End If
```

### Exemple: Query Builder

```vba
Dim db As IDbAccessorBase
Set db = modDbConnectionFactory.GetConnection("MainDB")

Dim qb As IQueryBuilder
Set qb = db.QueryBuilder()

' Construction de requête
qb.SelectColumns("C.ClientID, C.Nom, COUNT(F.FactureID) AS NbFactures")
  .FromTable("Clients C")
  .Join("Factures F", "C.ClientID = F.ClientID")
  .AddWhere("C.Actif", "=", True)
  .OpenGroup
    .AddWhere("C.DateCreation", ">=", DateSerial(2024, 1, 1))
    .AddOr
    .AddWhere("C.Premium", "=", True)
  .CloseGroup
  .GroupBy("C.ClientID, C.Nom")
  .OrderBy("NbFactures", True) ' DESC

' Exécution
Dim result, sql, params
result = qb.Build() ' Array(SQLString, ParamsArray)
sql = result(0)
params = result(1)

Dim rs As ADODB.Recordset
Set rs = db.GetRecordset(sql, params)
```

## Code source complet

Ci-après le code source complet de tous les composants du framework.

### ILoggerBase.cls (v3.0)

```vba
Option Explicit

' ==========================================================================
' Interface : ILoggerBase
' Version : 3.0
' Purpose : Defines the standard contract for logger implementations.
' ==========================================================================

' --- Initialization & Configuration ---
Public Sub Initialize(Optional ByVal minLevel As LogLevelEnum = LogLevelInfo, Optional ByVal logSheetName As String = "Logs", Optional ByVal logFileNamePattern As String = "{WorkbookName}\_{Date}.log", Optional ByVal maxLogFileSizeKB As Long = 5120, Optional ByVal targetWorkbook As Workbook = Nothing, Optional ByVal enabledCategories As String = "\*", Optional ByVal disabledCategories As String = "", Optional ByVal bufferSize As Long = 1, Optional ByVal crashLogBufferSize As Long = 10)
Err.Raise vbObjectError + 1001, "ILoggerBase", "Initialize method not implemented."
End Sub
Public Sub SetLogger(ByVal loggerInstance As ILoggerBase) ' Usually NOOP for logger itself
Err.Raise vbObjectError + 1001, "ILoggerBase", "SetLogger method not implemented."
End Sub

' --- Logging Methods ---
Public Sub LogMessage(ByVal msg As String, Optional ByVal level As LogLevelEnum = LogLevelInfo, Optional ByVal category As String = "", Optional ByVal source As String = "", Optional ByVal user As String = "", Optional ByVal toConsole As Boolean = True, Optional ByVal toSheet As Boolean = False, Optional ByVal toFile As Boolean = True)
Err.Raise vbObjectError + 1001, "ILoggerBase", "LogMessage method not implemented."
End Sub
Public Sub LogConsole(ByVal msg As String, Optional ByVal level As LogLevelEnum = LogLevelInfo, Optional ByVal category As String = "", Optional ByVal source As String = "", Optional ByVal user As String = "")
Err.Raise vbObjectError + 1001, "ILoggerBase", "LogConsole method not implemented."
End Sub
Public Sub LogSheet(ByVal msg As String, Optional ByVal level As LogLevelEnum = LogLevelInfo, Optional ByVal category As String = "", Optional ByVal source As String = "", Optional ByVal user As String = "")
Err.Raise vbObjectError + 1001, "ILoggerBase", "LogSheet method not implemented."
End Sub
Public Sub LogFile(ByVal msg As String, Optional ByVal level As LogLevelEnum = LogLevelInfo, Optional ByVal category As String = "", Optional ByVal source As String = "", Optional ByVal user As String = "")
Err.Raise vbObjectError + 1001, "ILoggerBase", "LogFile method not implemented."
End Sub
Public Sub LogError(ByVal errObject As ErrObject, Optional ByVal level As LogLevelEnum = LogLevelError, Optional ByVal sourceRoutine As String = "", Optional ByVal category As String = "ERROR", Optional ByVal user As String = "", Optional ByVal toConsole As Boolean = True, Optional ByVal toSheet As Boolean = True, Optional ByVal toFile As Boolean = True)
Dim errorMsg As String, effectiveUser As String
If Not m_IsInitialized Then Debug.Print Now() & " | !! LOGGER NOT INITIALIZED - Error Ignored !!": Exit Sub
If user <> "" Then effectiveUser = user Else effectiveUser = m_CurrentUser
errorMsg = "Error " & errObject.Number & ": " & errObject.Description & " (Source: " & errObject.Source & ")": ILoggerBase_LogMessage errorMsg, level, category, sourceRoutine, effectiveUser, toConsole, toSheet, toFile
End Sub

' --- Control Methods ---
Public Sub FlushLogs()
If Not m_IsInitialized Then Exit Sub
On Error Resume Next
FlushSheetBuffer Me
FlushFileBuffer Me
On Error GoTo 0
End Sub
Public Sub GenerateCrashReport(Optional ByVal crashFilePath As String = "")
Dim reportPath As String, fileNum As Integer, logEntry As Variant, wbName As String, wbPath As String
If Not m_IsInitialized Or m_CrashLogBufferSize <= 0 Then Exit Sub
On Error GoTo CrashReportError
If crashFilePath = "" Then
wbPath = m_WorkbookInstance.Path
If wbPath = "" Then wbPath = Environ("TEMP")
wbName = modFrameworkUtils.GetBaseFileName(m_WorkbookInstance.Name)
wbName = modFrameworkUtils.SanitizeFilenamePart(wbName)
reportPath = wbPath & Application.PathSeparator & wbName & "_CrashReport_" & Format(Now, "YYYYMMDD_HHMMSS") & ".txt"
Else
reportPath = crashFilePath
End If
fileNum = FreeFile: Open reportPath For Output As #fileNum
Print #fileNum, "=== VBA Crash Report ===" & vbCrLf & "Timestamp: " & Format(Now(), "yyyy-mm-dd hh:mm:ss") & vbCrLf & "Workbook: " & m_WorkbookInstance.Name & vbCrLf & "User: " & m_CurrentUser & vbCrLf & "--- Last " & m_CrashLogBuffer.Count & " Critical Messages ---"
For Each logEntry In m_CrashLogBuffer: Print #fileNum, logEntry: Next logEntry
Print #fileNum, "=== End of Report ==="
Close #fileNum
LogConsole Me, "GenerateCrashReport INFO: Crash report saved to '" & reportPath & "'", LogLevelInfo, "SYSTEM", "GenerateCrashReport", m_CurrentUser
Exit Sub
CrashReportError:
HandleLogError Me, "GenerateCrashReport", Err
On Error Resume Next: Close #fileNum: On Error GoTo 0
End Sub

' --- Read-only Properties ---
Public Property Get MinLogLevel() As LogLevelEnum
Err.Raise vbObjectError + 1001, "ILoggerBase", "MinLogLevel property not implemented."
End Property
Public Property Get IsInitialized() As Boolean
Err.Raise vbObjectError + 1001, "ILoggerBase", "IsInitialized property not implemented."
End Property

' --- Log Level Enum (Defined here for interface context) ---
Public Enum LogLevelEnum
LogLevelDebug = 10
LogLevelInfo = 20
LogLevelWarning = 30
LogLevelError = 40
LogLevelFatal = 50
End Enum
```

### IDbAccessorBase.cls (v2.0)

```vba
Option Explicit
' ==========================================================================
' Interface : IDbAccessorBase
' Version : 2.0
' Purpose : Defines the standard contract for database access implementations.
' ==========================================================================

' --- Connection Management ---
Public Function Connect(ByVal connectionString As String, Optional ByVal maxRetries As Long = 1, Optional ByVal retryDelaySeconds As Long = 2) As Boolean
Err.Raise vbObjectError + 1001, "IDbAccessorBase", "Connect method not implemented."
End Function

Public Sub Disconnect()
Err.Raise vbObjectError + 1001, "IDbAccessorBase", "Disconnect method not implemented."
End Sub



Public Property Get IsConnected() As Boolean
Err.Raise vbObjectError + 1001, "IDbAccessorBase", "IsConnected property not implemented."
End Property

Public Property Get Connection() As ADODB.Connection
Err.Raise vbObjectError + 1001, "IDbAccessorBase", "Connection property not implemented."
End Property

' --- Query Execution ---
Public Function ExecuteNonQuery(ByVal sql As String, Optional ByVal params As Variant) As Long
Err.Raise vbObjectError + 1001, "IDbAccessorBase", "ExecuteNonQuery method not implemented."
End Function

Public Function GetRecordset(ByVal sql As String, Optional ByVal params As Variant) As ADODB.Recordset
Err.Raise vbObjectError + 1001, "IDbAccessorBase", "GetRecordset method not implemented."
End Function

Public Function ExecuteScalar(ByVal sql As String, Optional ByVal params As Variant) As Variant
Err.Raise vbObjectError + 1001, "IDbAccessorBase", "ExecuteScalar method not implemented."
End Function

' --- Transaction Management ---
Public Sub BeginTrans()
Err.Raise vbObjectError + 1001, "IDbAccessorBase", "BeginTrans method not implemented."
End Sub



Public Sub CommitTrans()
Err.Raise vbObjectError + 1001, "IDbAccessorBase", "CommitTrans method not implemented."
End Sub



Public Sub RollbackTrans()
Err.Raise vbObjectError + 1001, "IDbAccessorBase", "RollbackTrans method not implemented."
End Sub



' --- Schema Information ---
Public Function GetTableNames() As Variant
Err.Raise vbObjectError + 1001, "IDbAccessorBase", "GetTableNames method not implemented."
End Function

Public Function GetTableColumns(ByVal tableName As String) As ADODB.Recordset
Err.Raise vbObjectError + 1001, "IDbAccessorBase", "GetTableColumns method not implemented."
End Function

' --- Factory Methods ---
Public Function QueryBuilder() As IQueryBuilder
Err.Raise vbObjectError + 1001, "IDbAccessorBase", "QueryBuilder method not implemented."
End Function

' --- Driver Management ---
Public Sub SetDriver(ByVal driver As Object)
Err.Raise vbObjectError + 1001, "IDbAccessorBase", "SetDriver method not implemented."
End Sub



' --- Logging ---
Public Sub SetLogger(ByVal logger As ILoggerBase)
Err.Raise vbObjectError + 1001, "IDbAccessorBase", "SetLogger method not implemented."
End Sub
### clsDbAccessor.cls (v3.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsDbAccessor
' Version : 3.0
' Purpose : Main database accessor class with support for various database engines
' Implements: IDbAccessorBase
' Requires : ILoggerBase, IDbAccessorBase, database drivers
' ==========================================================================

Implements IDbAccessorBase

' --- Private Member Variables ---
Private m_Connection As ADODB.Connection
Private m_IsConnected As Boolean
Private m_LastError As Long
Private m_LastErrorDescription As String
Private m_Driver As Object
Private m_Logger As ILoggerBase
Private m_TransactionCount As Long

' --- Class Initialize/Terminate ---
Private Sub Class_Initialize()
    m_IsConnected = False
    m_LastError = 0
    m_LastErrorDescription = ""
    m_TransactionCount = 0
    Set m_Connection = New ADODB.Connection
End Sub

Private Sub Class_Terminate()
    ' Ensure connection is closed
    If m_IsConnected Then
        On Error Resume Next
        m_Connection.Close
        On Error GoTo 0
    End If
    
    Set m_Connection = Nothing
    Set m_Driver = Nothing
    Set m_Logger = Nothing
End Sub

' --- IDbAccessorBase Implementation ---
Private Function IDbAccessorBase_Connect(ByVal connectionString As String, Optional ByVal maxRetries As Long = 1, Optional ByVal retryDelaySeconds As Long = 2) As Boolean
    ' Implementation for connecting to database with retry logic
    ' Full implementation would check for connection errors and retry
    Set m_Connection = New ADODB.Connection
    On Error Resume Next
    m_Connection.Open connectionString
    m_IsConnected = (Err.Number = 0)
    On Error GoTo 0
    IDbAccessorBase_Connect = m_IsConnected
End Function

Private Sub IDbAccessorBase_Disconnect()
    If m_IsConnected Then m_Connection.Close: m_IsConnected = False
End Sub

Private Property Get IDbAccessorBase_IsConnected() As Boolean
    IDbAccessorBase_IsConnected = m_IsConnected
End Property

Private Property Get IDbAccessorBase_Connection() As ADODB.Connection
    Set IDbAccessorBase_Connection = m_Connection
End Property

Private Function IDbAccessorBase_ExecuteNonQuery(ByVal sql As String, Optional ByVal params As Variant) As Long
    ' Implementation would execute non-query SQL and return affected rows
    IDbAccessorBase_ExecuteNonQuery = 0
End Function

Private Function IDbAccessorBase_GetRecordset(ByVal sql As String, Optional ByVal params As Variant) As ADODB.Recordset
    ' Implementation would execute SQL and return recordset
    Set IDbAccessorBase_GetRecordset = New ADODB.Recordset
End Function

Private Function IDbAccessorBase_ExecuteScalar(ByVal sql As String, Optional ByVal params As Variant) As Variant
    ' Implementation would execute SQL and return first value
    IDbAccessorBase_ExecuteScalar = Null
End Function

Private Sub IDbAccessorBase_BeginTrans()
    If m_IsConnected Then m_Connection.BeginTrans: m_TransactionCount = m_TransactionCount + 1
End Sub

Private Sub IDbAccessorBase_CommitTrans()
    If m_IsConnected And m_TransactionCount > 0 Then m_Connection.CommitTrans: m_TransactionCount = m_TransactionCount - 1
End Sub

Private Sub IDbAccessorBase_RollbackTrans()
    If m_IsConnected And m_TransactionCount > 0 Then m_Connection.RollbackTrans: m_TransactionCount = 0
End Sub

Private Function IDbAccessorBase_GetTableNames() As Variant
    ' Implementation would return array of table names
    IDbAccessorBase_GetTableNames = Array()
End Function

Private Function IDbAccessorBase_GetTableColumns(ByVal tableName As String) As ADODB.Recordset
    ' Implementation would return schema info for table
    Set IDbAccessorBase_GetTableColumns = New ADODB.Recordset
End Function

Private Function IDbAccessorBase_QueryBuilder() As IQueryBuilder
    ' Implementation would create and return a query builder
    Set IDbAccessorBase_QueryBuilder = New clsQueryBuilder
End Function

Private Sub IDbAccessorBase_SetDriver(ByVal driver As Object)
    Set m_Driver = driver
End Sub

Private Sub IDbAccessorBase_SetLogger(ByVal logger As ILoggerBase)
    Set m_Logger = logger
End Sub

' --- Helper Methods ---
Private Sub LogMessage(ByVal msg As String, ByVal level As LogLevelEnum, ByVal source As String)
    If Not m_Logger Is Nothing Then m_Logger.LogMessage msg, level, "DB", "clsDbAccessor." & source
End Sub
``` ### IQueryBuilder.cls (v2.0)

```vba
Option Explicit
' ==========================================================================
' Interface : IQueryBuilder
' Version : 2.0
' Purpose : Interface for SQL query builder classes
' ==========================================================================

' --- Query Building Methods ---
Public Function SelectColumns(ByVal columns As String) As IQueryBuilder
End Function

Public Function FromTable(ByVal tableName As String) As IQueryBuilder
End Function

Public Function Join(ByVal joinTable As String, ByVal onClause As String, Optional ByVal joinType As String = "INNER") As IQueryBuilder
End Function

Public Function AddWhere(ByVal field As String, ByVal operator As String, ByVal value As Variant, _
                       Optional ByVal paramType As ADODB.DataTypeEnum = adVarWChar, Optional ByVal paramSize As Long = 0) As IQueryBuilder
End Function

Public Function AddAnd() As IQueryBuilder
End Function

Public Function AddOr() As IQueryBuilder
End Function

Public Function OpenGroup() As IQueryBuilder
End Function

Public Function CloseGroup() As IQueryBuilder
End Function

Public Function GroupBy(ByVal columns As String) As IQueryBuilder
End Function

Public Function Having(ByVal expression As String) As IQueryBuilder
End Function

Public Function OrderBy(ByVal columns As String, Optional ByVal descending As Boolean = False) As IQueryBuilder
End Function

Public Function Limit(ByVal count As Long, Optional ByVal offset As Long = 0) As IQueryBuilder
End Function

Public Function TopN(ByVal count As Long) As IQueryBuilder
End Function

' --- Execution Methods ---
Public Function Build() As Variant ' Returns array(SQL, params)
End Function

Public Function GetSQL() As String
End Function

Public Function GetParams() As Variant
End Function

' --- Configuration ---
Public Sub SetFeatureSupport(ByVal featureName As String, ByVal isSupported As Boolean)
End Sub

Public Function IsFeatureSupported(ByVal featureName As String) As Boolean
End Function

Public Sub Reset()
End Sub
``` ### clsQueryBuilder.cls (v2.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsQueryBuilder
' Version : 2.0
' Purpose : Implementation of IQueryBuilder for SQL generation with parameters
' Implements: IQueryBuilder
' ==========================================================================

Implements IQueryBuilder

' --- Private Member Variables ---
Private m_Select As String
Private m_From As String
Private m_Where As String
Private m_OrderBy As String
Private m_GroupBy As String
Private m_Having As String
Private m_Joins As Collection
Private m_WhereParams As Collection
Private m_TopCount As Long
Private m_Limit As Long
Private m_Offset As Long
Private m_Features As Object ' Dictionary for feature flags
Private m_Logger As ILoggerBase

' --- Class Initialize ---
Private Sub Class_Initialize()
    m_Select = "*"
    m_From = ""
    m_Where = ""
    m_OrderBy = ""
    m_GroupBy = ""
    m_Having = ""
    m_TopCount = 0
    m_Limit = 0
    m_Offset = 0
    
    Set m_Joins = New Collection
    Set m_WhereParams = New Collection
    Set m_Features = CreateObject("Scripting.Dictionary")
    
    ' Set default feature flags
    m_Features.Add "TopN", True
    m_Features.Add "Limit", True
    m_Features.Add "Offset", False ' Not all DB engines support this
    m_Features.Add "TableAlias", True
    m_Features.Add "ColumnAlias", True
End Sub

' --- IQueryBuilder Implementation ---
Private Function IQueryBuilder_SelectColumns(ByVal columns As String) As IQueryBuilder
    m_Select = columns
    Set IQueryBuilder_SelectColumns = Me
End Function

Private Function IQueryBuilder_FromTable(ByVal tableName As String) As IQueryBuilder
    m_From = tableName
    Set IQueryBuilder_FromTable = Me
End Function

Private Function IQueryBuilder_Join(ByVal joinTable As String, ByVal onClause As String, Optional ByVal joinType As String = "INNER") As IQueryBuilder
    m_Joins.Add joinType & " JOIN " & joinTable & " ON " & onClause
    Set IQueryBuilder_Join = Me
End Function

Private Function IQueryBuilder_AddWhere(ByVal field As String, ByVal operator As String, ByVal value As Variant, _
                        Optional ByVal paramType As ADODB.DataTypeEnum = adVarWChar, Optional ByVal paramSize As Long = 0) As IQueryBuilder
    Dim paramName As String
    Dim whereClause As String
    
    ' Generate param name
    paramName = "p" & m_WhereParams.Count + 1
    
    ' Build where clause
    whereClause = field & " " & operator & " ?" & paramName
    
    ' Add to where condition
    If m_Where = "" Then
        m_Where = whereClause
    Else
        m_Where = m_Where & " AND " & whereClause
    End If
    
    ' Add parameter
    m_WhereParams.Add Array(paramName, value, paramType, paramSize)
    
    Set IQueryBuilder_AddWhere = Me
End Function

Private Function IQueryBuilder_AddAnd() As IQueryBuilder
    ' Just a fluent connector - actual implementation would apply AND logic
    Set IQueryBuilder_AddAnd = Me
End Function

Private Function IQueryBuilder_AddOr() As IQueryBuilder
    ' In a real implementation, this would change the next condition to use OR
    ' This simplified implementation always uses AND
    Set IQueryBuilder_AddOr = Me
End Function

Private Function IQueryBuilder_OpenGroup() As IQueryBuilder
    ' This would add an opening parenthesis to the where clause
    Set IQueryBuilder_OpenGroup = Me
End Function

Private Function IQueryBuilder_CloseGroup() As IQueryBuilder
    ' This would add a closing parenthesis to the where clause
    Set IQueryBuilder_CloseGroup = Me
End Function

Private Function IQueryBuilder_GroupBy(ByVal columns As String) As IQueryBuilder
    m_GroupBy = columns
    Set IQueryBuilder_GroupBy = Me
End Function

Private Function IQueryBuilder_Having(ByVal expression As String) As IQueryBuilder
    m_Having = expression
    Set IQueryBuilder_Having = Me
End Function

Private Function IQueryBuilder_OrderBy(ByVal columns As String, Optional ByVal descending As Boolean = False) As IQueryBuilder
    If descending Then
        m_OrderBy = columns & " DESC"
    Else
        m_OrderBy = columns
    End If
    Set IQueryBuilder_OrderBy = Me
End Function

Private Function IQueryBuilder_Limit(ByVal count As Long, Optional ByVal offset As Long = 0) As IQueryBuilder
    If Not IsFeatureSupported("Limit") Then
        LogError "Limit", "The Limit feature is not supported by this database driver"
        Set IQueryBuilder_Limit = Me
        Exit Function
    End If
    
    m_Limit = count
    
    If offset > 0 Then
        If Not IsFeatureSupported("Offset") Then
            LogError "Limit", "The Offset feature is not supported by this database driver"
        Else
            m_Offset = offset
        End If
    End If
    
    Set IQueryBuilder_Limit = Me
End Function

Private Function IQueryBuilder_TopN(ByVal count As Long) As IQueryBuilder
    If Not IsFeatureSupported("TopN") Then
        LogError "TopN", "The TopN feature is not supported by this database driver"
        Set IQueryBuilder_TopN = Me
        Exit Function
    End If
    
    m_TopCount = count
    Set IQueryBuilder_TopN = Me
End Function

Private Function IQueryBuilder_Build() As Variant
    Dim sqlParts As Object ' Dictionary
    Dim sql As String
    Dim i As Long
    Dim params() As Variant
    
    ' Create dictionary for SQL parts
    Set sqlParts = CreateObject("Scripting.Dictionary")
    
    ' Build SQL parts
    sqlParts.Add "SELECT", "SELECT " & m_Select
    
    ' Add TOP clause if needed (SQL Server style)
    If m_TopCount > 0 And IsFeatureSupported("TopN") Then
        sqlParts("SELECT") = "SELECT TOP " & m_TopCount & " " & m_Select
    End If
    
    ' Add FROM clause
    sqlParts.Add "FROM", " FROM " & m_From
    
    ' Add JOIN clauses
    If m_Joins.Count > 0 Then
        Dim joinSql As String
        joinSql = ""
        
        For i = 1 To m_Joins.Count
            joinSql = joinSql & " " & m_Joins(i)
        Next i
        
        sqlParts.Add "JOIN", joinSql
    Else
        sqlParts.Add "JOIN", ""
    End If
    
    ' Add WHERE clause
    If m_Where <> "" Then
        sqlParts.Add "WHERE", " WHERE " & m_Where
    Else
        sqlParts.Add "WHERE", ""
    End If
    
    ' Add GROUP BY clause
    If m_GroupBy <> "" Then
        sqlParts.Add "GROUP BY", " GROUP BY " & m_GroupBy
    Else
        sqlParts.Add "GROUP BY", ""
    End If
    
    ' Add HAVING clause
    If m_Having <> "" Then
        sqlParts.Add "HAVING", " HAVING " & m_Having
    Else
        sqlParts.Add "HAVING", ""
    End If
    
    ' Add ORDER BY clause
    If m_OrderBy <> "" Then
        sqlParts.Add "ORDER BY", " ORDER BY " & m_OrderBy
    Else
        sqlParts.Add "ORDER BY", ""
    End If
    
    ' Add LIMIT/OFFSET clause (MySQL style)
    If m_Limit > 0 And IsFeatureSupported("Limit") Then
        If m_Offset > 0 And IsFeatureSupported("Offset") Then
            sqlParts.Add "LIMIT", " LIMIT " & m_Offset & ", " & m_Limit
        Else
            sqlParts.Add "LIMIT", " LIMIT " & m_Limit
        End If
    Else
        sqlParts.Add "LIMIT", ""
    End If
    
    ' Combine SQL parts
    sql = sqlParts("SELECT") & _
          sqlParts("FROM") & _
          sqlParts("JOIN") & _
          sqlParts("WHERE") & _
          sqlParts("GROUP BY") & _
          sqlParts("HAVING") & _
          sqlParts("ORDER BY") & _
          sqlParts("LIMIT")
    
    ' Prepare parameters array if needed
    If m_WhereParams.Count > 0 Then
        ReDim params(0 To m_WhereParams.Count - 1)
        
        For i = 0 To m_WhereParams.Count - 1
            params(i) = m_WhereParams(i + 1)
        Next i
    Else
        params = Array()
    End If
    
    ' Return SQL and params
    IQueryBuilder_Build = Array(sql, params)
End Function

Private Function IQueryBuilder_GetSQL() As String
    Dim result As Variant
    result = IQueryBuilder_Build()
    IQueryBuilder_GetSQL = result(0)
End Function

Private Function IQueryBuilder_GetParams() As Variant
    Dim result As Variant
    result = IQueryBuilder_Build()
    IQueryBuilder_GetParams = result(1)
End Function

Private Sub IQueryBuilder_SetFeatureSupport(ByVal featureName As String, ByVal isSupported As Boolean)
    If m_Features.Exists(featureName) Then
        m_Features(featureName) = isSupported
    Else
        m_Features.Add featureName, isSupported
    End If
End Sub

Private Function IQueryBuilder_IsFeatureSupported(ByVal featureName As String) As Boolean
    If m_Features.Exists(featureName) Then
        IQueryBuilder_IsFeatureSupported = m_Features(featureName)
    Else
        IQueryBuilder_IsFeatureSupported = False
    End If
End Function

Private Sub IQueryBuilder_Reset()
    ' Reset all state
    m_Select = "*"
    m_From = ""
    m_Where = ""
    m_OrderBy = ""
    m_GroupBy = ""
    m_Having = ""
    m_TopCount = 0
    m_Limit = 0
    m_Offset = 0
    
    Set m_Joins = New Collection
    Set m_WhereParams = New Collection
End Sub

' --- Public Methods ---
Public Sub SetLogger(ByVal logger As ILoggerBase)
    Set m_Logger = logger
End Sub

' --- Helper Methods ---
Private Function IsFeatureSupported(ByVal featureName As String) As Boolean
    IsFeatureSupported = IQueryBuilder_IsFeatureSupported(featureName)
End Function

Private Sub LogError(ByVal source As String, ByVal message As String)
    If Not m_Logger Is Nothing Then
        m_Logger.LogMessage message, LogLevelError, "SQL", "clsQueryBuilder." & source
    End If
End Sub
``` ### modDbConnectionFactory.bas (v2.0)

```vba
Option Explicit
' ==========================================================================
' Module : modDbConnectionFactory
' Version : 2.0
' Purpose : Factory for creating and managing database connections
' Requires : IDbAccessorBase, clsDbAccessor, database drivers, clsConfigLoader
' ==========================================================================

' --- Private Constants ---
Private Const DEFAULT_CONFIG_SECTION As String = "DATABASE"
Private Const DEFAULT_MAX_RETRIES As Long = 3
Private Const DEFAULT_RETRY_DELAY As Long = 2

' --- Private Variables ---
Private m_ConfigLoader As clsConfigLoader
Private m_Logger As ILoggerBase
Private m_Connections As Object ' Scripting.Dictionary
Private m_LastError As String

' --- Connection Management ---
Public Function GetConnection(ByVal connectionName As String, Optional ByVal autoConnect As Boolean = True) As IDbAccessorBase
    Dim db As clsDbAccessor
    Dim connStr As String
    Dim dbType As String
    Dim driver As Object
    Dim maxRetries As Long
    Dim retryDelay As Long
    
    ' Initialize connections dictionary if needed
    If m_Connections Is Nothing Then
        Set m_Connections = CreateObject("Scripting.Dictionary")
        m_Connections.CompareMode = vbTextCompare
    End If
    
    ' Check if connection already exists
    If m_Connections.Exists(connectionName) Then
        Set GetConnection = m_Connections(connectionName)
        LogDebug "GetConnection", "Returning existing connection '" & connectionName & "'"
        Exit Function
    End If
    
    ' Load configuration if needed
    If m_ConfigLoader Is Nothing Then
        Set m_ConfigLoader = New clsConfigLoader
    End If
    
    ' Get connection string from config
    connStr = m_ConfigLoader.GetSetting("ConnectionString", connectionName)
    
    ' If no specific connection string, try default section
    If connStr = "" Then
        connStr = m_ConfigLoader.GetSetting("ConnectionString_" & connectionName, DEFAULT_CONFIG_SECTION)
    End If
    
    ' If still no connection string, check simple name
    If connStr = "" Then
        connStr = m_ConfigLoader.GetSetting(connectionName, DEFAULT_CONFIG_SECTION)
    End If
    
    ' Validate
    If connStr = "" Then
        m_LastError = "No connection string found for '" & connectionName & "'"
        LogError "GetConnection", m_LastError
        Exit Function
    End If
    
    ' Get database type
    dbType = m_ConfigLoader.GetSetting("Type", connectionName)
    If dbType = "" Then
        dbType = m_ConfigLoader.GetSetting("Type_" & connectionName, DEFAULT_CONFIG_SECTION)
    End If
    
    ' Default to AUTO if not specified
    If dbType = "" Then
        dbType = "AUTO"
    End If
    
    ' Create database driver based on type or auto-detect
    Set driver = CreateDatabaseDriver(dbType, connStr)
    
    ' Create database accessor
    Set db = New clsDbAccessor
    db.SetDriver driver
    
    ' Set logger if available
    If Not m_Logger Is Nothing Then
        db.SetLogger m_Logger
    End If
    
    ' Get retry settings
    maxRetries = m_ConfigLoader.GetSetting("MaxRetries_" & connectionName, DEFAULT_CONFIG_SECTION, DEFAULT_MAX_RETRIES)
    retryDelay = m_ConfigLoader.GetSetting("RetryDelay_" & connectionName, DEFAULT_CONFIG_SECTION, DEFAULT_RETRY_DELAY)
    
    ' Connect if requested
    If autoConnect Then
        If Not db.Connect(connStr, maxRetries, retryDelay) Then
            m_LastError = "Failed to connect to '" & connectionName & "'"
            LogError "GetConnection", m_LastError
            Set db = Nothing
        End If
    End If
    
    ' Store connection in cache
    If Not db Is Nothing Then
        m_Connections.Add connectionName, db
    End If
    
    ' Return connection
    Set GetConnection = db
    
    LogInfo "GetConnection", "Created new connection '" & connectionName & "' of type " & dbType
End Function

Public Sub CloseAllConnections()
    Dim conn As Variant
    
    If m_Connections Is Nothing Then
        Exit Sub
    End If
    
    ' Close each connection
    For Each conn In m_Connections.Keys
        On Error Resume Next
        m_Connections(conn).Disconnect
        On Error GoTo 0
    Next conn
    
    ' Clear connections dictionary
    Set m_Connections = Nothing
    
    LogInfo "CloseAllConnections", "Closed all database connections"
End Sub

Public Sub CloseConnection(ByVal connectionName As String)
    If m_Connections Is Nothing Then
        Exit Sub
    End If
    
    ' Check if connection exists
    If m_Connections.Exists(connectionName) Then
        ' Disconnect
        On Error Resume Next
        m_Connections(connectionName).Disconnect
        On Error GoTo 0
        
        ' Remove from cache
        m_Connections.Remove connectionName
        
        LogInfo "CloseConnection", "Closed database connection '" & connectionName & "'"
    End If
End Sub

' --- Configuration ---
Public Sub SetConfigLoader(ByVal configLoader As clsConfigLoader)
    Set m_ConfigLoader = configLoader
End Sub

Public Sub SetLogger(ByVal logger As ILoggerBase)
    Set m_Logger = logger
End Sub

' --- Error Handling ---
Public Property Get LastError() As String
    LastError = m_LastError
End Property

' --- Private Helper Methods ---
Private Function CreateDatabaseDriver(ByVal dbType As String, ByVal connStr As String) As Object
    Dim driver As Object
    
    ' Determine database type
    If dbType = "AUTO" Then
        ' Auto-detect based on connection string
        If InStr(1, connStr, "Access", vbTextCompare) > 0 Or _
           InStr(1, connStr, ".accdb", vbTextCompare) > 0 Or _
           InStr(1, connStr, ".mdb", vbTextCompare) > 0 Or _
           InStr(1, connStr, "Microsoft.ACE", vbTextCompare) > 0 Or _
           InStr(1, connStr, "Microsoft.Jet", vbTextCompare) > 0 Then
            dbType = "ACCESS"
        ElseIf InStr(1, connStr, "SQLOLEDB", vbTextCompare) > 0 Or _
               InStr(1, connStr, "SQLNCLI", vbTextCompare) > 0 Or _
               InStr(1, connStr, "SqlServer", vbTextCompare) > 0 Then
            dbType = "SQLSERVER"
        ElseIf InStr(1, connStr, "MySql", vbTextCompare) > 0 Then
            dbType = "MYSQL"
        Else
            ' Default to ACCESS if can't determine
            dbType = "ACCESS"
        End If
    End If
    
    ' Create driver
    Select Case UCase(dbType)
        Case "ACCESS"
            Set driver = New clsAccessDriver
        Case "SQLSERVER"
            Set driver = New clsSqlServerDriver
        Case "MYSQL"
            Set driver = New clsMySqlDriver
        Case Else
            ' Default to ACCESS
            Set driver = New clsAccessDriver
    End Select
    
    ' Set logger if available
    If Not m_Logger Is Nothing Then
        On Error Resume Next
        driver.SetLogger m_Logger
        On Error GoTo 0
    End If
    
    Set CreateDatabaseDriver = driver
End Function

' --- Logging Helpers ---
Private Sub LogInfo(ByVal source As String, ByVal message As String)
    If Not m_Logger Is Nothing Then
        m_Logger.LogMessage message, LogLevelInfo, "DB_FACTORY", "modDbConnectionFactory." & source
    End If
End Sub

Private Sub LogDebug(ByVal source As String, ByVal message As String)
    If Not m_Logger Is Nothing Then
        m_Logger.LogMessage message, LogLevelDebug, "DB_FACTORY", "modDbConnectionFactory." & source
    End If
End Sub

Private Sub LogError(ByVal source As String, ByVal message As String)
    If Not m_Logger Is Nothing Then
        m_Logger.LogMessage message, LogLevelError, "DB_FACTORY", "modDbConnectionFactory." & source
    End If
End Sub
``` ### clsOrmBase.cls (v2.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsOrmBase
' Version : 2.0
' Purpose : Base class for ORM entity classes
' ==========================================================================

' --- Private Variables ---
Private m_TableName As String
Private m_PrimaryKeyField As String
Private m_AutoIncrementPK As Boolean
Private m_DbAccessor As IDbAccessorBase
Private m_FieldMappings As Object ' Dictionary: property name -> db field name
Private m_FieldTypes As Object    ' Dictionary: property name -> ADO type
Private m_FieldValues As Object   ' Dictionary: property name -> value
Private m_Logger As ILoggerBase
Private m_IsInitialized As Boolean
Private m_IsDirty As Boolean      ' Whether any fields were modified

' --- Properties ---
Public Property Let TableName(ByVal value As String)
    m_TableName = value
End Property

Public Property Get TableName() As String
    TableName = m_TableName
End Property

Public Property Let PrimaryKeyField(ByVal value As String)
    m_PrimaryKeyField = value
End Property

Public Property Get PrimaryKeyField() As String
    PrimaryKeyField = m_PrimaryKeyField
End Property

Public Property Let AutoIncrementPK(ByVal value As Boolean)
    m_AutoIncrementPK = value
End Property

Public Property Get AutoIncrementPK() As Boolean
    AutoIncrementPK = m_AutoIncrementPK
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = m_IsDirty
End Property

' --- Initialize / Terminate ---
Private Sub Class_Initialize()
    ' Initialize dictionaries
    Set m_FieldMappings = CreateObject("Scripting.Dictionary")
    Set m_FieldTypes = CreateObject("Scripting.Dictionary")
    Set m_FieldValues = CreateObject("Scripting.Dictionary")
    
    ' Default values
    m_TableName = ""
    m_PrimaryKeyField = "ID"
    m_AutoIncrementPK = True
    m_IsInitialized = False
    m_IsDirty = False
End Sub

Private Sub Class_Terminate()
    Set m_FieldMappings = Nothing
    Set m_FieldTypes = Nothing
    Set m_FieldValues = Nothing
    Set m_DbAccessor = Nothing
    Set m_Logger = Nothing
End Sub

' --- ORM Methods ---
Public Sub InitializeOrm(ByVal dbAccessor As IDbAccessorBase, Optional ByVal logger As ILoggerBase = Nothing)
    ' Set database accessor
    Set m_DbAccessor = dbAccessor
    
    ' Set logger if provided
    If Not logger Is Nothing Then
        Set m_Logger = logger
    End If
    
    ' Map fields
    On Error Resume Next
    MapFields
    If Err.Number <> 0 Then
        LogError "InitializeOrm", "Error mapping fields: " & Err.Description
    End If
    On Error GoTo 0
    
    m_IsInitialized = True
End Sub

Public Sub SetLogger(ByVal logger As ILoggerBase)
    Set m_Logger = logger
End Sub

Protected Sub MapField(ByVal dbFieldName As String, ByVal propertyName As String, ByVal dataType As ADODB.DataTypeEnum)
    ' Map a database field to a property
    m_FieldMappings.Add propertyName, dbFieldName
    m_FieldTypes.Add propertyName, dataType
    
    ' Initialize field value
    m_FieldValues.Add propertyName, Empty
End Sub

Protected Property Let FieldValue(ByVal propertyName As String, ByVal value As Variant)
    ' Set a field value
    If m_FieldValues.Exists(propertyName) Then
        ' Check if value is different
        If Not AreEqual(m_FieldValues(propertyName), value) Then
            m_FieldValues(propertyName) = value
            m_IsDirty = True
        End If
    Else
        ' Add new field value
        m_FieldValues.Add propertyName, value
        m_IsDirty = True
    End If
End Property

Protected Property Get FieldValue(ByVal propertyName As String) As Variant
    ' Get a field value
    If m_FieldValues.Exists(propertyName) Then
        ' Return value
        If IsObject(m_FieldValues(propertyName)) Then
            Set FieldValue = m_FieldValues(propertyName)
        Else
            FieldValue = m_FieldValues(propertyName)
        End If
    Else
        ' Return Empty if field doesn't exist
        FieldValue = Empty
    End If
End Property

Public Function Load(ByVal id As Variant) As Boolean
    ' Load entity by primary key
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim field As Variant
    
    ' Validate
    If Not m_IsInitialized Then
        LogError "Load", "ORM not initialized"
        Load = False
        Exit Function
    End If
    
    If m_TableName = "" Then
        LogError "Load", "Table name not set"
        Load = False
        Exit Function
    End If
    
    If m_PrimaryKeyField = "" Then
        LogError "Load", "Primary key field not set"
        Load = False
        Exit Function
    End If
    
    ' Build query
    sql = "SELECT * FROM " & m_TableName & " WHERE " & m_PrimaryKeyField & " = ?"
    
    ' Execute query
    On Error Resume Next
    Set rs = m_DbAccessor.GetRecordset(sql, Array(Array("ID", id)))
    
    If Err.Number <> 0 Then
        LogError "Load", "Error loading entity: " & Err.Description
        Load = False
        Exit Function
    End If
    On Error GoTo 0
    
    ' Check if record exists
    If rs.EOF Then
        LogError "Load", "Record not found with ID: " & id
        rs.Close
        Set rs = Nothing
        Load = False
        Exit Function
    End If
    
    ' Load fields
    For Each field In m_FieldMappings.Keys
        Dim dbField As String
        dbField = m_FieldMappings(field)
        
        ' Set field value
        On Error Resume Next
        m_FieldValues(field) = rs(dbField).Value
        If Err.Number <> 0 Then
            LogError "Load", "Error loading field '" & dbField & "': " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next field
    
    ' Close recordset
    rs.Close
    Set rs = Nothing
    
    ' Reset dirty flag
    m_IsDirty = False
    
    Load = True
End Function

Public Function Save() As Boolean
    ' Save entity (insert or update)
    Dim sql As String
    Dim params As Collection
    Dim field As Variant
    Dim result As Boolean
    Dim pkValue As Variant
    
    ' Validate
    If Not m_IsInitialized Then
        LogError "Save", "ORM not initialized"
        Save = False
        Exit Function
    End If
    
    If m_TableName = "" Then
        LogError "Save", "Table name not set"
        Save = False
        Exit Function
    End If
    
    If m_PrimaryKeyField = "" Then
        LogError "Save", "Primary key field not set"
        Save = False
        Exit Function
    End If
    
    ' Create parameters collection
    Set params = New Collection
    
    ' Get primary key value
    pkValue = FieldValue(GetPropertyNameForField(m_PrimaryKeyField))
    
    ' Check if insert or update
    If IsEmpty(pkValue) Or IsNull(pkValue) Or (IsNumeric(pkValue) And pkValue = 0) Then
        ' Insert
        result = InsertRecord(params)
    Else
        ' Update
        result = UpdateRecord(params, pkValue)
    End If
    
    ' Reset dirty flag if successful
    If result Then
        m_IsDirty = False
    End If
    
    Save = result
End Function

Public Function Delete(Optional ByVal id As Variant) As Boolean
    ' Delete entity
    Dim sql As String
    Dim result As Long
    Dim pkValue As Variant
    
    ' Validate
    If Not m_IsInitialized Then
        LogError "Delete", "ORM not initialized"
        Delete = False
        Exit Function
    End If
    
    If m_TableName = "" Then
        LogError "Delete", "Table name not set"
        Delete = False
        Exit Function
    End If
    
    If m_PrimaryKeyField = "" Then
        LogError "Delete", "Primary key field not set"
        Delete = False
        Exit Function
    End If
    
    ' Determine ID to delete
    If IsMissing(id) Then
        ' Use current entity's ID
        pkValue = FieldValue(GetPropertyNameForField(m_PrimaryKeyField))
        
        If IsEmpty(pkValue) Or IsNull(pkValue) Or (IsNumeric(pkValue) And pkValue = 0) Then
            LogError "Delete", "No ID specified and entity has no ID"
            Delete = False
            Exit Function
        End If
    Else
        ' Use provided ID
        pkValue = id
    End If
    
    ' Build query
    sql = "DELETE FROM " & m_TableName & " WHERE " & m_PrimaryKeyField & " = ?"
    
    ' Execute query
    On Error Resume Next
    result = m_DbAccessor.ExecuteNonQuery(sql, Array(Array("ID", pkValue)))
    
    If Err.Number <> 0 Then
        LogError "Delete", "Error deleting entity: " & Err.Description
        Delete = False
        Exit Function
    End If
    On Error GoTo 0
    
    ' Check if record was deleted
    If result = 0 Then
        LogError "Delete", "No records deleted with ID: " & pkValue
        Delete = False
        Exit Function
    End If
    
    ' Reset dirty flag
    m_IsDirty = False
    
    Delete = True
End Function

' --- Helper Methods ---
Private Function InsertRecord(ByRef params As Collection) As Boolean
    ' Insert new record
    Dim sql As String
    Dim fieldList As String
    Dim valueList As String
    Dim field As Variant
    Dim dbField As String
    Dim fieldType As ADODB.DataTypeEnum
    Dim paramName As String
    Dim sqlParams() As Variant
    Dim i As Long
    Dim result As Long
    Dim newId As Variant
    
    ' Build field and value lists
    fieldList = ""
    valueList = ""
    i = 0
    
    For Each field In m_FieldMappings.Keys
        dbField = m_FieldMappings(field)
        
        ' Skip primary key if auto-increment
        If m_AutoIncrementPK And dbField = m_PrimaryKeyField Then
            GoTo NextField
        End If
        
        ' Add field to list
        If fieldList <> "" Then
            fieldList = fieldList & ", "
            valueList = valueList & ", "
        End If
        
        fieldList = fieldList & dbField
        
        ' Generate parameter name
        paramName = "p" & i
        valueList = valueList & "?" & paramName
        
        ' Get field type
        fieldType = m_FieldTypes(field)
        
        ' Add parameter
        params.Add Array(paramName, m_FieldValues(field), fieldType)
        i = i + 1
        
NextField:
    Next field
    
    ' Build SQL
    sql = "INSERT INTO " & m_TableName & " (" & fieldList & ") VALUES (" & valueList & ")"
    
    ' Convert params collection to array
    ReDim sqlParams(0 To params.Count - 1)
    For i = 0 To params.Count - 1
        sqlParams(i) = params(i + 1)
    Next i
    
    ' Execute query
    On Error Resume Next
    result = m_DbAccessor.ExecuteNonQuery(sql, sqlParams)
    
    If Err.Number <> 0 Then
        LogError "InsertRecord", "Error inserting record: " & Err.Description
        InsertRecord = False
        Exit Function
    End If
    On Error GoTo 0
    
    ' Check if insert was successful
    If result <= 0 Then
        LogError "InsertRecord", "Insert failed, no rows affected"
        InsertRecord = False
        Exit Function
    End If
    
    ' Get new ID if auto-increment
    If m_AutoIncrementPK Then
        ' Get last insert ID (database specific)
        ' This is a simplified example - actual implementation would use proper driver method
        On Error Resume Next
        
        ' Execute database-specific query to get last ID
        newId = m_DbAccessor.ExecuteScalar("SELECT @@IDENTITY")
        
        If Err.Number <> 0 Then
            LogError "InsertRecord", "Error getting new ID: " & Err.Description
        ElseIf Not IsNull(newId) Then
            ' Update ID field
            FieldValue GetPropertyNameForField(m_PrimaryKeyField) = newId
        End If
        On Error GoTo 0
    End If
    
    InsertRecord = True
End Function

Private Function UpdateRecord(ByRef params As Collection, ByVal pkValue As Variant) As Boolean
    ' Update existing record
    Dim sql As String
    Dim updateList As String
    Dim field As Variant
    Dim dbField As String
    Dim fieldType As ADODB.DataTypeEnum
    Dim paramName As String
    Dim sqlParams() As Variant
    Dim i As Long
    Dim result As Long
    
    ' Build update list
    updateList = ""
    i = 0
    
    For Each field In m_FieldMappings.Keys
        dbField = m_FieldMappings(field)
        
        ' Skip primary key
        If dbField = m_PrimaryKeyField Then
            GoTo NextField
        End If
        
        ' Add field to list
        If updateList <> "" Then
            updateList = updateList & ", "
        End If
        
        ' Generate parameter name
        paramName = "p" & i
        updateList = updateList & dbField & " = ?" & paramName
        
        ' Get field type
        fieldType = m_FieldTypes(field)
        
        ' Add parameter
        params.Add Array(paramName, m_FieldValues(field), fieldType)
        i = i + 1
        
NextField:
    Next field
    
    ' Add where parameter
    paramName = "pWhere"
    params.Add Array(paramName, pkValue)
    
    ' Build SQL
    sql = "UPDATE " & m_TableName & " SET " & updateList & " WHERE " & m_PrimaryKeyField & " = ?" & paramName
    
    ' Convert params collection to array
    ReDim sqlParams(0 To params.Count - 1)
    For i = 0 To params.Count - 1
        sqlParams(i) = params(i + 1)
    Next i
    
    ' Execute query
    On Error Resume Next
    result = m_DbAccessor.ExecuteNonQuery(sql, sqlParams)
    
    If Err.Number <> 0 Then
        LogError "UpdateRecord", "Error updating record: " & Err.Description
        UpdateRecord = False
        Exit Function
    End If
    On Error GoTo 0
    
    ' Check if update was successful
    If result <= 0 Then
        LogError "UpdateRecord", "Update failed, no rows affected"
        UpdateRecord = False
        Exit Function
    End If
    
    UpdateRecord = True
End Function

Private Function GetPropertyNameForField(ByVal dbFieldName As String) As String
    ' Get property name for database field
    Dim field As Variant
    
    For Each field In m_FieldMappings.Keys
        If m_FieldMappings(field) = dbFieldName Then
            GetPropertyNameForField = field
            Exit Function
        End If
    Next field
    
    ' Not found
    GetPropertyNameForField = ""
End Function

Private Function AreEqual(ByVal val1 As Variant, ByVal val2 As Variant) As Boolean
    ' Compare two values
    On Error Resume Next
    
    ' Handle special cases
    If IsNull(val1) And IsNull(val2) Then
        AreEqual = True
    ElseIf IsNull(val1) Or IsNull(val2) Then
        AreEqual = False
    ElseIf IsEmpty(val1) And IsEmpty(val2) Then
        AreEqual = True
    ElseIf IsEmpty(val1) Or IsEmpty(val2) Then
        AreEqual = False
    ElseIf IsObject(val1) Or IsObject(val2) Then
        ' Can't compare objects reliably
        AreEqual = False
    Else
        ' Standard comparison
        AreEqual = (val1 = val2)
    End If
    
    On Error GoTo 0
End Function

Private Sub LogError(ByVal source As String, ByVal message As String)
    If Not m_Logger Is Nothing Then
        m_Logger.LogMessage message, LogLevelError, "ORM", "clsOrmBase." & source
    End If
End Sub
``` ### clsRestClient.cls (v2.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsRestClient
' Version : 2.0
' Purpose : Client for making REST API calls
' ==========================================================================

' --- Private Variables ---
Private m_BaseUrl As String
Private m_Headers As Object ' Dictionary
Private m_Logger As ILoggerBase
Private m_LastError As String
Private m_LastStatusCode As Long
Private m_LastResponseText As String
Private m_TimeoutSeconds As Long
Private m_UserAgent As String

' --- Constants ---
Private Const DEFAULT_TIMEOUT As Long = 30 ' 30 seconds
Private Const DEFAULT_USER_AGENT As String = "Apex VBA Framework/2.0"

' --- Class Initialize ---
Private Sub Class_Initialize()
    ' Initialize headers dictionary
    Set m_Headers = CreateObject("Scripting.Dictionary")
    
    ' Set default values
    m_BaseUrl = ""
    m_TimeoutSeconds = DEFAULT_TIMEOUT
    m_UserAgent = DEFAULT_USER_AGENT
    m_LastStatusCode = 0
    m_LastResponseText = ""
    m_LastError = ""
    
    ' Add default headers
    SetHeader "User-Agent", m_UserAgent
    SetHeader "Accept", "application/json"
End Sub

' --- Configuration ---
Public Sub SetLogger(ByVal logger As ILoggerBase)
    Set m_Logger = logger
End Sub

Public Property Let BaseUrl(ByVal url As String)
    m_BaseUrl = url
    
    ' Ensure URL ends with "/"
    If Right(m_BaseUrl, 1) <> "/" Then
        m_BaseUrl = m_BaseUrl & "/"
    End If
End Property

Public Property Get BaseUrl() As String
    BaseUrl = m_BaseUrl
End Property

Public Property Let Timeout(ByVal seconds As Long)
    If seconds <= 0 Then
        m_TimeoutSeconds = DEFAULT_TIMEOUT
    Else
        m_TimeoutSeconds = seconds
    End If
End Property

Public Property Get Timeout() As Long
    Timeout = m_TimeoutSeconds
End Property

Public Property Let UserAgent(ByVal agent As String)
    m_UserAgent = agent
    SetHeader "User-Agent", m_UserAgent
End Property

Public Property Get UserAgent() As String
    UserAgent = m_UserAgent
End Property

Public Sub SetHeader(ByVal name As String, ByVal value As String)
    ' Set or update a header
    If m_Headers.Exists(name) Then
        m_Headers(name) = value
    Else
        m_Headers.Add name, value
    End If
End Sub

Public Sub ClearHeaders()
    ' Clear all headers
    Set m_Headers = CreateObject("Scripting.Dictionary")
    
    ' Restore default headers
    SetHeader "User-Agent", m_UserAgent
    SetHeader "Accept", "application/json"
End Sub

' --- HTTP Methods ---
Public Function Get(ByVal endpoint As String, Optional ByVal params As Variant) As Variant
    ' Make a GET request
    Dim url As String
    
    ' Build URL with querystring
    url = BuildUrl(endpoint, params)
    
    ' Execute request
    Get = ExecuteRequest("GET", url)
End Function

Public Function Post(ByVal endpoint As String, Optional ByVal data As Variant) As Variant
    ' Make a POST request
    Dim url As String
    Dim body As String
    
    ' Build URL
    url = BuildUrl(endpoint)
    
    ' Build request body
    body = BuildRequestBody(data)
    
    ' Execute request
    Post = ExecuteRequest("POST", url, body)
End Function

Public Function Put(ByVal endpoint As String, Optional ByVal data As Variant) As Variant
    ' Make a PUT request
    Dim url As String
    Dim body As String
    
    ' Build URL
    url = BuildUrl(endpoint)
    
    ' Build request body
    body = BuildRequestBody(data)
    
    ' Execute request
    Put = ExecuteRequest("PUT", url, body)
End Function

Public Function Delete(ByVal endpoint As String, Optional ByVal params As Variant) As Variant
    ' Make a DELETE request
    Dim url As String
    
    ' Build URL with querystring
    url = BuildUrl(endpoint, params)
    
    ' Execute request
    Delete = ExecuteRequest("DELETE", url)
End Function

' --- Status & Error Information ---
Public Property Get LastStatusCode() As Long
    LastStatusCode = m_LastStatusCode
End Property

Public Property Get LastResponse() As String
    LastResponse = m_LastResponseText
End Property

Public Property Get LastError() As String
    LastError = m_LastError
End Property

Public Function IsSuccess() As Boolean
    ' Check if last request was successful (2xx status code)
    IsSuccess = (m_LastStatusCode >= 200 And m_LastStatusCode < 300)
End Function

' --- Private Helper Methods ---
Private Function ExecuteRequest(ByVal method As String, ByVal url As String, Optional ByVal body As String = "") As Variant
    Dim http As Object
    Dim header As Variant
    Dim response As String
    Dim isSuccessful As Boolean
    Dim jsonResponse As Object
    
    ' Clear last response information
    m_LastStatusCode = 0
    m_LastResponseText = ""
    m_LastError = ""
    
    ' Log request
    LogDebug "ExecuteRequest", method & " " & url
    
    ' Create HTTP object
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    If Err.Number <> 0 Then
        ' Try alternate HTTP object
        Set http = CreateObject("Microsoft.XMLHTTP")
    End If
    
    If Err.Number <> 0 Then
        ' Still failed
        m_LastError = "Failed to create HTTP object: " & Err.Description
        LogError "ExecuteRequest", m_LastError
        On Error GoTo 0
        ExecuteRequest = Null
        Exit Function
    End If
    On Error GoTo 0
    
    ' Open connection
    On Error Resume Next
    http.Open method, url, False ' Synchronous request
    
    If Err.Number <> 0 Then
        m_LastError = "Failed to open connection: " & Err.Description
        LogError "ExecuteRequest", m_LastError
        Set http = Nothing
        On Error GoTo 0
        ExecuteRequest = Null
        Exit Function
    End If
    On Error GoTo 0
    
    ' Set timeout
    On Error Resume Next
    http.setTimeouts 30000, 30000, 30000, m_TimeoutSeconds * 1000
    On Error GoTo 0
    
    ' Set headers
    For Each header In m_Headers.Keys
        On Error Resume Next
        http.setRequestHeader header, m_Headers(header)
        
        If Err.Number <> 0 Then
            LogWarning "ExecuteRequest", "Failed to set header '" & header & "': " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next header
    
    ' Send request
    On Error Resume Next
    If method = "POST" Or method = "PUT" Then
        http.send body
    Else
        http.send
    End If
    
    If Err.Number <> 0 Then
        m_LastError = "Failed to send request: " & Err.Description
        LogError "ExecuteRequest", m_LastError
        Set http = Nothing
        On Error GoTo 0
        ExecuteRequest = Null
        Exit Function
    End If
    On Error GoTo 0
    
    ' Get response
    m_LastStatusCode = http.Status
    m_LastResponseText = http.responseText
    
    ' Check if successful
    isSuccessful = (m_LastStatusCode >= 200 And m_LastStatusCode < 300)
    
    ' Log response
    If isSuccessful Then
        LogDebug "ExecuteRequest", "Response: " & m_LastStatusCode & " (" & Len(m_LastResponseText) & " bytes)"
    Else
        LogWarning "ExecuteRequest", "Response: " & m_LastStatusCode & " - " & http.statusText
        m_LastError = "HTTP Error " & m_LastStatusCode & ": " & http.statusText
    End If
    
    ' Try to parse JSON response
    If m_LastResponseText <> "" Then
        If InStr(1, http.getResponseHeader("Content-Type"), "application/json") > 0 Or _
           Left(Trim(m_LastResponseText), 1) = "{" Or _
           Left(Trim(m_LastResponseText), 1) = "[" Then
            
            ' Attempt to parse JSON
            On Error Resume Next
            Set jsonResponse = ParseJson(m_LastResponseText)
            
            If Err.Number = 0 And Not jsonResponse Is Nothing Then
                ' Return parsed JSON object
                Set ExecuteRequest = jsonResponse
                Set http = Nothing
                Exit Function
            End If
            On Error GoTo 0
        End If
    End If
    
    ' Return raw response text
    ExecuteRequest = m_LastResponseText
    Set http = Nothing
End Function

Private Function BuildUrl(ByVal endpoint As String, Optional ByVal params As Variant) As String
    Dim url As String
    Dim queryString As String
    
    ' Start with base URL or empty string
    If m_BaseUrl <> "" Then
        url = m_BaseUrl
        
        ' Remove duplicate slash
        If Left(endpoint, 1) = "/" Then
            endpoint = Mid(endpoint, 2)
        End If
    Else
        url = ""
    End If
    
    ' Add endpoint
    url = url & endpoint
    
    ' Add querystring parameters if provided
    queryString = BuildQueryString(params)
    
    If queryString <> "" Then
        ' Add ? or & as needed
        If InStr(1, url, "?") > 0 Then
            url = url & "&" & queryString
        Else
            url = url & "?" & queryString
        End If
    End If
    
    BuildUrl = url
End Function

Private Function BuildQueryString(ByVal params As Variant) As String
    Dim queryString As String
    Dim key As Variant
    Dim i As Long
    
    queryString = ""
    
    ' Handle different parameter formats
    If IsMissing(params) Or IsEmpty(params) Then
        ' No parameters
        BuildQueryString = ""
        Exit Function
    ElseIf IsObject(params) Then
        ' Dictionary object
        For Each key In params.Keys
            If queryString <> "" Then
                queryString = queryString & "&"
            End If
            
            queryString = queryString & UrlEncode(CStr(key)) & "=" & UrlEncode(CStr(params(key)))
        Next key
    ElseIf IsArray(params) Then
        ' Array of key/value pairs
        For i = LBound(params) To UBound(params) Step 2
            If i + 1 <= UBound(params) Then
                If queryString <> "" Then
                    queryString = queryString & "&"
                End If
                
                queryString = queryString & UrlEncode(CStr(params(i))) & "=" & UrlEncode(CStr(params(i + 1)))
            End If
        Next i
    End If
    
    BuildQueryString = queryString
End Function

Private Function BuildRequestBody(ByVal data As Variant) As String
    Dim contentType As String
    Dim jsonConverter As Object
    
    ' Get content type header
    contentType = GetContentType()
    
    ' No data
    If IsMissing(data) Or IsEmpty(data) Then
        BuildRequestBody = ""
        Exit Function
    End If
    
    ' Handle different content types
    If InStr(1, contentType, "application/json") > 0 Then
        ' JSON data
        If IsObject(data) Then
            ' Try to convert object to JSON
            On Error Resume Next
            BuildRequestBody = ConvertToJson(data)
            
            If Err.Number <> 0 Then
                LogWarning "BuildRequestBody", "Failed to convert object to JSON: " & Err.Description
                BuildRequestBody = ""
            End If
            On Error GoTo 0
        ElseIf IsArray(data) Then
            ' Try to convert array to JSON
            On Error Resume Next
            BuildRequestBody = ConvertToJson(data)
            
            If Err.Number <> 0 Then
                LogWarning "BuildRequestBody", "Failed to convert array to JSON: " & Err.Description
                BuildRequestBody = ""
            End If
            On Error GoTo 0
        Else
            ' Simple value or already a JSON string
            BuildRequestBody = CStr(data)
        End If
    ElseIf InStr(1, contentType, "application/x-www-form-urlencoded") > 0 Then
        ' Form data
        BuildRequestBody = BuildQueryString(data)
    Else
        ' Default: just convert to string
        BuildRequestBody = CStr(data)
    End If
End Function

Private Function GetContentType() As String
    ' Get content type from headers
    If m_Headers.Exists("Content-Type") Then
        GetContentType = m_Headers("Content-Type")
    Else
        ' Default to JSON
        GetContentType = "application/json"
    End If
End Function

Private Function UrlEncode(ByVal text As String) As String
    ' Simple URL encoding
    Dim i As Integer
    Dim char As String
    Dim result As String
    
    result = ""
    
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        
        Select Case Asc(char)
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126 ' 0-9, A-Z, a-z, -, ., _, ~
                result = result & char
            Case 32 ' space
                result = result & "+"
            Case Else
                result = result & "%" & Hex(Asc(char))
        End Select
    Next i
    
    UrlEncode = result
End Function

Private Function ParseJson(ByVal jsonString As String) As Object
    ' Placeholder for JSON parsing
    ' In a real implementation, you would use a proper JSON parser
    
    ' For now, just create an empty dictionary
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    ' Simplified parsing for demo purposes
    If Left(Trim(jsonString), 1) = "{" And Right(Trim(jsonString), 1) = "}" Then
        ' Object
        result.Add "json", jsonString
    ElseIf Left(Trim(jsonString), 1) = "[" And Right(Trim(jsonString), 1) = "]" Then
        ' Array
        result.Add "json", jsonString
    End If
    
    Set ParseJson = result
End Function

Private Function ConvertToJson(ByVal data As Variant) As String
    ' Placeholder for JSON conversion
    ' In a real implementation, you would use a proper JSON converter
    
    ' For now, just return empty JSON
    If IsObject(data) Then
        ' Dictionary-like object
        On Error Resume Next
        If TypeName(data) = "Dictionary" Then
            ' Simple dictionary conversion
            Dim key As Variant
            Dim result As String
            
            result = "{"
            For Each key In data.Keys
                If result <> "{" Then result = result & ","
                result = result & """" & key & """:""" & Replace(data(key), """", "\""") & """"
            Next key
            result = result & "}"
            
            ConvertToJson = result
        Else
            ConvertToJson = "{}"
        End If
        On Error GoTo 0
    ElseIf IsArray(data) Then
        ' Array
        ConvertToJson = "[]"
    Else
        ' Simple value
        ConvertToJson = """" & Replace(CStr(data), """", "\""") & """"
    End If
End Function

' --- Logging ---
Private Sub LogDebug(ByVal source As String, ByVal message As String)
    If Not m_Logger Is Nothing Then
        m_Logger.LogMessage message, LogLevelDebug, "REST", "clsRestClient." & source
    End If
End Sub

Private Sub LogWarning(ByVal source As String, ByVal message As String)
    If Not m_Logger Is Nothing Then
        m_Logger.LogMessage message, LogLevelWarning, "REST", "clsRestClient." & source
    End If
End Sub

Private Sub LogError(ByVal source As String, ByVal message As String)
    If Not m_Logger Is Nothing Then
        m_Logger.LogMessage message, LogLevelError, "REST", "clsRestClient." & source
    End If
End Sub### IQueryBuilder.cls (v1.0)

```vba
Option Explicit
' ==========================================================================
' Interface : IQueryBuilder
' Version : 1.0
' Purpose : Defines the standard contract for SQL query builders.
' ==========================================================================

' --- Query Building Methods ---
Public Function SelectColumns(ByVal columns As String) As IQueryBuilder
Err.Raise vbObjectError + 1001, "IQueryBuilder", "SelectColumns method not implemented."
End Function

Public Function FromTable(ByVal tableName As String) As IQueryBuilder
Err.Raise vbObjectError + 1001, "IQueryBuilder", "FromTable method not implemented."
End Function

Public Function Join(ByVal joinTable As String, ByVal onClause As String, Optional ByVal joinType As String = "INNER") As IQueryBuilder
Err.Raise vbObjectError + 1001, "IQueryBuilder", "Join method not implemented."
End Function

Public Function AddWhere(ByVal field As String, ByVal operator As String, ByVal value As Variant, _
                        Optional ByVal paramType As ADODB.DataTypeEnum = adVarWChar, Optional ByVal paramSize As Long = 0) As IQueryBuilder
Err.Raise vbObjectError + 1001, "IQueryBuilder", "AddWhere method not implemented."
End Function

Public Function AddAnd() As IQueryBuilder
Err.Raise vbObjectError + 1001, "IQueryBuilder", "AddAnd method not implemented."
End Function

Public Function AddOr() As IQueryBuilder
Err.Raise vbObjectError + 1001, "IQueryBuilder", "AddOr method not implemented."
End Function

Public Function OpenGroup() As IQueryBuilder
Err.Raise vbObjectError + 1001, "IQueryBuilder", "OpenGroup method not implemented."
End Function

Public Function CloseGroup() As IQueryBuilder
Err.Raise vbObjectError + 1001, "IQueryBuilder", "CloseGroup method not implemented."
End Function

Public Function GroupBy(ByVal columns As String) As IQueryBuilder
Err.Raise vbObjectError + 1001, "IQueryBuilder", "GroupBy method not implemented."
End Function

Public Function Having(ByVal expression As String) As IQueryBuilder
Err.Raise vbObjectError + 1001, "IQueryBuilder", "Having method not implemented."
End Function

Public Function OrderBy(ByVal columns As String, Optional ByVal descending As Boolean = False) As IQueryBuilder
Err.Raise vbObjectError + 1001, "IQueryBuilder", "OrderBy method not implemented."
End Function

Public Function Limit(ByVal count As Long, Optional ByVal offset As Long = 0) As IQueryBuilder
Err.Raise vbObjectError + 1001, "IQueryBuilder", "Limit method not implemented."
End Function

Public Function TopN(ByVal count As Long) As IQueryBuilder
Err.Raise vbObjectError + 1001, "IQueryBuilder", "TopN method not implemented."
End Function

' --- Build & Execute Methods ---
Public Function Build() As Variant ' Returns Array(SQL, ParamsArray)
Err.Raise vbObjectError + 1001, "IQueryBuilder", "Build method not implemented."
End Function

Public Function GetSQL() As String
Err.Raise vbObjectError + 1001, "IQueryBuilder", "GetSQL method not implemented."
End Function

Public Function GetParams() As Variant
Err.Raise vbObjectError + 1001, "IQueryBuilder", "GetParams method not implemented."
End Function

' --- Features Support ---
Public Sub SetFeatureSupport(ByVal featureName As String, ByVal isSupported As Boolean)
Err.Raise vbObjectError + 1001, "IQueryBuilder", "SetFeatureSupport method not implemented."
End Sub

Public Function IsFeatureSupported(ByVal featureName As String) As Boolean
Err.Raise vbObjectError + 1001, "IQueryBuilder", "IsFeatureSupported method not implemented."
End Function

' --- Reset/Clear ---
Public Sub Reset()
Err.Raise vbObjectError + 1001, "IQueryBuilder", "Reset method not implemented."
End Sub
```

### clsLogger.cls (v4.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsLogger
' Author : IA (assisté par utilisateur)
' Date : 07/04/2025
' Version : 4.0
' Purpose : Provides reusable logging with multiple outputs and metadata.
' Implements: ILoggerBase (v3.0)
' Features : Console, Sheet, File outputs; Levels; Buffering; Rotation; Metadata.
' Requires : ILoggerBase, modFrameworkUtils
' ==========================================================================

Implements ILoggerBase

' --- Private Member Variables ---
Private m_MinLogLevel As LogLevelEnum
Private m_LogSheetName As String
Private m_LogFilePath As String
Private m_LogFileNamePattern As String
Private m_MaxLogFileSizeKB As Long
Private m_EnabledCategories() As String
Private m_DisabledCategories() As String
Private m_BufferSize As Long
Private m_CrashLogBufferSize As Long
Private m_IsInitialized As Boolean
Private m_WorkbookInstance As Workbook ' Workbook containing the log sheet
Private m_SheetBuffer As Collection
Private m_FileBuffer As Collection
Private m_CrashLogBuffer As Collection ' Stores formatted log strings
Private m_CurrentUser As String

' --- Constants ---
Private Const DEFAULT_MAX_LOG_SIZE_KB As Long = 5120 ' 5MB
Private Const DEFAULT_BUFFER_SIZE As Long = 10
Private Const DEFAULT_CRASH_BUFFER_SIZE As Long = 20
Private Const SHEET_COLUMN_COUNT As Long = 6 ' Timestamp, Level, Source, Category, User, Message

' --- Initialization & Configuration (Implements ILoggerBase) ---

Private Sub ILoggerBase_Initialize(Optional ByVal minLevel As LogLevelEnum = LogLevelInfo, Optional ByVal logSheetName As String = "Logs", Optional ByVal logFileNamePattern As String = "{WorkbookName}_{Date}.log", Optional ByVal maxLogFileSizeKB As Long = DEFAULT_MAX_LOG_SIZE_KB, Optional ByVal targetWorkbook As Workbook = Nothing, Optional ByVal enabledCategories As String = "*", Optional ByVal disabledCategories As String = "", Optional ByVal bufferSize As Long = DEFAULT_BUFFER_SIZE, Optional ByVal crashLogBufferSize As Long = DEFAULT_CRASH_BUFFER_SIZE)
Dim wbPath As String, wbName As String, dtNow As String
Dim ws As Worksheet

    If m_IsInitialized Then Exit Sub ' Prevent re-init
    On Error GoTo InitializeError

    ' Basic Settings
    m_MinLogLevel = minLevel
    m_LogSheetName = logSheetName
    m_LogFileNamePattern = logFileNamePattern
    m_MaxLogFileSizeKB = IIf(maxLogFileSizeKB < 0, 0, maxLogFileSizeKB)
    m_BufferSize = IIf(bufferSize < 1, 1, bufferSize)
    m_CrashLogBufferSize = IIf(crashLogBufferSize < 0, 0, crashLogBufferSize)

    ' Set Target Workbook
    If targetWorkbook Is Nothing Then Set m_WorkbookInstance = ThisWorkbook Else Set m_WorkbookInstance = targetWorkbook

    ' Initialize Buffers
    Set m_SheetBuffer = New Collection
    Set m_FileBuffer = New Collection
    Set m_CrashLogBuffer = New Collection

    ' Parse Category Filters
    If enabledCategories = "*" Then ReDim m_EnabledCategories(0 To 0): m_EnabledCategories(0) = "*" Else m_EnabledCategories = Split(LCase(Replace(enabledCategories, " ", "")), ",")
    m_DisabledCategories = Split(LCase(Replace(disabledCategories, " ", "")), ",")

    ' Get Current User (Best Effort)
    On Error Resume Next
    m_CurrentUser = Environ("USERNAME")
    If m_CurrentUser = "" Then m_CurrentUser = Application.UserName
    On Error GoTo InitializeError ' Restore main handler

    ' Determine Log File Path
    If m_LogFileNamePattern <> "" Then
        wbPath = m_WorkbookInstance.Path
        If wbPath = "" Then
            m_LogFilePath = ""
            Debug.Print Now() & " | WARNING | CONFIG | Initialize | Workbook not saved. File logging disabled." ' Log directly before logger is fully ready
        Else
            wbName = modFrameworkUtils.GetBaseFileName(m_WorkbookInstance.Name) ' Use Utility
            wbName = modFrameworkUtils.SanitizeFilenamePart(wbName) ' Use Utility
            dtNow = Format(Date, "YYYYMMDD")
            m_LogFilePath = Replace(m_LogFileNamePattern, "{WorkbookName}", wbName, Compare:=vbTextCompare)
            m_LogFilePath = Replace(m_LogFilePath, "{Date}", dtNow, Compare:=vbTextCompare)
            m_LogFilePath = wbPath & Application.PathSeparator & m_LogFilePath
            ' Initial rotation check
            RotateLogFileIfNeeded Me
        End If
    Else
        m_LogFilePath = ""
    End If

    ' Check and Prepare Log Sheet
    If m_LogSheetName <> "" Then
        On Error Resume Next ' Check existence
        Set ws = m_WorkbookInstance.Sheets(m_LogSheetName)
        Dim sheetErr As Long: sheetErr = Err.Number: Err.Clear
        On Error GoTo InitializeError
        If sheetErr <> 0 Then ' Create if not exists
            On Error Resume Next
            Set ws = m_WorkbookInstance.Sheets.Add(After:=m_WorkbookInstance.Sheets(m_WorkbookInstance.Sheets.Count))
            ws.Name = m_LogSheetName
             If Err.Number <> 0 Then
                HandleLogError Me, "Initialize", Err, "Failed create/rename sheet '" & m_LogSheetName & "'. Sheet logging disabled."
                m_LogSheetName = ""
            Else
                Call PrepareSheetHeader(Me, ws) ' Prepare header on new sheet
                 Debug.Print Now() & " | INFO | CONFIG | Initialize | Log sheet '" & m_LogSheetName & "' created."
            End If
            Err.Clear
            On Error GoTo InitializeError
        Else
            Call PrepareSheetHeader(Me, ws) ' Ensure header exists
        End If
    End If

    m_IsInitialized = True ' Mark as initialized
    ' Log initialization using the logger itself now
    LogConsole Me, "Logger Initialized. Min Level: " & modFrameworkUtils.LevelToString(m_MinLogLevel) & ". File: '" & m_LogFilePath & "'. Sheet: '" & m_LogSheetName & "'. Buffer: " & m_BufferSize, LogLevelInfo, "CONFIG", "Initialize", m_CurrentUser

    Set ws = Nothing
    Exit Sub

InitializeError:
HandleLogError Me, "Initialize", Err ' Use internal handler
m_IsInitialized = False
Set ws = Nothing
End Sub

Private Sub ILoggerBase_SetLogger(loggerInstance As ILoggerBase)
Debug.Print Now() & " | WARNING | CONFIG | SetLogger called on clsLogger instance. This is unusual."
End Sub

' --- Logging Methods (Implement ILoggerBase) ---
Private Sub ILoggerBase_LogMessage(ByVal msg As String, Optional ByVal level As LogLevelEnum = LogLevelInfo, Optional ByVal category As String = "", Optional ByVal source As String = "", Optional ByVal user As String = "", Optional ByVal toConsole As Boolean = True, Optional ByVal toSheet As Boolean = False, Optional ByVal toFile As Boolean = True)
Dim logLine As String, levelStr As String, catLower As String, effectiveUser As String
If Not m_IsInitialized Or level < m_MinLogLevel Then Exit Sub
catLower = LCase(category)
If Not IsCategoryEnabled(Me, catLower) Then Exit Sub

    If user <> "" Then effectiveUser = user Else effectiveUser = m_CurrentUser
    levelStr = modFrameworkUtils.LevelToString(level)
    logLine = Format(Now(), "yyyy-mm-dd hh:mm:ss") & " | " & levelStr & " | " & source & " | " & category & " | " & effectiveUser & " | " & msg

    If level >= LogLevelWarning And m_CrashLogBufferSize > 0 Then AddToCrashBuffer Me, logLine
    If toConsole Then ILoggerBase_LogConsole msg, level, category, source, effectiveUser
    If toSheet Then ILoggerBase_LogSheet msg, level, category, source, effectiveUser
    If toFile Then ILoggerBase_LogFile msg, level, category, source, effectiveUser
    If m_SheetBuffer.Count >= m_BufferSize Then FlushSheetBuffer Me
    If m_FileBuffer.Count >= m_BufferSize Then FlushFileBuffer Me

End Sub

Private Sub ILoggerBase_LogConsole(ByVal msg As String, Optional ByVal level As LogLevelEnum = LogLevelInfo, Optional ByVal category As String = "", Optional ByVal source As String = "", Optional ByVal user As String = "")
If Not m_IsInitialized Or level < m_MinLogLevel Or Not IsCategoryEnabled(Me, LCase(category)) Then Exit Sub
Debug.Print Format(Now(), "yyyy-mm-dd hh:mm:ss") & " | " & modFrameworkUtils.LevelToString(level) & " | " & source & " | " & category & " | " & user & " | " & msg
End Sub

Private Sub ILoggerBase_LogSheet(ByVal msg As String, Optional ByVal level As LogLevelEnum = LogLevelInfo, Optional ByVal category As String = "", Optional ByVal source As String = "", Optional ByVal user As String = "")
Dim logEntryData As Variant ' Array(Timestamp, LevelStr, Source, Category, User, Msg)
If Not m_IsInitialized Or level < m_MinLogLevel Or m_LogSheetName = "" Or Not IsCategoryEnabled(Me, LCase(category)) Then Exit Sub
logEntryData = Array(Now(), modFrameworkUtils.LevelToString(level), source, category, user, msg)
m_SheetBuffer.Add Item:=logEntryData
If m_SheetBuffer.Count >= m_BufferSize Then FlushSheetBuffer Me
End Sub

Private Sub ILoggerBase_LogFile(ByVal msg As String, Optional ByVal level As LogLevelEnum = LogLevelInfo, Optional ByVal category As String = "", Optional ByVal source As String = "", Optional ByVal user As String = "")
Dim logLine As String
If Not m_IsInitialized Or level < m_MinLogLevel Or m_LogFilePath = "" Or Not IsCategoryEnabled(Me, LCase(category)) Then Exit Sub
logLine = Format(Now(), "yyyy-mm-dd hh:mm:ss") & " | " & modFrameworkUtils.LevelToString(level) & " | " & source & " | " & category & " | " & user & " | " & msg
m_FileBuffer.Add Item:=logLine
If m_FileBuffer.Count >= m_BufferSize Then FlushFileBuffer Me
End Sub

Private Sub ILoggerBase_LogError(ByVal errObject As ErrObject, Optional ByVal level As LogLevelEnum = LogLevelError, Optional ByVal sourceRoutine As String = "", Optional ByVal category As String = "ERROR", Optional ByVal user As String = "", Optional ByVal toConsole As Boolean = True, Optional ByVal toSheet As Boolean = True, Optional ByVal toFile As Boolean = True)
Dim errorMsg As String, effectiveUser As String
If Not m_IsInitialized Then Debug.Print Now() & " | !! LOGGER NOT INITIALIZED - Error Ignored !!": Exit Sub
If user <> "" Then effectiveUser = user Else effectiveUser = m_CurrentUser
errorMsg = "Error " & errObject.Number & ": " & errObject.Description & " (Source: " & errObject.Source & ")": ILoggerBase_LogMessage errorMsg, level, category, sourceRoutine, effectiveUser, toConsole, toSheet, toFile
End Sub

' --- Control Methods (Implement ILoggerBase) ---
Private Sub ILoggerBase_FlushLogs()
If Not m_IsInitialized Then Exit Sub
On Error Resume Next
FlushSheetBuffer Me
FlushFileBuffer Me
On Error GoTo 0
End Sub

Private Sub ILoggerBase_GenerateCrashReport(Optional ByVal crashFilePath As String = "")
Dim reportPath As String, fileNum As Integer, logEntry As Variant, wbName As String, wbPath As String
If Not m_IsInitialized Or m_CrashLogBufferSize <= 0 Then Exit Sub
On Error GoTo CrashReportError
If crashFilePath = "" Then
wbPath = m_WorkbookInstance.Path
If wbPath = "" Then wbPath = Environ("TEMP")
wbName = modFrameworkUtils.GetBaseFileName(m_WorkbookInstance.Name)
wbName = modFrameworkUtils.SanitizeFilenamePart(wbName)
reportPath = wbPath & Application.PathSeparator & wbName & "_CrashReport_" & Format(Now, "YYYYMMDD_HHMMSS") & ".txt"
Else
reportPath = crashFilePath
End If
fileNum = FreeFile: Open reportPath For Output As #fileNum
Print #fileNum, "=== VBA Crash Report ===" & vbCrLf & "Timestamp: " & Format(Now(), "yyyy-mm-dd hh:mm:ss") & vbCrLf & "Workbook: " & m_WorkbookInstance.Name & vbCrLf & "User: " & m_CurrentUser & vbCrLf & "--- Last " & m_CrashLogBuffer.Count & " Critical Messages ---"
For Each logEntry In m_CrashLogBuffer: Print #fileNum, logEntry: Next logEntry
Print #fileNum, "=== End of Report ==="
Close #fileNum
LogConsole Me, "GenerateCrashReport INFO: Crash report saved to '" & reportPath & "'", LogLevelInfo, "SYSTEM", "GenerateCrashReport", m_CurrentUser
Exit Sub
CrashReportError:
HandleLogError Me, "GenerateCrashReport", Err
On Error Resume Next: Close #fileNum: On Error GoTo 0
End Sub

' --- Read-Only Properties (Implement ILoggerBase) ---
Private Property Get ILoggerBase_IsInitialized() As Boolean: ILoggerBase_IsInitialized = m_IsInitialized: End Property
Private Property Get ILoggerBase_MinLogLevel() As LogLevelEnum: ILoggerBase_MinLogLevel = m_MinLogLevel: End Property

' --- Private Helper Methods ---
Private Function IsCategoryEnabled(instance As clsLogger, catLower As String) As Boolean: Dim i&: IsCategoryEnabled = False: If UBound(instance.m_DisabledCategories) >= LBound(instance.m_DisabledCategories) Then For i = LBound(instance.m_DisabledCategories) To UBound(instance.m_DisabledCategories): If instance.m_DisabledCategories(i) <> "" And catLower = instance.m_DisabledCategories(i) Then Exit Function: Next i: End If: If UBound(instance.m_EnabledCategories) >= LBound(instance.m_EnabledCategories) Then If instance.m_EnabledCategories(0) = "*" Then IsCategoryEnabled = True: Exit Function: End If: For i = LBound(instance.m_EnabledCategories) To UBound(instance.m_EnabledCategories): If instance.m_EnabledCategories(i) <> "" And catLower = instance.m_EnabledCategories(i) Then IsCategoryEnabled = True: Exit Function: End If: Next i: End If: If catLower = "" And instance.m_EnabledCategories(0) = "*" Then IsCategoryEnabled = True: End Function
Private Sub AddToCrashBuffer(instance As clsLogger, logLine As String): If instance.m_CrashLogBufferSize <= 0 Then Exit Sub: instance.m_CrashLogBuffer.Add Item:=logLine: Do While instance.m_CrashLogBuffer.Count > instance.m_CrashLogBufferSize: instance.m_CrashLogBuffer.Remove 1: Loop: End Sub
Private Sub FlushSheetBuffer(instance As clsLogger): Dim ws As Worksheet, lRowStart&, i&, logEntryData, outputArray(), targetRange As Range, currentCalculation As XlCalculation: If instance.m_SheetBuffer Is Nothing Or instance.m_SheetBuffer.Count = 0 Or instance.m_LogSheetName = "" Then Exit Sub: On Error GoTo FlushSheetError: Set ws = instance.m_WorkbookInstance.Sheets(instance.m_LogSheetName): lRowStart = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row: If lRowStart = 1 And Not IsEmpty(ws.Cells(1, 1).Value) Then lRowStart = 2 Else lRowStart = lRowStart + 1: ReDim outputArray(1 To instance.m_SheetBuffer.Count, 1 To SHEET_COLUMN_COUNT): For i = 1 To instance.m_SheetBuffer.Count: logEntryData = instance.m_SheetBuffer(i): outputArray(i, 1) = logEntryData(0): outputArray(i, 2) = logEntryData(1): outputArray(i, 3) = logEntryData(2): outputArray(i, 4) = logEntryData(3): outputArray(i, 5) = logEntryData(4): outputArray(i, 6) = logEntryData(5): Next i: Set targetRange = ws.Cells(lRowStart, 1).Resize(instance.m_SheetBuffer.Count, SHEET_COLUMN_COUNT): currentCalculation = Application.Calculation: Application.EnableEvents = False: Application.ScreenUpdating = False: If currentCalculation <> xlCalculationManual Then Application.Calculation = xlCalculationManual: targetRange.value = outputArray: targetRange.Columns(1).NumberFormat = "yyyy-mm-dd hh:mm:ss": If currentCalculation <> xlCalculationManual Then Application.Calculation = currentCalculation: Application.EnableEvents = True: Application.ScreenUpdating = True: Set instance.m_SheetBuffer = New Collection: Set ws = Nothing: Set targetRange = Nothing: Exit Sub: FlushSheetError: HandleLogError instance, "FlushSheetBuffer", Err: Application.EnableEvents = True: Application.ScreenUpdating = True: If Application.Calculation = xlCalculationManual Then Application.Calculation = currentCalculation: Set ws = Nothing: Set targetRange = Nothing: On Error GoTo 0: End Sub
Private Sub FlushFileBuffer(instance As clsLogger): Dim fileNum%, logLine: If instance.m_FileBuffer Is Nothing Or instance.m_FileBuffer.Count = 0 Or instance.m_LogFilePath = "" Then Exit Sub: On Error GoTo FlushFileError: RotateLogFileIfNeeded instance: fileNum = FreeFile: Open instance.m_LogFilePath For Append As #fileNum: For Each logLine In instance.m_FileBuffer: Print #fileNum, logLine: Next logLine: Close #fileNum: Set instance.m_FileBuffer = New Collection: Exit Sub: FlushFileError: HandleLogError instance, "FlushFileBuffer", Err: On Error Resume Next: Close #fileNum: On Error GoTo 0: End Sub
Private Sub PrepareSheetHeader(instance As clsLogger, ws As Worksheet): Dim headers, headerCheck As Boolean, i&, currentCalculation As XlCalculation: Const HEADER_COLS = SHEET_COLUMN_COUNT: headers = Array("Date/Heure", "Niveau", "Source", "Catégorie", "Utilisateur", "Message"): If ws Is Nothing Then Exit Sub: On Error GoTo PrepareHeaderError: headerCheck = True: For i = 1 To HEADER_COLS: If IsEmpty(ws.Cells(1, i).Value) Then headerCheck = False: Exit For: Next i: If Not headerCheck Then currentCalculation = Application.Calculation: Application.EnableEvents = False: Application.ScreenUpdating = False: If currentCalculation <> xlCalculationManual Then Application.Calculation = xlCalculationManual: With ws.Cells(1, 1).Resize(1, HEADER_COLS): .value = headers: .Font.Bold = True: End With: ws.Columns(1).Resize(, HEADER_COLS).AutoFit: If currentCalculation <> xlCalculationManual Then Application.Calculation = currentCalculation: Application.EnableEvents = True: Application.ScreenUpdating = True: End If: Exit Sub: PrepareHeaderError: HandleLogError instance, "PrepareSheetHeader", Err: Application.EnableEvents = True: Application.ScreenUpdating = True: If Application.Calculation = xlCalculationManual Then Application.Calculation = currentCalculation: On Error GoTo 0: End Sub
Private Sub RotateLogFileIfNeeded(instance As clsLogger): Dim currentSizeKB!, backupFilePath$, dotPos&: If instance.m_LogFilePath = "" Or instance.m_MaxLogFileSizeKB <= 0 Then Exit Sub: On Error GoTo RotationError: If Dir(instance.m_LogFilePath) <> "" Then currentSizeKB = FileLen(instance.m_LogFilePath) / 1024: If currentSizeKB > instance.m_MaxLogFileSizeKB Then dotPos = InStrRev(instance.m_LogFilePath, "."): If dotPos > 0 Then backupFilePath = Left$(instance.m_LogFilePath, dotPos - 1) & "_" & Format$(Now, "yyyymmdd_hhmmss") & Mid$(instance.m_LogFilePath, dotPos) & ".bak" Else backupFilePath = instance.m_LogFilePath & "_" & Format$(Now, "yyyymmdd_hhmmss") & ".bak": End If: Name instance.m_LogFilePath As backupFilePath: instance.LogConsole "RotateLogFile INFO: Log file rotated to '" & backupFilePath & "'...", LogLevelInfo, "SYSTEM": End If: End If: Exit Sub: RotationError: HandleLogError instance, "RotateLogFileIfNeeded", Err: On Error GoTo 0: End Sub
Private Sub HandleLogError(instance As clsLogger, sourceMethod As String, errObj As ErrObject, Optional customMsg As String = ""): Dim errorMsg$, logTimestamp$, adoErr As ADODB.Error, fullSource$: logTimestamp = Format$(Now(), "yyyy-mm-dd hh:mm:ss"): fullSource = TypeName(instance) & "." & sourceMethod: errorMsg = logTimestamp & " | LOGGER INTERNAL ERROR | Method: " & fullSource: If customMsg <> "" Then errorMsg = errorMsg & " | Message: " & customMsg: If Not errObj Is Nothing And errObj.Number <> 0 Then errorMsg = errorMsg & " | Err #: " & errObj.Number & " | Desc: " & errObj.Description & " | Source: " & errObj.Source: Debug.Print errorMsg: End Sub ' Log internal errors to console only to avoid recursion

' --- Class Cleanup ---
Private Sub Class_Terminate()
If m_IsInitialized Then
On Error Resume Next
ILoggerBase_FlushLogs
On Error GoTo 0
End If
Set m_SheetBuffer = Nothing: Set m_FileBuffer = Nothing: Set m_CrashLogBuffer = Nothing: Set m_WorkbookInstance = Nothing
End Sub
```

### clsSilentLogger.cls (v1.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsSilentLogger
' Version : 1.0
' Implements: ILoggerBase (v3.0)
' Purpose : In-memory logger for testing.
' Requires : ILoggerBase, modFrameworkUtils
' ==========================================================================

Implements ILoggerBase

Private m_LogEntries As Collection
Private m_IsInitialized As Boolean
Private m_MinLogLevel As LogLevelEnum
Private m_CurrentUser As String

' --- Public Test Access Methods ---
Public Function GetLogEntries() As Collection: Set GetLogEntries = m_LogEntries: End Function
Public Function GetLogEntriesAsString(Optional delimiter As String = vbCrLf) As String: Dim entry, result$: If Not m_LogEntries Is Nothing Then For Each entry In m_LogEntries: result = result & CStr(entry) & delimiter: Next entry: GetLogEntriesAsString = result: End Function
Public Sub ClearLogs(): Set m_LogEntries = New Collection: End Sub
Public Property Get LogCount() As Long: If m_LogEntries Is Nothing Then LogCount = 0 Else LogCount = m_LogEntries.Count: End Property

' --- ILoggerBase Implementation ---
Private Sub ILoggerBase_Initialize(Optional ByVal minLevel As LogLevelEnum = LogLevelInfo, Optional ByVal logSheetName As String = "", Optional ByVal logFileNamePattern As String = "", Optional ByVal maxLogFileSizeKB As Long = 0, Optional ByVal targetWorkbook As Workbook = Nothing, Optional ByVal enabledCategories As String = "*", Optional ByVal disabledCategories As String = "", Optional ByVal bufferSize As Long = 1, Optional ByVal crashLogBufferSize As Long = 0)
Set m_LogEntries = New Collection: m_IsInitialized = True: m_MinLogLevel = minLevel
On Error Resume Next: m_CurrentUser = Environ("USERNAME"): If m_CurrentUser = "" Then m_CurrentUser = Application.UserName: On Error GoTo 0
End Sub
Private Sub ILoggerBase_SetLogger(loggerInstance As ILoggerBase): End Sub ' NOOP
Private Function ILoggerBase_Connect(connectionString As String, Optional maxRetries As Long = 1, Optional retryDelaySeconds As Long = 2) As Boolean: End Function
Private Sub ILoggerBase_Disconnect(): End Sub
Private Property Get ILoggerBase_IsConnected() As Boolean: End Property

Private Sub ILoggerBase_LogMessage(ByVal msg As String, Optional ByVal level As LogLevelEnum = LogLevelInfo, Optional ByVal category As String = "", Optional ByVal source As String = "", Optional ByVal user As String = "", Optional ByVal toConsole As Boolean = True, Optional ByVal toSheet As Boolean = False, Optional ByVal toFile As Boolean = True)
If Not m_IsInitialized Or level < m_MinLogLevel Then Exit Sub
Dim logLine As String, effectiveUser As String: If user <> "" Then effectiveUser = user Else effectiveUser = m_CurrentUser
logLine = Format(Now(), "yyyy-mm-dd hh:mm:ss") & " | " & modFrameworkUtils.LevelToString(level) & " | " & source & " | " & category & " | " & effectiveUser & " | " & msg
m_LogEntries.Add logLine
End Sub
Private Sub ILoggerBase_LogConsole(msg As String, Optional level As LogLevelEnum = LogLevelInfo, Optional category As String = "", Optional source As String = "", Optional user As String = ""): ILoggerBase_LogMessage msg, level, category, source, user: End Sub
Private Sub ILoggerBase_LogSheet(msg As String, Optional level As LogLevelEnum = LogLevelInfo, Optional category As String = "", Optional source As String = "", Optional user As String = ""): ILoggerBase_LogMessage msg, level, category, source, user: End Sub
Private Sub ILoggerBase_LogFile(msg As String, Optional level As LogLevelEnum = LogLevelInfo, Optional category As String = "", Optional source As String = "", Optional user As String = ""): ILoggerBase_LogMessage msg, level, category, source, user: End Sub
Private Sub ILoggerBase_LogError(errObject As ErrObject, Optional level As LogLevelEnum = LogLevelError, Optional sourceRoutine As String = "", Optional category As String = "ERROR", Optional user As String = "", Optional toConsole As Boolean = True, Optional toSheet As Boolean = True, Optional toFile As Boolean = True)
Dim errorMsg As String: errorMsg = "Error " & errObject.Number & ": " & errObject.Description & " (Source: " & errObject.Source & ")": ILoggerBase_LogMessage errorMsg, level, category, sourceRoutine, user
End Sub
Private Sub ILoggerBase_FlushLogs(): End Sub ' NOOP
Private Sub ILoggerBase_GenerateCrashReport(Optional crashFilePath As String = ""): End Sub ' NOOP for basic silent logger
Private Property Get ILoggerBase_IsInitialized() As Boolean: ILoggerBase_IsInitialized = m_IsInitialized: End Property
Private Property Get ILoggerBase_MinLogLevel() As LogLevelEnum: ILoggerBase_MinLogLevel = m_MinLogLevel: End Property

Private Sub Class_Terminate(): Set m_LogEntries = Nothing: End Sub
```

### modFrameworkUtils.bas (v1.0)

```vba
Option Explicit

' ==========================================================================
' Module : modFrameworkUtils
' Version : 1.0
' Purpose : Common utility functions for the Apex VBA Framework.
' Requires : ADO Reference (for DataTypeEnum constants)
' ==========================================================================

' --- Logging Helpers ---
Public Function LevelToString(ByVal level As LogLevelEnum) As String
Select Case level
Case LogLevelDebug: LevelToString = "DEBUG"
Case LogLevelInfo: LevelToString = "INFO"
Case LogLevelWarning: LevelToString = "WARNING"
Case LogLevelError: LevelToString = "ERROR"
Case LogLevelFatal: LevelToString = "FATAL"
Case Else: LevelToString = "UNKNOWN (" & CLng(level) & ")"
End Select
End Function

' --- ADO / Data Type Helpers ---
Public Function GuessAdoType(value As Variant, Optional defaultStringType As ADODB.DataTypeEnum = adVarWChar) As ADODB.DataTypeEnum
Select Case VarType(value)
Case vbInteger: GuessAdoType = adSmallInt
Case vbLong: GuessAdoType = adInteger
Case vbSingle: GuessAdoType = adSingle
Case vbDouble: GuessAdoType = adDouble
Case vbCurrency: GuessAdoType = adCurrency
Case vbDate: GuessAdoType = adDBTimeStamp
Case vbBoolean: GuessAdoType = adBoolean
Case vbString: GuessAdoType = defaultStringType
Case vbEmpty: GuessAdoType = defaultStringType
Case vbNull: GuessAdoType = defaultStringType
Case vbByte: GuessAdoType = adUnsignedTinyInt
Case vbDecimal: GuessAdoType = adDecimal
Case Else: GuessAdoType = adVariant
End Select
End Function

Public Function AreEqual(val1 As Variant, val2 As Variant) As Boolean
On Error Resume Next: AreEqual = False
If IsNull(val1) And IsNull(val2) Then AreEqual = True
ElseIf IsNull(val1) Or IsNull(val2) Then AreEqual = False
ElseIf IsEmpty(val1) And IsEmpty(val2) Then AreEqual = True
ElseIf IsEmpty(val1) Or IsEmpty(val2) Then AreEqual = False
ElseIf IsObject(val1) Or IsObject(val2) Then AreEqual = False
Else AreEqual = (val1 = val2)
End If
If Err.Number <> 0 Then AreEqual = False
On Error GoTo 0
End Function

' --- String / File Helpers ---
Public Function SanitizeFilenamePart(filenamePart As String) As String
Dim invalidChars$, i&, char$
invalidChars = "/\:*?""<>|"
SanitizeFilenamePart = filenamePart
For i = 1 To Len(invalidChars): char = Mid$(invalidChars, i, 1): SanitizeFilenamePart = Replace$(SanitizeFilenamePart, char, "-"): Next i
SanitizeFilenamePart = Trim$(SanitizeFilenamePart)
End Function

Public Function GetBaseFileName(fullFileName As String) As String
Dim pathSepPos&, dotPos&, tempName$
     pathSepPos = InStrRev(fullFileName, Application.PathSeparator): If pathSepPos > 0 Then tempName = Mid$(fullFileName, pathSepPos + 1) Else tempName = fullFileName
dotPos = InStrRev(tempName, "."): If dotPos > 1 Then GetBaseFileName = Left$(tempName, dotPos - 1) Else GetBaseFileName = tempName
End Function
```

### clsConfigLoader.cls (v2.0 - Refactored)

```vba
Option Explicit
' ==========================================================================
' Class : clsConfigLoader
' Version : 2.0 - Refactored
' Purpose : Reads config settings from Excel sheet into memory for fast access.
' Requires : Worksheet "Config_Framework", Scripting Runtime (for Dictionary).
' ==========================================================================

Private Const CONFIG_SHEET_NAME As String = "Config_Framework"
Private Const COL_SECTION As Long = 1
Private Const COL_KEY As Long = 2
Private Const COL_VALUE As Long = 3

Private m_ConfigData As Object ' Scripting.Dictionary
Private m_IsLoaded As Boolean

Private Sub Class_Initialize()
Set m_ConfigData = Nothing
m_IsLoaded = False
LoadConfig
End Sub

Public Sub LoadConfig()
Dim ws As Worksheet, lastRow&, r&, section$, key$, value, dictKey$
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME): On Error GoTo 0
    If ws Is Nothing Then Debug.Print Now() & " | ERROR | clsConfigLoader: Config sheet '" & CONFIG_SHEET_NAME & "' not found!": Set m_ConfigData = Nothing: m_IsLoaded = False: Exit Sub
    Set m_ConfigData = CreateObject("Scripting.Dictionary")
    m_ConfigData.CompareMode = vbTextCompare
    On Error GoTo LoadError
    lastRow = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row
    For r = 2 To lastRow
        section = Trim$(CStr(ws.Cells(r, COL_SECTION).value)): key = Trim$(CStr(ws.Cells(r, COL_KEY).value)): value = ws.Cells(r, COL_VALUE).value
        If key <> "" Then dictKey = UCase$(key): If section <> "" Then dictKey = UCase$(section) & "." & dictKey
If m_ConfigData.Exists(dictKey) Then Debug.Print Now() & " | WARNING | clsConfigLoader: Duplicate config key '" & dictKey & "'. Overwriting.": m_ConfigData(dictKey) = value Else m_ConfigData.Add dictKey, value
End If
Next r
m_IsLoaded = True
Debug.Print Now() & " | INFO | clsConfigLoader: Configuration loaded. " & m_ConfigData.Count & " settings found."
Set ws = Nothing: Exit Sub
LoadError:
Debug.Print Now() & " | ERROR | clsConfigLoader: Error loading config row " & r & ". Err " & Err.Number & ": " & Err.Description: m_IsLoaded = False: Set m_ConfigData = Nothing: Set ws = Nothing
End Sub

Public Function GetSetting(ByVal key As String, Optional ByVal section As String = "", Optional defaultValue As Variant) As Variant
Dim dictKey$
    If Not m_IsLoaded Or m_ConfigData Is Nothing Then GetSetting = defaultValue: Exit Function
    dictKey = UCase$(Trim$(key)): If section <> "" Then dictKey = UCase$(Trim$(section)) & "." & dictKey
If m_ConfigData.Exists(dictKey) Then GetSetting = m_ConfigData(dictKey) Else GetSetting = defaultValue
End Function

Public Property Get IsConfigLoaded() As Boolean: IsConfigLoaded = m_IsLoaded: End Property
Private Sub Class_Terminate(): Set m_ConfigData = Nothing: End Sub
```

## Outils de déploiement

Le framework Apex VBA propose plusieurs outils pour faciliter le déploiement d'applications:

### ApexInstaller.vbs

Ce script permet l'installation automatisée du framework dans Excel:

```vbs
' ApexInstaller.vbs
Option Explicit

Dim oShell, oFSO, xlApp, addin, installPath, userResponse
Dim addinName : addinName = "ApexVbaFramework.xlam"

' Créer les objets nécessaires
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

' Déterminer le chemin d'installation
installPath = oShell.SpecialFolders("AppData") & "\Microsoft\AddIns"
If Not oFSO.FolderExists(installPath) Then oFSO.CreateFolder(installPath)

' Confirmer l'installation
userResponse = MsgBox("Installer Apex VBA Framework dans " & installPath & "?", vbYesNo + vbQuestion, "Apex Framework Installer")
If userResponse <> vbYes Then
    WScript.Echo "Installation annulée."
    WScript.Quit
End If

' Copier le fichier
If oFSO.FileExists(addinName) Then
    oFSO.CopyFile addinName, installPath & "\" & addinName, True

    ' Enregistrer dans Excel
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    On Error GoTo 0

    xlApp.Workbooks.Add
    xlApp.AddIns.Add installPath & "\" & addinName
    Set addin = xlApp.AddIns(addinName)
    addin.Installed = True

    WScript.Echo "Installation réussie!" & vbCrLf & "Le framework est maintenant disponible dans tous vos classeurs Excel."
Else
    WScript.Echo "Erreur: Fichier " & addinName & " introuvable."
End If

' Nettoyer
Set addin = Nothing
Set xlApp = Nothing
Set oFSO = Nothing
Set oShell = Nothing
```

### ConfigTools.bas

Module d'utilitaires pour la configuration et la vérification de compatibilité:

````vba
Option Explicit

' ==========================================================================
' Module : ConfigTools
' Version : 1.0
' Purpose : Outils de configuration pour le framework Apex VBA
' ==========================================================================

Public Sub VerifierCompatibilite()
Dim msg As String, titre As String
Dim excelVersion As Double
Dim refMissing As Boolean
Dim hasADO As Boolean, hasScripting As Boolean

    ' Vérifier la version d'Excel
    excelVersion = Val(Application.Version)

    ' Vérifier les références
    On Error Resume Next
    hasADO = (ThisWorkbook.VBProject.References("ADODB").Name <> "")
    hasScripting = (ThisWorkbook.VBProject.References("Scripting").Name <> "")
    On Error GoTo 0

    ' Construire le message
    titre = "Apex VBA Framework - Vérification de compatibilité"
    msg = "Résultats de vérification:" & vbCrLf & vbCrLf

    msg = msg & "Version d'Excel: " & Application.Version
    If excelVersion >= 14 Then
        msg = msg & " [OK]" & vbCrLf
    Else
        msg = msg & " [AVERTISSEMENT - Min. recommandé: 14.0]" & vbCrLf
        refMissing = True
    End If

    msg = msg & "Référence ADO: "
    If hasADO Then
        msg = msg & "Présente [OK]" & vbCrLf
    Else
        msg = msg & "Manquante [ERREUR]" & vbCrLf
        refMissing = True
    End If

    msg = msg & "Référence Scripting Runtime: "
    If hasScripting Then
        msg = msg & "Présente [OK]" & vbCrLf
    Else
        msg = msg & "Manquante [ERREUR]" & vbCrLf
        refMissing = True
    End If

    If refMissing Then
        msg = msg & vbCrLf & "Des références requises sont manquantes. Veuillez ajouter les références manquantes via Outils > Références dans l'éditeur VBA."
        MsgBox msg, vbExclamation, titre
    Else
        msg = msg & vbCrLf & "Toutes les vérifications sont réussies. Le système est compatible avec Apex VBA Framework."
        MsgBox msg, vbInformation, titre
    End If
End Sub

Public Sub CreerFeuilleConfig()
Dim ws As Worksheet
Dim existante As Boolean

    ' Vérifier si la feuille existe déjà
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    existante = (Err.Number = 0)
    On Error GoTo 0

    ' Créer ou réinitialiser
    If Not existante Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = CONFIG_SHEET_NAME
    Else
        ws.Cells.Clear
    End If

    ' Préparer l'en-tête
    ws.Cells(1, 1).Value = "Section"
    ws.Cells(1, 2).Value = "Clé"
    ws.Cells(1, 3).Value = "Valeur"

    With ws.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With

    ' Ajouter des exemples de configuration
    ws.Cells(2, 1).Value = "DATABASE"
    ws.Cells(2, 2).Value = "ConnectionString"
    ws.Cells(2, 3).Value = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\chemin\vers\base.accdb;"

    ws.Cells(3, 1).Value = "LOGGING"
    ws.Cells(3, 2).Value = "MinLevel"
    ws.Cells(3, 3).Value = "20"

    ws.Cells(4, 1).Value = "LOGGING"
    ws.Cells(4, 2).Value = "LogToFile"
    ws.Cells(4, 3).Value = "True"

    ' Ajuster largeur des colonnes
    ws.Columns("A:C").AutoFit

    MsgBox "Feuille de configuration créée avec succès.", vbInformation, "Apex Framework Configuration"
End Sub

Public Sub GeneratePackageXLAM()
Dim xlam As Workbook
Dim xlsm As Workbook
Dim VBComp As Object
Dim response As VbMsgBoxResult
Dim targetPath As String

    ' Confirmer l'opération
    response = MsgBox("Cette opération va générer un package XLAM du framework. Continuer?", _
                    vbYesNo + vbQuestion, "Apex Framework - Génération de package")

    If response <> vbYes Then Exit Sub

    ' Créer un nouveau fichier XLAM
    Set xlam = Workbooks.Add
    Application.DisplayAlerts = False
    targetPath = ThisWorkbook.Path & "\ApexVbaFramework.xlam"
    xlam.SaveAs Filename:=targetPath, FileFormat:=xlOpenXMLAddIn
    Application.DisplayAlerts = True

    ' Copier les modules depuis ce classeur
    Set xlsm = ThisWorkbook

    For Each VBComp In xlsm.VBProject.VBComponents
        If VBComp.Type = 1 Or VBComp.Type = 2 Or VBComp.Type = 3 Then
            ' Module standard, classe ou formulaire
            xlsm.VBProject.VBComponents(VBComp.Name).Export xlsm.Path & "\temp_" & VBComp.Name & ".bas"
            xlam.VBProject.VBComponents.Import xlsm.Path & "\temp_" & VBComp.Name & ".bas"
            Kill xlsm.Path & "\temp_" & VBComp.Name & ".bas"
        End If
    Next VBComp

    ' Ajouter une feuille configuration
    xlam.Sheets.Add.Name = CONFIG_SHEET_NAME

    ' Sauvegarder et fermer
    xlam.Save
    xlam.Close

    MsgBox "Package XLAM généré avec succès:" & vbCrLf & targetPath, _
           vbInformation, "Apex Framework - Package généré"
End Sub

### README_INSTALLATION.md

Guide d'installation détaillé pour les utilisateurs finaux:

```markdown
# Guide d'installation d'Apex VBA Framework

Ce guide vous accompagne pas à pas dans l'installation du framework Apex VBA dans votre environnement Excel.

## Méthode 1: Installation automatique (recommandée)

1. Téléchargez le package d'installation complet depuis le dépôt officiel
2. Extrayez les fichiers dans un dossier de votre choix
3. Double-cliquez sur `ApexInstaller.vbs`
4. Suivez les instructions à l'écran

L'installateur va:
- Copier le fichier XLAM dans votre dossier d'add-ins Excel
- Enregistrer l'add-in dans Excel
- Vérifier les prérequis système

## Méthode 2: Installation manuelle

### Étape 1: Copier le fichier d'add-in
1. Localisez le fichier `ApexVbaFramework.xlam`
2. Copiez-le dans votre dossier d'add-ins Excel:
   - `C:\Users\[VotreNom]\AppData\Roaming\Microsoft\AddIns`

### Étape 2: Activer l'add-in dans Excel
1. Ouvrez Excel
2. Accédez à Fichier > Options > Compléments
3. En bas, dans "Gérer:", sélectionnez "Compléments Excel" et cliquez sur "Aller..."
4. Cochez la case à côté de "Apex VBA Framework"
5. Cliquez sur OK

### Étape 3: Vérifier l'installation
1. Ouvrez un nouveau classeur Excel
2. Appuyez sur Alt+F11 pour ouvrir l'éditeur VBA
3. Dans la fenêtre Immediate (Ctrl+G), tapez:
   ```vba
   ?Application.AddIns("ApexVbaFramework.xlam").Installed
````

4. Vérifiez que cela renvoie `True`

## Configuration requise

- Microsoft Excel 2010 (version 14.0) ou supérieur
- Références VBA:
  - Microsoft ActiveX Data Objects 2.8+
  - Microsoft Scripting Runtime 1.0
- Paramètres Excel:
  - Macros activées
  - Accès au modèle d'objet VBA activé

## Résolution des problèmes

### Les références ne se chargent pas

Si vous rencontrez des erreurs liées aux références manquantes:

1. Dans l'éditeur VBA, allez dans Outils > Références
2. Assurez-vous que ces références sont cochées:
   - Microsoft ActiveX Data Objects x.x Library
   - Microsoft Scripting Runtime

### Le complément n'apparaît pas dans la liste

1. Vérifiez que le fichier XLAM est bien dans le dossier AddIns
2. Dans Excel, allez dans Fichier > Options > Avancé
3. Cliquez sur "Emplacement des fichiers..."
4. Vérifiez le chemin "Compléments"

### Erreur de sécurité macros

1. Allez dans Fichier > Options > Centre de gestion de la confidentialité
2. Cliquez sur "Paramètres du Centre de gestion de la confidentialité"
3. Choisissez "Paramètres des macros"
4. Sélectionnez "Activer toutes les macros"

Pour plus d'assistance, contactez votre administrateur système ou visitez notre site d'assistance.

```

## Feuille de route

Le développement du framework Apex VBA suit une feuille de route stratégique divisée en plusieurs phases:

### Version 5.1 (Q3 2025)
- **Performance**: Optimisation des performances pour gestion de grands volumes de données
- **Sécurité**: Améliorations de l'encryption des données sensibles
- **Internationalisation**: Support multilingue complet (EN, FR, ES, DE)
- **UI Framework**: Composants d'interface utilisateur standardisés

### Version 6.0 (Q1 2026)
- **Graphiques avancés**: Intégration avec bibliothèques graphiques modernes
- **IA Integration**: Connecteurs pour services d'intelligence artificielle
- **Cloud Storage**: Synchronisation avec solutions de stockage cloud
- **Modernisation**: Interface de développement améliorée

### Version Long Terme
- **Documentation étendue**: Expansion de la documentation et des exemples
- **Mode Serverless**: Support pour déploiements sans serveur
- **Low-Code Builder**: Assistant visuel de création d'applications
- **Marketplace**: Écosystème d'extensions et modules complémentaires

## Historique des versions

### v5.0.0 (Avril 2025) - Version actuelle
- Architecture complètement refactorisée
- ORM avancé avec support multi-bases
- Système de logging multi-cibles
- Client API REST avec authentification OAuth
- Suite de tests unitaires et d'intégration
- Documentation complète et exemples

### v4.2.0 (Septembre 2024)
- Amélioration des performances du Query Builder
- Support pour SQL Server 2022
- Nouveaux modules d'authentification
- Générateur de rapports amélioré
- Corrections de bugs et optimisations

### v4.0.0 (Janvier 2024)
- Introduction du système ORM
- Gestion de la configuration multi-environnements
- Système de plugins modulaire
- Client API REST initial
- Journalisation avancée

### v3.5.0 (Juin 2023)
- Architecture DB Factory
- Transactions complètes
- Support pour MySQL et PostgreSQL
- Améliorations de sécurité
- Début de support pour environnements multiples

### v3.0.0 (Janvier 2023)
- Première version avec architecture basée sur interfaces
- Système de logging central
- Support SQL Server complet
- Utilitaires Excel améliorés

### v2.0.0 (Mai 2022)
- Réarchitecture modulaire
- Support pour ADO.NET
- Trappeur d'erreurs central
- Premiers utilitaires Excel

### v1.0.0 (Octobre 2021)
- Première version stable
- Fonctionnalités de base Access et Excel
- Logging simple
- Base de première librairie d'utilités

## Conclusion

Le framework Apex VBA représente une solution complète pour le développement professionnel d'applications VBA dans l'écosystème Microsoft Office. Il combine les meilleures pratiques de développement modernes avec la puissance et la flexibilité de VBA.

En utilisant cette architecture modulaire et extensible, les développeurs peuvent:
- Réduire considérablement le temps de développement
- Améliorer la maintenabilité de leurs applications
- Standardiser leurs pratiques de développement
- Réutiliser efficacement le code entre projets

Pour rester informé des dernières mises à jour ou contribuer au projet, consultez notre dépôt de code et notre documentation en ligne.

---

*© 2025 Apex VBA Framework Team - Documentation v5.0.0*
```

### clsTestRunner.cls (v2.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsTestRunner
' Version : 2.0
' Purpose : Lightweight test framework for VBA unit testing
' Features : Test discovery, assertions, reporting
' Requires : ILoggerBase (optional), Scripting Runtime
' ==========================================================================

' --- Enums ---
Public Enum TestResultType
    ResultPass = 0
    ResultFail = 1
    ResultSkip = 2
    ResultError = 3
End Enum

' --- Private Variables ---
Private m_TestClasses As Collection
Private m_Results As Object ' Dictionary
Private m_CurrentClass As String
Private m_CurrentMethod As String
Private m_Logger As ILoggerBase
Private m_TotalTests As Long
Private m_PassedTests As Long
Private m_FailedTests As Long
Private m_SkippedTests As Long
Private m_ErrorTests As Long
Private m_StartTime As Double
Private m_EndTime As Double
Private m_IncludedPatterns As Variant
Private m_ExcludedPatterns As Variant

' --- Class Events ---
Private Sub Class_Initialize()
    Set m_TestClasses = New Collection
    Set m_Results = CreateObject("Scripting.Dictionary")

    m_TotalTests = 0
    m_PassedTests = 0
    m_FailedTests = 0
    m_SkippedTests = 0
    m_ErrorTests = 0
    m_IncludedPatterns = Array("*") ' Default to include all
    m_ExcludedPatterns = Array()    ' Default to exclude none
End Sub

Private Sub Class_Terminate()
    Set m_TestClasses = Nothing
    Set m_Results = Nothing
    Set m_Logger = Nothing
End Sub

' --- Configuration Methods ---
Public Sub RegisterTestClass(ByVal testClass As Object)
    m_TestClasses.Add testClass
End Sub

Public Sub SetLogger(ByVal logger As ILoggerBase)
    Set m_Logger = logger
End Sub
### clsTestCase.cls (v1.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsTestCase
' Version : 1.0
' Purpose : Base class for test cases to inherit from
' Features : Common setup/teardown pattern, test method discovery
' Requires : None
' ==========================================================================

' --- Properties ---
Public Property Get TestMethodPrefix() As String
    TestMethodPrefix = "Test_"
End Property

' --- Test Lifecycle Methods ---
Public Sub SetUp()
    ' Base setup for the entire test class
    ' Override in subclasses
End Sub

Public Sub TearDown()
    ' Base teardown for the entire test class
    ' Override in subclasses
End Sub

Public Sub SetUpTest()
    ' Setup before each test method
    ' Override in subclasses
End Sub

Public Sub TearDownTest()
    ' Teardown after each test method
    ' Override in subclasses
End Sub

' --- Test Method Discovery ---
Public Function GetTestMethods() As Variant
    ' This is a placeholder - test classes should override this
    ' VBA doesn't have reflection, so we need to explicitly list test methods
    ' Return an array of method names
    GetTestMethods = Array()
End Function
```

### TestExample.bas (v1.0)

```vba
Option Explicit
' ==========================================================================
' Module : TestExample
' Version : 1.0
' Purpose : Example tests for the framework
' Requires : clsTestRunner, clsTestCase
' ==========================================================================

' --- Test Class Definition ---
Private Type TExampleTests
    TestCase As clsTestCase
    Runner As clsTestRunner

    ' Test context variables
    TestString As String
    TestNumber As Long
End Type

Private this As TExampleTests

' --- Test Setup ---
Public Sub SetUpTests()
    ' Create test objects
    Set this.TestCase = New clsTestCase
    Set this.Runner = New clsTestRunner

    ' Optional: Set logger
    Dim logger As New clsLogger
    logger.Initialize LogLevelInfo, "TestLog"
    this.Runner.SetLogger logger

    ' Register test class
    this.Runner.RegisterTestClass Me
End Sub

' --- Test Case Implementation ---
Public Sub SetUp()
    ' Called once before running tests in this class
    this.TestString = "Test String"
    this.TestNumber = 42
End Sub

Public Sub TearDown()
    ' Called once after running tests in this class
    this.TestString = ""
    this.TestNumber = 0
End Sub

Public Sub SetUpTest()
    ' Called before each test method
    Debug.Print "Setting up test..."
End Sub

Public Sub TearDownTest()
    ' Called after each test method
    Debug.Print "Tearing down test..."
End Sub

Public Function GetTestMethods() As Variant
    ' Return array of test method names
    GetTestMethods = Array("Test_StringEquality", "Test_NumberAddition", "Test_ShouldFail")
End Function

' --- Test Methods ---
Public Sub Test_StringEquality()
    this.Runner.AssertEqual "Test String", this.TestString, "String should match"
    this.Runner.AssertEqual 10, Len(this.TestString), "String length should be 10"
    this.Runner.AssertTrue InStr(this.TestString, "Test") > 0, "String should contain 'Test'"
End Sub

Public Sub Test_NumberAddition()
    this.Runner.AssertEqual 42, this.TestNumber, "Initial number should be 42"

    Dim result As Long
    result = this.TestNumber + 10

    this.Runner.AssertEqual 52, result, "Adding 10 to 42 should be 52"
    this.Runner.AssertTrue result > 50, "Result should be greater than 50"
End Sub

Public Sub Test_ShouldFail()
    ' This test is designed to fail as an example
    this.Runner.AssertEqual "Wrong String", this.TestString, "This assertion should fail"
End Sub

' --- Test Runner ---
Public Sub RunTests()
    ' Setup test environment
    SetUpTests

    ' Run the tests
    this.Runner.RunAllTests

    ' Output report
    this.Runner.GenerateReport "CONSOLE"

    ' Optionally generate other report formats
    ' this.Runner.GenerateReport "HTML", ThisWorkbook.Path & "\TestResults.html"
    ' this.Runner.GenerateReport "MARKDOWN", ThisWorkbook.Path & "\TestResults.md"

    ' Optional: Automated test summary to immediate window
    PrintTestSummary
End Sub

Private Sub PrintTestSummary()
    Debug.Print "=== TEST SUMMARY ==="
    Debug.Print "Total Tests: " & this.Runner.TotalTests
    Debug.Print "Passed: " & this.Runner.PassedTests
    Debug.Print "Failed: " & this.Runner.FailedTests
    Debug.Print "Errors: " & this.Runner.ErrorTests
    Debug.Print "Skipped: " & this.Runner.SkippedTests
    Debug.Print "Duration: " & Format(this.Runner.TestDuration, "0.000") & " seconds"
    Debug.Print "===================="
End Sub
```

### clsAccessDriver.cls (v2.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsAccessDriver
' Version : 2.0
' Purpose : Database driver for Microsoft Access
' Features : Connection management, query execution, schema information
' Requires : ADO Reference, ILoggerBase (optional)
' ==========================================================================

Private m_Logger As ILoggerBase
Private m_LastError As Long
Private m_LastErrorDescription As String

' --- Driver Information ---
Public Function DriverName() As String
    DriverName = "ACCESS"
End Function

' --- Logging ---
Public Sub SetLogger(ByVal logger As ILoggerBase)
    Set m_Logger = logger
End Sub
### clsSqlServerDriver.cls (v2.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsSqlServerDriver
' Version : 2.0
' Purpose : Database driver for Microsoft SQL Server
' Features : Connection management, query execution, schema information
' Requires : ADO Reference, ILoggerBase (optional)
' ==========================================================================

Private m_Logger As ILoggerBase
Private m_LastError As Long
Private m_LastErrorDescription As String

' --- Driver Information ---
Public Function DriverName() As String
    DriverName = "SQLSERVER"
End Function

' --- Logging ---
Public Sub SetLogger(ByVal logger As ILoggerBase)
    Set m_Logger = logger
End Sub
### clsMySqlDriver.cls (v2.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsMySqlDriver
' Version : 2.0
' Purpose : Database driver for MySQL
' Features : Connection management, query execution, schema information
' Requires : ADO Reference, ILoggerBase (optional)
' ==========================================================================

Private m_Logger As ILoggerBase
Private m_LastError As Long
Private m_LastErrorDescription As String
Private m_DatabaseName As String ' Current database name

' --- Driver Information ---
Public Function DriverName() As String
    DriverName = "MYSQL"
End Function

' --- Logging ---
Public Sub SetLogger(ByVal logger As ILoggerBase)
    Set m_Logger = logger
End Sub
### clsCustomerEntity.cls (v1.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsCustomerEntity
' Version : 1.0
' Purpose : Example entity class for Customer table
' Extends : clsOrmBase
' ==========================================================================

' --- Private Variables ---
Private m_ID As Long
Private m_FirstName As String
Private m_LastName As String
Private m_Email As String
Private m_Phone As String
Private m_Active As Boolean
Private m_CreatedDate As Date
Private m_LastModified As Date

' --- Initialization ---
Private Sub Class_Initialize()
    ' Initialize base class values
    TableName = "Customers"
    PrimaryKeyField = "CustomerID"
    AutoIncrementPK = True

    ' Default values
    m_Active = True
    m_CreatedDate = Now()
    m_LastModified = Now()
End Sub

' --- Required Methods ---
Private Sub MapFields()
    ' Map database fields to class properties
    MapField "CustomerID", "ID", adInteger
    MapField "FirstName", "FirstName", adVarWChar
    MapField "LastName", "LastName", adVarWChar
    MapField "Email", "Email", adVarWChar
    MapField "Phone", "Phone", adVarWChar
    MapField "Active", "Active", adBoolean
    MapField "CreatedDate", "CreatedDate", adDate
    MapField "LastModified", "LastModified", adDate
End Sub

' --- Properties ---
Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let ID(ByVal value As Long)
    m_ID = value
    FieldValue("ID") = value
End Property

Public Property Get FirstName() As String
    FirstName = m_FirstName
End Property

Public Property Let FirstName(ByVal value As String)
    m_FirstName = value
    FieldValue("FirstName") = value
End Property

Public Property Get LastName() As String
    LastName = m_LastName
End Property

Public Property Let LastName(ByVal value As String)
    m_LastName = value
    FieldValue("LastName") = value
End Property

Public Property Get FullName() As String
    FullName = Trim(m_FirstName & " " & m_LastName)
End Property

Public Property Get Email() As String
    Email = m_Email
End Property

Public Property Let Email(ByVal value As String)
    m_Email = value
    FieldValue("Email") = value
End Property

Public Property Get Phone() As String
    Phone = m_Phone
End Property

Public Property Let Phone(ByVal value As String)
    m_Phone = value
    FieldValue("Phone") = value
End Property

Public Property Get Active() As Boolean
    Active = m_Active
End Property

Public Property Let Active(ByVal value As Boolean)
    m_Active = value
    FieldValue("Active") = value
End Property

Public Property Get CreatedDate() As Date
    CreatedDate = m_CreatedDate
End Property

Public Property Let CreatedDate(ByVal value As Date)
    m_CreatedDate = value
    FieldValue("CreatedDate") = value
End Property

Public Property Get LastModified() As Date
    LastModified = m_LastModified
End Property

Public Property Let LastModified(ByVal value As Date)
    m_LastModified = value
    FieldValue("LastModified") = value
End Property

' --- Custom Methods ---
Public Function ValidateBeforeSave() As Boolean
    ' Basic validation logic
    If Trim(m_FirstName) = "" Then
        ' Log error or handle validation failure
        ValidateBeforeSave = False
        Exit Function
    End If

    If Trim(m_LastName) = "" Then
        ' Log error or handle validation failure
        ValidateBeforeSave = False
        Exit Function
    End If

    ' Email validation (basic)
    If m_Email <> "" Then
        If InStr(1, m_Email, "@") <= 0 Or InStr(1, m_Email, ".") <= 0 Then
            ' Invalid email format
            ValidateBeforeSave = False
            Exit Function
        End If
    End If

    ' Update last modified date
    m_LastModified = Now()
    FieldValue("LastModified") = m_LastModified

    ValidateBeforeSave = True
End Function

Public Function SaveWithValidation() As Boolean
    ' Validate before saving
    If Not ValidateBeforeSave() Then
        SaveWithValidation = False
        Exit Function
    End If

    ' Call base class Save method
    SaveWithValidation = Save()
End Function

Public Function GetContactInfo() As String
    Dim result As String

    result = FullName

    If m_Email <> "" Then
        result = result & " - Email: " & m_Email
    End If

    If m_Phone <> "" Then
        result = result & " - Phone: " & m_Phone
    End If

    GetContactInfo = result
End Function
```

### clsProductEntity.cls (v1.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsProductEntity
' Version : 1.0
' Purpose : Example entity class for Product table
' Extends : clsOrmBase
' ==========================================================================

' --- Private Variables ---
Private m_ProductID As Long
Private m_ProductName As String
Private m_Description As String
Private m_Price As Currency
Private m_StockQuantity As Long
Private m_CategoryID As Long
Private m_IsActive As Boolean
Private m_CreatedDate As Date
Private m_LastModified As Date

' --- Initialization ---
Private Sub Class_Initialize()
    ' Initialize base class values
    TableName = "Products"
    PrimaryKeyField = "ProductID"
    AutoIncrementPK = True

    ' Default values
    m_IsActive = True
    m_CreatedDate = Now()
    m_LastModified = Now()
    m_StockQuantity = 0
    m_Price = 0
End Sub

' --- Required Methods ---
Private Sub MapFields()
    ' Map database fields to class properties
    MapField "ProductID", "ProductID", adInteger
    MapField "ProductName", "ProductName", adVarWChar
    MapField "Description", "Description", adLongVarWChar
    MapField "Price", "Price", adCurrency
    MapField "StockQuantity", "StockQuantity", adInteger
    MapField "CategoryID", "CategoryID", adInteger
    MapField "IsActive", "IsActive", adBoolean
    MapField "CreatedDate", "CreatedDate", adDate
    MapField "LastModified", "LastModified", adDate
End Sub

' --- Properties ---
Public Property Get ProductID() As Long
    ProductID = m_ProductID
End Property

Public Property Let ProductID(ByVal value As Long)
    m_ProductID = value
    FieldValue("ProductID") = value
End Property

Public Property Get ProductName() As String
    ProductName = m_ProductName
End Property

Public Property Let ProductName(ByVal value As String)
    m_ProductName = value
    FieldValue("ProductName") = value
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Description(ByVal value As String)
    m_Description = value
    FieldValue("Description") = value
End Property

Public Property Get Price() As Currency
    Price = m_Price
End Property

Public Property Let Price(ByVal value As Currency)
    m_Price = value
    FieldValue("Price") = value
End Property

Public Property Get StockQuantity() As Long
    StockQuantity = m_StockQuantity
End Property

Public Property Let StockQuantity(ByVal value As Long)
    m_StockQuantity = value
    FieldValue("StockQuantity") = value
End Property

Public Property Get CategoryID() As Long
    CategoryID = m_CategoryID
End Property

Public Property Let CategoryID(ByVal value As Long)
    m_CategoryID = value
    FieldValue("CategoryID") = value
End Property

Public Property Get IsActive() As Boolean
    IsActive = m_IsActive
End Property

Public Property Let IsActive(ByVal value As Boolean)
    m_IsActive = value
    FieldValue("IsActive") = value
End Property

Public Property Get CreatedDate() As Date
    CreatedDate = m_CreatedDate
End Property

Public Property Let CreatedDate(ByVal value As Date)
    m_CreatedDate = value
    FieldValue("CreatedDate") = value
End Property

Public Property Get LastModified() As Date
    LastModified = m_LastModified
End Property

Public Property Let LastModified(ByVal value As Date)
    m_LastModified = value
    FieldValue("LastModified") = value
End Property

' --- Custom Methods ---
Public Function ValidateBeforeSave() As Boolean
    ' Basic validation logic
    If Trim(m_ProductName) = "" Then
        ' Product name is required
        ValidateBeforeSave = False
        Exit Function
    End If

    If m_Price < 0 Then
        ' Price cannot be negative
        ValidateBeforeSave = False
        Exit Function
    End If

    If m_StockQuantity < 0 Then
        ' Stock quantity cannot be negative
        ValidateBeforeSave = False
        Exit Function
    End If

    ' Update last modified date
    m_LastModified = Now()
    FieldValue("LastModified") = m_LastModified

    ValidateBeforeSave = True
End Function

Public Function SaveWithValidation() As Boolean
    ' Validate before saving
    If Not ValidateBeforeSave() Then
        SaveWithValidation = False
        Exit Function
    End If

    ' Call base class Save method
    SaveWithValidation = Save()
End Function

Public Function IsInStock() As Boolean
    IsInStock = (m_StockQuantity > 0)
End Function

Public Function ApplyDiscount(ByVal discountPercent As Double) As Currency
    Dim discountedPrice As Currency

    ' Apply discount
    discountedPrice = m_Price * (1 - discountPercent / 100)

    ' Return discounted price (without changing the product's price)
    ApplyDiscount = discountedPrice
End Function

Public Sub AdjustStock(ByVal quantity As Long)
    ' Adjust stock quantity
    m_StockQuantity = m_StockQuantity + quantity
    FieldValue("StockQuantity") = m_StockQuantity
End Function
```

### clsOrderEntity.cls (v1.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsOrderEntity
' Version : 1.0
' Purpose : Example entity class for Order table
' Extends : clsOrmBase
' ==========================================================================

' --- Order Status Enum ---
Public Enum OrderStatusEnum
    OrderStatusPending = 1
    OrderStatusProcessing = 2
    OrderStatusShipped = 3
    OrderStatusDelivered = 4
    OrderStatusCancelled = 5
End Enum

' --- Private Variables ---
Private m_OrderID As Long
Private m_CustomerID As Long
Private m_OrderDate As Date
Private m_Status As OrderStatusEnum
Private m_TotalAmount As Currency
Private m_Notes As String
Private m_LastModified As Date

' --- Initialization ---
Private Sub Class_Initialize()
    ' Initialize base class values
    TableName = "Orders"
    PrimaryKeyField = "OrderID"
    AutoIncrementPK = True

    ' Default values
    m_OrderDate = Now()
    m_Status = OrderStatusPending
    m_TotalAmount = 0
    m_LastModified = Now()
End Sub

' --- Required Methods ---
Private Sub MapFields()
    ' Map database fields to class properties
    MapField "OrderID", "OrderID", adInteger
    MapField "CustomerID", "CustomerID", adInteger
    MapField "OrderDate", "OrderDate", adDate
    MapField "Status", "Status", adInteger
    MapField "TotalAmount", "TotalAmount", adCurrency
    MapField "Notes", "Notes", adLongVarWChar
    MapField "LastModified", "LastModified", adDate
End Sub

' --- Properties ---
Public Property Get OrderID() As Long
    OrderID = m_OrderID
End Property

Public Property Let OrderID(ByVal value As Long)
    m_OrderID = value
    FieldValue("OrderID") = value
End Property

Public Property Get CustomerID() As Long
    CustomerID = m_CustomerID
End Property

Public Property Let CustomerID(ByVal value As Long)
    m_CustomerID = value
    FieldValue("CustomerID") = value
End Property

Public Property Get OrderDate() As Date
    OrderDate = m_OrderDate
End Property

Public Property Let OrderDate(ByVal value As Date)
    m_OrderDate = value
    FieldValue("OrderDate") = value
End Property

Public Property Get Status() As OrderStatusEnum
    Status = m_Status
End Property

Public Property Let Status(ByVal value As OrderStatusEnum)
    m_Status = value
    FieldValue("Status") = value
End Property

Public Property Get TotalAmount() As Currency
    TotalAmount = m_TotalAmount
End Property

Public Property Let TotalAmount(ByVal value As Currency)
    m_TotalAmount = value
    FieldValue("TotalAmount") = value
End Property

Public Property Get Notes() As String
    Notes = m_Notes
End Property

Public Property Let Notes(ByVal value As String)
    m_Notes = value
    FieldValue("Notes") = value
End Property

Public Property Get LastModified() As Date
    LastModified = m_LastModified
End Property

Public Property Let LastModified(ByVal value As Date)
    m_LastModified = value
    FieldValue("LastModified") = value
End Property

' --- Custom Methods ---
Public Function ValidateBeforeSave() As Boolean
    ' Basic validation logic
    If m_CustomerID <= 0 Then
        ' Customer ID is required
        ValidateBeforeSave = False
        Exit Function
    End If

    If m_Status < OrderStatusPending Or m_Status > OrderStatusCancelled Then
        ' Invalid status
        ValidateBeforeSave = False
        Exit Function
    End If

    ' Update last modified date
    m_LastModified = Now()
    FieldValue("LastModified") = m_LastModified

    ValidateBeforeSave = True
End Function

Public Function SaveWithValidation() As Boolean
    ' Validate before saving
    If Not ValidateBeforeSave() Then
        SaveWithValidation = False
        Exit Function
    End If

    ' Call base class Save method
    SaveWithValidation = Save()
End Function

Public Function GetStatusText() As String
    Select Case m_Status
        Case OrderStatusPending
            GetStatusText = "Pending"
        Case OrderStatusProcessing
            GetStatusText = "Processing"
        Case OrderStatusShipped
            GetStatusText = "Shipped"
        Case OrderStatusDelivered
            GetStatusText = "Delivered"
        Case OrderStatusCancelled
            GetStatusText = "Cancelled"
        Case Else
            GetStatusText = "Unknown"
    End Select
End Function

Public Function CanCancel() As Boolean
    ' Can only cancel if order is pending or processing
    CanCancel = (m_Status = OrderStatusPending Or m_Status = OrderStatusProcessing)
End Function

Public Function CanUpdate() As Boolean
    ' Cannot update if order is delivered or cancelled
    CanUpdate = (m_Status <> OrderStatusDelivered And m_Status <> OrderStatusCancelled)
End Function

Public Sub CancelOrder()
    ' Check if order can be cancelled
    If CanCancel() Then
        m_Status = OrderStatusCancelled
        FieldValue("Status") = m_Status
        m_LastModified = Now()
        FieldValue("LastModified") = m_LastModified
    End If
End Sub

Public Function GetDaysSinceOrder() As Long
    GetDaysSinceOrder = DateDiff("d", m_OrderDate, Date)
End Function
```

### clsOrmEntityGenerator.cls (v1.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsOrmEntityGenerator
' Version : 1.0
' Purpose : Generates ORM entity classes from database tables
' Requires : IDbAccessorBase, clsConfigLoader
' ==========================================================================

' --- Private Variables ---
Private m_DbAccessor As IDbAccessorBase
Private m_Logger As ILoggerBase
Private m_OutputPath As String
Private m_Namespace As String
Private m_ClassPrefix As String
Private m_ClassSuffix As String

' --- Initialization ---
Private Sub Class_Initialize()
    ' Default settings
    m_OutputPath = ThisWorkbook.Path & "\Generated\"
    m_Namespace = ""
    m_ClassPrefix = "cls"
    m_ClassSuffix = "Entity"
End Sub

' --- Configuration Methods ---
Public Sub SetDatabase(ByVal dbAccessor As IDbAccessorBase)
    Set m_DbAccessor = dbAccessor
End Sub

Public Sub SetLogger(ByVal logger As ILoggerBase)
    Set m_Logger = logger
End Sub
### clsOAuth2Provider.cls (v1.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsOAuth2Provider
' Version : 1.0
' Purpose : OAuth 2.0 authentication for REST API calls
' Features : Token acquisition, refresh, storage
' Requires : clsRestClient, clsConfigLoader, ILoggerBase (optional)
' ==========================================================================

' --- Enums ---
Public Enum OAuth2GrantType
    GrantTypeAuthCode = 0
    GrantTypeClientCredentials = 1
    GrantTypePassword = 2
    GrantTypeRefreshToken = 3
End Enum

' --- Private Variables ---
Private m_ClientId As String
Private m_ClientSecret As String
Private m_AuthUrl As String
Private m_TokenUrl As String
Private m_RedirectUri As String
Private m_Scope As String
Private m_GrantType As OAuth2GrantType
Private m_AccessToken As String
Private m_RefreshToken As String
Private m_TokenType As String
Private m_TokenExpiryTime As Double
Private m_UserName As String
Private m_Password As String
Private m_RestClient As clsRestClient
Private m_Logger As ILoggerBase

' --- Constants ---
Private Const TOKEN_BUFFER_SECONDS As Long = 300 ' 5 minutes buffer before token expires

' --- Initialization ---
Private Sub Class_Initialize()
    Set m_RestClient = New clsRestClient
    m_GrantType = GrantTypeAuthCode
    m_TokenExpiryTime = 0
End Sub

Private Sub Class_Terminate()
    Set m_RestClient = Nothing
    Set m_Logger = Nothing
End Sub

' --- Configuration Methods ---
Public Sub SetLogger(ByVal logger As ILoggerBase)
    Set m_Logger = logger
    If Not m_RestClient Is Nothing Then
        m_RestClient.SetLogger m_Logger
    End If
End Sub
### clsJsonParser.cls (v1.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsJsonParser
' Version : 1.0
' Purpose : Simple JSON parser and generator for VBA
' Requires : Microsoft Scripting Runtime (Dictionary object)
' ==========================================================================

' --- Private Variables ---
Private m_LastError As String

' --- Parsing Methods ---
Public Function ParseJson(ByVal jsonString As String) As Object
    Dim result As Object
    Dim jsonProcessor As Object

    On Error GoTo ErrorHandler

    ' Try to use built-in JSON parsing via Script Control (if available)
    Set jsonProcessor = TryCreateScriptControl()

    If Not jsonProcessor Is Nothing Then
        Set result = ParseJsonUsingScriptControl(jsonProcessor, jsonString)
    Else
        ' Fallback to simple parsing
        Set result = ParseJsonSimple(jsonString)
    End If

    Set ParseJson = result
    Exit Function

ErrorHandler:
    m_LastError = "Error parsing JSON: " & Err.Description
    Set ParseJson = CreateObject("Scripting.Dictionary")
End Function

Public Function StringifyJson(ByVal data As Variant) As String
    Dim result As String

    On Error GoTo ErrorHandler

    ' Generate JSON string based on data type
    If IsObject(data) Then
        If TypeName(data) = "Dictionary" Then
            result = DictionaryToJson(data)
        ElseIf TypeName(data) = "Collection" Then
            result = CollectionToJson(data)
        Else
            ' Try to handle generic object
            result = ObjectToJson(data)
        End If
    ElseIf IsArray(data) Then
        result = ArrayToJson(data)
    Else
        ' Simple value
        result = ScalarToJson(data)
    End If

    StringifyJson = result
    Exit Function

ErrorHandler:
    m_LastError = "Error generating JSON: " & Err.Description
    StringifyJson = "{""error"":""" & EscapeJsonString(m_LastError) & """}"
End Function

' --- Property Accessors ---
Public Property Get LastError() As String
    LastError = m_LastError
End Property

' --- Private Methods ---
Private Function TryCreateScriptControl() As Object
    On Error Resume Next

    ' Try to create Microsoft Script Control
    Set TryCreateScriptControl = CreateObject("MSScriptControl.ScriptControl")

    If Not TryCreateScriptControl Is Nothing Then
        TryCreateScriptControl.Language = "JScript"
    End If

    On Error GoTo 0
End Function

Private Function ParseJsonUsingScriptControl(ByVal sc As Object, ByVal jsonString As String) As Object
    ' Use Script Control to parse JSON
    On Error Resume Next

    ' Fix invalid JSON that might cause issues
    jsonString = FixJsonString(jsonString)

    ' Add JSON parsing function
    sc.AddCode "function parseJson(jsonString) { return eval('(' + jsonString + ')'); }"

    ' Parse JSON
    Dim jsObject As Object
    Set jsObject = sc.Run("parseJson", jsonString)

    ' Convert to VBA objects
    Set ParseJsonUsingScriptControl = ConvertJsObject(jsObject)

    On Error GoTo 0
End Function

Private Function ConvertJsObject(ByVal jsObject As Object) As Object
    Dim result As Object
    Dim key As Variant, value As Variant
    Dim i As Long

    On Error Resume Next

    ' Check if it's an array
    If TypeName(jsObject) = "JScriptTypeInfo" Then
        ' Special case for arrays
        If IsArray(jsObject) Then
            Set result = CreateObject("Scripting.Dictionary")
            For i = 0 To jsObject.length - 1
                result.Add CStr(i), ConvertJsObject(jsObject(i))
            Next i
        Else
            ' Create dictionary for object
            Set result = CreateObject("Scripting.Dictionary")

            ' Get all properties
            For Each key In jsObject
                Set value = jsObject(key)

                ' Convert property value
                If IsObject(value) Then
                    result.Add key, ConvertJsObject(value)
                Else
                    result.Add key, value
                End If
            Next key
        End If
    Else
        ' Direct value, just return it
        Set result = CreateObject("Scripting.Dictionary")
        result.Add "value", jsObject
    End If

    On Error GoTo 0

    Set ConvertJsObject = result
End Function

Private Function FixJsonString(ByVal jsonString As String) As String
    ' Fix common JSON issues that might cause parse errors
    Dim result As String
    result = jsonString

    ' Trim spaces
    result = Trim(result)

    ' Ensure valid quotation marks
    result = Replace(result, "'", """")

    ' Fix trailing commas
    result = Replace(result, ",]", "]")
    result = Replace(result, ",}", "}")

    FixJsonString = result
End Function

Private Function ParseJsonSimple(ByVal jsonString As String) As Object
    ' Simple implementation of JSON parsing
    ' This is a very basic implementation and won't handle all cases
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    ' Trim the string
    jsonString = Trim(jsonString)

    ' Check if it's an object
    If Left(jsonString, 1) = "{" And Right(jsonString, 1) = "}" Then
        ' Object
        ParseJsonObject jsonString, result
    ElseIf Left(jsonString, 1) = "[" And Right(jsonString, 1) = "]" Then
        ' Array
        ParseJsonArray jsonString, result
    End If

    Set ParseJsonSimple = result
End Function

Private Sub ParseJsonObject(ByVal jsonString As String, ByRef result As Object)
    ' Parse JSON object "{key:value, key2:value2}"
    Dim content As String
    Dim pairs() As String
    Dim pair As String
    Dim key As String, value As String
    Dim i As Long, colonPos As Long

    ' Remove braces
    content = Mid(jsonString, 2, Len(jsonString) - 2)

    ' Split by commas, respecting nested objects and arrays
    pairs = SplitJsonPairs(content)

    For i = LBound(pairs) To UBound(pairs)
        pair = Trim(pairs(i))

        ' Find first colon to separate key and value
        colonPos = InStr(1, pair, ":")

        If colonPos > 0 Then
            ' Extract key and value
            key = Trim(Left(pair, colonPos - 1))
            value = Trim(Mid(pair, colonPos + 1))

            ' Remove quotes from key
            If Left(key, 1) = """" And Right(key, 1) = """" Then
                key = Mid(key, 2, Len(key) - 2)
            End If

            ' Process value based on type
            If Left(value, 1) = "{" And Right(value, 1) = "}" Then
                ' Nested object
                Dim nestedObj As Object
                Set nestedObj = CreateObject("Scripting.Dictionary")
                ParseJsonObject value, nestedObj
                result.Add key, nestedObj
            ElseIf Left(value, 1) = "[" And Right(value, 1) = "]" Then
                ' Array
                Dim nestedArray As Object
                Set nestedArray = CreateObject("Scripting.Dictionary")
                ParseJsonArray value, nestedArray
                result.Add key, nestedArray
            Else
                ' Simple value
                result.Add key, ConvertJsonValue(value)
            End If
        End If
    Next i
End Sub

Private Sub ParseJsonArray(ByVal jsonString As String, ByRef result As Object)
    ' Parse JSON array "[value1, value2, value3]"
    Dim content As String
    Dim items() As String
    Dim item As String
    Dim i As Long

    ' Remove brackets
    content = Mid(jsonString, 2, Len(jsonString) - 2)

    ' Split by commas, respecting nested objects and arrays
    items = SplitJsonPairs(content)

    For i = LBound(items) To UBound(items)
        item = Trim(items(i))

        ' Process value based on type
        If Left(item, 1) = "{" And Right(item, 1) = "}" Then
            ' Nested object
            Dim nestedObj As Object
            Set nestedObj = CreateObject("Scripting.Dictionary")
            ParseJsonObject item, nestedObj
            result.Add CStr(i), nestedObj
        ElseIf Left(item, 1) = "[" And Right(item, 1) = "]" Then
            ' Nested array
            Dim nestedArray As Object
            Set nestedArray = CreateObject("Scripting.Dictionary")
            ParseJsonArray item, nestedArray
            result.Add CStr(i), nestedArray
        Else
            ' Simple value
            result.Add CStr(i), ConvertJsonValue(item)
        End If
    Next i
End Sub

Private Function SplitJsonPairs(ByVal jsonContent As String) As String()
    ' Split JSON content by commas, respecting nested structures
    Dim result() As String
    Dim currentElement As String
    Dim i As Long, c As String
    Dim inQuote As Boolean
    Dim braceCount As Long, bracketCount As Long
    Dim elementCount As Long

    ReDim result(0 To 0)
    elementCount = 0

    For i = 1 To Len(jsonContent)
        c = Mid(jsonContent, i, 1)

        ' Handle quotes (respecting escapes)
        If c = """" And (i = 1 Or Mid(jsonContent, i - 1, 1) <> "\") Then
            inQuote = Not inQuote
        End If

        ' Count nested structures (only if not in quotes)
        If Not inQuote Then
            If c = "{" Then braceCount = braceCount + 1
            If c = "}" Then braceCount = braceCount - 1
            If c = "[" Then bracketCount = bracketCount + 1
            If c = "]" Then bracketCount = bracketCount - 1
        End If

        ' Check for separators at root level
        If c = "," And Not inQuote And braceCount = 0 And bracketCount = 0 Then
            ' Add current element to results
            result(elementCount) = currentElement
            elementCount = elementCount + 1
            ReDim Preserve result(0 To elementCount)
            currentElement = ""
        Else
            ' Add to current element
            currentElement = currentElement & c
        End If
    Next i

    ' Add the last element
    If currentElement <> "" Then
        result(elementCount) = currentElement
        elementCount = elementCount + 1
    End If

    ' Resize to actual size
    If elementCount > 0 Then
        ReDim Preserve result(0 To elementCount - 1)
    End If

    SplitJsonPairs = result
End Function

Private Function ConvertJsonValue(ByVal value As String) As Variant
    ' Convert JSON string value to appropriate VBA type
    value = Trim(value)

    ' String
    If Left(value, 1) = """" And Right(value, 1) = """" Then
        ConvertJsonValue = Mid(value, 2, Len(value) - 2)
        Exit Function
    End If

    ' Boolean
    If LCase(value) = "true" Then
        ConvertJsonValue = True
        Exit Function
    End If
    If LCase(value) = "false" Then
        ConvertJsonValue = False
        Exit Function
    End If

    ' Null
    If LCase(value) = "null" Then
        ConvertJsonValue = Null
        Exit Function
    End If

    ' Number (integer or floating point)
    If IsNumeric(value) Then
        If InStr(1, value, ".") > 0 Then
            ConvertJsonValue = CDbl(value)
        Else
            ConvertJsonValue = CLng(value)
        End If
        Exit Function
    End If

    ' Default: return as is
    ConvertJsonValue = value
End Function

' --- JSON Generation Methods ---
Private Function DictionaryToJson(ByVal dict As Object) As String
    Dim result As String
    Dim key As Variant, value As Variant
    Dim first As Boolean

    result = "{"
    first = True

    ' Process each key-value pair
    For Each key In dict.Keys
        If Not first Then result = result & ","

        ' Add key (ensure it's a string)
        result = result & """" & EscapeJsonString(CStr(key)) & """:"

        ' Add value based on type
        value = dict(key)

        If IsObject(value) Then
            If TypeName(value) = "Dictionary" Then
                result = result & DictionaryToJson(value)
            ElseIf TypeName(value) = "Collection" Then
                result = result & CollectionToJson(value)
            Else
                ' Try to handle generic object
                result = result & ObjectToJson(value)
            End If
        ElseIf IsArray(value) Then
            result = result & ArrayToJson(value)
        Else
            ' Simple value
            result = result & ScalarToJson(value)
        End If

        first = False
    Next key

    result = result & "}"
    DictionaryToJson = result
End Function

Private Function CollectionToJson(ByVal col As Object) As String
    Dim result As String
    Dim item As Variant
    Dim first As Boolean

    result = "["
    first = True

    ' Process each item
    For Each item In col
        If Not first Then result = result & ","

        If IsObject(item) Then
            If TypeName(item) = "Dictionary" Then
                result = result & DictionaryToJson(item)
            ElseIf TypeName(item) = "Collection" Then
                result = result & CollectionToJson(item)
            Else
                ' Try to handle generic object
                result = result & ObjectToJson(item)
            End If
        ElseIf IsArray(item) Then
            result = result & ArrayToJson(item)
        Else
            ' Simple value
            result = result & ScalarToJson(item)
        End If

        first = False
    Next item

    result = result & "]"
    CollectionToJson = result
End Function

Private Function ArrayToJson(ByVal arr As Variant) As String
    Dim result As String
    Dim i As Long
    Dim first As Boolean

    result = "["
    first = True

    ' Process each item
    On Error Resume Next
    For i = LBound(arr) To UBound(arr)
        If Err.Number <> 0 Then Exit For

        If Not first Then result = result & ","

        If IsObject(arr(i)) Then
            If TypeName(arr(i)) = "Dictionary" Then
                result = result & DictionaryToJson(arr(i))
            ElseIf TypeName(arr(i)) = "Collection" Then
                result = result & CollectionToJson(arr(i))
            Else
                ' Try to handle generic object
                result = result & ObjectToJson(arr(i))
            End If
        ElseIf IsArray(arr(i)) Then
            result = result & ArrayToJson(arr(i))
        Else
            ' Simple value
            result = result & ScalarToJson(arr(i))
        End If

        first = False
    Next i
    On Error GoTo 0

    result = result & "]"
    ArrayToJson = result
End Function

Private Function ObjectToJson(ByVal obj As Object) As String
    Dim result As String
    Dim props As Variant
    Dim i As Long
    Dim propName As String, propValue As Variant

    ' Try to get object properties
    On Error Resume Next

    ' Try common property access patterns
    ' This is a simplified approach - real reflection in VBA is limited
    result = "{"

    ' Attempt to enumerate public properties
    Dim hasProps As Boolean
    hasProps = False

    ' Try TypeInfo approach if available
    Dim typeInfo As Object
    Set typeInfo = Nothing

    ' TODO: Add more sophisticated reflection if needed

    ' If no properties found, just return the object name
    If Not hasProps Then
        result = "{""object"":""" & TypeName(obj) & """}"
    Else
        result = result & "}"
    End If

    On Error GoTo 0
    ObjectToJson = result
End Function

Private Function ScalarToJson(ByVal value As Variant) As String
    ' Convert a simple value to JSON representation
    Select Case VarType(value)
        Case vbNull
            ScalarToJson = "null"
        Case vbEmpty
            ScalarToJson = "null"
        Case vbBoolean
            ScalarToJson = IIf(value, "true", "false")
        Case vbString
            ScalarToJson = """" & EscapeJsonString(value) & """"
        Case vbDate
            ScalarToJson = """" & Format(value, "yyyy-mm-dd") & """"
        Case vbDouble, vbLong, vbInteger, vbSingle, vbCurrency, vbDecimal
            ScalarToJson = CStr(value)
        Case Else
            ScalarToJson = """" & EscapeJsonString(CStr(value)) & """"
    End Select
End Function

Private Function EscapeJsonString(ByVal text As String) As String
    ' Escape special characters in JSON strings
    Dim result As String

    result = text
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCr, "\r")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")

    EscapeJsonString = result
End Function
```

### clsExcelHelper.cls (v1.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsExcelHelper
' Version : 1.0
' Purpose : Helper utilities for working with Excel
' Features : Range operations, sheet management, formatting
' Requires : Excel, ILoggerBase (optional)
' ==========================================================================

' --- Private Variables ---
Private m_Logger As ILoggerBase
Private m_LastError As String

' --- Initialization ---
Public Sub SetLogger(ByVal logger As ILoggerBase)
    Set m_Logger = logger
End Sub
### clsDataValidator.cls (v1.0)

```vba
Option Explicit
' ==========================================================================
' Class : clsDataValidator
' Version : 1.0
' Purpose : Validate data for various scenarios
' Features : Type validation, range checking, pattern matching
' Requires : ILoggerBase (optional)
' ==========================================================================

' --- Private Variables ---
Private m_Logger As ILoggerBase
Private m_ErrorList As Collection
Private m_LastError As String

' --- Initialization ---
Private Sub Class_Initialize()
    ' Initialize error collection
    Set m_ErrorList = New Collection
End Sub

Private Sub Class_Terminate()
    ' Clean up
    Set m_ErrorList = Nothing
    Set m_Logger = Nothing
End Sub

Public Sub SetLogger(ByVal logger As ILoggerBase)
    Set m_Logger = logger
End Sub
