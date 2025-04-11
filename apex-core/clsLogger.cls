' Migrated to apex-core - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit
' ==========================================================================
' Class : clsLogger
' Version : 1.0
' Purpose : Système de journalisation centralisé pour le framework
' Date : 10/04/2025
' ==========================================================================

' --- Énumérations ---
Public Enum LogLevel
    lvlDebug = 0    ' Informations détaillées pour le débogage
    lvlInfo = 1     ' Informations générales sur le fonctionnement
    lvlWarning = 2  ' Avertissements qui n'empêchent pas le fonctionnement
    lvlError = 3    ' Erreurs qui affectent le fonctionnement
    lvlFatal = 4    ' Erreurs critiques qui nécessitent l'arrêt
End Enum

Public Enum LogDestination
    destConsole = 1         ' Sortie dans la fenêtre immédiate
    destFile = 2            ' Sortie dans un fichier
    destSheet = 4           ' Sortie dans une feuille Excel
    destDatabase = 8        ' Sortie dans une base de données
    destCustom = 16         ' Sortie personnalisée via un callback
    destEventLog = 32       ' Sortie dans le journal d'événements Windows
    destAll = 63           ' Toutes les destinations
End Enum

' --- Types personnalisés ---
Private Type LoggerConfig
    Enabled As Boolean              ' Activation du logger
    MinLevel As LogLevel            ' Niveau minimum à journaliser
    Destinations As Long            ' Destinations de journalisation (combinaison de LogDestination)
    LogFilePath As String           ' Chemin du fichier de log
    LogFileName As String           ' Nom du fichier de log
    LogFileMaxSize As Long          ' Taille maximale du fichier de log en octets
    IncludeTimestamp As Boolean     ' Inclure l'horodatage
    IncludeLevel As Boolean         ' Inclure le niveau de log
    IncludeSource As Boolean        ' Inclure la source du message
    IncludeLineNumber As Boolean    ' Inclure le numéro de ligne (si disponible)
    AutoFlush As Boolean            ' Vider le buffer après chaque message
    MaxBufferSize As Long           ' Taille maximum du buffer (en nombre de messages)
    WorksheetName As String         ' Nom de la feuille pour la journalisation dans Excel
    WorkbookName As String          ' Nom du classeur pour la journalisation dans Excel
    ConnectionString As String      ' Chaîne de connexion pour la journalisation en base de données
    RotationEnabled As Boolean      ' Activer la rotation des logs
    RotationInterval As Long        ' Intervalle de rotation (en jours)
    AsyncEnabled As Boolean         ' Journalisation asynchrone
End Type

' --- Variables privées ---
Private m_Config As LoggerConfig
Private m_Buffer() As String
Private m_BufferCount As Long
Private m_FileNum As Integer
Private m_LastError As String
Private m_IsInitialized As Boolean
Private m_LogLevelNames(0 To 4) As String
Private m_CustomCallback As Object

' --- Initialisation ---
Private Sub Class_Initialize()
    ' Initialisation par défaut
    InitializeLogLevelNames
    
    With m_Config
        .Enabled = True
        .MinLevel = lvlInfo
        .Destinations = destConsole
        .LogFilePath = "logs\"
        .LogFileName = "application_log.txt"
        .LogFileMaxSize = 5242880 ' 5 Mo
        .IncludeTimestamp = True
        .IncludeLevel = True
        .IncludeSource = True
        .IncludeLineNumber = False
        .AutoFlush = True
        .MaxBufferSize = 100
        .WorksheetName = "Logs"
        .WorkbookName = ""
        .ConnectionString = ""
        .RotationEnabled = True
        .RotationInterval = 7 ' Rotation hebdomadaire
        .AsyncEnabled = False
    End With
    
    ' Initialiser le buffer
    ReDim m_Buffer(1 To m_Config.MaxBufferSize)
    m_BufferCount = 0
    m_FileNum = 0
    m_LastError = ""
    
    ' Marquer comme initialisé
    m_IsInitialized = True
End Sub

Private Sub Class_Terminate()
    ' Nettoyage à la destruction
    If m_BufferCount > 0 Then
        ' Vider le buffer avant de terminer
        FlushBuffer
    End If
    
    ' Fermer le fichier s'il est ouvert
    If m_FileNum <> 0 Then
        Close #m_FileNum
        m_FileNum = 0
    End If
End Sub

' --- Propriétés ---
Public Property Get Enabled() As Boolean
    Enabled = m_Config.Enabled
End Property

Public Property Let Enabled(ByVal value As Boolean)
    m_Config.Enabled = value
End Property

Public Property Get MinLevel() As LogLevel
    MinLevel = m_Config.MinLevel
End Property

Public Property Let MinLevel(ByVal value As LogLevel)
    m_Config.MinLevel = value
End Property

Public Property Get Destinations() As Long
    Destinations = m_Config.Destinations
End Property

Public Property Let Destinations(ByVal value As Long)
    m_Config.Destinations = value
End Property

Public Property Get LogFilePath() As String
    LogFilePath = m_Config.LogFilePath
End Property

Public Property Let LogFilePath(ByVal value As String)
    m_Config.LogFilePath = value
End Property

Public Property Get LogFileName() As String
    LogFileName = m_Config.LogFileName
End Property

Public Property Let LogFileName(ByVal value As String)
    m_Config.LogFileName = value
End Property

Public Property Get AutoFlush() As Boolean
    AutoFlush = m_Config.AutoFlush
End Property

Public Property Let AutoFlush(ByVal value As Boolean)
    m_Config.AutoFlush = value
End Property

Public Property Get LastError() As String
    LastError = m_LastError
End Property

Public Property Set CustomCallback(ByVal value As Object)
    Set m_CustomCallback = value
End Property

' --- Méthodes publiques ---
Public Sub LogDebug(ByVal message As String, Optional ByVal source As String = "", Optional ByVal lineNumber As Long = 0)
    ' Journalise un message de débogage
    LogMessage message, lvlDebug, source, lineNumber
End Sub

Public Sub LogInfo(ByVal message As String, Optional ByVal source As String = "", Optional ByVal lineNumber As Long = 0)
    ' Journalise une information
    LogMessage message, lvlInfo, source, lineNumber
End Sub

Public Sub LogWarning(ByVal message As String, Optional ByVal source As String = "", Optional ByVal lineNumber As Long = 0)
    ' Journalise un avertissement
    LogMessage message, lvlWarning, source, lineNumber
End Sub

Public Sub LogError(ByVal message As String, Optional ByVal source As String = "", Optional ByVal lineNumber As Long = 0)
    ' Journalise une erreur
    LogMessage message, lvlError, source, lineNumber
End Sub

Public Sub LogFatal(ByVal message As String, Optional ByVal source As String = "", Optional ByVal lineNumber As Long = 0)
    ' Journalise une erreur fatale
    LogMessage message, lvlFatal, source, lineNumber
End Sub

Public Sub LogMessage(ByVal message As String, ByVal level As LogLevel, Optional ByVal source As String = "", Optional ByVal lineNumber As Long = 0)
    ' Journalise un message avec un niveau spécifié
    
    ' Vérifier si la journalisation est activée et si le niveau est suffisant
    If Not m_Config.Enabled Or level < m_Config.MinLevel Then
        Exit Sub
    End If
    
    ' Formater le message
    Dim formattedMessage As String
    formattedMessage = FormatLogMessage(message, level, source, lineNumber)
    
    ' Ajouter au buffer
    m_BufferCount = m_BufferCount + 1
    m_Buffer(m_BufferCount) = formattedMessage
    
    ' Vider le buffer si nécessaire
    If m_Config.AutoFlush Or m_BufferCount >= m_Config.MaxBufferSize Then
        FlushBuffer
    End If
End Sub

Public Sub FlushBuffer()
    ' Vide le buffer de journalisation vers les destinations configurées
    Dim i As Long
    
    ' Pas besoin de continuer si le buffer est vide
    If m_BufferCount = 0 Then Exit Sub
    
    ' Log vers les destinations configurées
    For i = 1 To m_BufferCount
        ' Console
        If (m_Config.Destinations And destConsole) = destConsole Then
            LogToConsole m_Buffer(i)
        End If
        
        ' Fichier
        If (m_Config.Destinations And destFile) = destFile Then
            LogToFile m_Buffer(i)
        End If
        
        ' Feuille Excel
        If (m_Config.Destinations And destSheet) = destSheet Then
            LogToSheet m_Buffer(i)
        End If
        
        ' Base de données
        If (m_Config.Destinations And destDatabase) = destDatabase Then
            LogToDatabase m_Buffer(i)
        End If
        
        ' Callback personnalisé
        If (m_Config.Destinations And destCustom) = destCustom Then
            LogToCustom m_Buffer(i)
        End If
        
        ' Journal d'événements Windows
        If (m_Config.Destinations And destEventLog) = destEventLog Then
            LogToEventLog m_Buffer(i)
        End If
    Next i
    
    ' Réinitialiser le buffer
    m_BufferCount = 0
End Sub

Public Sub ClearBuffer()
    ' Vide le buffer sans écrire
    m_BufferCount = 0
End Sub

Public Function LoadConfigFromFile(ByVal filePath As String) As Boolean
    ' Charge la configuration depuis un fichier INI
    On Error GoTo ErrorHandler
    
    ' TODO: Implémenter le chargement de configuration
    LoadConfigFromFile = True
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors du chargement de la configuration: " & Err.Description
    LoadConfigFromFile = False
End Function

Public Function SaveConfigToFile(ByVal filePath As String) As Boolean
    ' Sauvegarde la configuration dans un fichier INI
    On Error GoTo ErrorHandler
    
    ' TODO: Implémenter la sauvegarde de configuration
    SaveConfigToFile = True
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la sauvegarde de la configuration: " & Err.Description
    SaveConfigToFile = False
End Function

' --- Méthodes privées ---
Private Sub InitializeLogLevelNames()
    ' Initialise les noms des niveaux de log
    m_LogLevelNames(lvlDebug) = "DEBUG"
    m_LogLevelNames(lvlInfo) = "INFO"
    m_LogLevelNames(lvlWarning) = "WARNING"
    m_LogLevelNames(lvlError) = "ERROR"
    m_LogLevelNames(lvlFatal) = "FATAL"
End Sub

Private Function FormatLogMessage(ByVal message As String, ByVal level As LogLevel, ByVal source As String, ByVal lineNumber As Long) As String
    ' Formate un message de log selon la configuration
    Dim result As String
    
    ' Timestamp
    If m_Config.IncludeTimestamp Then
        result = Format(Now, "yyyy-mm-dd hh:nn:ss") & " | "
    End If
    
    ' Niveau
    If m_Config.IncludeLevel Then
        result = result & m_LogLevelNames(level) & " | "
    End If
    
    ' Source
    If m_Config.IncludeSource And source <> "" Then
        result = result & source & " | "
    End If
    
    ' Numéro de ligne
    If m_Config.IncludeLineNumber And lineNumber <> 0 Then
        result = result & "Line " & lineNumber & " | "
    End If
    
    ' Message
    result = result & message
    
    FormatLogMessage = result
End Function

Private Sub LogToConsole(ByVal message As String)
    ' Écrit un message dans la fenêtre immédiate
    On Error Resume Next
    Debug.Print message
End Sub

Private Sub LogToFile(ByVal message As String)
    ' Écrit un message dans un fichier
    On Error GoTo ErrorHandler
    
    Dim filePath As String
    
    ' Créer le dossier des logs si nécessaire
    If m_Config.LogFilePath <> "" Then
        On Error Resume Next
        MkDir m_Config.LogFilePath
        On Error GoTo ErrorHandler
    End If
    
    ' Construire le chemin complet
    filePath = m_Config.LogFilePath & m_Config.LogFileName
    
    ' Vérifier la rotation des logs
    If m_Config.RotationEnabled Then
        CheckRotation filePath
    End If
    
    ' Vérifier si le fichier est déjà ouvert
    If m_FileNum = 0 Then
        m_FileNum = FreeFile
        Open filePath For Append As #m_FileNum
    End If
    
    ' Écrire le message
    Print #m_FileNum, message
    
    Exit Sub
    
ErrorHandler:
    m_LastError = "Erreur d'écriture dans le fichier de log: " & Err.Description
End Sub

Private Sub LogToSheet(ByVal message As String)
    ' Écrit un message dans une feuille Excel
    On Error GoTo ErrorHandler
    
    ' TODO: Implémenter la journalisation dans une feuille Excel
    
    Exit Sub
    
ErrorHandler:
    m_LastError = "Erreur d'écriture dans la feuille de log: " & Err.Description
End Sub

Private Sub LogToDatabase(ByVal message As String)
    ' Écrit un message dans une base de données
    On Error GoTo ErrorHandler
    
    ' TODO: Implémenter la journalisation dans une base de données
    
    Exit Sub
    
ErrorHandler:
    m_LastError = "Erreur d'écriture dans la base de données: " & Err.Description
End Sub

Private Sub LogToCustom(ByVal message As String)
    ' Envoie le message à un gestionnaire personnalisé
    On Error GoTo ErrorHandler
    
    If Not m_CustomCallback Is Nothing Then
        ' TODO: Implémenter le callback
        ' Supposons que le callback a une méthode ProcessLogMessage
        ' m_CustomCallback.ProcessLogMessage message
    End If
    
    Exit Sub
    
ErrorHandler:
    m_LastError = "Erreur lors de l'appel du gestionnaire personnalisé: " & Err.Description
End Sub

Private Sub LogToEventLog(ByVal message As String)
    ' Écrit un message dans le journal d'événements Windows
    On Error GoTo ErrorHandler
    
    ' TODO: Implémenter la journalisation dans le journal d'événements Windows
    
    Exit Sub
    
ErrorHandler:
    m_LastError = "Erreur d'écriture dans le journal d'événements: " & Err.Description
End Sub

Private Sub CheckRotation(ByVal filePath As String)
    ' Vérifie s'il faut faire une rotation des logs
    On Error GoTo ErrorHandler
    
    If Not m_Config.RotationEnabled Then Exit Sub
    
    ' Vérifier si le fichier existe
    If Dir(filePath) = "" Then Exit Sub
    
    ' Vérifier la taille du fichier
    If FileLen(filePath) >= m_Config.LogFileMaxSize Then
        RotateLogFile filePath
    End If
    
    Exit Sub
    
ErrorHandler:
    m_LastError = "Erreur lors de la vérification de rotation: " & Err.Description
End Sub

Private Sub RotateLogFile(ByVal filePath As String)
    ' Effectue la rotation du fichier de log
    On Error GoTo ErrorHandler
    
    ' Fermer le fichier s'il est ouvert
    If m_FileNum <> 0 Then
        Close #m_FileNum
        m_FileNum = 0
    End If
    
    ' Renommer le fichier avec un timestamp
    Dim newPath As String
    newPath = Left$(filePath, InStrRev(filePath, ".") - 1) & "_" & _
              Format(Now, "yyyymmdd_hhnnss") & ".txt"
    
    Name filePath As newPath
    
    Exit Sub
    
ErrorHandler:
    m_LastError = "Erreur lors de la rotation du fichier: " & Err.Description
End Sub 