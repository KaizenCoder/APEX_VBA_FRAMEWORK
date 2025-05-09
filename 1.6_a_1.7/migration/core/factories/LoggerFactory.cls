VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@ModuleDescription("Factory pour la cr�ation et gestion des loggers")
'@PredeclaredId
'@Folder("APEX.Core.Factories")
'@Version 1.7.0
'@Author APEX Framework Team
Option Explicit

Implements ILoggerFactoryBase

' =========================================================
' LoggerFactory
' Impl�mentation de la factory pour les loggers
' =========================================================

' Configuration par d�faut
Private Const DEFAULT_LOGGER_NAME As String = "DefaultLogger"
Private Const DEFAULT_LOG_LEVEL As LogLevel = LogInfo

' Collections priv�es
Private m_loggers As Object        ' Collection de loggers
Private m_defaultLoggerName As String
Private m_globalLogLevel As LogLevel

' =========================================================
' Initialisation et instanciation
' =========================================================
Private Sub Class_Initialize()
    ' Initialiser les collections
    Set m_loggers = CreateObject("Scripting.Dictionary")
    
    ' Configuration par d�faut
    m_defaultLoggerName = DEFAULT_LOGGER_NAME
    m_globalLogLevel = DEFAULT_LOG_LEVEL
    
    ' Cr�er le logger par d�faut
    Dim defaultLogger As ILoggerBase
    Set defaultLogger = New Logger ' Logger par d�faut, � impl�menter
    m_loggers.Add DEFAULT_LOGGER_NAME, defaultLogger
End Sub

' =========================================================
' M�thodes de factory
' =========================================================
Public Function CreateStandardLogger(ByVal name As String, Optional ByVal config As Object = Nothing) As ILoggerBase
    ' M�thode publique pour cr�er un logger standard
    Dim logger As New Logger ' � impl�menter
    logger.Initialize config
    Set CreateStandardLogger = logger
End Function

Public Function CreateFileLogger(ByVal name As String, ByVal filePath As String, Optional ByVal config As Object = Nothing) As ILoggerBase
    ' M�thode publique pour cr�er un logger de fichier
    Dim logger As New FileLogger ' � impl�menter
    If config Is Nothing Then
        ' Cr�er une configuration basique
        Set config = New Dictionary ' � impl�menter ou utiliser Scripting.Dictionary
        config("FilePath") = filePath
    End If
    logger.Initialize config
    Set CreateFileLogger = logger
End Function

Public Function CreateExcelLogger(ByVal name As String, ByVal ws As Object, Optional ByVal config As Object = Nothing) As ILoggerBase
    ' M�thode publique pour cr�er un logger Excel
    Dim logger As New ExcelLogger ' � impl�menter
    If config Is Nothing Then
        ' Cr�er une configuration basique
        Set config = New Dictionary ' � impl�menter ou utiliser Scripting.Dictionary
        Set config("Worksheet") = ws
    End If
    logger.Initialize config
    Set CreateExcelLogger = logger
End Function

' =========================================================
' Impl�mentation de ILoggerFactoryBase
' =========================================================

' Factory methods
Private Function ILoggerFactoryBase_CreateLogger(ByVal name As String, Optional ByVal config As Object = Nothing) As ILoggerBase
    ' V�rifier si le logger existe d�j�
    If m_loggers.Exists(name) Then
        ' Logger existe d�j�, retourner une erreur ou remplacer?
        ' Pour l'instant, on remplace
        m_loggers.Remove name
    End If
    
    ' Cr�er un nouveau logger standard par d�faut
    Dim logger As ILoggerBase
    Set logger = CreateStandardLogger(name, config)
    
    ' Ajouter � la collection
    m_loggers.Add name, logger
    
    ' Retourner le logger
    Set ILoggerFactoryBase_CreateLogger = logger
End Function

Private Function ILoggerFactoryBase_GetLogger(ByVal name As String) As ILoggerBase
    ' V�rifier si le logger existe
    If m_loggers.Exists(name) Then
        Set ILoggerFactoryBase_GetLogger = m_loggers(name)
    Else
        ' Cr�er un nouveau logger avec ce nom
        Set ILoggerFactoryBase_GetLogger = ILoggerFactoryBase_CreateLogger(name)
    End If
End Function

Private Function ILoggerFactoryBase_GetDefaultLogger() As ILoggerBase
    ' R�cup�rer le logger par d�faut
    Set ILoggerFactoryBase_GetDefaultLogger = m_loggers(m_defaultLoggerName)
End Function

Private Function ILoggerFactoryBase_SetDefaultLogger(ByVal name As String) As Boolean
    ' V�rifier si le logger existe
    If m_loggers.Exists(name) Then
        m_defaultLoggerName = name
        ILoggerFactoryBase_SetDefaultLogger = True
    Else
        ILoggerFactoryBase_SetDefaultLogger = False
    End If
End Function

' Configuration
Private Function ILoggerFactoryBase_Configure(ByVal config As Object) As Boolean
    ' Configurer la factory (� impl�menter selon le format de configuration)
    On Error GoTo ErrorHandler
    
    ' Exemple de configuration
    ' Configurer le niveau de log global si sp�cifi�
    If Not config Is Nothing Then
        If TypeOf config Is Dictionary Or TypeOf config Is Object Then
            ' V�rifier les propri�t�s disponibles
            On Error Resume Next
            Dim logLevel As Variant
            logLevel = config("LogLevel")
            If Err.Number = 0 Then
                m_globalLogLevel = logLevel
            End If
            On Error GoTo ErrorHandler
            
            ' Autres configurations...
        End If
    End If
    
    ILoggerFactoryBase_Configure = True
    Exit Function
    
ErrorHandler:
    ILoggerFactoryBase_Configure = False
End Function

Private Function ILoggerFactoryBase_SetGlobalLogLevel(ByVal level As LogLevel) As Boolean
    ' D�finir le niveau de log global
    m_globalLogLevel = level
    
    ' Appliquer � tous les loggers existants
    Dim key As Variant
    For Each key In m_loggers.Keys
        Dim logger As ILoggerBase
        Set logger = m_loggers(key)
        logger.SetLogLevel level
    Next key
    
    ILoggerFactoryBase_SetGlobalLogLevel = True
End Function

' Logger management
Private Function ILoggerFactoryBase_GetLoggerNames() As Variant
    ' R�cup�rer les noms des loggers
    ILoggerFactoryBase_GetLoggerNames = m_loggers.Keys
End Function

Private Sub ILoggerFactoryBase_FlushAll()
    ' Vider tous les buffers
    Dim key As Variant
    For Each key In m_loggers.Keys
        Dim logger As ILoggerBase
        Set logger = m_loggers(key)
        logger.Flush
    Next key
End Sub

Private Function ILoggerFactoryBase_RemoveLogger(ByVal name As String) As Boolean
    ' Supprimer un logger
    If m_loggers.Exists(name) Then
        ' Ne pas supprimer le logger par d�faut
        If name = m_defaultLoggerName Then
            ILoggerFactoryBase_RemoveLogger = False
            Exit Function
        End If
        
        m_loggers.Remove name
        ILoggerFactoryBase_RemoveLogger = True
    Else
        ILoggerFactoryBase_RemoveLogger = False
    End If
End Function 