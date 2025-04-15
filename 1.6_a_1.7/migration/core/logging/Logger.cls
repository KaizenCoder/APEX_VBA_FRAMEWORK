VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@ModuleDescription("Implémentation standard du logger")
'@Folder("APEX.Core.Logging")
'@Version 1.7.0
'@Author APEX Framework Team
Option Explicit

Implements ILoggerBase

' =========================================================
' Logger
' Implémentation standard du logger pour le framework APEX
' =========================================================

' Configuration et état
Private m_name As String
Private m_logLevel As LogLevel
Private m_config As Object
Private m_isInitialized As Boolean
Private m_buffer As Collection

' =========================================================
' Initialisation et configuration
' =========================================================
Private Sub Class_Initialize()
    m_logLevel = LogInfo ' Niveau par défaut
    m_isInitialized = False
    Set m_buffer = New Collection
End Sub

' =========================================================
' Implémentation de ILoggerBase
' =========================================================

' Configuration
Private Function ILoggerBase_Initialize(Optional ByVal config As Object = Nothing) As Boolean
    On Error GoTo ErrorHandler
    
    ' Stocker la configuration
    Set m_config = config
    
    ' Configurer le logger si config fournie
    If Not config Is Nothing Then
        ' Extraire les paramètres de config
        On Error Resume Next
        
        ' Nom du logger
        If IsObject(config) Then
            Dim loggerName As Variant
            loggerName = config("Name")
            If Err.Number = 0 Then
                m_name = CStr(loggerName)
            End If
            
            ' Niveau de log
            Err.Clear
            Dim logLevel As Variant
            logLevel = config("LogLevel")
            If Err.Number = 0 Then
                m_logLevel = logLevel
            End If
        End If
        On Error GoTo ErrorHandler
    End If
    
    ' Générer un nom par défaut si non spécifié
    If Len(m_name) = 0 Then
        m_name = "Logger_" & Format(Now, "yyyymmddhhnnss")
    End If
    
    m_isInitialized = True
    ILoggerBase_Initialize = True
    Exit Function
    
ErrorHandler:
    ILoggerBase_Initialize = False
End Function

Private Function ILoggerBase_SetLogLevel(ByVal level As LogLevel) As Boolean
    m_logLevel = level
    ILoggerBase_SetLogLevel = True
End Function

Private Function ILoggerBase_GetLogLevel() As LogLevel
    ILoggerBase_GetLogLevel = m_logLevel
End Function

Private Function ILoggerBase_GetName() As String
    ILoggerBase_GetName = m_name
End Function

' Logging methods
Private Sub ILoggerBase_LogMessage(ByVal level As LogLevel, ByVal message As String, Optional ByVal source As String = "", Optional ByVal extraData As Variant = Null)
    ' Vérifier si ce niveau est activé
    If Not ILoggerBase_IsEnabled(level) Then Exit Sub
    
    ' Formater le message
    Dim formattedMsg As String
    formattedMsg = FormatLogMessage(level, message, source, extraData)
    
    ' Ajouter au buffer
    m_buffer.Add formattedMsg
    
    ' Le logger standard écrit dans Debug.Print par défaut
    Debug.Print formattedMsg
    
    ' Vider le buffer automatiquement à certaines conditions
    If level >= LogError Or m_buffer.Count >= 100 Then
        ILoggerBase_Flush
    End If
End Sub

Private Sub ILoggerBase_Trace(ByVal message As String, Optional ByVal source As String = "", Optional ByVal extraData As Variant = Null)
    ILoggerBase_LogMessage LogTrace, message, source, extraData
End Sub

Private Sub ILoggerBase_Debug(ByVal message As String, Optional ByVal source As String = "", Optional ByVal extraData As Variant = Null)
    ILoggerBase_LogMessage LogDebug, message, source, extraData
End Sub

Private Sub ILoggerBase_Info(ByVal message As String, Optional ByVal source As String = "", Optional ByVal extraData As Variant = Null)
    ILoggerBase_LogMessage LogInfo, message, source, extraData
End Sub

Private Sub ILoggerBase_Warning(ByVal message As String, Optional ByVal source As String = "", Optional ByVal extraData As Variant = Null)
    ILoggerBase_LogMessage LogWarning, message, source, extraData
End Sub

Private Sub ILoggerBase_Error(ByVal message As String, Optional ByVal source As String = "", Optional ByVal extraData As Variant = Null)
    ILoggerBase_LogMessage LogError, message, source, extraData
End Sub

Private Sub ILoggerBase_Fatal(ByVal message As String, Optional ByVal source As String = "", Optional ByVal extraData As Variant = Null)
    ILoggerBase_LogMessage LogFatal, message, source, extraData
End Sub

Private Sub ILoggerBase_LogException(ByVal ex As Object, Optional ByVal level As LogLevel = LogError, Optional ByVal additionalInfo As String = "")
    If Not ILoggerBase_IsEnabled(level) Then Exit Sub
    
    ' Extraire les informations de l'exception
    Dim errMsg As String
    Dim errSrc As String
    Dim errNum As Long
    
    On Error Resume Next
    errMsg = ex.Description
    errSrc = ex.Source
    errNum = ex.Number
    
    ' Formater le message d'exception
    Dim exMsg As String
    exMsg = "Exception #" & errNum & ": " & errMsg
    If Len(additionalInfo) > 0 Then
        exMsg = exMsg & " | " & additionalInfo
    End If
    
    ' Logger l'exception
    ILoggerBase_LogMessage level, exMsg, errSrc
End Sub

' Buffer management
Private Sub ILoggerBase_Flush()
    ' La classe de base écrit uniquement en Debug.Print
    ' Cette méthode vide simplement le buffer
    Set m_buffer = New Collection
End Sub

Private Function ILoggerBase_IsEnabled(ByVal level As LogLevel) As Boolean
    ILoggerBase_IsEnabled = (level >= m_logLevel)
End Function

' =========================================================
' Méthodes privées
' =========================================================
Private Function FormatLogMessage(ByVal level As LogLevel, ByVal message As String, ByVal source As String, ByVal extraData As Variant) As String
    ' Obtenir le timestamp actuel
    Dim timestamp As String
    timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    ' Obtenir le texte du niveau de log
    Dim levelText As String
    Select Case level
        Case LogTrace: levelText = "TRACE"
        Case LogDebug: levelText = "DEBUG"
        Case LogInfo: levelText = "INFO "
        Case LogWarning: levelText = "WARN "
        Case LogError: levelText = "ERROR"
        Case LogFatal: levelText = "FATAL"
        Case Else: levelText = "?????"
    End Select
    
    ' Formater le message de base
    Dim formattedMsg As String
    formattedMsg = "[" & timestamp & "][" & levelText & "] "
    
    ' Ajouter la source si fournie
    If Len(source) > 0 Then
        formattedMsg = formattedMsg & "[" & source & "] "
    End If
    
    ' Ajouter le message principal
    formattedMsg = formattedMsg & message
    
    ' Ajouter les données supplémentaires si fournies
    If Not IsNull(extraData) Then
        formattedMsg = formattedMsg & " | Data: " & FormatExtraData(extraData)
    End If
    
    FormatLogMessage = formattedMsg
End Function

Private Function FormatExtraData(ByVal data As Variant) As String
    ' Formater les données supplémentaires selon leur type
    On Error Resume Next
    
    If IsEmpty(data) Then
        FormatExtraData = "<empty>"
    ElseIf IsNull(data) Then
        FormatExtraData = "<null>"
    ElseIf IsArray(data) Then
        FormatExtraData = "<array>"
    ElseIf IsObject(data) Then
        If TypeName(data) = "Dictionary" Then
            FormatExtraData = "<dictionary>"
        Else
            FormatExtraData = "<" & TypeName(data) & ">"
        End If
    Else
        FormatExtraData = CStr(data)
    End If
    
    On Error GoTo 0
End Function 