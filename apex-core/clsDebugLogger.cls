' Classe: clsDebugLogger
' Date: 12/04/2025
' Description: Logger qui redirige vers Debug.Print
' Fait partie de l'architecture d'interopérabilité Apex-Excel

Option Explicit

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team


Implements ILoggerBase

' Variables privées
Private m_MinLogLevel As LogLevelEnum

' Initialisation
Private Sub Class_Initialize()
    m_MinLogLevel = LogLevelInfo
End Sub

' Implémentation de l'interface ILoggerBase
'@Description: 
'@Param: 
'@Returns: 

Private Sub ILoggerBase_Initialize(Optional ByVal minLevel As LogLevelEnum = LogLevelInfo, _
    Optional ByVal logSheetName As String = "Logs", _
    Optional ByVal logFileNamePattern As String = "{WorkbookName}\_{Date}.log", _
    Optional ByVal maxLogFileSizeKB As Long = 5120, _
    Optional ByVal targetWorkbook As Workbook = Nothing, _
    Optional ByVal enabledCategories As String = "\*", _
    Optional ByVal disabledCategories As String = "", _
    Optional ByVal bufferSize As Long = 1, _
    Optional ByVal crashLogBufferSize As Long = 10)
    
    m_MinLogLevel = minLevel
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ILoggerBase_SetLogger(ByVal loggerInstance As ILoggerBase)
    ' NOOP pour cette implémentation
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ILoggerBase_LogMessage(ByVal msg As String, _
    Optional ByVal level As LogLevelEnum = LogLevelInfo, _
    Optional ByVal category As String = "", _
    Optional ByVal source As String = "", _
    Optional ByVal user As String = "", _
    Optional ByVal toConsole As Boolean = True, _
    Optional ByVal toSheet As Boolean = False, _
    Optional ByVal toFile As Boolean = True)
    
    ' Vérifier si le niveau est suffisant
    If level < m_MinLogLevel Then Exit Sub
    
    ' Formater le message
    Dim formattedMsg As String
    formattedMsg = Format(Now, "yyyy-mm-dd hh:nn:ss") & " [" & GetLevelName(level) & "]"
    
    If category <> "" Then
        formattedMsg = formattedMsg & " [" & category & "]"
    End If
    
    If source <> "" Then
        formattedMsg = formattedMsg & " [" & source & "]"
    End If
    
    formattedMsg = formattedMsg & " " & msg
    
    ' Toujours rediriger vers Debug.Print
    Debug.Print formattedMsg
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ILoggerBase_LogConsole(ByVal msg As String, _
    Optional ByVal level As LogLevelEnum = LogLevelInfo, _
    Optional ByVal category As String = "", _
    Optional ByVal source As String = "", _
    Optional ByVal user As String = "")
    
    ILoggerBase_LogMessage msg, level, category, source, user, True, False, False
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ILoggerBase_LogSheet(ByVal msg As String, _
    Optional ByVal level As LogLevelEnum = LogLevelInfo, _
    Optional ByVal category As String = "", _
    Optional ByVal source As String = "", _
    Optional ByVal user As String = "")
    
    ' Cette implémentation ne gère pas la journalisation dans une feuille
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ILoggerBase_LogFile(ByVal msg As String, _
    Optional ByVal level As LogLevelEnum = LogLevelInfo, _
    Optional ByVal category As String = "", _
    Optional ByVal source As String = "", _
    Optional ByVal user As String = "")
    
    ' Cette implémentation ne gère pas la journalisation dans un fichier
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ILoggerBase_LogError(ByVal errObject As ErrObject, _
    Optional ByVal level As LogLevelEnum = LogLevelError, _
    Optional ByVal sourceRoutine As String = "", _
    Optional ByVal category As String = "ERROR", _
    Optional ByVal user As String = "", _
    Optional ByVal toConsole As Boolean = True, _
    Optional ByVal toSheet As Boolean = True, _
    Optional ByVal toFile As Boolean = True)
    
    ' Construire un message d'erreur formaté
    Dim errorMsg As String
    errorMsg = "Error #" & errObject.Number & ": " & errObject.Description
    
    If sourceRoutine <> "" Then
        errorMsg = errorMsg & " in " & sourceRoutine
    End If
    
    ILoggerBase_LogMessage errorMsg, level, category, sourceRoutine, user, toConsole, False, False
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ILoggerBase_FlushLogs()
    ' Rien à vider - Debug.Print est immédiat
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ILoggerBase_GenerateCrashReport(Optional ByVal crashFilePath As String = "")
    ' Non implémenté pour cette classe
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Property Get ILoggerBase_MinLogLevel() As LogLevelEnum
    ILoggerBase_MinLogLevel = m_MinLogLevel
End Property

Private Property Get ILoggerBase_IsInitialized() As Boolean
    ILoggerBase_IsInitialized = True
End Property

' Méthodes d'assistance
'@Description: 
'@Param: 
'@Returns: 

Private Function GetLevelName(ByVal level As LogLevelEnum) As String
    Select Case level
        Case LogLevelDebug: GetLevelName = "DEBUG"
        Case LogLevelInfo: GetLevelName = "INFO"
        Case LogLevelWarning: GetLevelName = "WARNING"
        Case LogLevelError: GetLevelName = "ERROR"
        Case LogLevelFatal: GetLevelName = "FATAL"
        Case Else: GetLevelName = "UNKNOWN"
    End Select
End Function