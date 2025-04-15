VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@ModuleDescription("Implémentation de logger pour fichiers")
'@Folder("APEX.Core.Logging")
'@Version 1.7.0
'@Author APEX Framework Team
Option Explicit

Implements ILoggerBase

' =========================================================
' FileLogger
' Implémentation de logger pour fichier
' =========================================================

' Configuration et état
Private m_name As String
Private m_logLevel As LogLevel
Private m_config As Object
Private m_isInitialized As Boolean
Private m_buffer As Collection
Private m_filePath As String
Private m_fileNum As Integer
Private m_isFileOpen As Boolean
Private m_autoFlush As Boolean
Private m_maxFileSize As Long
Private m_rotateFiles As Boolean

' Constantes
Private Const DEFAULT_MAX_FILE_SIZE As Long = 5242880 ' 5 MB par défaut
Private Const DEFAULT_AUTO_FLUSH As Boolean = True

' =========================================================
' Initialisation et configuration
' =========================================================
Private Sub Class_Initialize()
    m_logLevel = LogInfo ' Niveau par défaut
    m_isInitialized = False
    m_isFileOpen = False
    m_autoFlush = DEFAULT_AUTO_FLUSH
    m_maxFileSize = DEFAULT_MAX_FILE_SIZE
    m_rotateFiles = True
    m_fileNum = FreeFile ' Obtenir un numéro de fichier disponible
    Set m_buffer = New Collection
End Sub

Private Sub Class_Terminate()
    ' Fermer le fichier si nécessaire
    If m_isFileOpen Then
        Close m_fileNum
        m_isFileOpen = False
    End If
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
            
            ' Chemin du fichier (obligatoire)
            Err.Clear
            Dim filePath As Variant
            filePath = config("FilePath")
            If Err.Number = 0 Then
                m_filePath = CStr(filePath)
            End If
            
            ' Auto Flush
            Err.Clear
            Dim autoFlush As Variant
            autoFlush = config("AutoFlush")
            If Err.Number = 0 Then
                m_autoFlush = CBool(autoFlush)
            End If
            
            ' Taille maximale du fichier
            Err.Clear
            Dim maxFileSize As Variant
            maxFileSize = config("MaxFileSize")
            If Err.Number = 0 Then
                m_maxFileSize = CLng(maxFileSize)
            End If
            
            ' Rotation des fichiers
            Err.Clear
            Dim rotateFiles As Variant
            rotateFiles = config("RotateFiles")
            If Err.Number = 0 Then
                m_rotateFiles = CBool(rotateFiles)
            End If
        End If
        On Error GoTo ErrorHandler
    End If
    
    ' Générer un nom par défaut si non spécifié
    If Len(m_name) = 0 Then
        m_name = "FileLogger_" & Format(Now, "yyyymmddhhnnss")
    End If
    
    ' Vérifier que le chemin du fichier est spécifié
    If Len(m_filePath) = 0 Then
        Err.Raise vbObjectError + 1000, "FileLogger.Initialize", "Chemin de fichier non spécifié"
    End If
    
    ' Créer le fichier de log s'il n'existe pas
    If Not FileExists(m_filePath) Then
        ' Créer le répertoire parent si nécessaire
        CreateParentDirectory m_filePath
        
        ' Créer le fichier vide
        Open m_filePath For Output As m_fileNum
        Close m_fileNum
    End If
    
    ' Vérifier si le fichier doit être rotaté
    CheckFileRotation
    
    ' Ouvrir le fichier pour ajouter des logs
    Open m_filePath For Append As m_fileNum
    m_isFileOpen = True
    
    ' Écrire un en-tête de log
    Print #m_fileNum, "--- Session de logging démarrée le " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ---"
    
    m_isInitialized = True
    ILoggerBase_Initialize = True
    Exit Function
    
ErrorHandler:
    If m_isFileOpen Then
        Close m_fileNum
        m_isFileOpen = False
    End If
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
    
    ' Vérifier que le fichier est ouvert
    If Not m_isFileOpen Then
        On Error Resume Next
        Open m_filePath For Append As m_fileNum
        If Err.Number <> 0 Then Exit Sub ' Impossible d'ouvrir le fichier
        m_isFileOpen = True
        On Error GoTo 0
    End If
    
    ' Formater le message
    Dim formattedMsg As String
    formattedMsg = FormatLogMessage(level, message, source, extraData)
    
    ' Ajouter au buffer
    m_buffer.Add formattedMsg
    
    ' Écrire dans le fichier
    On Error Resume Next
    Print #m_fileNum, formattedMsg
    
    ' Vérifier rotation du fichier
    CheckFileRotation
    
    ' Flush si nécessaire
    If m_autoFlush Or level >= LogError Or m_buffer.Count >= 100 Then
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
    ' Vider le buffer et s'assurer que tout est écrit dans le fichier
    Set m_buffer = New Collection
    
    ' Forcer l'écriture dans le fichier
    If m_isFileOpen Then
        Close m_fileNum
        Open m_filePath For Append As m_fileNum
    End If
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

Private Function FileExists(ByVal filePath As String) As Boolean
    ' Vérifier si un fichier existe
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

Private Sub CreateParentDirectory(ByVal filePath As String)
    ' Créer le répertoire parent si nécessaire
    Dim folderPath As String
    Dim lastSep As Long
    
    ' Trouver le dernier séparateur de répertoire
    lastSep = InStrRev(filePath, "\")
    If lastSep > 0 Then
        folderPath = Left(filePath, lastSep - 1)
        
        ' Créer le répertoire s'il n'existe pas
        On Error Resume Next
        If Len(Dir(folderPath, vbDirectory)) = 0 Then
            MkDir folderPath
        End If
        On Error GoTo 0
    End If
End Sub

Private Sub CheckFileRotation()
    ' Vérifier si le fichier doit être rotaté
    If Not m_rotateFiles Then Exit Sub
    
    On Error Resume Next
    
    ' Vérifier la taille du fichier
    Dim fileSize As Long
    fileSize = FileLen(m_filePath)
    
    ' Rotater si nécessaire
    If fileSize > m_maxFileSize Then
        ' Fermer le fichier actuel
        If m_isFileOpen Then
            Close m_fileNum
            m_isFileOpen = False
        End If
        
        ' Générer un nom de fichier pour la rotation
        Dim newFilePath As String
        newFilePath = Left(m_filePath, InStrRev(m_filePath, ".") - 1) & "_" & _
                      Format(Now, "yyyymmdd_hhnnss") & Mid(m_filePath, InStrRev(m_filePath, "."))
        
        ' Renommer le fichier actuel
        Name m_filePath As newFilePath
        
        ' Ouvrir un nouveau fichier
        Open m_filePath For Output As m_fileNum
        Print #m_fileNum, "--- Nouveau fichier créé après rotation le " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ---"
        m_isFileOpen = True
    End If
    
    On Error GoTo 0
End Sub 