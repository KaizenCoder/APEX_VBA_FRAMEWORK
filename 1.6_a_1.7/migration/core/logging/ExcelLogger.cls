VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@ModuleDescription("Implémentation de logger pour Excel")
'@Folder("APEX.Core.Logging")
'@Version 1.7.0
'@Author APEX Framework Team
Option Explicit

Implements ILoggerBase

' =========================================================
' ExcelLogger
' Implémentation de logger pour feuille Excel
' =========================================================

' Configuration et état
Private m_name As String
Private m_logLevel As LogLevel
Private m_config As Object
Private m_isInitialized As Boolean
Private m_buffer As Collection
Private m_worksheet As Object  ' Feuille Excel (Worksheet)
Private m_maxRows As Long
Private m_currentRow As Long
Private m_formatCells As Boolean
Private m_lastCell As String

' Constantes
Private Const DEFAULT_MAX_ROWS As Long = 10000
Private Const DEFAULT_FORMAT_CELLS As Boolean = True
Private Const COL_TIMESTAMP As Long = 1
Private Const COL_LEVEL As Long = 2
Private Const COL_SOURCE As Long = 3
Private Const COL_MESSAGE As Long = 4
Private Const COL_DATA As Long = 5

' =========================================================
' Initialisation et configuration
' =========================================================
Private Sub Class_Initialize()
    m_logLevel = LogInfo ' Niveau par défaut
    m_isInitialized = False
    m_maxRows = DEFAULT_MAX_ROWS
    m_formatCells = DEFAULT_FORMAT_CELLS
    m_currentRow = 1 ' Première ligne
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
            
            ' Feuille Excel (obligatoire)
            Err.Clear
            Dim ws As Variant
            Set ws = config("Worksheet")
            If Err.Number = 0 Then
                Set m_worksheet = ws
            End If
            
            ' Nombre maximum de lignes
            Err.Clear
            Dim maxRows As Variant
            maxRows = config("MaxRows")
            If Err.Number = 0 Then
                m_maxRows = CLng(maxRows)
            End If
            
            ' Formater les cellules
            Err.Clear
            Dim formatCells As Variant
            formatCells = config("FormatCells")
            If Err.Number = 0 Then
                m_formatCells = CBool(formatCells)
            End If
        End If
        On Error GoTo ErrorHandler
    End If
    
    ' Générer un nom par défaut si non spécifié
    If Len(m_name) = 0 Then
        m_name = "ExcelLogger_" & Format(Now, "yyyymmddhhnnss")
    End If
    
    ' Vérifier que la feuille est spécifiée
    If m_worksheet Is Nothing Then
        Err.Raise vbObjectError + 1000, "ExcelLogger.Initialize", "Feuille Excel non spécifiée"
    End If
    
    ' Trouver la dernière ligne utilisée dans la feuille
    m_currentRow = GetFirstAvailableRow
    
    ' Créer les en-têtes si c'est une nouvelle feuille
    If m_currentRow = 1 Then
        CreateHeaderRow
        m_currentRow = 2 ' Commencer à la ligne 2
    End If
    
    ' Ajouter une ligne pour indiquer le démarrage d'une nouvelle session
    m_worksheet.Cells(m_currentRow, COL_TIMESTAMP).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    m_worksheet.Cells(m_currentRow, COL_LEVEL).Value = "INFO"
    m_worksheet.Cells(m_currentRow, COL_MESSAGE).Value = "--- Session de logging démarrée ---"
    
    ' Formater la ligne si nécessaire
    If m_formatCells Then
        FormatRow m_currentRow, LogInfo
    End If
    
    m_currentRow = m_currentRow + 1
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
    
    ' Vérifier si nous avons atteint le maximum de lignes
    If m_currentRow > m_maxRows Then
        ' Supprimer une partie des anciennes lignes
        RemoveOldRows
    End If
    
    On Error Resume Next
    
    ' Ajouter le message au buffer
    Dim logEntry As New Collection
    logEntry.Add Format(Now, "yyyy-mm-dd hh:nn:ss") ' Timestamp
    logEntry.Add level                             ' Niveau
    logEntry.Add source                            ' Source
    logEntry.Add message                           ' Message
    logEntry.Add extraData                         ' Données supplémentaires
    m_buffer.Add logEntry
    
    ' Écrire dans Excel
    m_worksheet.Cells(m_currentRow, COL_TIMESTAMP).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    m_worksheet.Cells(m_currentRow, COL_LEVEL).Value = GetLevelText(level)
    m_worksheet.Cells(m_currentRow, COL_SOURCE).Value = source
    m_worksheet.Cells(m_currentRow, COL_MESSAGE).Value = message
    
    ' Ajouter les données supplémentaires si fournies
    If Not IsNull(extraData) Then
        m_worksheet.Cells(m_currentRow, COL_DATA).Value = FormatExtraData(extraData)
    End If
    
    ' Formater la ligne si nécessaire
    If m_formatCells Then
        FormatRow m_currentRow, level
    End If
    
    ' Incrémenter le compteur de ligne
    m_currentRow = m_currentRow + 1
    m_lastCell = "E" & m_currentRow - 1
    
    ' Flush si nécessaire (niveaux d'erreur élevés)
    If level >= LogError Then
        ILoggerBase_Flush
    End If
    
    On Error GoTo 0
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
    ' Vider le buffer et s'assurer que tout est visible dans Excel
    Set m_buffer = New Collection
    
    ' S'assurer que les données sont visibles dans Excel
    On Error Resume Next
    Application.ScreenUpdating = True
    
    ' Défiler vers la dernière ligne
    If m_lastCell <> "" Then
        m_worksheet.Range(m_lastCell).Select
    End If
    
    On Error GoTo 0
End Sub

Private Function ILoggerBase_IsEnabled(ByVal level As LogLevel) As Boolean
    ILoggerBase_IsEnabled = (level >= m_logLevel)
End Function

' =========================================================
' Méthodes privées
' =========================================================
Private Function GetLevelText(ByVal level As LogLevel) As String
    ' Obtenir le texte du niveau de log
    Select Case level
        Case LogTrace: GetLevelText = "TRACE"
        Case LogDebug: GetLevelText = "DEBUG"
        Case LogInfo: GetLevelText = "INFO"
        Case LogWarning: GetLevelText = "WARNING"
        Case LogError: GetLevelText = "ERROR"
        Case LogFatal: GetLevelText = "FATAL"
        Case Else: GetLevelText = "UNKNOWN"
    End Select
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

Private Function GetFirstAvailableRow() As Long
    ' Trouver la première ligne disponible dans la feuille
    On Error Resume Next
    
    ' Chercher dans la colonne A (timestamp)
    Dim lastRow As Long
    lastRow = m_worksheet.Cells(m_worksheet.Rows.Count, COL_TIMESTAMP).End(xlUp).Row
    
    ' Si la feuille est vide, lastRow sera 1
    If lastRow = 1 And Len(m_worksheet.Cells(1, COL_TIMESTAMP).Value) = 0 Then
        GetFirstAvailableRow = 1
    Else
        GetFirstAvailableRow = lastRow + 1
    End If
    
    On Error GoTo 0
End Function

Private Sub CreateHeaderRow()
    ' Créer la ligne d'en-tête
    On Error Resume Next
    
    m_worksheet.Cells(1, COL_TIMESTAMP).Value = "Timestamp"
    m_worksheet.Cells(1, COL_LEVEL).Value = "Niveau"
    m_worksheet.Cells(1, COL_SOURCE).Value = "Source"
    m_worksheet.Cells(1, COL_MESSAGE).Value = "Message"
    m_worksheet.Cells(1, COL_DATA).Value = "Données"
    
    ' Formater l'en-tête
    If m_formatCells Then
        With m_worksheet.Range(m_worksheet.Cells(1, COL_TIMESTAMP), m_worksheet.Cells(1, COL_DATA))
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
        End With
        
        ' Ajuster les largeurs des colonnes
        m_worksheet.Columns(COL_TIMESTAMP).ColumnWidth = 20
        m_worksheet.Columns(COL_LEVEL).ColumnWidth = 10
        m_worksheet.Columns(COL_SOURCE).ColumnWidth = 20
        m_worksheet.Columns(COL_MESSAGE).ColumnWidth = 60
        m_worksheet.Columns(COL_DATA).ColumnWidth = 30
    End If
    
    On Error GoTo 0
End Sub

Private Sub FormatRow(ByVal row As Long, ByVal level As LogLevel)
    ' Formater une ligne selon le niveau de log
    On Error Resume Next
    
    Dim levelCell As Object
    Set levelCell = m_worksheet.Cells(row, COL_LEVEL)
    
    ' Réinitialiser le formatage
    levelCell.Interior.ColorIndex = xlNone
    
    ' Appliquer la couleur selon le niveau
    Select Case level
        Case LogTrace
            ' Pas de formatage spécial
        Case LogDebug
            levelCell.Interior.Color = RGB(240, 240, 240)
        Case LogInfo
            levelCell.Interior.Color = RGB(220, 240, 220)
        Case LogWarning
            levelCell.Interior.Color = RGB(255, 240, 200)
            levelCell.Font.Bold = True
        Case LogError
            levelCell.Interior.Color = RGB(255, 200, 200)
            levelCell.Font.Bold = True
        Case LogFatal
            levelCell.Interior.Color = RGB(255, 150, 150)
            levelCell.Font.Bold = True
            ' Formater toute la ligne
            m_worksheet.Range(m_worksheet.Cells(row, COL_TIMESTAMP), _
                             m_worksheet.Cells(row, COL_DATA)).Interior.Color = RGB(255, 200, 200)
    End Select
    
    On Error GoTo 0
End Sub

Private Sub RemoveOldRows()
    ' Supprimer les anciennes lignes lorsque la limite est atteinte
    On Error Resume Next
    
    ' Conserver les 3/4 des lignes récentes
    Dim keepRows As Long
    keepRows = m_maxRows * 3 / 4
    
    ' Calculer combien de lignes supprimer
    Dim removeRows As Long
    removeRows = m_currentRow - keepRows
    
    If removeRows > 0 Then
        ' Supprimer les lignes (mais pas la ligne d'en-tête)
        m_worksheet.Rows("2:" & (removeRows + 1)).Delete
        
        ' Ajuster le compteur de ligne
        m_currentRow = m_currentRow - removeRows
    End If
    
    On Error GoTo 0
End Sub 