VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSheetLogger"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe: clsSheetLogger
' Date: 12/04/2025
' Description: Logger qui écrit dans une feuille Excel avec rotation automatique
' Fait partie de l'architecture d'interopérabilité Apex-Excel

Option Explicit

Implements ILoggerBase

' Constantes
Private Const DEFAULT_LOG_SHEET_NAME As String = "Logs"
Private Const MAX_LOG_ROWS As Long = 1000

' Variables privées
Private m_MinLogLevel As LogLevelEnum
Private m_LogSheetName As String
Private m_TargetWorkbook As Workbook
Private m_CurrentRow As Long
Private m_Initialized As Boolean

' Initialisation
Private Sub Class_Initialize()
    m_MinLogLevel = LogLevelInfo
    m_LogSheetName = DEFAULT_LOG_SHEET_NAME
    m_CurrentRow = 2 ' 1ère ligne pour les en-têtes
    m_Initialized = False
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
    m_LogSheetName = logSheetName
    
    ' Si aucun classeur n'est spécifié, utiliser celui actif
    If targetWorkbook Is Nothing Then
        Set m_TargetWorkbook = ActiveWorkbook
    Else
        Set m_TargetWorkbook = targetWorkbook
    End If
    
    ' S'assurer que la feuille de journalisation existe
    If Not SheetExists(m_LogSheetName) Then
        CreateLogSheet
    End If
    
    m_Initialized = True
    
    ' Trouver la première ligne vide
    FindFirstEmptyRow
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ILoggerBase_SetLogger(ByVal loggerInstance As ILoggerBase)
    ' Non applicable pour cette implémentation
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
    
    ' Vérifier le niveau et l'initialisation
    If level < m_MinLogLevel Or Not m_Initialized Then Exit Sub
    
    ' Log to sheet if required
    If toSheet Then
        WriteToSheet msg, level, category, source, user
    End If
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ILoggerBase_LogConsole(ByVal msg As String, _
    Optional ByVal level As LogLevelEnum = LogLevelInfo, _
    Optional ByVal category As String = "", _
    Optional ByVal source As String = "", _
    Optional ByVal user As String = "")
    
    ' Non implémenté - cette classe ne gère que les feuilles
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ILoggerBase_LogSheet(ByVal msg As String, _
    Optional ByVal level As LogLevelEnum = LogLevelInfo, _
    Optional ByVal category As String = "", _
    Optional ByVal source As String = "", _
    Optional ByVal user As String = "")
    
    If Not m_Initialized Or level < m_MinLogLevel Then Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    WriteToSheet msg, level, category, source, user
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ILoggerBase_LogFile(ByVal msg As String, _
    Optional ByVal level As LogLevelEnum = LogLevelInfo, _
    Optional ByVal category As String = "", _
    Optional ByVal source As String = "", _
    Optional ByVal user As String = "")
    
    ' Non implémenté - cette classe ne gère que les feuilles
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
    
    If Not m_Initialized Or level < m_MinLogLevel Then Exit Sub
    
    ' Construire un message d'erreur formaté
    Dim errorMsg As String
    errorMsg = "Error #" & errObject.Number & ": " & errObject.Description
    
    If sourceRoutine <> "" Then
        errorMsg = errorMsg & " in " & sourceRoutine
    End If
    
    If toSheet Then
        WriteToSheet errorMsg, level, category, sourceRoutine, user
    End If
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ILoggerBase_FlushLogs()
    ' S'assurer que le classeur est sauvegardé
    On Error Resume Next
    m_TargetWorkbook.Save
    On Error GoTo 0
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
    ILoggerBase_IsInitialized = m_Initialized
End Property

' Méthodes privées
'@Description: 
'@Param: 
'@Returns: 

Private Sub WriteToSheet(ByVal msg As String, ByVal level As LogLevelEnum, _
                        ByVal category As String, ByVal source As String, _
                        ByVal user As String)
    On Error Resume Next
    
    ' Vérifier si la rotation est nécessaire
    If m_CurrentRow > MAX_LOG_ROWS Then
        RotateLogSheet
    End If
    
    Dim ws As Worksheet
    Set ws = m_TargetWorkbook.Sheets(m_LogSheetName)
    
    ' Écrire les données de log
    ws.Cells(m_CurrentRow, 1).Value = Now
    ws.Cells(m_CurrentRow, 2).Value = GetLevelName(level)
    ws.Cells(m_CurrentRow, 3).Value = category
    ws.Cells(m_CurrentRow, 4).Value = source
    ws.Cells(m_CurrentRow, 5).Value = user
    ws.Cells(m_CurrentRow, 6).Value = msg
    
    ' Appliquer un format conditionnel basé sur le niveau
    FormatLogRow m_CurrentRow, level
    
    ' Incrémenter le compteur de ligne
    m_CurrentRow = m_CurrentRow + 1
    
    On Error GoTo 0
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub CreateLogSheet()
    On Error Resume Next
    
    ' Créer une nouvelle feuille de logs
    Dim ws As Worksheet
    Set ws = m_TargetWorkbook.Worksheets.Add(After:=m_TargetWorkbook.Sheets(m_TargetWorkbook.Sheets.Count))
    ws.Name = m_LogSheetName
    
    ' Configurer les en-têtes
    ws.Cells(1, 1).Value = "Horodatage"
    ws.Cells(1, 2).Value = "Niveau"
    ws.Cells(1, 3).Value = "Catégorie"
    ws.Cells(1, 4).Value = "Source"
    ws.Cells(1, 5).Value = "Utilisateur"
    ws.Cells(1, 6).Value = "Message"
    
    ' Formater les en-têtes
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 6)).Font.Bold = True
    
    ' Ajuster les largeurs de colonne
    ws.Columns("A:A").ColumnWidth = 20 ' Horodatage
    ws.Columns("B:B").ColumnWidth = 10 ' Niveau
    ws.Columns("C:C").ColumnWidth = 15 ' Catégorie
    ws.Columns("D:D").ColumnWidth = 20 ' Source
    ws.Columns("E:E").ColumnWidth = 15 ' Utilisateur
    ws.Columns("F:F").ColumnWidth = 80 ' Message
    
    ' Réinitialiser à la première ligne de log
    m_CurrentRow = 2
    
    On Error GoTo 0
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    
    Set ws = m_TargetWorkbook.Sheets(sheetName)
    SheetExists = (Not ws Is Nothing)
    
    On Error GoTo 0
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Sub FindFirstEmptyRow()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = m_TargetWorkbook.Sheets(m_LogSheetName)
    
    ' Trouver la première ligne vide
    m_CurrentRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' S'assurer d'être au moins à la ligne 2
    If m_CurrentRow < 2 Then
        m_CurrentRow = 2
    End If
    
    On Error GoTo 0
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub RotateLogSheet()
    On Error Resume Next
    
    ' Créer une copie d'archive
    Dim archiveSheetName As String
    archiveSheetName = m_LogSheetName & "_Archive_" & Format(Now, "yyyymmdd_hhnnss")
    
    ' Créer une nouvelle feuille de logs
    CreateLogSheet
    
    ' Renommer l'ancienne feuille
    m_TargetWorkbook.Sheets(m_LogSheetName & "(2)").Name = archiveSheetName
    
    On Error GoTo 0
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub FormatLogRow(ByVal rowNum As Long, ByVal level As LogLevelEnum)
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = m_TargetWorkbook.Sheets(m_LogSheetName)
    
    ' Appliquer un format conditionnel basé sur le niveau
    Dim rowRange As Range
    Set rowRange = ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 6))
    
    Select Case level
        Case LogLevelError, LogLevelFatal
            rowRange.Interior.Color = RGB(255, 200, 200) ' Rouge pâle
        Case LogLevelWarning
            rowRange.Interior.Color = RGB(255, 255, 200) ' Jaune pâle
        Case Else
            ' Pas de format spécial pour les niveaux info et debug
    End Select
    
    On Error GoTo 0
End'@Description: 
'@Param: 
'@Returns: 

 Sub

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