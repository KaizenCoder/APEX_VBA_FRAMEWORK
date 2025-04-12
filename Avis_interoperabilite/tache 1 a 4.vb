2024-07-26 21:00
Chat : 052

### ?? Contexte & Analyse
- ?? Critique : Impl�mentation de l'architecture d'interop�rabilit� APEX-Excel
- ?? Important : Respect des priorit�s et sp�cifications techniques d�finies
- ?? Standard : Structure coh�rente avec les conventions VBA et le framework APEX

### ? Action & Impact
- ?? Action : D�veloppement complet de l'architecture d'interop�rabilit� [REF:INTEROP-005]
- ?? R�sultat : Impl�mentation des 5 t�ches prioritaires [?? 45min]
- ?? Impact : Base technique solide pour l'industrialisation du framework

### ?? Synth�se & R�flexion
L'impl�mentation d'une architecture fond�e sur des interfaces abstraites repr�sente un investissement strat�gique majeur pour le framework APEX. En �tablissant ces patterns d'interop�rabilit� d�s les premi�res phases du d�veloppement, nous cr�ons non seulement une base technique solide mais �galement un mod�le conceptuel qui guidera l'ensemble des d�veloppements futurs. Cette approche, bien que demandant un effort initial plus important, g�n�rera des b�n�fices exponentiels en termes de maintenabilit�, testabilit� et extensibilit� tout au long du cycle de vie du produit.

Je vais maintenant impl�menter l'architecture d'interop�rabilit� Apex-Excel selon les sp�cifications du prompt. Voici le code source pour chaque composant requis, organis� par t�ches prioritaires.
T�che 1 : Syst�me de logging
1. Interface ILoggerBase.cls

' ILoggerBase.cls
' Description: Interface fondamentale pour tous les types de loggers du framework APEX.
' Permet l'abstraction et le remplacement des m�canismes de journalisation.
Option Explicit

' Journalise un message avec un niveau de s�v�rit� sp�cifi�
Public Sub Log(ByVal level As String, ByVal message As String)
End Sub

' Journalise un message informatif
Public Sub Info(ByVal message As String)
End Sub

' Journalise un avertissement
Public Sub Warn(ByVal message As String)
End Sub

' Journalise une erreur
Public Sub Error(ByVal message As String)
End Sub

' Journalise un message de d�bogage
Public Sub Debug(ByVal message As String)
End Sub


2. Impl�mentation clsDebugLogger.cls

' clsDebugLogger.cls
' Description: Logger qui redirige les messages vers Debug.Print
' Utile pour le d�veloppement et le d�bogage.
Option Explicit
Implements ILoggerBase

Private Sub ILoggerBase_Log(ByVal level As String, ByVal message As String)
    Debug.Print "[" & UCase(level) & "] " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " - " & message
End Sub

Private Sub ILoggerBase_Info(ByVal message As String)
    ILoggerBase_Log "INFO", message
End Sub

Private Sub ILoggerBase_Warn(ByVal message As String)
    ILoggerBase_Log "WARN", message
End Sub

Private Sub ILoggerBase_Error(ByVal message As String)
    ILoggerBase_Log "ERROR", message
End Sub

Private Sub ILoggerBase_Debug(ByVal message As String)
    ILoggerBase_Log "DEBUG", message
End Sub

' clsSheetLogger.cls
' Description: Logger qui �crit les messages dans une feuille Excel
' avec rotation automatique des logs
Option Explicit
Implements ILoggerBase

Private targetSheet As Worksheet
Private nextRow As Long
Private Const MAX_ROWS As Long = 1000

Public Sub Init(ByVal ws As Worksheet)
    Set targetSheet = ws
    PrepareSheet
End Sub

Private Sub PrepareSheet()
    On Error Resume Next
    
    If targetSheet Is Nothing Then Exit Sub
    
    ' Initialiser les en-t�tes si feuille vide
    If targetSheet.Cells(1, 1).Value = "" Then
        targetSheet.Cells(1, 1).Value = "Horodatage"
        targetSheet.Cells(1, 2).Value = "Niveau"
        targetSheet.Cells(1, 3).Value = "Message"
        
        ' Mise en forme des en-t�tes
        With targetSheet.Range("A1:C1")
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
        End With
    End If
    
    ' Trouver la prochaine ligne disponible
    nextRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row + 1
    
    ' V�rifier si rotation n�cessaire
    If nextRow > MAX_ROWS Then
        RotateLog
        nextRow = 2 ' Apr�s rotation, on recommence � la ligne 2
    End If
    
    On Error GoTo 0
End Sub

Private Sub RotateLog()
    On Error Resume Next
    
    ' G�n�rer un nom bas� sur la date et l'heure
    Dim ts As String
    ts = Format(Now, "yyyymmdd_hhnnss")
    
    Dim newSheetName As String
    newSheetName = "Logs_" & ts
    
    ' Limiter � 31 caract�res (limite Excel)
    If Len(newSheetName) > 31 Then newSheetName = Left(newSheetName, 31)
    
    ' Copier la feuille actuelle et la renommer
    targetSheet.Copy After:=targetSheet
    ActiveSheet.Name = newSheetName
    
    ' Effacer le contenu de la feuille principale (sauf les en-t�tes)
    targetSheet.Range("A2:C" & targetSheet.Rows.Count).ClearContents
    
    On Error GoTo 0
End Sub

Private Sub ILoggerBase_Log(ByVal level As String, ByVal message As String)
    On Error Resume Next
    
    If targetSheet Is Nothing Then Exit Sub
    
    ' �crire les donn�es
    targetSheet.Cells(nextRow, 1).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    targetSheet.Cells(nextRow, 2).Value = UCase(level)
    targetSheet.Cells(nextRow, 3).Value = message
    
    ' Mise en forme conditionnelle par niveau
    Select Case UCase(level)
        Case "ERROR"
            targetSheet.Cells(nextRow, 2).Interior.Color = RGB(255, 200, 200)
        Case "WARN"
            targetSheet.Cells(nextRow, 2).Interior.Color = RGB(255, 255, 200)
    End Select
    
    ' Incr�menter et v�rifier rotation
    nextRow = nextRow + 1
    If nextRow > MAX_ROWS Then RotateLog
    
    On Error GoTo 0
End Sub

Private Sub ILoggerBase_Info(ByVal message As String)
    ILoggerBase_Log "INFO", message
End Sub

Private Sub ILoggerBase_Warn(ByVal message As String)
    ILoggerBase_Log "WARN", message
End Sub

Private Sub ILoggerBase_Error(ByVal message As String)
    ILoggerBase_Log "ERROR", message
End Sub

Private Sub ILoggerBase_Debug(ByVal message As String)
    ILoggerBase_Log "DEBUG", message
End Sub




3. Impl�mentation clsSheetLogger.cls

' clsSheetLogger.cls
' Description: Logger qui �crit les messages dans une feuille Excel
' avec rotation automatique des logs
Option Explicit
Implements ILoggerBase

Private targetSheet As Worksheet
Private nextRow As Long
Private Const MAX_ROWS As Long = 1000

Public Sub Init(ByVal ws As Worksheet)
    Set targetSheet = ws
    PrepareSheet
End Sub

Private Sub PrepareSheet()
    On Error Resume Next
    
    If targetSheet Is Nothing Then Exit Sub
    
    ' Initialiser les en-t�tes si feuille vide
    If targetSheet.Cells(1, 1).Value = "" Then
        targetSheet.Cells(1, 1).Value = "Horodatage"
        targetSheet.Cells(1, 2).Value = "Niveau"
        targetSheet.Cells(1, 3).Value = "Message"
        
        ' Mise en forme des en-t�tes
        With targetSheet.Range("A1:C1")
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
        End With
    End If
    
    ' Trouver la prochaine ligne disponible
    nextRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row + 1
    
    ' V�rifier si rotation n�cessaire
    If nextRow > MAX_ROWS Then
        RotateLog
        nextRow = 2 ' Apr�s rotation, on recommence � la ligne 2
    End If
    
    On Error GoTo 0
End Sub

Private Sub RotateLog()
    On Error Resume Next
    
    ' G�n�rer un nom bas� sur la date et l'heure
    Dim ts As String
    ts = Format(Now, "yyyymmdd_hhnnss")
    
    Dim newSheetName As String
    newSheetName = "Logs_" & ts
    
    ' Limiter � 31 caract�res (limite Excel)
    If Len(newSheetName) > 31 Then newSheetName = Left(newSheetName, 31)
    
    ' Copier la feuille actuelle et la renommer
    targetSheet.Copy After:=targetSheet
    ActiveSheet.Name = newSheetName
    
    ' Effacer le contenu de la feuille principale (sauf les en-t�tes)
    targetSheet.Range("A2:C" & targetSheet.Rows.Count).ClearContents
    
    On Error GoTo 0
End Sub

Private Sub ILoggerBase_Log(ByVal level As String, ByVal message As String)
    On Error Resume Next
    
    If targetSheet Is Nothing Then Exit Sub
    
    ' �crire les donn�es
    targetSheet.Cells(nextRow, 1).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    targetSheet.Cells(nextRow, 2).Value = UCase(level)
    targetSheet.Cells(nextRow, 3).Value = message
    
    ' Mise en forme conditionnelle par niveau
    Select Case UCase(level)
        Case "ERROR"
            targetSheet.Cells(nextRow, 2).Interior.Color = RGB(255, 200, 200)
        Case "WARN"
            targetSheet.Cells(nextRow, 2).Interior.Color = RGB(255, 255, 200)
    End Select
    
    ' Incr�menter et v�rifier rotation
    nextRow = nextRow + 1
    If nextRow > MAX_ROWS Then RotateLog
    
    On Error GoTo 0
End Sub

Private Sub ILoggerBase_Info(ByVal message As String)
    ILoggerBase_Log "INFO", message
End Sub

Private Sub ILoggerBase_Warn(ByVal message As String)
    ILoggerBase_Log "WARN", message
End Sub

Private Sub ILoggerBase_Error(ByVal message As String)
    ILoggerBase_Log "ERROR", message
End Sub

Private Sub ILoggerBase_Debug(ByVal message As String)
    ILoggerBase_Log "DEBUG", message
End Sub

4. Impl�mentation clsFileLogger.cls

' clsFileLogger.cls
' Description: Logger qui �crit les messages dans un fichier texte externe
Option Explicit
Implements ILoggerBase

Private logFilePath As String
Private fileNum As Integer

Public Sub Init(ByVal path As String)
    On Error Resume Next
    
    logFilePath = path
    
    ' V�rifier si le dossier existe, sinon le cr�er
    Dim folderPath As String
    folderPath = Left(path, InStrRev(path, "\"))
    
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
    
    ' Ouvrir le fichier en mode Append
    fileNum = FreeFile
    Open logFilePath For Append As #fileNum
    
    ' �crire un en-t�te si nouveau fichier
    Dim fileLen As Long
    fileLen = LOF(fileNum)
    
    If fileLen = 0 Then
        Print #fileNum, "=== DEBUT DU JOURNAL DE LOGS - " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    Else
        Print #fileNum, ""
        Print #fileNum, "=== REPRISE DU JOURNAL - " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    End If
    
    On Error GoTo 0
End Sub

Private Sub ILoggerBase_Log(ByVal level As String, ByVal message As String)
    On Error Resume Next
    
    If fileNum = 0 Then Exit Sub
    
    Print #fileNum, "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "] [" & UCase(level) & "] " & message
    
    On Error GoTo 0
End Sub

Private Sub ILoggerBase_Info(ByVal message As String)
    ILoggerBase_Log "INFO", message
End Sub

Private Sub ILoggerBase_Warn(ByVal message As String)
    ILoggerBase_Log "WARN", message
End Sub

Private Sub ILoggerBase_Error(ByVal message As String)
    ILoggerBase_Log "ERROR", message
End Sub

Private Sub ILoggerBase_Debug(ByVal message As String)
    ILoggerBase_Log "DEBUG", message
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    If fileNum > 0 Then
        Print #fileNum, "=== FIN DE SESSION - " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
        Close #fileNum
    End If
    
    On Error GoTo 0
End Sub



5. Impl�mentation clsTestLogger.cls

' clsTestLogger.cls
' Description: Logger sp�cialis� pour les tests qui stocke les messages en m�moire
' pour v�rification ult�rieure
Option Explicit
Implements ILoggerBase

Private logs As Collection

Private Sub Class_Initialize()
    Set logs = New Collection
End Sub

Private Sub ILoggerBase_Log(ByVal level As String, ByVal message As String)
    logs.Add "[" & UCase(level) & "] " & message
End Sub

Private Sub ILoggerBase_Info(ByVal message As String)
    ILoggerBase_Log "INFO", message
End Sub

Private Sub ILoggerBase_Warn(ByVal message As String)
    ILoggerBase_Log "WARN", message
End Sub

Private Sub ILoggerBase_Error(ByVal message As String)
    ILoggerBase_Log "ERROR", message
End Sub

Private Sub ILoggerBase_Debug(ByVal message As String)
    ILoggerBase_Log "DEBUG", message
End Sub

' M�thodes sp�cifiques aux tests
Public Function GetLogs() As Collection
    Set GetLogs = logs
End Function

Public Function Contains(ByVal text As String) As Boolean
    Dim item As Variant
    For Each item In logs
        If InStr(1, item, text, vbTextCompare) > 0 Then
            Contains = True
            Exit Function
        End If
    Next
    Contains = False
End Function

Public Function Count() As Long
    Count = logs.Count
End Function

Public Sub Clear()
    Set logs = New Collection
End Sub

6. Impl�mentation clsCompositeLogger.cls
' clsCompositeLogger.cls
' Description: Logger qui combine plusieurs loggers (pattern Composite)
' Permet d'�crire simultan�ment dans plusieurs destinations
Option Explicit
Implements ILoggerBase

Private loggers As Collection

Private Sub Class_Initialize()
    Set loggers = New Collection
End Sub

Public Sub AddLogger(ByVal logger As ILoggerBase)
    loggers.Add logger
End Sub

Public Function GetLoggersCount() As Long
    GetLoggersCount = loggers.Count
End Function

Private Sub ILoggerBase_Log(ByVal level As String, ByVal message As String)
    Dim logger As ILoggerBase
    Dim i As Long
    
    For i = 1 To loggers.Count
        Set logger = loggers(i)
        logger.Log level, message
    Next i
End Sub

Private Sub ILoggerBase_Info(ByVal message As String)
    Dim logger As ILoggerBase
    Dim i As Long
    
    For i = 1 To loggers.Count
        Set logger = loggers(i)
        logger.Info message
    Next i
End Sub

Private Sub ILoggerBase_Warn(ByVal message As String)
    Dim logger As ILoggerBase
    Dim i As Long
    
    For i = 1 To loggers.Count
        Set logger = loggers(i)
        logger.Warn message
    Next i
End Sub

Private Sub ILoggerBase_Error(ByVal message As String)
    Dim logger As ILoggerBase
    Dim i As Long
    
    For i = 1 To loggers.Count
        Set logger = loggers(i)
        logger.Error message
    Next i
End Sub

Private Sub ILoggerBase_Debug(ByVal message As String)
    Dim logger As ILoggerBase
    Dim i As Long
    
    For i = 1 To loggers.Count
        Set logger = loggers(i)
        logger.Debug message
    Next i
End Sub


7. Factory modLogFactory.bas
' modLogFactory.bas
' Description: Centralise la cr�ation et la gestion des loggers dans l'application
Option Explicit

' D�finition des environnements possibles
Public Enum LoggerEnvironment
    LOGGER_DEV = 0
    LOGGER_TEST = 1
    LOGGER_PROD = 2
End Enum

' Logger par d�faut au niveau de l'application
Private defaultLogger As ILoggerBase

' R�cup�re le logger par d�faut (le cr�e si n�cessaire)
Public Function GetLogger() As ILoggerBase
    If defaultLogger Is Nothing Then
        Set defaultLogger = CreateDefaultLogger
    End If
    Set GetLogger = defaultLogger
End Function

' D�finit explicitement le logger par d�faut
Public Sub SetLogger(ByVal logger As ILoggerBase)
    Set defaultLogger = logger
End Sub

' Cr�e un logger Debug.Print standard
Public Function CreateDebugLogger() As ILoggerBase
    Dim logger As New clsDebugLogger
    Set CreateDebugLogger = logger
End Function

' Cr�e un logger Excel
Public Function CreateSheetLogger(ByVal ws As Worksheet) As ILoggerBase
    Dim logger As New clsSheetLogger
    logger.Init ws
    Set CreateSheetLogger = logger
End Function

' Cr�e un logger fichier texte
Public Function CreateFileLogger(ByVal path As String) As ILoggerBase
    Dim logger As New clsFileLogger
    logger.Init path
    Set CreateFileLogger = logger
End Function

' Cr�e un logger pour les tests
Public Function CreateTestLogger() As ILoggerBase
    Dim logger As New clsTestLogger
    Set CreateTestLogger = logger
End Function

' Cr�e un logger composite � partir de plusieurs loggers
Public Function CreateCompositeLogger(ParamArray loggers() As Variant) As ILoggerBase
    Dim composite As New clsCompositeLogger
    Dim i As Long
    
    For i = LBound(loggers) To UBound(loggers)
        composite.AddLogger loggers(i)
    Next i
    
    Set CreateCompositeLogger = composite
End Function

' Cr�e un logger adapt� � l'environnement sp�cifi�
Public Function CreateDefaultLogger(Optional ByVal env As LoggerEnvironment = LOGGER_DEV) As ILoggerBase
    Select Case env
        Case LOGGER_DEV
            ' En dev: console + fichier local
            Dim devLogger As ILoggerBase
            Set devLogger = CreateCompositeLogger( _
                CreateDebugLogger, _
                CreateFileLogger(ThisWorkbook.Path & "\logs\apex_dev.log"))
            Set CreateDefaultLogger = devLogger
            
        Case LOGGER_TEST
            ' En test: feuille + fichier
            On Error Resume Next
            Dim testWs As Worksheet
            
            ' Chercher ou cr�er la feuille de logs
            Set testWs = ThisWorkbook.Sheets("Logs")
            If testWs Is Nothing Then
                Set testWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                testWs.Name = "Logs"
            End If
            
            Dim testLogger As ILoggerBase
            Set testLogger = CreateCompositeLogger( _
                CreateSheetLogger(testWs), _
                CreateFileLogger(ThisWorkbook.Path & "\logs\apex_test.log"))
            Set CreateDefaultLogger = testLogger
            On Error GoTo 0
            
        Case LOGGER_PROD
            ' En prod: fichier seulement avec moins de verbosit�
            Dim prodLogger As ILoggerBase
            Set prodLogger = CreateFileLogger(ThisWorkbook.Path & "\logs\apex_prod.log")
            Set CreateDefaultLogger = prodLogger
            
        Case Else
            ' Par d�faut: logger console simple
            Set CreateDefaultLogger = CreateDebugLogger
    End Select
End Function

' D�tecte automatiquement l'environnement
Public Function DetectEnvironment() As LoggerEnvironment
    On Error GoTo DefaultDev
    
    ' Recherche d'une cellule ou plage nomm�e "ENV"
    Dim envStr As String
    envStr = UCase(Trim(ThisWorkbook.Names("ENV").RefersToRange.Value))
    
    Select Case envStr
        Case "DEV"
            DetectEnvironment = LOGGER_DEV
        Case "TEST"
            DetectEnvironment = LOGGER_TEST
        Case "PROD"
            DetectEnvironment = LOGGER_PROD
        Case Else
            GoTo DefaultDev
    End Select
    
    Exit Function
    
DefaultDev:
    DetectEnvironment = LOGGER_DEV
End Function

T�che 2 : Framework d'acc�s aux donn�es Excel
1. Interface IWorkbookAccessor.cls

' IWorkbookAccessor.cls
' Description: Interface pour acc�der aux classeurs Excel de mani�re abstraite
Option Explicit

' R�cup�re une feuille par son nom
Public Function GetSheet(ByVal name As String) As ISheetAccessor
End Function

' R�cup�re la feuille active
Public Function GetActiveSheet() As ISheetAccessor
End Function

' Enregistre le classeur sous un chemin sp�cifi�
Public Function SaveAs(ByVal path As String) As Boolean
End Function

2. Interface ISheetAccessor.cls
' ISheetAccessor.cls
' Description: Interface pour acc�der aux feuilles Excel de mani�re abstraite
Option Explicit

' R�cup�re une cellule par ses coordonn�es
Public Function GetCell(ByVal row As Long, ByVal col As Long) As ICellAccessor
End Function

' Lit une plage de cellules
Public Function ReadRange(ByVal startRow As Long, ByVal startCol As Long, _
                         ByVal endRow As Long, ByVal endCol As Long) As Variant
End Function

' �crit des donn�es dans une plage
Public Sub WriteRange(ByVal startRow As Long, ByVal startCol As Long, _
                     ByRef values As Variant)
End Sub

' R�cup�re le nom de la feuille
Public Function Name() As String
End Function

3. Interface ICellAccessor.cls
' ISheetAccessor.cls
' Description: Interface pour acc�der aux feuilles Excel de mani�re abstraite
Option Explicit

' R�cup�re une cellule par ses coordonn�es
Public Function GetCell(ByVal row As Long, ByVal col As Long) As ICellAccessor
End Function

' Lit une plage de cellules
Public Function ReadRange(ByVal startRow As Long, ByVal startCol As Long, _
                         ByVal endRow As Long, ByVal endCol As Long) As Variant
End Function

' �crit des donn�es dans une plage
Public Sub WriteRange(ByVal startRow As Long, ByVal startCol As Long, _
                     ByRef values As Variant)
End Sub

' R�cup�re le nom de la feuille
Public Function Name() As String
End Function

3. Interface ICellAccessor.cls
' ICellAccessor.cls
' Description: Interface pour acc�der aux cellules Excel de mani�re abstraite
Option Explicit

' R�cup�re la valeur de la cellule
Public Property Get Value() As Variant
End Property

' D�finit la valeur de la cellule
Public Property Let Value(ByVal newVal As Variant)
End Property

4. Impl�mentation clsExcelWorkbookAccessor.cls
' clsExcelWorkbookAccessor.cls
' Description: Impl�mentation concr�te de IWorkbookAccessor pour Excel
Option Explicit
Implements IWorkbookAccessor

Private xlWb As Workbook

Public Sub Init(ByVal wb As Workbook)
    Set xlWb = wb
End Sub

Private Function IWorkbookAccessor_GetSheet(ByVal name As String) As ISheetAccessor
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = xlWb.Worksheets(name)
    
    If ws Is Nothing Then
        ' Feuille non trouv�e
        Set IWorkbookAccessor_GetSheet = Nothing
        Exit Function
    End If
    
    Dim accessor As New clsExcelSheetAccessor
    accessor.Init ws
    
    Set IWorkbookAccessor_GetSheet = accessor
    
    On Error GoTo 0
End Function

Private Function IWorkbookAccessor_GetActiveSheet() As ISheetAccessor
    On Error Resume Next
    
    If xlWb.ActiveSheet Is Nothing Then
        Set IWorkbookAccessor_GetActiveSheet = Nothing
        Exit Function
    End If
    
    Dim accessor As New clsExcelSheetAccessor
    accessor.Init xlWb.ActiveSheet
    
    Set IWorkbookAccessor_GetActiveSheet = accessor
    
    On Error GoTo 0
End Function

Private Function IWorkbookAccessor_SaveAs(ByVal path As String) As Boolean
    On Error Resume Next
    
    xlWb.SaveAs path
    IWorkbookAccessor_SaveAs = (Err.Number = 0)
    
    On Error GoTo 0
End Function

5. Impl�mentation clsExcelSheetAccessor.cls
' clsExcelSheetAccessor.cls
' Description: Impl�mentation concr�te de ISheetAccessor pour Excel
Option Explicit
Implements ISheetAccessor

Private xlSheet As Worksheet

Public Sub Init(ByVal sheet As Worksheet)
    Set xlSheet = sheet
End Sub

Private Function ISheetAccessor_GetCell(ByVal row As Long, ByVal col As Long) As ICellAccessor
    On Error Resume Next
    
    Dim cell As New clsExcelCellAccessor
    cell.Init xlSheet.Cells(row, col)
    
    Set ISheetAccessor_GetCell = cell
    
    On Error GoTo 0
End Function

Private Function ISheetAccessor_ReadRange(ByVal startRow As Long, ByVal startCol As Long, _
                                         ByVal endRow As Long, ByVal endCol As Long) As Variant
    On Error Resume Next
    
    ISheetAccessor_ReadRange = xlSheet.Range( _
        xlSheet.Cells(startRow, startCol), _
        xlSheet.Cells(endRow, endCol)).Value
    
    On Error GoTo 0
End Function

Private Sub ISheetAccessor_WriteRange(ByVal startRow As Long, ByVal startCol As Long, _
                                    ByRef values As Variant)
    On Error Resume Next
    
    If Not IsArray(values) Then
        ' Valeur simple
        xlSheet.Cells(startRow, startCol).Value = values
    Else
        ' Tableau de valeurs
        Dim targetRange As Range
        Set targetRange = xlSheet.Range( _
            xlSheet.Cells(startRow, startCol), _
            xlSheet.Cells(startRow + UBound(values, 1) - LBound(values, 1), _
                         startCol + UBound(values, 2) - LBound(values, 2)))
        targetRange.Value = values
    End If
    
    On Error GoTo 0
End Sub

Private Function ISheetAccessor_Name() As String
    ISheetAccessor_Name = xlSheet.Name
End Function


6. Impl�mentation clsExcelCellAccessor.cls
' clsExcelCellAccessor.cls
' Description: Impl�mentation concr�te de ICellAccessor pour Excel
Option Explicit
Implements ICellAccessor

Private xlCell As Range

Public Sub Init(ByVal cell As Range)
    Set xlCell = cell
End Sub

Private Property Get ICellAccessor_Value() As Variant
    ICellAccessor_Value = xlCell.Value
End Property

Private Property Let ICellAccessor_Value(ByVal newVal As Variant)
    xlCell.Value = newVal
End Property

7. Impl�mentation clsMockWorkbookAccessor.cls
' clsMockWorkbookAccessor.cls
' Description: Mock de IWorkbookAccessor pour les tests
Option Explicit
Implements IWorkbookAccessor

Private mockSheets As Collection
Private mockName As String

Private Sub Class_Initialize()
    Set mockSheets = New Collection
    mockName = "MockWorkbook"
End Sub

Public Sub AddMockSheet(ByVal sheetName As String)
    On Error Resume Next
    
    Dim mockSheet As New clsMockSheetAccessor
    mockSheet.Init sheetName
    
    ' V�rifier si la feuille existe d�j�
    Dim i As Long
    For i = 1 To mockSheets.Count
        If TypeOf mockSheets(i) Is clsMockSheetAccessor Then
            If mockSheets(i).Name = sheetName Then
                ' Remplacer la feuille existante
                mockSheets.Remove i
                Exit For
            End If
        End If
    Next i
    
    mockSheets.Add mockSheet, sheetName
    
    On Error GoTo 0
End Sub

Public Sub SetName(ByVal name As String)
    mockName = name
End Sub

Private Function IWorkbookAccessor_GetSheet(ByVal name As String) As ISheetAccessor
    On Error Resume Next
    
    Dim sheet As ISheetAccessor
    Set sheet = mockSheets(name)
    
    Set IWorkbookAccessor_GetSheet = sheet
    
    On Error GoTo 0
End Function

Private Function IWorkbookAccessor_GetActiveSheet() As ISheetAccessor
    ' Par d�faut, retourne la premi�re feuille
    If mockSheets.Count > 0 Then
        Set IWorkbookAccessor_GetActiveSheet = mockSheets(1)
    Else
        Set IWorkbookAccessor_GetActiveSheet = Nothing
    End If
End Function

Private Function IWorkbookAccessor_SaveAs(ByVal path As String) As Boolean
    ' Simule toujours un succ�s
    IWorkbookAccessor_SaveAs = True
End Function

8. Impl�mentation clsMockSheetAccessor.cls
' clsMockSheetAccessor.cls
' Description: Mock de ISheetAccessor pour les tests
Option Explicit
Implements ISheetAccessor

Private sheetName As String
Private cellValues As Object ' Dictionary pour stocker les valeurs des cellules

Private Sub Class_Initialize()
    Set cellValues = CreateObject("Scripting.Dictionary")
End Sub

Public Sub Init(ByVal name As String)
    sheetName = name
End Sub

' M�thode helper pour construire une cl� de cellule
Private Function CellKey(ByVal row As Long, ByVal col As Long) As String
    CellKey = row & "_" & col
End Function

Private Function ISheetAccessor_GetCell(ByVal row As Long, ByVal col As Long) As ICellAccessor
    Dim cell As New clsMockCellAccessor
    cell.Init Me, row, col
    
    Set ISheetAccessor_GetCell = cell
End Function

Private Function ISheetAccessor_ReadRange(ByVal startRow As Long, ByVal startCol As Long, _
                                         ByVal endRow As Long, ByVal endCol As Long) As Variant
    ' Cr�er un tableau pour stocker les valeurs
    Dim result() As Variant
    ReDim result(startRow To endRow, startCol To endCol)
    
    ' Remplir le tableau avec les valeurs stock�es
    Dim r As Long, c As Long
    For r = startRow To endRow
        For c = startCol To endCol
            Dim key As String
            key = CellKey(r, c)
            
            If cellValues.Exists(key) Then
                result(r, c) = cellValues(key)
            Else
                result(r, c) = Empty
            End If
        Next c
    Next r
    
    ISheetAccessor_ReadRange = result
End Function

Private Sub ISheetAccessor_WriteRange(ByVal startRow As Long, ByVal startCol As Long, _
                                     ByRef values As Variant)
    ' G�rer le cas d'une valeur simple
    If Not IsArray(values) Then
        cellValues(CellKey(startRow, startCol)) = values
        Exit Sub
    End If
    
    ' G�rer le cas d'un tableau
    Dim r As Long, c As Long
    Dim rowOffset As Long, colOffset As Long
    
    For r = LBound(values, 1) To UBound(values, 1)
        rowOffset = r - LBound(values, 1)
        
        For c = LBound(values, 2) To UBound(values, 2)
            colOffset = c - LBound(values, 2)
            
            cellValues(CellKey(startRow + rowOffset, startCol + colOffset)) = values(r, c)
        Next c
    Next r
End Sub

Private Function ISheetAccessor_Name() As String
    ISheetAccessor_Name = sheetName
End Function

' M�thodes sp�cifiques pour les tests
Public Sub SetCellValue(ByVal row As Long, ByVal col As Long, ByVal value As Variant)
    cellValues(CellKey(row, col)) = value
End Sub

Public Function GetCellValue(ByVal row As Long, ByVal col As Long) As Variant
    Dim key As String
    key = CellKey(row, col)
    
    If cellValues.Exists(key) Then
        GetCellValue = cellValues(key)
    Else
        GetCellValue = Empty
    End If
End Function

9. Impl�mentation clsMockCellAccessor.cls
' clsMockCellAccessor.cls
' Description: Mock de ICellAccessor pour les tests
Option Explicit
Implements ICellAccessor

Private parentSheet As clsMockSheetAccessor
Private cellRow As Long
Private cellCol As Long

Public Sub Init(ByVal sheet As clsMockSheetAccessor, ByVal row As Long, ByVal col As Long)
    Set parentSheet = sheet
    cellRow = row
    cellCol = col
End Sub

Private Property Get ICellAccessor_Value() As Variant
    ICellAccessor_Value = parentSheet.GetCellValue(cellRow, cellCol)
End Property

Private Property Let ICellAccessor_Value(ByVal newVal As Variant)
    parentSheet.SetCellValue cellRow, cellCol, newVal
End Property

T�che 3 : Contexte d'application
1. Interface IAppContext.cls
' IAppContext.cls
' Description: Interface pour le contexte global de l'application
Option Explicit

' R�cup�re le logger configur�
Public Property Get Logger() As ILoggerBase
End Property

' R�cup�re l'environnement courant
Public Property Get Environment() As LoggerEnvironment
End Property

' R�cup�re une valeur de configuration
Public Property Get Config(ByVal key As String) As String
End Property

' R�cup�re un accesseur pour un classeur
Public Function GetWorkbookAccessor(ByVal wb As Workbook) As IWorkbookAccessor
End Function

' L�ve une erreur personnalis�e avec journalisation
Public Sub RaiseError(ByVal code As Long, ByVal source As String, ByVal message As String)
End Sub

' Journalise une exception
Public Sub ReportException(ByVal source As String)
End Sub

2. Impl�mentation clsAppContext.cls
' clsAppContext.cls
' Description: Impl�mentation du contexte d'application central
Option Explicit
Implements IAppContext

Private appLogger As ILoggerBase
Private appEnv As LoggerEnvironment
Private appConfig As Object ' Dictionary
Private appWorkbooks As Object ' Dictionary d'accesseurs de classeurs

Private Sub Class_Initialize()
    Set appConfig = CreateObject("Scripting.Dictionary")
    Set appWorkbooks = CreateObject("Scripting.Dictionary")
    
    ' Initialiser avec l'environnement d�tect�
    Init DetectEnvironment
End Sub

Public Sub Init(Optional ByVal env As LoggerEnvironment = LOGGER_DEV)
    appEnv = env
    Set appLogger = CreateDefaultLogger(env)
    
    LoadDefaultConfig
    
    appLogger.Info "Contexte d'application initialis� - Environnement : " & EnvToString(appEnv)
End Sub

Private Sub LoadDefaultConfig()
    ' Configuration par d�faut
    appConfig.Add "LogFilePath", ThisWorkbook.Path & "\logs\apex_" & EnvToString(appEnv) & ".log"
    appConfig.Add "TempFolder", Environ("TEMP")
    appConfig.Add "DefaultSheet", "Donn�es"
    appConfig.Add "DateFormat", "yyyy-mm-dd"
    appConfig.Add "MaxCacheSize", "100"
    
    ' Tenter de charger la configuration depuis une feuille Excel si elle existe
    On Error Resume Next
    
    Dim configSheet As Worksheet
    Set configSheet = ThisWorkbook.Sheets("Config")
    
    If Not configSheet Is Nothing Then
        Dim lastRow As Long
        lastRow = configSheet.Cells(configSheet.Rows.Count, 1).End(xlUp).Row
        
        Dim i As Long
        For i = 2 To lastRow ' En supposant que la ligne 1 contient des en-t�tes
            Dim key As String, value As String
            
            key = Trim(configSheet.Cells(i, 1).Value)
            value = Trim(configSheet.Cells(i, 2).Value)
            
            If key <> "" Then
                If appConfig.Exists(key) Then
                    appConfig(key) = value
                Else
                    appConfig.Add key, value
                End If
            End If
        Next i
        
        appLogger.Info "Configuration charg�e depuis la feuille Config"
    End If
    
    On Error GoTo 0
End Sub

Private Function EnvToString(env As LoggerEnvironment) As String
    Select Case env
        Case LOGGER_DEV
            EnvToString = "dev"
        Case LOGGER_TEST
            EnvToString = "test"
        Case LOGGER_PROD
            EnvToString = "prod"
        Case Else
            EnvToString = "unknown"
    End Select
End Function

Private Property Get IAppContext_Logger() As ILoggerBase
    Set IAppContext_Logger = appLogger
End Property

Private Property Get IAppContext_Environment() As LoggerEnvironment
    IAppContext_Environment = appEnv
End Property

Private Property Get IAppContext_Config(ByVal key As String) As String
    If appConfig.Exists(key) Then
        IAppContext_Config = appConfig(key)
    Else
        IAppContext_Config = ""
    End If
End Property

Private Function IAppContext_GetWorkbookAccessor(ByVal wb As Workbook) As IWorkbookAccessor
    On Error Resume Next
    
    Dim wbPath As String
    wbPath = wb.Path & "\" & wb.Name
    
    ' V�rifier si l'accesseur existe d�j�
    If appWorkbooks.Exists(wbPath) Then
        Set IAppContext_GetWorkbookAccessor = appWorkbooks(wbPath)
    Else
        ' Cr�er un nouvel accesseur
        Dim accessor As New clsExcelWorkbookAccessor
        accessor.Init wb
        
        ' Stocker l'accesseur pour r�utilisation
        appWorkbooks.Add wbPath, accessor
        
        Set IAppContext_GetWorkbookAccessor = accessor
    End If
    
    On Error GoTo 0
End Function

Private Sub IAppContext_RaiseError(ByVal code As Long, ByVal source As String, ByVal message As String)
    ' Journaliser l'erreur
    appLogger.Error "[" & source & "] " & message & " (Code: " & code & ")"
    
    ' Lever l'erreur
    Err.Raise vbObjectError + code, source, message
End Sub

Private Sub IAppContext_ReportException(ByVal source As String)
    Dim msg As String
    msg = "Erreur #" & Err.Number & " dans " & Err.Source & " - " & Err.Description
    
    If source <> "" Then
        msg = "[" & source & "] " & msg
    End If
    
    appLogger.Error msg
End Sub

' Lib�rer les ressources
Private Sub Class_Terminate()
    ' Clear dictionary
    appConfig.RemoveAll
    appWorkbooks.RemoveAll
    
    Set appConfig = Nothing
    Set appWorkbooks = Nothing
    Set appLogger = Nothing
End Sub

T�che 4 : Template de module m�tier standardis�
' modTraitementStandard.bas
' Description: Template pour les modules m�tier standardis�s
Option Explicit

Private ctx As IAppContext

' Point d'entr�e principal du module
Public Sub RunTraitementStandard(ByVal injectedCtx As IAppContext)
    Set ctx = injectedCtx
    On Error GoTo GestionErreur

    ctx.Logger.Info "=== D�BUT TraitementStandard ==="

    Call �tape_Initialisation
    Call �tape_ChargementDesDonn�es
    Call �tape_Traitement
    Call �tape_Export

    ctx.Logger.Info "=== FIN TraitementStandard ==="
    Exit Sub

GestionErreur:
    ctx.ReportException "RunTraitementStandard"
    ' Informer l'utilisateur
    MsgBox "Une erreur est survenue lors du traitement. Consultez les logs pour plus d'informations.", _
           vbExclamation, "Erreur traitement"
End Sub

' �tape 1 : Initialisation des donn�es et v�rifications
Private Sub �tape_Initialisation()
    ctx.Logger.Debug "Initialisation des variables et v�rifications pr�alables"
    
    ' R�cup�rer le nom de la feuille depuis la config
    Dim sheetName As String
    sheetName = ctx.Config("DefaultSheet")
    
    If sheetName = "" Then
        ctx.RaiseError 1001, "�tape_Initialisation", "Aucune feuille de donn�es d�finie dans la configuration."
    End If
    
    ' V�rifier que la feuille existe
    Dim workbook As IWorkbookAccessor
    Set workbook = ctx.GetWorkbookAccessor(ThisWorkbook)
    
    If workbook.GetSheet(sheetName) Is Nothing Then
        ctx.RaiseError 1002, "�tape_Initialisation", "La feuille """ & sheetName & """ est introuvable."
    End If
    
    ctx.Logger.Info "Initialisation r�ussie - Feuille de donn�es: " & sheetName
End Sub

' �tape 2 : Chargement des donn�es
Private Sub �tape_ChargementDesDonn�es()
    ctx.Logger.Debug "Chargement des donn�es depuis " & ctx.Config("DefaultSheet")
    
    ' R�cup�rer la feuille
    Dim workbook As IWorkbookAccessor
    Set workbook = ctx.GetWorkbookAccessor(ThisWorkbook)
    
    Dim sheet As ISheetAccessor
    Set sheet = workbook.GetSheet(ctx.Config("DefaultSheet"))
    
    ' D�terminer la plage de donn�es
    ' [Code � personnaliser selon le format des donn�es]
    
    ' Exemple: lire les donn�es
    Dim data As Variant
    data = sheet.ReadRange(2, 1, 10, 5) ' Exemple: lignes 2-10, colonnes A-E
    
    ' Validation des donn�es charg�es
    If Not IsArray(data) Then
        ctx.RaiseError 2001, "�tape_ChargementDesDonn�es", "Aucune donn�e n'a pu �tre charg�e."
    End If
    
    ctx.Logger.Info "Donn�es charg�es avec succ�s"
End Sub

' �tape 3 : Traitement m�tier
Private Sub �tape_Traitement()
    ctx.Logger.Debug "Ex�cution du traitement m�tier principal"
    
    ' Ici, impl�menter le traitement sp�cifique
    ' [Code � personnaliser selon le traitement requis]
    
    ' Exemple de traitement fictif
    Dim i As Long
    For i = 1 To 5
        ctx.Logger.Debug "Traitement de l'�l�ment " & i
        
        ' Simulation d'une dur�e de traitement
        Application.Wait Now + TimeSerial(0, 0, 1)
    Next i
    
    ctx.Logger.Info "Traitement ex�cut� avec succ�s"
End Sub

' �tape 4 : Export des r�sultats
Private Sub �tape_Export()
    ctx.Logger.Debug "Export des r�sultats"
    
    ' Ici, impl�menter l'export des r�sultats
    ' [Code � personnaliser selon le format d'export]
    
    ' Exemple d'export fictif vers une feuille de r�sultats
    On Error Resume Next
    
    Dim resultSheet As Worksheet
    
    ' V�rifier si la feuille existe, sinon la cr�er
    Set resultSheet = ThisWorkbook.Sheets("R�sultats")
    If resultSheet Is Nothing Then
        Set resultSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        resultSheet.Name = "R�sultats"
    End If
    
    ' Cr�er un accesseur pour la feuille
    Dim workbook As IWorkbookAccessor
    Set workbook = ctx.GetWorkbookAccessor(ThisWorkbook)
    
    Dim sheet As ISheetAccessor
    Set sheet = workbook.GetSheet("R�sultats")
    
    ' �crire des en-t�tes
    sheet.GetCell(1, 1).Value = "Date d'ex�cution"
    sheet.GetCell(1, 2).Value = "Utilisateur"
    sheet.GetCell(1, 3).Value = "Statut"
    
    ' �crire le r�sultat
    sheet.GetCell(2, 1).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    sheet.GetCell(2, 2).Value = Environ("USERNAME")
    sheet.GetCell(2, 3).Value = "Traitement termin�"
    
    On Error GoTo 0
    
    ctx.Logger.Info "Export des r�sultats termin�"
End Sub


