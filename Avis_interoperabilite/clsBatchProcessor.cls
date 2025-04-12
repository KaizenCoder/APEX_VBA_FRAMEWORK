VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBatchProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ==========================================================================
' Classe : clsBatchProcessor
' Version : 1.0
' Purpose : Processeur de données par lots pour optimiser les performances
' Author  : APEX Framework Team
' Date    : 2024-04-11
' ==========================================================================

Private Const CLASS_NAME As String = "clsBatchProcessor"
Private Const LOG_CATEGORY As String = "BatchProcessor"

' Type de fonction de traitement pour un lot
Public Enum BatchProcessingMode
    bpmReadOnly = 0      ' Traitement en lecture seule
    bpmReadWrite = 1     ' Traitement avec écriture
End Enum

' Configuration
Private m_batchSize As Long               ' Taille d'un lot (nombre de lignes)
Private m_processingMode As BatchProcessingMode
Private m_statusCallback As Object        ' Objet avec méthode UpdateStatus(percent, message)
Private m_statusCallbackMethod As String  ' Nom de la méthode de callback

' Services
Private m_logger As ILogger
Private m_sheetAccessor As clsCachedSheetAccessor

' ============================================================================
' Function: Initialize
' Purpose: Initialise le processeur par lots
' ============================================================================
Public Function Initialize(ByVal sheetAccessor As Object, _
                         Optional ByVal batchSize As Long = 1000, _
                         Optional ByVal processingMode As BatchProcessingMode = bpmReadOnly) As clsBatchProcessor
    If sheetAccessor Is Nothing Then
        ReportError ERR_INVALID_ARGUMENT, "Initialize: sheetAccessor cannot be Nothing", CLASS_NAME
        Exit Function
    End If
    
    ' Récupérer les services
    Dim appContext As IApplicationContext
    Set appContext = GetApplicationContext()
    Set m_logger = appContext.Logger
    
    ' Si l'accesseur fourni n'est pas déjà un CachedSheetAccessor, l'encapsuler
    If TypeName(sheetAccessor) = "clsCachedSheetAccessor" Then
        Set m_sheetAccessor = sheetAccessor
    ElseIf TypeName(sheetAccessor) = "ISheetAccessor" Then
        Set m_sheetAccessor = New clsCachedSheetAccessor
        m_sheetAccessor.Initialize sheetAccessor
    Else
        ReportError ERR_INVALID_ARGUMENT, "Initialize: sheetAccessor must implement ISheetAccessor", CLASS_NAME
        Exit Function
    End If
    
    ' Valider et configurer la taille du lot
    If batchSize <= 0 Then batchSize = 1000
    m_batchSize = batchSize
    m_processingMode = processingMode
    
    m_logger.LogInfo "BatchProcessor initialisé avec taille de lot=" & m_batchSize & _
                    ", mode=" & IIf(m_processingMode = bpmReadOnly, "Lecture seule", "Lecture/Écriture"), _
                    LOG_CATEGORY
    
    Set Initialize = Me
End Function

' ============================================================================
' Property: BatchSize
' Purpose: Obtient ou définit la taille du lot
' ============================================================================
Public Property Get BatchSize() As Long
    BatchSize = m_batchSize
End Property

Public Property Let BatchSize(ByVal value As Long)
    If value <= 0 Then value = 1000
    m_batchSize = value
End Property

' ============================================================================
' Property: ProcessingMode
' Purpose: Obtient ou définit le mode de traitement
' ============================================================================
Public Property Get ProcessingMode() As BatchProcessingMode
    ProcessingMode = m_processingMode
End Property

Public Property Let ProcessingMode(ByVal value As BatchProcessingMode)
    m_processingMode = value
End Property

' ============================================================================
' Method: SetStatusCallback
' Purpose: Définit un callback pour les mises à jour de progression
' ============================================================================
Public Sub SetStatusCallback(ByVal callbackObject As Object, ByVal callbackMethod As String)
    Set m_statusCallback = callbackObject
    m_statusCallbackMethod = callbackMethod
End Sub

' ============================================================================
' Method: ProcessRange
' Purpose: Traite une plage de données par lots avec fonction de callback
' Parameters:
'   - startRow, startCol  : Début de la plage
'   - endRow, endCol      : Fin de la plage
'   - callbackObject      : Objet contenant la méthode de callback
'   - processMethod       : Nom de la méthode à appeler pour chaque lot
'   - additionalParam     : Paramètre supplémentaire à passer au callback (Optional)
' ============================================================================
Public Sub ProcessRange(ByVal startRow As Long, ByVal startCol As Long, _
                      ByVal endRow As Long, ByVal endCol As Long, _
                      ByVal callbackObject As Object, ByVal processMethod As String, _
                      Optional ByVal additionalParam As Variant)
    Dim currentRow As Long
    Dim batchStartRow As Long, batchEndRow As Long
    Dim batchData As Variant
    Dim totalRows As Long, processedRows As Long
    Dim hasAdditionalParam As Boolean
    Dim statusMessage As String
    Dim startTime As Double, endTime As Double
    Dim elapsedSeconds As Double, rowsPerSecond As Double
    
    On Error GoTo ErrorHandler
    
    ' Vérifier si le paramètre additionnel a été fourni
    hasAdditionalParam = Not IsMissing(additionalParam)
    
    ' Calculer le nombre total de lignes à traiter
    totalRows = endRow - startRow + 1
    processedRows = 0
    
    m_logger.LogInfo "Début traitement par lots - Plage R" & startRow & "C" & startCol & ":R" & endRow & "C" & endCol & _
                   ", Taille lot=" & m_batchSize & ", Mode=" & m_processingMode, LOG_CATEGORY
    
    ' Enregistrer l'heure de début
    startTime = Timer
    
    ' Traiter chaque lot
    currentRow = startRow
    Do While currentRow <= endRow
        ' Calculer les limites du lot courant
        batchStartRow = currentRow
        batchEndRow = WorksheetFunction.Min(currentRow + m_batchSize - 1, endRow)
        
        ' Mettre à jour le statut
        processedRows = processedRows + (batchEndRow - batchStartRow + 1)
        statusMessage = "Traitement lot " & Format((processedRows / totalRows), "0%") & _
                      " (" & Format(processedRows, "#,##0") & "/" & Format(totalRows, "#,##0") & " lignes)"
        
        UpdateStatus processedRows / totalRows, statusMessage
        
        ' Lire les données du lot
        batchData = m_sheetAccessor.ReadRange(batchStartRow, startCol, batchEndRow, endCol)
        
        ' Appeler la méthode de traitement avec les paramètres appropriés
        If hasAdditionalParam Then
            CallByName callbackObject, processMethod, VbMethod, batchData, batchStartRow, startCol, additionalParam
        Else
            CallByName callbackObject, processMethod, VbMethod, batchData, batchStartRow, startCol
        End If
        
        ' Si en mode lecture/écriture, rafraîchir le cache après chaque lot
        If m_processingMode = bpmReadWrite Then
            m_sheetAccessor.RefreshCache
        End If
        
        ' Passer au lot suivant
        currentRow = batchEndRow + 1
        
        ' Vérifier si l'utilisateur a appuyé sur Echap pour annuler
        DoEvents
        If Application.EnableCancelKey = xlErrorHandler Then
            On Error Resume Next
            If GetAsyncKeyState(VK_ESCAPE) <> 0 Then
                If MsgBox("Voulez-vous annuler le traitement par lots ?", vbYesNo + vbQuestion, "Annulation") = vbYes Then
                    m_logger.LogWarning "Traitement par lots annulé par l'utilisateur à " & Format(processedRows / totalRows, "0.0%"), LOG_CATEGORY
                    Exit Sub
                End If
            End If
            On Error GoTo ErrorHandler
        End If
    Loop
    
    ' Calculer et logger les statistiques
    endTime = Timer
    elapsedSeconds = endTime - startTime
    If elapsedSeconds > 0 Then
        rowsPerSecond = totalRows / elapsedSeconds
    Else
        rowsPerSecond = 0
    End If
    
    m_logger.LogInfo "Traitement par lots terminé - " & Format(totalRows, "#,##0") & " lignes en " & _
                   Format(elapsedSeconds, "0.00") & " secondes (" & Format(rowsPerSecond, "#,##0") & " lignes/sec)", _
                   LOG_CATEGORY
    
    ' Statut final
    UpdateStatus 1, "Traitement terminé - " & Format(totalRows, "#,##0") & " lignes traitées"
    
    Exit Sub
    
ErrorHandler:
    m_logger.LogError "Erreur ProcessRange: " & Err.Number & " - " & Err.Description & " à la ligne " & currentRow, LOG_CATEGORY
    ReportError Err.Number, "ProcessRange: " & Err.Description & " à la ligne " & currentRow, CLASS_NAME
End Sub

' ============================================================================
' Method: UpdateStatus
' Purpose: Met à jour la progression via le callback et/ou la barre d'état
' ============================================================================
Private Sub UpdateStatus(ByVal percentComplete As Double, ByVal statusMessage As String)
    ' Mettre à jour via le callback s'il est défini
    If Not m_statusCallback Is Nothing And Len(m_statusCallbackMethod) > 0 Then
        On Error Resume Next
        CallByName m_statusCallback, m_statusCallbackMethod, VbMethod, percentComplete, statusMessage
        On Error GoTo 0
    End If
    
    ' Mettre à jour la barre d'état d'Excel
    On Error Resume Next
    Application.StatusBar = statusMessage
    On Error GoTo 0
    
    ' Permettre à l'interface de se rafraîchir
    DoEvents
End Sub

' ============================================================================
' Method: Class_Terminate
' Purpose: Nettoyage à la destruction de l'objet
' ============================================================================
Private Sub Class_Terminate()
    ' Restaurer la barre d'état
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
    
    ' Libérer les références
    Set m_statusCallback = Nothing
    Set m_sheetAccessor = Nothing
    Set m_logger = Nothing
End Sub 