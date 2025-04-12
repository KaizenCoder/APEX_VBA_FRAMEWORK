VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUnitOfWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Interoperabilite.Integration")
'@ModuleDescription("Implémentation du pattern Unit of Work pour la gestion des transactions")
Option Explicit

'----------------------------------------------------------------------------------------
' Constantes
'----------------------------------------------------------------------------------------
Private Const MODULE_NAME As String = "clsUnitOfWork"
Private Const MAX_OPERATIONS As Long = 100

'----------------------------------------------------------------------------------------
' Énumérations
'----------------------------------------------------------------------------------------
Public Enum UnitOfWorkState
    uowIdle = 0
    uowActive = 1
    uowCommitting = 2
    uowRollingBack = 3
    uowFailed = 4
    uowCommitted = 5
End Enum

Public Enum OperationType
    opInsert = 1
    opUpdate = 2
    opDelete = 3
    opCustom = 4
End Enum

'----------------------------------------------------------------------------------------
' Types personnalisés
'----------------------------------------------------------------------------------------
Private Type OperationRegistration
    EntityName As String
    EntityId As Variant
    OperationType As OperationType
    OriginalState As Variant
    NewState As Variant
    CallbackModule As String
    CallbackMethod As String
    Parameters As Variant
    IsRegistered As Boolean
End Type

'----------------------------------------------------------------------------------------
' Variables privées
'----------------------------------------------------------------------------------------
Private m_operations() As OperationRegistration
Private m_operationCount As Long
Private m_state As UnitOfWorkState
Private m_isInitialized As Boolean
Private m_logger As Object
Private m_lastError As String
Private m_transactionDepth As Long

'----------------------------------------------------------------------------------------
' Initialisation
'----------------------------------------------------------------------------------------

'@Description("Initialise l'unité de travail")
Public Function Initialize(Optional ByVal logger As Object = Nothing) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier si déjà initialisé
    If m_isInitialized Then
        m_lastError = "L'unité de travail est déjà initialisée"
        Initialize = True
        Exit Function
    End If
    
    ' Initialiser le tableau d'opérations
    ReDim m_operations(1 To MAX_OPERATIONS) As OperationRegistration
    m_operationCount = 0
    
    ' Stocker le logger s'il est fourni
    If Not logger Is Nothing Then
        Set m_logger = logger
    End If
    
    ' Initialiser l'état
    m_state = uowIdle
    m_transactionDepth = 0
    
    ' Marquer comme initialisé
    m_isInitialized = True
    
    ' Log de l'initialisation
    LogInfo "Unité de travail initialisée avec succès"
    
    Initialize = True
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de l'initialisation de l'unité de travail: " & Err.Description
    Initialize = False
End Function

'----------------------------------------------------------------------------------------
' Gestion des transactions
'----------------------------------------------------------------------------------------

'@Description("Démarre une nouvelle transaction ou incrémente le niveau d'imbrication")
Public Function BeginTransaction() As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier l'initialisation
    If Not m_isInitialized Then
        m_lastError = "L'unité de travail n'est pas initialisée"
        BeginTransaction = False
        Exit Function
    End If
    
    ' Si c'est la première transaction, initialiser l'état
    If m_transactionDepth = 0 Then
        m_state = uowActive
        ReDim m_operations(1 To MAX_OPERATIONS) As OperationRegistration
        m_operationCount = 0
        LogInfo "Nouvelle transaction démarrée"
    End If
    
    ' Incrémenter le niveau d'imbrication
    m_transactionDepth = m_transactionDepth + 1
    LogInfo "Niveau de transaction: " & m_transactionDepth
    
    BeginTransaction = True
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors du démarrage de la transaction: " & Err.Description
    LogError "Erreur lors du démarrage de la transaction: " & Err.Description
    BeginTransaction = False
End Function

'@Description("Valide la transaction actuelle, si c'est la dernière imbriquée")
Public Function CommitTransaction() As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier l'initialisation et l'état
    If Not m_isInitialized Then
        m_lastError = "L'unité de travail n'est pas initialisée"
        CommitTransaction = False
        Exit Function
    End If
    
    If m_transactionDepth = 0 Then
        m_lastError = "Aucune transaction active à valider"
        LogWarning "Tentative de validation sans transaction active"
        CommitTransaction = False
        Exit Function
    End If
    
    ' Décrémenter le niveau d'imbrication
    m_transactionDepth = m_transactionDepth - 1
    
    ' Si c'est la dernière transaction imbriquée, exécuter le commit
    If m_transactionDepth = 0 Then
        m_state = uowCommitting
        LogInfo "Validation des modifications en cours..."
        
        ' Exécuter toutes les opérations enregistrées
        If Not ExecuteOperations() Then
            ' En cas d'échec, marquer comme échoué et effectuer un rollback
            m_state = uowFailed
            RollbackTransactionInternal
            m_lastError = "Échec de la validation: " & m_lastError
            LogError "Échec de la validation: " & m_lastError
            CommitTransaction = False
            Exit Function
        End If
        
        ' Marquer comme validé et nettoyer
        m_state = uowCommitted
        CleanupTransactionData
        LogInfo "Transaction validée avec succès"
    Else
        LogInfo "Validation partielle - niveau de transaction restant: " & m_transactionDepth
    End If
    
    CommitTransaction = True
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de la validation de la transaction: " & Err.Description
    LogError "Erreur lors de la validation de la transaction: " & Err.Description
    
    ' En cas d'erreur, effectuer un rollback
    m_state = uowFailed
    If m_transactionDepth > 0 Then
        RollbackTransactionInternal
    End If
    
    CommitTransaction = False
End Function

'@Description("Annule la transaction actuelle et toutes les transactions imbriquées")
Public Function RollbackTransaction() As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier l'initialisation et l'état
    If Not m_isInitialized Then
        m_lastError = "L'unité de travail n'est pas initialisée"
        RollbackTransaction = False
        Exit Function
    End If
    
    If m_transactionDepth = 0 Then
        m_lastError = "Aucune transaction active à annuler"
        LogWarning "Tentative d'annulation sans transaction active"
        RollbackTransaction = False
        Exit Function
    End If
    
    ' Appeler l'implémentation interne du rollback
    RollbackTransaction = RollbackTransactionInternal()
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de l'annulation de la transaction: " & Err.Description
    LogError "Erreur lors de l'annulation de la transaction: " & Err.Description
    RollbackTransaction = False
End Function

'@Description("Implémentation interne de l'annulation de la transaction")
Private Function RollbackTransactionInternal() As Boolean
    On Error Resume Next
    
    m_state = uowRollingBack
    LogInfo "Annulation des modifications en cours..."
    
    ' Réinitialiser le niveau d'imbrication à 0
    m_transactionDepth = 0
    
    ' Nettoyer les données de la transaction
    CleanupTransactionData
    
    ' Réinitialiser l'état
    m_state = uowIdle
    
    LogInfo "Transaction annulée avec succès"
    RollbackTransactionInternal = True
End Function

'@Description("Exécute toutes les opérations enregistrées")
Private Function ExecuteOperations() As Boolean
    On Error GoTo ErrorHandler
    
    Dim i As Long
    
    ' Parcourir toutes les opérations enregistrées
    For i = 1 To m_operationCount
        If m_operations(i).IsRegistered Then
            ' Si une méthode de callback est définie, l'exécuter
            If Len(m_operations(i).CallbackModule) > 0 And Len(m_operations(i).CallbackMethod) > 0 Then
                Dim callbackResult As Variant
                
                ' Construire l'appel complet pour Application.Run
                Dim callbackPath As String
                callbackPath = m_operations(i).CallbackModule & "." & m_operations(i).CallbackMethod
                
                ' Exécuter le callback avec les paramètres
                If IsArray(m_operations(i).Parameters) Then
                    ' Si des paramètres sont fournis, les passer au callback
                    callbackResult = Application.Run(callbackPath, m_operations(i).Parameters)
                Else
                    ' Sinon, exécuter le callback sans paramètres
                    callbackResult = Application.Run(callbackPath)
                End If
                
                ' Vérifier le résultat du callback (on suppose qu'il renvoie un booléen)
                If Not CBool(callbackResult) Then
                    m_lastError = "Échec de l'exécution de l'opération " & i & " pour l'entité " & m_operations(i).EntityName
                    LogError m_lastError
                    ExecuteOperations = False
                    Exit Function
                End If
            End If
        End If
    Next i
    
    ExecuteOperations = True
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de l'exécution des opérations: " & Err.Description
    LogError "Erreur lors de l'exécution des opérations: " & Err.Description
    ExecuteOperations = False
End Function

'@Description("Nettoie les données de la transaction après validation ou annulation")
Private Sub CleanupTransactionData()
    On Error Resume Next
    
    ' Réinitialiser le tableau d'opérations
    ReDim m_operations(1 To MAX_OPERATIONS) As OperationRegistration
    m_operationCount = 0
    
    ' Réinitialiser l'état
    m_state = uowIdle
End Sub

'----------------------------------------------------------------------------------------
' Enregistrement des opérations
'----------------------------------------------------------------------------------------

'@Description("Enregistre une opération d'insertion")
Public Function RegisterInsert(ByVal entityName As String, ByVal entityId As Variant, _
                              ByVal newState As Variant, _
                              Optional ByVal callbackModule As String = "", _
                              Optional ByVal callbackMethod As String = "", _
                              Optional ByVal parameters As Variant = Null) As Boolean
    RegisterInsert = RegisterOperation(entityName, entityId, opInsert, Empty, newState, _
                                     callbackModule, callbackMethod, parameters)
End Function

'@Description("Enregistre une opération de mise à jour")
Public Function RegisterUpdate(ByVal entityName As String, ByVal entityId As Variant, _
                              ByVal originalState As Variant, ByVal newState As Variant, _
                              Optional ByVal callbackModule As String = "", _
                              Optional ByVal callbackMethod As String = "", _
                              Optional ByVal parameters As Variant = Null) As Boolean
    RegisterUpdate = RegisterOperation(entityName, entityId, opUpdate, originalState, newState, _
                                     callbackModule, callbackMethod, parameters)
End Function

'@Description("Enregistre une opération de suppression")
Public Function RegisterDelete(ByVal entityName As String, ByVal entityId As Variant, _
                              ByVal originalState As Variant, _
                              Optional ByVal callbackModule As String = "", _
                              Optional ByVal callbackMethod As String = "", _
                              Optional ByVal parameters As Variant = Null) As Boolean
    RegisterDelete = RegisterOperation(entityName, entityId, opDelete, originalState, Empty, _
                                     callbackModule, callbackMethod, parameters)
End Function

'@Description("Enregistre une opération personnalisée")
Public Function RegisterCustom(ByVal entityName As String, Optional ByVal entityId As Variant = Null, _
                              Optional ByVal originalState As Variant = Null, Optional ByVal newState As Variant = Null, _
                              Optional ByVal callbackModule As String = "", _
                              Optional ByVal callbackMethod As String = "", _
                              Optional ByVal parameters As Variant = Null) As Boolean
    RegisterCustom = RegisterOperation(entityName, entityId, opCustom, originalState, newState, _
                                     callbackModule, callbackMethod, parameters)
End Function

'@Description("Enregistre une opération générique")
Private Function RegisterOperation(ByVal entityName As String, ByVal entityId As Variant, _
                                  ByVal operationType As OperationType, _
                                  ByVal originalState As Variant, ByVal newState As Variant, _
                                  ByVal callbackModule As String, ByVal callbackMethod As String, _
                                  ByVal parameters As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier l'initialisation et l'état
    If Not m_isInitialized Then
        m_lastError = "L'unité de travail n'est pas initialisée"
        RegisterOperation = False
        Exit Function
    End If
    
    If m_state <> uowActive Then
        m_lastError = "Aucune transaction active pour enregistrer une opération"
        LogWarning "Tentative d'enregistrement d'opération sans transaction active"
        RegisterOperation = False
        Exit Function
    End If
    
    ' Vérifier les arguments
    If Len(entityName) = 0 Then
        m_lastError = "Le nom de l'entité ne peut pas être vide"
        RegisterOperation = False
        Exit Function
    End If
    
    ' Vérifier si nous avons atteint la limite d'opérations
    If m_operationCount >= MAX_OPERATIONS Then
        m_lastError = "Nombre maximum d'opérations atteint"
        LogError "Limite d'opérations atteinte lors de l'enregistrement pour " & entityName
        RegisterOperation = False
        Exit Function
    End If
    
    ' Incrémenter le compteur et ajouter l'opération
    m_operationCount = m_operationCount + 1
    With m_operations(m_operationCount)
        .EntityName = entityName
        .EntityId = entityId
        .OperationType = operationType
        .OriginalState = originalState
        .NewState = newState
        .CallbackModule = callbackModule
        .CallbackMethod = callbackMethod
        .Parameters = parameters
        .IsRegistered = True
    End With
    
    ' Log de l'enregistrement
    LogInfo "Opération " & OperationTypeToString(operationType) & " enregistrée pour " & entityName
    
    RegisterOperation = True
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de l'enregistrement de l'opération: " & Err.Description
    LogError "Erreur lors de l'enregistrement de l'opération: " & Err.Description
    RegisterOperation = False
End Function

'@Description("Convertit un type d'opération en chaîne")
Private Function OperationTypeToString(ByVal opType As OperationType) As String
    Select Case opType
        Case opInsert
            OperationTypeToString = "Insertion"
        Case opUpdate
            OperationTypeToString = "Mise à jour"
        Case opDelete
            OperationTypeToString = "Suppression"
        Case opCustom
            OperationTypeToString = "Personnalisée"
        Case Else
            OperationTypeToString = "Inconnue"
    End Select
End Function

'----------------------------------------------------------------------------------------
' Propriétés
'----------------------------------------------------------------------------------------

'@Description("Retourne l'état actuel de l'unité de travail")
Public Property Get State() As UnitOfWorkState
    State = m_state
End Property

'@Description("Retourne le nombre d'opérations enregistrées")
Public Property Get OperationCount() As Long
    OperationCount = m_operationCount
End Property

'@Description("Indique si une transaction est active")
Public Property Get HasActiveTransaction() As Boolean
    HasActiveTransaction = (m_transactionDepth > 0)
End Property

'@Description("Retourne le niveau d'imbrication des transactions")
Public Property Get TransactionDepth() As Long
    TransactionDepth = m_transactionDepth
End Property

'@Description("Retourne la dernière erreur survenue")
Public Property Get LastError() As String
    LastError = m_lastError
End Property

'----------------------------------------------------------------------------------------
' Journalisation
'----------------------------------------------------------------------------------------

'@Description("Enregistre un message d'information")
Private Sub LogInfo(ByVal message As String)
    On Error Resume Next
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo MODULE_NAME, message
    Else
        Debug.Print "[INFO] [" & MODULE_NAME & "] " & message
    End If
End Sub

'@Description("Enregistre un avertissement")
Private Sub LogWarning(ByVal message As String)
    On Error Resume Next
    
    If Not m_logger Is Nothing Then
        m_logger.LogWarning MODULE_NAME, message
    Else
        Debug.Print "[WARNING] [" & MODULE_NAME & "] " & message
    End If
End Sub

'@Description("Enregistre une erreur")
Private Sub LogError(ByVal message As String)
    On Error Resume Next
    
    If Not m_logger Is Nothing Then
        m_logger.LogError MODULE_NAME, message
    Else
        Debug.Print "[ERROR] [" & MODULE_NAME & "] " & message
    End If
End Sub

'----------------------------------------------------------------------------------------
' Classe Events
'----------------------------------------------------------------------------------------

Private Sub Class_Initialize()
    m_isInitialized = False
    m_operationCount = 0
    m_state = uowIdle
    m_transactionDepth = 0
    m_lastError = ""
End Sub

Private Sub Class_Terminate()
    ' Si une transaction est active lors de la destruction, effectuer un rollback
    If m_transactionDepth > 0 Then
        RollbackTransactionInternal
    End If
End Sub 