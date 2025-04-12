VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTransactionManager"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("APEX.Metier.Database.Transaction")
Option Explicit

'==========================================================================
' Classe    : clsTransactionManager
' Purpose   : Gestion des transactions distribuées
' Author    : APEX Framework Team
' Date      : 2024-04-11
' Reference : DB-001
'==========================================================================

Private Type TTransactionState
    IsActive As Boolean
    HasError As Boolean
    ErrorMessage As String
    StartTime As Date
    SavepointCount As Long
    TransactionID As String
    IsolationLevel As Long
    Connections As Collection
End Type

Private Type TTransactionStats
    TotalTransactions As Long
    SuccessfulTransactions As Long
    FailedTransactions As Long
    AverageTransactionTime As Double
    LongestTransaction As Double
    CurrentTransactions As Long
End Type

Private this As TTransactionState
Private m_Stats As TTransactionStats
Private m_Logger As ILoggerBase
Private m_MetricsCollector As clsMetricsCollector

' Constantes d'isolation
Private Const TxUnspecified As Long = 0
Private Const TxReadUncommitted As Long = 1
Private Const TxReadCommitted As Long = 2
Private Const TxRepeatableRead As Long = 3
Private Const TxSerializable As Long = 4

'@Description("Initialise le gestionnaire de transactions")
'@Description: 
'@Param: 
'@Returns: 

Public Function Initialize(Optional ByVal logger As ILoggerBase = Nothing, _
                         Optional ByVal metricsCollector As clsMetricsCollector = Nothing) As Boolean
    ' Initialisation des dépendances
    Set m_Logger = logger
    Set m_MetricsCollector = metricsCollector
    
    ' Initialisation de l'état
    Set this.Connections = New Collection
    this.IsolationLevel = TxReadCommitted ' Par défaut
    
    Initialize = True
End Function

'@Description("Démarre une nouvelle transaction")
'@Description: 
'@Param: 
'@Returns: 

Public Function BeginTransaction(Optional ByVal isolationLevel As Long = -1) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier si une transaction est déjà active
    If this.IsActive Then
        LogError "BeginTransaction", "Une transaction est déjà active"
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Initialiser l'état
    this.IsActive = True
    this.HasError = False
    this.ErrorMessage = ""
    this.StartTime = Now
    this.SavepointCount = 0
    this.TransactionID = GenerateTransactionID
    
    ' Définir le niveau d'isolation
    If isolationLevel >= 0 Then
        this.IsolationLevel = isolationLevel
    End If
    
    ' Métriques
    m_Stats.CurrentTransactions = m_Stats.CurrentTransactions + 1
    m_Stats.TotalTransactions = m_Stats.TotalTransactions + 1
    
    If Not m_MetricsCollector Is Nothing Then
        m_MetricsCollector.RecordMetric "TransactionStart", 1
        m_MetricsCollector.RecordMetric "ActiveTransactions", m_Stats.CurrentTransactions
    End If
    
    LogInfo "Transaction démarrée: " & this.TransactionID
    BeginTransaction = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    LogError "BeginTransaction", Err.Description
    BeginTransaction = False
End Function

'@Description("Valide la transaction en cours")
'@Description: 
'@Param: 
'@Returns: 

Public Function CommitTransaction() As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier si une transaction est active
    If Not this.IsActive Then
        LogError "CommitTransaction", "Aucune transaction active"
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Vérifier les erreurs
    If this.HasError Then
        LogError "CommitTransaction", "La transaction contient des erreurs, rollback nécessaire"
        RollbackTransaction
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Valider sur toutes les connexions
    Dim conn As ADODB.Connection
    For Each conn In this.Connections
        conn.CommitTrans
    Next conn
    
    ' Métriques
    UpdateTransactionMetrics True
    
    ' Réinitialiser l'état
    ResetTransactionState
    
    LogInfo "Transaction validée: " & this.TransactionID
    CommitTransaction = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    LogError "CommitTransaction", Err.Description
    this.HasError = True
    this.ErrorMessage = Err.Description
    RollbackTransaction
    CommitTransaction = False
End Function

'@Description("Annule la transaction en cours")
'@Description: 
'@Param: 
'@Returns: 

Public Function RollbackTransaction() As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier si une transaction est active
    If Not this.IsActive Then
        LogError "RollbackTransaction", "Aucune transaction active"
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Annuler sur toutes les connexions
    Dim conn As ADODB.Connection
    For Each conn In this.Connections
        conn.RollbackTrans
    Next conn
    
    ' Métriques
    UpdateTransactionMetrics False
    
    ' Réinitialiser l'état
    ResetTransactionState
    
    LogInfo "Transaction annulée: " & this.TransactionID
    RollbackTransaction = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    LogError "RollbackTransaction", Err.Description
    RollbackTransaction = False
End Function

'@Description("Crée un point de sauvegarde")
'@Description: 
'@Param: 
'@Returns: 

Public Function CreateSavepoint() As String
    On Error GoTo ErrorHandler
    
    ' Vérifier si une transaction est active
    If Not this.IsActive Then
        LogError "CreateSavepoint", "Aucune transaction active"
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Générer le nom du savepoint
    this.SavepointCount = this.SavepointCount + 1
    Dim savepointName As String
    savepointName = "SP_" & this.TransactionID & "_" & this.SavepointCount
    
    ' Créer le savepoint sur toutes les connexions
    Dim conn As ADODB.Connection
    For Each conn In this.Connections
        conn.Execute "SAVEPOINT " & savepointName
    Next conn
    
    LogInfo "Point de sauvegarde créé: " & savepointName
    CreateSavepoint = savepointName
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    LogError "CreateSavepoint", Err.Description
    CreateSavepoint = ""
End Function

'@Description("Restaure un point de sauvegarde")
'@Description: 
'@Param: 
'@Returns: 

Public Function RollbackToSavepoint(ByVal savepointName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier si une transaction est active
    If Not this.IsActive Then
        LogError "RollbackToSavepoint", "Aucune transaction active"
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Restaurer le savepoint sur toutes les connexions
    Dim conn As ADODB.Connection
    For Each conn In this.Connections
        conn.Execute "ROLLBACK TO SAVEPOINT " & savepointName
    Next conn
    
    LogInfo "Restauration au point de sauvegarde: " & savepointName
    RollbackToSavepoint = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    LogError "RollbackToSavepoint", Err.Description
    RollbackToSavepoint = False
End Function

'@Description("Ajoute une connexion à la transaction")
'@Description: 
'@Param: 
'@Returns: 

Public Function AddConnection(ByVal connection As ADODB.Connection) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier si la connexion existe déjà
    Dim conn As ADODB.Connection
    For Each conn In this.Connections
        If conn Is connection Then
            Exit'@Description: 
'@Param: 
'@Returns: 

 Function
        End If
    Next conn
    
    ' Configurer la connexion
    connection.IsolationLevel = this.IsolationLevel
    connection.BeginTrans
    
    ' Ajouter la connexion
    this.Connections.Add connection
    
    AddConnection = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    LogError "AddConnection", Err.Description
    AddConnection = False
End Function

'@Description("Obtient les statistiques des transactions")
'@Description: 
'@Param: 
'@Returns: 

Public Function GetStats() As String
    Dim successRate As Double
    If m_Stats.TotalTransactions > 0 Then
        successRate = m_Stats.SuccessfulTransactions / m_Stats.TotalTransactions
    End If
    
    GetStats = "=== Statistiques Transactions ===" & vbNewLine & _
               "Total: " & m_Stats.TotalTransactions & vbNewLine & _
               "Réussies: " & m_Stats.SuccessfulTransactions & vbNewLine & _
               "Échouées: " & m_Stats.FailedTransactions & vbNewLine & _
               "Taux de succès: " & Format(successRate, "0.0%") & vbNewLine & _
               "Temps moyen: " & Format(m_Stats.AverageTransactionTime, "0.000") & " s" & vbNewLine & _
               "Plus longue: " & Format(m_Stats.LongestTransaction, "0.000") & " s" & vbNewLine & _
               "En cours: " & m_Stats.CurrentTransactions
End Function

'@Description("Génère un identifiant unique pour la transaction")
'@Description: 
'@Param: 
'@Returns: 

Private Function GenerateTransactionID() As String
    GenerateTransactionID = "TX_" & Format(Now, "yyyymmddhhnnss") & "_" & Format(Timer * 1000, "000")
End Function

'@Description("Met à jour les métriques de transaction")
'@Description: 
'@Param: 
'@Returns: 

Private Sub UpdateTransactionMetrics(ByVal success As Boolean)
    ' Calculer la durée
    Dim duration As Double
    duration = DateDiff("s", this.StartTime, Now)
    
    ' Mettre à jour les statistiques
    If success Then
        m_Stats.SuccessfulTransactions = m_Stats.SuccessfulTransactions + 1
    Else
        m_Stats.FailedTransactions = m_Stats.FailedTransactions + 1
    End If
    
    m_Stats.CurrentTransactions = m_Stats.CurrentTransactions - 1
    
    ' Moyenne mobile
    m_Stats.AverageTransactionTime = (m_Stats.AverageTransactionTime * (m_Stats.TotalTransactions - 1) + duration) / m_Stats.TotalTransactions
    
    ' Plus longue transaction
    If duration > m_Stats.LongestTransaction Then
        m_Stats.LongestTransaction = duration
    End If
    
    ' Métriques
    If Not m_MetricsCollector Is Nothing Then
        m_MetricsCollector.RecordMetric "TransactionDuration", duration
        m_MetricsCollector.RecordMetric "TransactionSuccess", IIf(success, 1, 0)
        m_MetricsCollector.RecordMetric "ActiveTransactions", m_Stats.CurrentTransactions
    End If
End Sub

'@Description("Réinitialise l'état de la transaction")
'@Description: 
'@Param: 
'@Returns: 

Private Sub ResetTransactionState()
    this.IsActive = False
    this.HasError = False
    this.ErrorMessage = ""
    this.SavepointCount = 0
    this.TransactionID = ""
    Set this.Connections = New Collection
End Sub

'@Description("Journalise une erreur")
'@Description: 
'@Param: 
'@Returns: 

Private Sub LogError(ByVal method As String, ByVal description As String)
    If Not m_Logger Is Nothing Then
        m_Logger.Log "ERROR", "TransactionManager." & method, description
    End If
End Sub

'@Description("Journalise une information")
'@Description: 
'@Param: 
'@Returns: 

Private Sub LogInfo(ByVal message As String)
    If Not m_Logger Is Nothing Then
        m_Logger.Log "INFO", "TransactionManager", message
    End If
End Sub 