VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockTestHelper"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("APEX.Tests.Utils")
Option Explicit

'@Description("Configuration pour les tests de charge")
Private Type TStressConfig
    DbErrorRate As Double
    DbLatency As Long
    NetworkLatency As Long
    NetworkPacketLoss As Double
    NetworkErrorRate As Double
    NetworkTimeout As Long
    IterationCount As Long
End Type

'@Description("Configuration pour les tests de performance")
Private Type TPerformanceConfig
    EnableMetrics As Boolean
    MetricsCategory As String
    SamplingRate As Long
    ReportInterval As Long
End Type

Private Type TTestContext
    Db As MockDatabaseConnection
    Network As MockNetworkConnection
    Metrics As MockMetricsCollector
    Logger As MockLogger
    StressConfig As TStressConfig
    PerformanceConfig As TPerformanceConfig
End Type

Private this As TTestContext

'@Description("Initialise le contexte de test avec configuration par défaut")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Initialize()
    'Créer les instances
    Set this.Db = New MockDatabaseConnection
    Set this.Network = New MockNetworkConnection
    Set this.Metrics = New MockMetricsCollector
    Set this.Logger = New MockLogger
    
    'Configuration par défaut pour les tests de charge
    With this.StressConfig
        .DbErrorRate = 0.1
        .DbLatency = 50
        .NetworkLatency = 100
        .NetworkPacketLoss = 0.05
        .NetworkErrorRate = 0.05
        .NetworkTimeout = 1000
        .IterationCount = 100
    End With
    
    'Configuration par défaut pour les métriques
    With this.PerformanceConfig
        .EnableMetrics = True
        .MetricsCategory = "test"
        .SamplingRate = 1
        .ReportInterval = 10
    End With
    
    'Initialiser les composants
    this.Db.Initialize
    this.Network.Initialize this.Logger, this.Metrics
    this.Metrics.Initialize
End Sub

'@Description("Configure les paramètres pour un test de charge")
'@Description: 
'@Param: 
'@Returns: 

Public Sub ConfigureStressTest(Optional ByVal dbErrorRate As Double = 0.2, _
                             Optional ByVal dbLatency As Long = 100, _
                             Optional ByVal networkLatency As Long = 200, _
                             Optional ByVal networkPacketLoss As Double = 0.1, _
                             Optional ByVal networkErrorRate As Double = 0.1, _
                             Optional ByVal networkTimeout As Long = 1000, _
                             Optional ByVal iterationCount As Long = 100)
    With this.StressConfig
        .DbErrorRate = dbErrorRate
        .DbLatency = dbLatency
        .NetworkLatency = networkLatency
        .NetworkPacketLoss = networkPacketLoss
        .NetworkErrorRate = networkErrorRate
        .NetworkTimeout = networkTimeout
        .IterationCount = iterationCount
    End With
    
    'Appliquer la configuration
    this.Db.SetErrorRate dbErrorRate
    this.Db.SetLatency dbLatency
    this.Network.Configure networkLatency, networkPacketLoss, networkErrorRate, networkTimeout
End Sub

'@Description("Configure les paramètres pour un test de performance")
'@Description: 
'@Param: 
'@Returns: 

Public Sub ConfigurePerformanceTest(Optional ByVal enableMetrics As Boolean = True, _
                                  Optional ByVal category As String = "performance", _
                                  Optional ByVal samplingRate As Long = 1, _
                                  Optional ByVal reportInterval As Long = 10)
    With this.PerformanceConfig
        .EnableMetrics = enableMetrics
        .MetricsCategory = category
        .SamplingRate = samplingRate
        .ReportInterval = reportInterval
    End With
End Sub

'@Description("Exécute une opération avec mesure de performance")
'@Description: 
'@Param: 
'@Returns: 

Public Function ExecuteWithMetrics(ByVal operation As String, _
                                 ByVal params As String) As Variant
    If Not this.PerformanceConfig.EnableMetrics Then
        ExecuteWithMetrics = ExecuteOperation(operation, params)
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    'Mesurer le temps d'exécution
    Dim startTime As Date
    startTime = Now
    
    On Error GoTo ErrorHandler
    
    'Exécuter l'opération
    ExecuteWithMetrics = ExecuteOperation(operation, params)
    
    'Enregistrer les métriques
    Dim duration As Double
    duration = DateDiff("s", startTime, Now)
    
    this.Metrics.RecordMetric operation & "_duration", _
                             duration, _
                             this.PerformanceConfig.MetricsCategory, _
                             "operation=" & operation & ",params=" & params
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    'Enregistrer l'erreur
    this.Metrics.RecordMetric operation & "_error", 1, _
                             this.PerformanceConfig.MetricsCategory, _
                             "error=" & Err.Number & ",description=" & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'@Description("Exécute une opération spécifique")
'@Description: 
'@Param: 
'@Returns: 

Private Function ExecuteOperation(ByVal operation As String, _
                                ByVal params As String) As Variant
    Select Case LCase$(operation)
        Case "db"
            ExecuteOperation = this.Db.ExecuteQuery(params)
            
        Case "network"
            ExecuteOperation = this.Network.SendRequest(params)
            
        Case Else
            Err.Raise 5, "MockTestHelper", "Opération non supportée: " & operation
    End Select
End Function

'@Description("Génère un rapport complet des tests")
'@Description: 
'@Param: 
'@Returns: 

Public Function GenerateReport() As String
    Dim report As String
    
    report = "=== Rapport de Test ===" & vbNewLine & vbNewLine
    
    'Configuration
    report = report & "Configuration:" & vbNewLine
    With this.StressConfig
        report = report & "- DB Error Rate: " & .DbErrorRate & vbNewLine
        report = report & "- DB Latency: " & .DbLatency & "ms" & vbNewLine
        report = report & "- Network Latency: " & .NetworkLatency & "ms" & vbNewLine
        report = report & "- Network Packet Loss: " & .NetworkPacketLoss & vbNewLine
        report = report & "- Network Error Rate: " & .NetworkErrorRate & vbNewLine
        report = report & "- Network Timeout: " & .NetworkTimeout & "ms" & vbNewLine
        report = report & "- Iterations: " & .IterationCount & vbNewLine
    End With
    
    report = report & vbNewLine & "Métriques:" & vbNewLine
    report = report & this.Metrics.GetMetricsReport
    
    report = report & vbNewLine & "Base de données:" & vbNewLine
    report = report & this.Db.GetStats
    
    report = report & vbNewLine & "Réseau:" & vbNewLine
    report = report & this.Network.GetStats
    
    GenerateReport = report
End Function

'@Description("Nettoie le contexte de test")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Cleanup()
    Set this.Db = Nothing
    Set this.Network = Nothing
    Set this.Metrics = Nothing
    Set this.Logger = Nothing
End Sub

'@Description("Accesseurs des composants mock")
Public Property Get Database() As MockDatabaseConnection
    Set Database = this.Db
End Property

Public Property Get Network() As MockNetworkConnection
    Set Network = this.Network
End Property

Public Property Get Metrics() As MockMetricsCollector
    Set Metrics = this.Metrics
End Property

Public Property Get Logger() As MockLogger
    Set Logger = this.Logger
End Property 