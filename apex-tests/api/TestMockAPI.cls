VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestMockAPI"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@TestModule
'@Folder("APEX.Tests.API")
Option Explicit

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
'@Description: 
'@Param: 
'@Returns: 

Private Sub ModuleInitialize()
    'Cette méthode s'exécute une fois par module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
'@Description: 
'@Param: 
'@Returns: 

Private Sub ModuleCleanup()
    'Cette méthode s'exécute une fois par module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestMethod("API")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestMockTestHelperInitialization()
    On Error GoTo TestFail
    
    'Arrange
    Dim helper As New MockTestHelper
    
    'Act
    helper.Initialize
    
    'Assert
    Assert.IsNotNothing helper.Database, "Database devrait être initialisé"
    Assert.IsNotNothing helper.Network, "Network devrait être initialisé"
    Assert.IsNotNothing helper.Metrics, "Metrics devrait être initialisé"
    Assert.IsNotNothing helper.Logger, "Logger devrait être initialisé"
    
    'Cleanup
    helper.Cleanup
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("API")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestMockDatabaseConnection()
    On Error GoTo TestFail
    
    'Arrange
    Dim db As New MockDatabaseConnection
    db.Initialize
    
    'Act & Assert
    'Test SetErrorRate
    db.SetErrorRate 0.5
    
    'Test SetLatency
    db.SetLatency 100
    
    'Test EnableDeadlockSimulation
    db.EnableDeadlockSimulation True
    
    'Test ExecuteQuery
    On Error Resume Next
    db.ExecuteQuery "SELECT 1"
    Assert.IsTrue Err.Number = 0 Or Err.Number = ERR_DB_ERROR, _
                 "Devrait réussir ou échouer avec ERR_DB_ERROR"
    On Error GoTo TestFail
    
    'Test GetStats
    Dim stats As String
    stats = db.GetStats
    Assert.IsTrue Len(stats) > 0, "Les statistiques ne devraient pas être vides"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("API")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestMockNetworkConnection()
    On Error GoTo TestFail
    
    'Arrange
    Dim network As New MockNetworkConnection
    Dim logger As New MockLogger
    Dim metrics As New MockMetricsCollector
    
    'Act
    network.Initialize logger, metrics
    
    'Configure
    network.Configure 100, 0.1, 0.1, 1000
    
    'Test SendRequest
    On Error Resume Next
    network.SendRequest "GET /api/test"
    Assert.IsTrue Err.Number = 0 Or Err.Number = ERR_NETWORK_ERROR, _
                 "Devrait réussir ou échouer avec ERR_NETWORK_ERROR"
    On Error GoTo TestFail
    
    'Test GetStats
    Dim stats As String
    stats = network.GetStats
    Assert.IsTrue Len(stats) > 0, "Les statistiques ne devraient pas être vides"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("API")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestMockMetricsCollector()
    On Error GoTo TestFail
    
    'Arrange
    Dim metrics As New MockMetricsCollector
    metrics.Initialize
    
    'Act
    'Test RecordMetric
    metrics.RecordMetric "test_metric", 42, "test_category", "tag1=value1"
    metrics.RecordMetric "test_metric", 44, "test_category", "tag1=value1"
    
    'Test GetAverageMetric
    Dim avg As Double
    avg = metrics.GetAverageMetric("test_metric", "test_category")
    Assert.AreEqual 43, avg, "La moyenne devrait être 43"
    
    'Test GetTotalMetric
    Dim total As Double
    total = metrics.GetTotalMetric("test_metric", "test_category")
    Assert.AreEqual 86, total, "Le total devrait être 86"
    
    'Test EnableMetricAggregation
    metrics.EnableMetricAggregation False
    On Error Resume Next
    metrics.GetAverageMetric "test_metric"
    Assert.IsTrue Err.Number <> 0, "Devrait échouer quand l'agrégation est désactivée"
    On Error GoTo TestFail
    
    'Test GetMetricsReport
    Dim report As String
    report = metrics.GetMetricsReport
    Assert.IsTrue InStr(report, "test_metric") > 0, "Le rapport devrait contenir la métrique"
    Assert.IsTrue InStr(report, "test_category") > 0, "Le rapport devrait contenir la catégorie"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("API")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestMockLogger()
    On Error GoTo TestFail
    
    'Arrange
    Dim logger As New MockLogger
    
    'Act
    'Test différents niveaux de log
    logger.LogInfo "Test info"
    logger.LogDebug "Test debug"
    logger.LogError "Test error"
    
    'Test ContainsMessage
    Assert.IsTrue logger.ContainsMessage("Test info"), "Devrait contenir le message info"
    Assert.IsTrue logger.ContainsMessage("Test debug"), "Devrait contenir le message debug"
    Assert.IsTrue logger.ContainsMessage("Test error"), "Devrait contenir le message error"
    
    'Test GetLogs
    Dim logs As String
    logs = logger.GetLogs
    Assert.IsTrue InStr(logs, "Test info") > 0, "Les logs devraient contenir le message info"
    Assert.IsTrue InStr(logs, "Test debug") > 0, "Les logs devraient contenir le message debug"
    Assert.IsTrue InStr(logs, "Test error") > 0, "Les logs devraient contenir le message error"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("API")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestPerformanceConfiguration()
    On Error GoTo TestFail
    
    'Arrange
    Dim helper As New MockTestHelper
    helper.Initialize
    
    'Act
    helper.ConfigurePerformanceTest True, "perf_test", 2, 5
    
    'Assert
    'Exécuter une opération et vérifier les métriques
    helper.ExecuteWithMetrics "db", "SELECT 1"
    
    Dim report As String
    report = helper.Metrics.GetMetricsReport
    
    Assert.IsTrue InStr(report, "perf_test") > 0, "Le rapport devrait contenir la catégorie"
    
    'Cleanup
    helper.Cleanup
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("API")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestStressConfiguration()
    On Error GoTo TestFail
    
    'Arrange
    Dim helper As New MockTestHelper
    helper.Initialize
    
    'Act
    helper.ConfigureStressTest 0.2, 100, 200, 0.1, 0.1, 1000, 50
    
    'Assert
    'Exécuter quelques opérations
    Dim i As Long
    Dim errorCount As Long
    
    For i = 1 To 10
        On Error Resume Next
        helper.ExecuteWithMetrics "db", "SELECT " & i
        If Err.Number <> 0 Then errorCount = errorCount + 1
        On Error GoTo TestFail
    Next i
    
    'Vérifier que le taux d'erreur est proche de la configuration
    Dim errorRate As Double
    errorRate = errorCount / 10
    
    'Tolérance de ±30% pour le taux d'erreur (nature aléatoire)
    Assert.IsTrue Abs(errorRate - 0.2) <= 0.06, _
                 "Le taux d'erreur devrait être proche de 20%"
    
    'Cleanup
    helper.Cleanup
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("API")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestErrorHandling()
    On Error GoTo TestFail
    
    'Arrange
    Dim helper As New MockTestHelper
    helper.Initialize
    
    'Act & Assert
    'Test paramètres invalides
    On Error Resume Next
    
    'Taux d'erreur invalide
    helper.Database.SetErrorRate 2
    Assert.IsTrue Err.Number = ERR_INVALID_PARAMETER, _
                 "Devrait échouer avec ERR_INVALID_PARAMETER"
    
    'Latence négative
    helper.Database.SetLatency -1
    Assert.IsTrue Err.Number = ERR_INVALID_PARAMETER, _
                 "Devrait échouer avec ERR_INVALID_PARAMETER"
    
    'Timeout invalide
    helper.Network.Configure 100, 0.1, 0.1, -1
    Assert.IsTrue Err.Number = ERR_INVALID_PARAMETER, _
                 "Devrait échouer avec ERR_INVALID_PARAMETER"
    
    On Error GoTo TestFail
    
    'Cleanup
    helper.Cleanup
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("API")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestMetricsAggregation()
    On Error GoTo TestFail
    
    'Arrange
    Dim metrics As New MockMetricsCollector
    metrics.Initialize
    
    'Act
    'Enregistrer des métriques dans différentes catégories
    metrics.RecordMetric "test", 10, "cat1"
    metrics.RecordMetric "test", 20, "cat1"
    metrics.RecordMetric "test", 30, "cat2"
    metrics.RecordMetric "test", 40, "cat2"
    
    'Assert
    'Vérifier les moyennes par catégorie
    Dim avgCat1 As Double
    avgCat1 = metrics.GetAverageMetric("test", "cat1")
    Assert.AreEqual 15, avgCat1, "Moyenne cat1 devrait être 15"
    
    Dim avgCat2 As Double
    avgCat2 = metrics.GetAverageMetric("test", "cat2")
    Assert.AreEqual 35, avgCat2, "Moyenne cat2 devrait être 35"
    
    'Vérifier la moyenne globale
    Dim avgTotal As Double
    avgTotal = metrics.GetAverageMetric("test")
    Assert.AreEqual 25, avgTotal, "Moyenne totale devrait être 25"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("API")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestReportGeneration()
    On Error GoTo TestFail
    
    'Arrange
    Dim helper As New MockTestHelper
    helper.Initialize
    
    'Act
    'Générer quelques données
    helper.Logger.LogInfo "Test info"
    helper.Logger.LogError "Test error"
    
    helper.ExecuteWithMetrics "db", "SELECT 1"
    helper.ExecuteWithMetrics "network", "GET /api/test"
    
    'Assert
    Dim report As String
    report = helper.GenerateReport
    
    'Vérifier les sections du rapport
    Assert.IsTrue InStr(report, "Configuration:") > 0, "Devrait contenir la configuration"
    Assert.IsTrue InStr(report, "Métriques:") > 0, "Devrait contenir les métriques"
    Assert.IsTrue InStr(report, "Base de données:") > 0, "Devrait contenir les stats DB"
    Assert.IsTrue InStr(report, "Réseau:") > 0, "Devrait contenir les stats réseau"
    
    'Cleanup
    helper.Cleanup
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub 