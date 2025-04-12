VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestMockMetricsCollector"

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
'@Folder("APEX.Tests.Mocks")
Option Explicit

Private Assert As Object
Private Fakes As Object
Private metrics As MockMetricsCollector

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

'@TestInitialize
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestInitialize()
    'Cette méthode s'exécute avant chaque test.
    Set metrics = New MockMetricsCollector
    metrics.Initialize
End Sub

'@TestCleanup
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestCleanup()
    'Cette méthode s'exécute après chaque test.
    Set metrics = Nothing
End Sub

'@TestMethod("Metrics")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestRecordMetric()
    On Error GoTo TestFail
    
    'Arrange
    Const METRIC_NAME As String = "test_metric"
    Const METRIC_VALUE As Double = 42
    Const CATEGORY As String = "test_category"
    Const TAGS As String = "tag1,tag2"
    
    'Act
    metrics.RecordMetric METRIC_NAME, METRIC_VALUE, CATEGORY, TAGS
    
    'Assert
    Dim total As Double
    total = metrics.GetTotalMetric(METRIC_NAME, CATEGORY)
    Assert.AreEqual METRIC_VALUE, total, "La valeur totale devrait correspondre"
    
    'Vérifier le rapport
    Dim report As String
    report = metrics.GetMetricsReport()
    Assert.IsTrue InStr(report, METRIC_NAME) > 0, "Le nom de la métrique devrait apparaître"
    Assert.IsTrue InStr(report, CATEGORY) > 0, "La catégorie devrait apparaître"
    Assert.IsTrue InStr(report, TAGS) > 0, "Les tags devraient apparaître"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Metrics")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestAggregation()
    On Error GoTo TestFail
    
    'Arrange
    Const METRIC_NAME As String = "test_metric"
    Const CATEGORY As String = "test_category"
    
    'Act
    metrics.RecordMetric METRIC_NAME, 10, CATEGORY
    metrics.RecordMetric METRIC_NAME, 20, CATEGORY
    metrics.RecordMetric METRIC_NAME, 30, CATEGORY
    
    'Assert
    Dim avg As Double
    avg = metrics.GetAverageMetric(METRIC_NAME, CATEGORY)
    Assert.AreEqual 20, avg, "La moyenne devrait être 20"
    
    Dim total As Double
    total = metrics.GetTotalMetric(METRIC_NAME, CATEGORY)
    Assert.AreEqual 60, total, "Le total devrait être 60"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Metrics")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestDisabledAggregation()
    On Error GoTo TestFail
    
    'Arrange
    metrics.EnableMetricAggregation False
    
    'Act & Assert
    On Error Resume Next
    metrics.GetAverageMetric "test"
    
    'Vérifier que l'erreur est levée
    Assert.AreEqual 5, Err.Number, "Une erreur devrait être levée"
    Assert.IsTrue InStr(Err.Description, "agrégation est désactivée") > 0, _
                 "Le message d'erreur devrait mentionner l'agrégation"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Metrics")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestReset()
    On Error GoTo TestFail
    
    'Arrange
    metrics.RecordMetric "metric1", 10
    metrics.RecordMetric "metric2", 20
    
    'Act
    metrics.Reset
    
    'Assert
    Dim report As String
    report = metrics.GetMetricsReport()
    
    Assert.IsTrue InStr(report, "Total des entrées: 0") > 0, "Le compteur devrait être à zéro"
    Assert.IsTrue InStr(report, "Catégories: 0") > 0, "Les catégories devraient être vides"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Metrics")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestMultipleCategories()
    On Error GoTo TestFail
    
    'Arrange
    Const METRIC_NAME As String = "test_metric"
    
    'Act
    metrics.RecordMetric METRIC_NAME, 10, "cat1"
    metrics.RecordMetric METRIC_NAME, 20, "cat2"
    metrics.RecordMetric METRIC_NAME, 30, "cat1"
    
    'Assert
    Dim avgCat1 As Double
    avgCat1 = metrics.GetAverageMetric(METRIC_NAME, "cat1")
    Assert.AreEqual 20, avgCat1, "La moyenne pour cat1 devrait être 20"
    
    Dim avgCat2 As Double
    avgCat2 = metrics.GetAverageMetric(METRIC_NAME, "cat2")
    Assert.AreEqual 20, avgCat2, "La moyenne pour cat2 devrait être 20"
    
    Dim avgTotal As Double
    avgTotal = metrics.GetAverageMetric(METRIC_NAME)
    Assert.AreEqual 20, avgTotal, "La moyenne totale devrait être 20"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub 