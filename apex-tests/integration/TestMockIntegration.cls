VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestMockIntegration"

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
'@Folder("APEX.Tests.Integration")
Option Explicit

Private mockDb As MockDatabaseConnection
Private metrics As MockMetricsCollector

'@ModuleInitialize
'@Description: 
'@Param: 
'@Returns: 

Private Sub ModuleInitialize()
    'Cette méthode s'exécute une fois par module.
    Set mockDb = New MockDatabaseConnection
    Set metrics = New MockMetricsCollector
End Sub

'@ModuleCleanup
'@Description: 
'@Param: 
'@Returns: 

Private Sub ModuleCleanup()
    'Cette méthode s'exécute une fois par module.
    Set mockDb = Nothing
    Set metrics = Nothing
End Sub

'@TestInitialize
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestInitialize()
    'Cette méthode s'exécute avant chaque test.
    mockDb.Initialize
    metrics.Initialize
End Sub

'@TestCleanup
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestCleanup()
    'Cette méthode s'exécute après chaque test.
End Sub

'@TestMethod("Integration")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestNetworkResilienceWithCache()
    On Error GoTo TestFail
    
    'Configurer le mock pour simuler des problèmes réseau
    mockDb.SetErrorRate 0.3 'Taux d'erreur de 30%
    mockDb.SetLatency 500 'Latence de 500ms
    
    'Exécuter une série de requêtes
    Dim i As Long
    For i = 1 To 10
        On Error Resume Next
        mockDb.ExecuteQuery "SELECT * FROM TestTable WHERE ID = " & i
        If Err.Number <> 0 Then
            metrics.RecordMetric "query_error", 1
        Else
            metrics.RecordMetric "query_success", 1
        End If
        On Error GoTo TestFail
    Next i
    
    'Vérifier les métriques
    Assert.IsTrue metrics.GetMetricCount("query_error") > 0, "Aucune erreur détectée"
    Assert.IsTrue metrics.GetMetricCount("query_success") > 0, "Aucun succès détecté"
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    
TestFail:
    Assert.Fail "Test échoué: " & Err.Description
End Sub

'@TestMethod("Integration")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestStressTestWithMetrics()
    On Error GoTo TestFail
    
    'Configurer le mock pour le stress test
    mockDb.SetErrorRate 0.1 'Taux d'erreur de 10%
    mockDb.SetLatency 100 'Latence de 100ms
    
    'Exécuter 1000 requêtes
    Dim i As Long
    For i = 1 To 1000
        On Error Resume Next
        If i Mod 2 = 0 Then
            mockDb.ExecuteQuery "INSERT INTO TestTable (ID) VALUES (" & i & ")"
            metrics.RecordMetric "insert_operation", 1
        Else
            mockDb.ExecuteQuery "SELECT * FROM TestTable WHERE ID = " & i
            metrics.RecordMetric "select_operation", 1
        End If
        If Err.Number <> 0 Then
            metrics.RecordMetric "operation_error", 1
        End If
        On Error GoTo TestFail
    Next i
    
    'Vérifier les métriques
    Assert.IsTrue metrics.GetMetricCount("insert_operation") = 500, "Nombre incorrect d'insertions"
    Assert.IsTrue metrics.GetMetricCount("select_operation") = 500, "Nombre incorrect de sélections"
    Assert.IsTrue metrics.GetMetricCount("operation_error") > 0, "Aucune erreur détectée"
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    
TestFail:
    Assert.Fail "Test échoué: " & Err.Description
End Sub

'@TestMethod("Integration")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestConcurrencyAndDeadlockDetection()
    On Error GoTo TestFail
    
    'Configurer le mock pour simuler des deadlocks
    mockDb.SetErrorRate 0.05 'Taux d'erreur de 5%
    mockDb.SetLatency 200 'Latence de 200ms
    mockDb.EnableDeadlockSimulation True
    
    'Simuler des opérations concurrentes
    Dim i As Long
    For i = 1 To 100
        On Error Resume Next
        'Simuler deux transactions qui peuvent créer un deadlock
        mockDb.ExecuteQuery "UPDATE Table1 SET Value = " & i & " WHERE ID = 1"
        mockDb.ExecuteQuery "UPDATE Table2 SET Value = " & i & " WHERE ID = 1"
        
        If Err.Number <> 0 Then
            If InStr(1, Err.Description, "deadlock", vbTextCompare) > 0 Then
                metrics.RecordMetric "deadlock_detected", 1
            Else
                metrics.RecordMetric "other_error", 1
            End If
        End If
        On Error GoTo TestFail
    Next i
    
    'Vérifier les métriques
    Assert.IsTrue metrics.GetMetricCount("deadlock_detected") > 0, "Aucun deadlock détecté"
    
    'Vérifier le rapport de métriques
    Debug.Print metrics.GetSummary
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    
TestFail:
    Assert.Fail "Test échoué: " & Err.Description
End Sub 