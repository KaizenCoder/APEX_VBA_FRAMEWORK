VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestMockPerformanceScenarios"

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
'@Folder("APEX.Tests.Integration.Performance")
Option Explicit

Private Assert As Object
Private Fakes As Object
Private helper As MockTestHelper
Private Const TEST_BACKUP_PATH As String = "D:\Dev\Apex_VBA_FRAMEWORK\apex-tests\backup\"

'@ModuleInitialize
'@Description: 
'@Param: 
'@Returns: 

Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
'@Description: 
'@Param: 
'@Returns: 

Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestInitialize()
    Set helper = New MockTestHelper
    helper.Initialize
    helper.Performance.SetBackupPath TEST_BACKUP_PATH
End Sub

'@TestCleanup
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestCleanup()
    If Not helper Is Nothing Then
        helper.Cleanup
        Set helper = Nothing
    End If
End Sub

'@TestMethod("Performance")
'@Description("Test de charge avec montée progressive")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestLoadScaling()
    On Error GoTo TestFail
    
    'Configuration
    With helper.Performance
        .SetScalingProfile "GRADUAL"
        .SetMaxUsers 100
        .SetRampUpTime 10 'secondes
        .EnableDetailedMetrics True
    End With
    
    'Paramètres
    Dim targetRPS As Long
    targetRPS = 50 'Requêtes par seconde
    
    'Simuler montée en charge
    Dim currentUsers As Long
    Dim totalRequests As Long
    Dim startTime As Date
    startTime = Now
    
    Do While currentUsers < helper.Performance.GetMaxUsers()
        'Augmenter utilisateurs
        currentUsers = helper.Performance.GetCurrentUsers()
        
        'Simuler charge
        Dim i As Long
        For i = 1 To targetRPS
            On Error Resume Next
            
            'Exécuter requête
            helper.ExecuteWithLoad "db", "SELECT * FROM TestData WHERE ID = " & i
            
            If Err.Number = 0 Then
                helper.Metrics.RecordMetric "request_success", 1, "load_test", _
                    "users=" & currentUsers
                totalRequests = totalRequests + 1
            Else
                helper.Metrics.RecordMetric "request_error", 1, "load_test", _
                    "users=" & currentUsers & ",error=" & Err.Description
            End If
            
            On Error GoTo TestFail
        Next i
        
        'Vérifier métriques
        Dim currentRPS As Double
        currentRPS = helper.Performance.GetCurrentRPS()
        
        'Enregistrer métriques
        helper.Metrics.RecordMetric "current_rps", currentRPS, "load_test", _
            "users=" & currentUsers
        helper.Metrics.RecordMetric "response_time", _
            helper.Performance.GetAverageResponseTime(), "load_test", _
            "users=" & currentUsers
            
        'Vérifier limites
        If helper.Performance.GetMaxResponseTime() > 5000 Then 'ms
            helper.Logger.LogWarning "Temps de réponse maximum dépassé"
            Exit Do
        End If
        
        'Pause pour simulation réaliste
        helper.Performance.WaitForNextCycle
    Loop
    
    'Analyse résultats
    Dim avgRPS As Double
    Dim maxRPS As Double
    Dim avgResponseTime As Double
    Dim errorRate As Double
    
    avgRPS = helper.Metrics.GetAverageMetric("current_rps", "load_test")
    maxRPS = helper.Metrics.GetMaxMetric("current_rps", "load_test")
    avgResponseTime = helper.Metrics.GetAverageMetric("response_time", "load_test")
    errorRate = helper.Metrics.GetErrorRate("load_test")
    
    'Assertions
    Assert.IsTrue avgRPS > 0, "RPS moyen devrait être positif"
    Assert.IsTrue errorRate < 0.1, "Taux d'erreur devrait être < 10%"
    
    'Rapport
    helper.Logger.LogInfo "Test de charge terminé:" & vbNewLine & _
                         "- RPS moyen: " & avgRPS & vbNewLine & _
                         "- RPS max: " & maxRPS & vbNewLine & _
                         "- Temps réponse moyen: " & avgResponseTime & "ms" & vbNewLine & _
                         "- Taux d'erreur: " & Format(errorRate, "0.00%")
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Performance")
'@Description("Test de stress avec pics de charge")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestStressSpikes()
    On Error GoTo TestFail
    
    'Configuration
    With helper.Performance
        .SetStressProfile "SPIKES"
        .SetBaseline 50 'utilisateurs
        .SetSpikeFactor 5 'multiplicateur pour pics
        .SetSpikeDuration 5 'secondes
        .EnableResourceMonitoring True
    End With
    
    'Paramètres
    Const NUM_SPIKES As Long = 3
    Dim spikeInterval As Long
    spikeInterval = 15 'secondes
    
    'Exécuter test de stress
    Dim currentSpike As Long
    For currentSpike = 1 To NUM_SPIKES
        'Phase normale
        helper.Logger.LogInfo "Phase normale #" & currentSpike
        ExecuteNormalLoad 10
        
        'Phase de pic
        helper.Logger.LogInfo "Pic de charge #" & currentSpike
        ExecuteSpikeLoad 5
        
        'Récupération
        helper.Logger.LogInfo "Phase de récupération #" & currentSpike
        helper.Performance.WaitForRecovery
        
        'Vérifier santé système
        If Not helper.Performance.IsSystemHealthy Then
            helper.Logger.LogError "Santé système compromise après pic #" & currentSpike
            Exit For
        End If
    Next currentSpike
    
    'Analyse résultats
    Dim baselineMetrics As Collection
    Dim spikeMetrics As Collection
    Set baselineMetrics = helper.Performance.GetBaselineMetrics()
    Set spikeMetrics = helper.Performance.GetSpikeMetrics()
    
    'Vérifier dégradation
    Dim baselineResponseTime As Double
    Dim spikeResponseTime As Double
    Dim degradationFactor As Double
    
    baselineResponseTime = baselineMetrics("avg_response_time")
    spikeResponseTime = spikeMetrics("avg_response_time")
    degradationFactor = spikeResponseTime / baselineResponseTime
    
    'Assertions
    Assert.IsTrue degradationFactor < 10, _
        "Dégradation performance excessive: " & degradationFactor & "x"
    
    'Rapport
    helper.Logger.LogInfo "Test de stress terminé:" & vbNewLine & _
                         "- Pics exécutés: " & currentSpike & "/" & NUM_SPIKES & vbNewLine & _
                         "- Temps réponse base: " & baselineResponseTime & "ms" & vbNewLine & _
                         "- Temps réponse pic: " & spikeResponseTime & "ms" & vbNewLine & _
                         "- Facteur dégradation: " & Format(degradationFactor, "0.00") & "x"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Performance")
'@Description("Test de benchmark des opérations critiques")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestCriticalOperationsBenchmark()
    On Error GoTo TestFail
    
    'Configuration
    With helper.Performance
        .SetBenchmarkMode True
        .EnableProfiling True
        .SetWarmupIterations 3
        .SetBenchmarkIterations 10
    End With
    
    'Définir opérations critiques
    Dim operations As Collection
    Set operations = New Collection
    With operations
        .Add Array("db_query", "SELECT * FROM LargeTable WHERE Status = 'Active'")
        .Add Array("db_insert", "INSERT INTO AuditLog (Timestamp, Action) VALUES (Now, 'Test')")
        .Add Array("cache_read", "GET user_preferences_123")
        .Add Array("cache_write", "SET session_token_456")
        .Add Array("compute", "CALCULATE_METRICS")
    End With
    
    'Exécuter benchmarks
    Dim op As Variant
    For Each op In operations
        Dim opName As String
        Dim opCommand As String
        opName = op(0)
        opCommand = op(1)
        
        'Phase d'échauffement
        helper.Logger.LogInfo "Échauffement: " & opName
        helper.Performance.StartWarmup opName
        
        Dim w As Long
        For w = 1 To helper.Performance.GetWarmupIterations()
            On Error Resume Next
            helper.ExecuteWithBenchmark opName, opCommand
            On Error GoTo TestFail
        Next w
        
        helper.Performance.EndWarmup opName
        
        'Benchmark
        helper.Logger.LogInfo "Benchmark: " & opName
        helper.Performance.StartBenchmark opName
        
        Dim b As Long
        For b = 1 To helper.Performance.GetBenchmarkIterations()
            On Error Resume Next
            helper.ExecuteWithBenchmark opName, opCommand
            
            If Err.Number = 0 Then
                helper.Metrics.RecordMetric "bench_success", 1, "benchmark_test", _
                    "op=" & opName & ",iter=" & b
            Else
                helper.Metrics.RecordMetric "bench_error", 1, "benchmark_test", _
                    "op=" & opName & ",iter=" & b & ",error=" & Err.Description
            End If
            
            On Error GoTo TestFail
        Next b
        
        helper.Performance.EndBenchmark opName
        
        'Analyser résultats
        Dim benchResult As Dictionary
        Set benchResult = helper.Performance.GetBenchmarkResults(opName)
        
        'Vérifier performances
        Dim avgTime As Double
        Dim maxTime As Double
        avgTime = benchResult("average_time")
        maxTime = benchResult("max_time")
        
        'Vérifier seuils
        Select Case opName
            Case "db_query"
                Assert.IsTrue avgTime < 100, "Query trop lente: " & avgTime & "ms"
            Case "cache_read"
                Assert.IsTrue avgTime < 10, "Cache read trop lent: " & avgTime & "ms"
            Case "compute"
                Assert.IsTrue avgTime < 500, "Calcul trop lent: " & avgTime & "ms"
        End Select
        
        'Enregistrer résultats
        helper.Performance.SaveBenchmarkResults opName, _
            TEST_BACKUP_PATH & "benchmarks\" & opName & "_results.json"
    Next op
    
    'Rapport final
    helper.Performance.GenerateBenchmarkReport _
        TEST_BACKUP_PATH & "benchmarks\benchmark_report.html"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ExecuteNormalLoad(ByVal durationSeconds As Long)
    Dim endTime As Date
    endTime = DateAdd("s", durationSeconds, Now)
    
    Do While Now < endTime
        On Error Resume Next
        helper.ExecuteWithLoad "normal", "NORMAL_OPERATION"
        
        If Err.Number = 0 Then
            helper.Metrics.RecordMetric "normal_success", 1, "stress_test"
        Else
            helper.Metrics.RecordMetric "normal_error", 1, "stress_test"
        End If
        
        On Error GoTo 0
        helper.Performance.WaitForNextCycle
    Loop
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ExecuteSpikeLoad(ByVal durationSeconds As Long)
    Dim endTime As Date
    endTime = DateAdd("s", durationSeconds, Now)
    
    Do While Now < endTime
        On Error Resume Next
        helper.ExecuteWithLoad "spike", "SPIKE_OPERATION"
        
        If Err.Number = 0 Then
            helper.Metrics.RecordMetric "spike_success", 1, "stress_test"
        Else
            helper.Metrics.RecordMetric "spike_error", 1, "stress_test"
        End If
        
        On Error GoTo 0
        helper.Performance.WaitForNextCycle
    Loop
End Sub 