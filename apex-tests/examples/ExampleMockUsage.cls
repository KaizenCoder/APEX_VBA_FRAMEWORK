VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExampleMockUsage"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("APEX.Tests.Examples")
Option Explicit

'@Description("Exemple 1: Test simple avec métriques")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Example1_SimpleMetrics()
    'Créer et initialiser le helper
    Dim helper As New MockTestHelper
    helper.Initialize
    
    'Configurer pour un test simple
    helper.ConfigurePerformanceTest True, "example1", 1, 5
    
    'Exécuter quelques opérations
    helper.ExecuteWithMetrics "db", "SELECT * FROM Users"
    helper.ExecuteWithMetrics "db", "SELECT * FROM Orders"
    helper.ExecuteWithMetrics "network", "GET /api/status"
    
    'Afficher le rapport
    Debug.Print helper.GenerateReport
    
    'Nettoyer
    helper.Cleanup
End Sub

'@Description("Exemple 2: Test de charge avec erreurs")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Example2_StressTest()
    'Créer et initialiser le helper
    Dim helper As New MockTestHelper
    helper.Initialize
    
    'Configurer pour un test de charge
    helper.ConfigureStressTest _
        dbErrorRate:=0.2, _
        dbLatency:=100, _
        networkLatency:=200, _
        networkPacketLoss:=0.1, _
        networkErrorRate:=0.1, _
        networkTimeout:=1000, _
        iterationCount:=50
        
    'Exécuter les opérations avec gestion d'erreurs
    Dim i As Long
    For i = 1 To 50
        On Error Resume Next
        
        'Opération DB
        helper.ExecuteWithMetrics "db", "SELECT * FROM Table" & i
        If Err.Number <> 0 Then
            helper.Metrics.RecordMetric "custom_error", 1, "example2", "source=db"
        End If
        
        'Opération réseau
        helper.ExecuteWithMetrics "network", "GET /api/data/" & i
        If Err.Number <> 0 Then
            helper.Metrics.RecordMetric "custom_error", 1, "example2", "source=network"
        End If
        
        On Error GoTo 0
    Next i
    
    'Afficher le rapport
    Debug.Print helper.GenerateReport
    
    'Nettoyer
    helper.Cleanup
End Sub

'@Description("Exemple 3: Test de performance avec catégories")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Example3_PerformanceCategories()
    'Créer et initialiser le helper
    Dim helper As New MockTestHelper
    helper.Initialize
    
    'Configurer pour un test de performance
    helper.ConfigurePerformanceTest True, "example3", 1, 1
    
    'Test des différentes catégories d'opérations
    Dim categories As Variant
    categories = Array("users", "orders", "products")
    
    Dim category As Variant
    For Each category In categories
        'Simuler des opérations DB
        helper.Database.SetLatency 50 * (1 + Rnd)
        helper.ExecuteWithMetrics "db", "SELECT * FROM " & category
        helper.Metrics.RecordMetric "db_operation", 1, CStr(category)
        
        'Simuler des opérations réseau
        helper.Network.Configure 100 * (1 + Rnd), 0.05, 0.05, 1000
        helper.ExecuteWithMetrics "network", "GET /api/" & category
        helper.Metrics.RecordMetric "network_operation", 1, CStr(category)
    Next category
    
    'Afficher les statistiques par catégorie
    Dim report As String
    report = "=== Statistiques par Catégorie ===" & vbNewLine
    
    For Each category In categories
        report = report & vbNewLine & "Catégorie: " & category & vbNewLine
        report = report & "- DB Ops: " & _
                 helper.Metrics.GetTotalMetric("db_operation", CStr(category)) & vbNewLine
        report = report & "- Network Ops: " & _
                 helper.Metrics.GetTotalMetric("network_operation", CStr(category)) & vbNewLine
    Next category
    
    Debug.Print report
    Debug.Print helper.GenerateReport
    
    'Nettoyer
    helper.Cleanup
End Sub

'@Description("Exemple 4: Test avec logging personnalisé")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Example4_CustomLogging()
    'Créer et initialiser le helper
    Dim helper As New MockTestHelper
    helper.Initialize
    
    'Configurer le test
    helper.ConfigurePerformanceTest True, "example4"
    
    'Simuler différents niveaux de log
    helper.Logger.LogInfo "Démarrage du test"
    helper.Logger.LogDebug "Configuration chargée"
    
    On Error Resume Next
    
    'Simuler une erreur DB
    helper.Database.SetErrorRate 1 'Force l'erreur
    helper.ExecuteWithMetrics "db", "SELECT * FROM NonExistentTable"
    If Err.Number <> 0 Then
        helper.Logger.LogError "Erreur DB: " & Err.Description
        helper.Metrics.RecordMetric "error_count", 1, "example4", "type=db"
    End If
    
    'Simuler une erreur réseau
    helper.Network.Configure 100, 0, 1, 1000 'Force l'erreur
    helper.ExecuteWithMetrics "network", "GET /api/invalid"
    If Err.Number <> 0 Then
        helper.Logger.LogError "Erreur Réseau: " & Err.Description
        helper.Metrics.RecordMetric "error_count", 1, "example4", "type=network"
    End If
    
    On Error GoTo 0
    
    'Afficher les logs et métriques
    Debug.Print "=== Logs ===" & vbNewLine
    Debug.Print helper.Logger.GetLogs
    Debug.Print vbNewLine & "=== Métriques ===" & vbNewLine
    Debug.Print helper.Metrics.GetMetricsReport
    
    'Nettoyer
    helper.Cleanup
End Sub

'@Description("Exemple 5: Test complet avec tous les composants")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Example5_CompleteTest()
    'Créer et initialiser le helper
    Dim helper As New MockTestHelper
    helper.Initialize
    
    'Configuration complète
    With helper
        'Configuration DB
        .Database.SetErrorRate 0.1
        .Database.SetLatency 75
        .Database.EnableDeadlockSimulation True
        
        'Configuration réseau
        .Network.Configure 150, 0.05, 0.05, 2000
        
        'Configuration métriques
        .ConfigurePerformanceTest True, "example5", 1, 5
        
        'Logging initial
        .Logger.LogInfo "Démarrage du test complet"
    End With
    
    'Exécution des opérations
    Dim operations As Variant
    operations = Array("create", "read", "update", "delete")
    
    Dim op As Variant
    For Each op In operations
        'Logger l'opération
        helper.Logger.LogDebug "Exécution de l'opération: " & op
        
        On Error Resume Next
        
        'Opération DB
        helper.ExecuteWithMetrics "db", "EXEC sp_" & op & "_data"
        If Err.Number <> 0 Then
            helper.Logger.LogError "Erreur DB (" & op & "): " & Err.Description
            helper.Metrics.RecordMetric "operation_error", 1, "example5", _
                                      "type=db,operation=" & op
        Else
            helper.Metrics.RecordMetric "operation_success", 1, "example5", _
                                      "type=db,operation=" & op
        End If
        
        'Opération réseau
        helper.ExecuteWithMetrics "network", op & " /api/data"
        If Err.Number <> 0 Then
            helper.Logger.LogError "Erreur Réseau (" & op & "): " & Err.Description
            helper.Metrics.RecordMetric "operation_error", 1, "example5", _
                                      "type=network,operation=" & op
        Else
            helper.Metrics.RecordMetric "operation_success", 1, "example5", _
                                      "type=network,operation=" & op
        End If
        
        On Error GoTo 0
    Next op
    
    'Analyse des résultats
    Dim successRate As Double
    Dim totalOps As Double
    Dim totalErrors As Double
    
    totalOps = helper.Metrics.GetTotalMetric("operation_success", "example5") + _
               helper.Metrics.GetTotalMetric("operation_error", "example5")
    totalErrors = helper.Metrics.GetTotalMetric("operation_error", "example5")
    
    If totalOps > 0 Then
        successRate = (totalOps - totalErrors) / totalOps * 100
        helper.Logger.LogInfo "Taux de succès: " & Format(successRate, "0.00") & "%"
    End If
    
    'Afficher le rapport complet
    Debug.Print "=== Rapport Complet ===" & vbNewLine
    Debug.Print "Logs:" & vbNewLine
    Debug.Print helper.Logger.GetLogs & vbNewLine
    Debug.Print "Rapport détaillé:" & vbNewLine
    Debug.Print helper.GenerateReport
    
    'Nettoyer
    helper.Cleanup
End Sub 