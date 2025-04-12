VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestMockIntegrationScenarios"

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

Private Assert As Object
Private Fakes As Object
Private helper As MockTestHelper

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
    Set helper = New MockTestHelper
    helper.Initialize
End Sub

'@TestCleanup
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestCleanup()
    'Cette méthode s'exécute après chaque test.
    If Not helper Is Nothing Then
        helper.Cleanup
        Set helper = Nothing
    End If
End Sub

'@TestMethod("Integration")
'@Description("Simule un scénario de synchronisation de données avec retries")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestDataSyncScenario()
    On Error GoTo TestFail
    
    'Configuration
    With helper
        'DB avec latence élevée et erreurs occasionnelles
        .Database.SetLatency 200
        .Database.SetErrorRate 0.3
        
        'Réseau instable
        .Network.Configure 300, 0.2, 0.2, 2000
        
        'Métriques détaillées
        .ConfigurePerformanceTest True, "sync_test", 1, 1
        
        'Logging initial
        .Logger.LogInfo "Démarrage synchronisation"
    End With
    
    'Simuler la synchronisation
    Dim i As Long
    Dim retryCount As Long
    Dim maxRetries As Long
    maxRetries = 3
    
    For i = 1 To 5 'Synchroniser 5 éléments
        retryCount = 0
        Do
            On Error Resume Next
            
            'Lecture données source
            helper.ExecuteWithMetrics "db", "SELECT * FROM Source WHERE ID = " & i
            If Err.Number = 0 Then
                'Envoi données
                helper.ExecuteWithMetrics "network", "POST /api/sync/" & i
                If Err.Number = 0 Then
                    'Mise à jour statut
                    helper.ExecuteWithMetrics "db", "UPDATE Source SET Synced = 1 WHERE ID = " & i
                End If
            End If
            
            'Gestion des erreurs
            If Err.Number <> 0 Then
                helper.Logger.LogError "Erreur sync #" & i & " (retry " & retryCount & "): " & Err.Description
                helper.Metrics.RecordMetric "sync_retry", 1, "sync_test", "item=" & i
                retryCount = retryCount + 1
                If retryCount >= maxRetries Then
                    helper.Metrics.RecordMetric "sync_fail", 1, "sync_test", "item=" & i
                    Exit Do
                End If
            Else
                helper.Metrics.RecordMetric "sync_success", 1, "sync_test", "item=" & i
                Exit Do
            End If
            
            On Error GoTo TestFail
        Loop
    Next i
    
    'Analyse des résultats
    Dim successCount As Double
    Dim failCount As Double
    Dim retryTotal As Double
    
    successCount = helper.Metrics.GetTotalMetric("sync_success", "sync_test")
    failCount = helper.Metrics.GetTotalMetric("sync_fail", "sync_test")
    retryTotal = helper.Metrics.GetTotalMetric("sync_retry", "sync_test")
    
    'Assertions
    Assert.IsTrue successCount > 0, "Devrait avoir des synchronisations réussies"
    Assert.IsTrue successCount + failCount = 5, "Total devrait être 5"
    helper.Logger.LogInfo "Synchronisation terminée: " & successCount & " succès, " & _
                         failCount & " échecs, " & retryTotal & " retries"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Integration")
'@Description("Simule un scénario de traitement batch avec deadlocks")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestBatchProcessingScenario()
    On Error GoTo TestFail
    
    'Configuration
    With helper
        'DB avec deadlocks
        .Database.SetLatency 100
        .Database.SetErrorRate 0.1
        .Database.EnableDeadlockSimulation True
        
        'Métriques
        .ConfigurePerformanceTest True, "batch_test", 1, 1
    End With
    
    'Paramètres batch
    Const BATCH_SIZE As Long = 10
    Const NUM_BATCHES As Long = 5
    
    'Traitement des batches
    Dim batchId As Long
    Dim itemId As Long
    Dim deadlockRetries As Long
    
    For batchId = 1 To NUM_BATCHES
        helper.Logger.LogInfo "Traitement batch " & batchId
        
        'Début transaction
        On Error Resume Next
        helper.ExecuteWithMetrics "db", "BEGIN TRANSACTION"
        
        'Traitement items
        For itemId = 1 To BATCH_SIZE
            'Simuler traitement
            helper.ExecuteWithMetrics "db", "UPDATE Items SET Status = 'Processing' " & _
                                         "WHERE BatchId = " & batchId & " AND ItemId = " & itemId
            
            If Err.Number <> 0 Then
                'Deadlock détecté
                If InStr(Err.Description, "deadlock") > 0 Then
                    helper.Logger.LogError "Deadlock détecté batch " & batchId & " item " & itemId
                    helper.Metrics.RecordMetric "deadlock", 1, "batch_test", _
                                              "batch=" & batchId & ",item=" & itemId
                    
                    'Retry après deadlock
                    deadlockRetries = deadlockRetries + 1
                    If deadlockRetries <= 3 Then
                        helper.ExecuteWithMetrics "db", "ROLLBACK"
                        itemId = itemId - 1 'Réessayer cet item
                        Continue For
                    End If
                End If
                
                'Autre erreur ou trop de retries
                helper.ExecuteWithMetrics "db", "ROLLBACK"
                helper.Metrics.RecordMetric "batch_error", 1, "batch_test", "batch=" & batchId
                Exit For
            End If
            
            'Traitement réussi
            helper.Metrics.RecordMetric "item_success", 1, "batch_test", _
                                      "batch=" & batchId & ",item=" & itemId
        Next itemId
        
        'Commit si pas d'erreur
        If Err.Number = 0 Then
            helper.ExecuteWithMetrics "db", "COMMIT"
            helper.Metrics.RecordMetric "batch_success", 1, "batch_test", "batch=" & batchId
        End If
        
        On Error GoTo TestFail
    Next batchId
    
    'Analyse résultats
    Dim batchSuccess As Double
    Dim batchErrors As Double
    Dim deadlocks As Double
    Dim itemSuccess As Double
    
    batchSuccess = helper.Metrics.GetTotalMetric("batch_success", "batch_test")
    batchErrors = helper.Metrics.GetTotalMetric("batch_error", "batch_test")
    deadlocks = helper.Metrics.GetTotalMetric("deadlock", "batch_test")
    itemSuccess = helper.Metrics.GetTotalMetric("item_success", "batch_test")
    
    'Assertions
    Assert.IsTrue batchSuccess + batchErrors = NUM_BATCHES, "Nombre total de batches incorrect"
    Assert.IsTrue batchSuccess > 0, "Devrait avoir des batches réussis"
    Assert.IsTrue itemSuccess > 0, "Devrait avoir des items traités"
    
    'Rapport
    helper.Logger.LogInfo "Traitement terminé:" & vbNewLine & _
                         "- Batches réussis: " & batchSuccess & vbNewLine & _
                         "- Batches échoués: " & batchErrors & vbNewLine & _
                         "- Deadlocks: " & deadlocks & vbNewLine & _
                         "- Items traités: " & itemSuccess
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Integration")
'@Description("Simule un scénario de cache avec invalidation")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestCacheInvalidationScenario()
    On Error GoTo TestFail
    
    'Configuration
    With helper
        'DB lente
        .Database.SetLatency 500
        .Database.SetErrorRate 0.1
        
        'Métriques
        .ConfigurePerformanceTest True, "cache_test", 1, 1
    End With
    
    'Structure pour simuler le cache
    Dim cache As New Collection
    Dim cacheHits As Long
    Dim cacheMisses As Long
    
    'Fonction de lecture avec cache
    Dim'@Description: 
'@Param: 
'@Returns: 

 Function GetDataWithCache(ByVal key As String) As String
        On Error Resume Next
        
        'Vérifier cache
        GetDataWithCache = cache(key)
        If Err.Number = 0 Then
            'Cache hit
            helper.Metrics.RecordMetric "cache_hit", 1, "cache_test", "key=" & key
            cacheHits = cacheHits + 1
        Else
            'Cache miss
            helper.Metrics.RecordMetric "cache_miss", 1, "cache_test", "key=" & key
            cacheMisses = cacheMisses + 1
            
            'Lire depuis DB
            helper.ExecuteWithMetrics "db", "SELECT Data FROM Cache WHERE Key = '" & key & "'"
            If Err.Number = 0 Then
                'Stocker dans le cache
                GetDataWithCache = "Data_" & key
                cache.Add GetDataWithCache, key
                helper.Metrics.RecordMetric "cache_store", 1, "cache_test", "key=" & key
            End If
        End If
        
        On Error GoTo TestFail
    End Function
    
    'Simuler accès aux données
    Dim i As Long
    Dim key As String
    Dim data As String
    
    'Premier passage - remplir le cache
    For i = 1 To 10
        key = "Key" & i
        data = GetDataWithCache(key)
        Assert.IsTrue Len(data) > 0, "Données devraient être récupérées pour " & key
    Next i
    
    'Deuxième passage - utiliser le cache
    For i = 1 To 10
        key = "Key" & i
        data = GetDataWithCache(key)
        Assert.IsTrue Len(data) > 0, "Données devraient être en cache pour " & key
    Next i
    
    'Invalider certaines entrées
    For i = 1 To 5
        key = "Key" & i
        cache.Remove key
        helper.Metrics.RecordMetric "cache_invalidate", 1, "cache_test", "key=" & key
    Next i
    
    'Troisième passage - cache partiel
    For i = 1 To 10
        key = "Key" & i
        data = GetDataWithCache(key)
        Assert.IsTrue Len(data) > 0, "Données devraient être récupérées pour " & key
    Next i
    
    'Analyse résultats
    Dim hits As Double
    Dim misses As Double
    Dim stores As Double
    Dim invalidations As Double
    
    hits = helper.Metrics.GetTotalMetric("cache_hit", "cache_test")
    misses = helper.Metrics.GetTotalMetric("cache_miss", "cache_test")
    stores = helper.Metrics.GetTotalMetric("cache_store", "cache_test")
    invalidations = helper.Metrics.GetTotalMetric("cache_invalidate", "cache_test")
    
    'Assertions
    Assert.AreEqual 10, hits, "Nombre de cache hits incorrect"
    Assert.AreEqual 15, misses, "Nombre de cache misses incorrect"
    Assert.AreEqual 15, stores, "Nombre de cache stores incorrect"
    Assert.AreEqual 5, invalidations, "Nombre d'invalidations incorrect"
    
    'Rapport
    helper.Logger.LogInfo "Test cache terminé:" & vbNewLine & _
                         "- Hits: " & hits & vbNewLine & _
                         "- Misses: " & misses & vbNewLine & _
                         "- Stores: " & stores & vbNewLine & _
                         "- Invalidations: " & invalidations
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Integration")
'@Description("Simule un scénario de réplication avec conflits")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestReplicationScenario()
    On Error GoTo TestFail
    
    'Configuration
    With helper
        'DB avec latence variable
        .Database.SetLatency 150
        .Database.SetErrorRate 0.15
        
        'Réseau instable
        .Network.Configure 250, 0.1, 0.1, 3000
        
        'Métriques
        .ConfigurePerformanceTest True, "repl_test", 1, 1
    End With
    
    'Structure pour simuler les versions
    Dim versions As New Collection
    
    'Simuler réplication entre noeuds
    Const NUM_NODES As Long = 3
    Dim nodeId As Long
    Dim recordId As Long
    
    'Fonction de mise à jour avec gestion de version
    Dim'@Description: 
'@Param: 
'@Returns: 

 Function UpdateRecord(ByVal node As Long, _
                            ByVal record As Long, _
                            ByVal version As Long) As Boolean
        On Error Resume Next
        
        'Vérifier version actuelle
        Dim currentVersion As Long
        currentVersion = versions("R" & record)
        
        If version > currentVersion Then
            'Mise à jour autorisée
            helper.ExecuteWithMetrics "db", "UPDATE Records SET Version = " & version & _
                                         " WHERE Id = " & record & " AND Node = " & node
            
            If Err.Number = 0 Then
                versions("R" & record) = version
                helper.Metrics.RecordMetric "repl_success", 1, "repl_test", _
                                          "node=" & node & ",record=" & record
                UpdateRecord = True
            Else
                helper.Metrics.RecordMetric "repl_error", 1, "repl_test", _
                                          "node=" & node & ",record=" & record
            End If
        Else
            'Conflit de version
            helper.Metrics.RecordMetric "repl_conflict", 1, "repl_test", _
                                      "node=" & node & ",record=" & record
        End If
        
        On Error GoTo TestFail
    End Function
    
    'Initialiser versions
    For recordId = 1 To 5
        versions.Add 0, "R" & recordId
    Next recordId
    
    'Simuler mises à jour depuis différents noeuds
    Dim iteration As Long
    For iteration = 1 To 3
        For nodeId = 1 To NUM_NODES
            For recordId = 1 To 5
                'Simuler latence réseau
                helper.ExecuteWithMetrics "network", "SYNC /node/" & nodeId & "/record/" & recordId
                
                If Err.Number = 0 Then
                    'Tenter mise à jour
                    Dim newVersion As Long
                    newVersion = iteration * 100 + nodeId
                    
                    If UpdateRecord(nodeId, recordId, newVersion) Then
                        'Propager aux autres noeuds
                        Dim targetNode As Long
                        For targetNode = 1 To NUM_NODES
                            If targetNode <> nodeId Then
                                helper.ExecuteWithMetrics "network", _
                                    "REPLICATE /node/" & targetNode & "/record/" & recordId & _
                                    "/version/" & newVersion
                            End If
                        Next targetNode
                    End If
                End If
            Next recordId
        Next nodeId
    Next iteration
    
    'Analyse résultats
    Dim successes As Double
    Dim errors As Double
    Dim conflicts As Double
    
    successes = helper.Metrics.GetTotalMetric("repl_success", "repl_test")
    errors = helper.Metrics.GetTotalMetric("repl_error", "repl_test")
    conflicts = helper.Metrics.GetTotalMetric("repl_conflict", "repl_test")
    
    'Assertions
    Assert.IsTrue successes > 0, "Devrait avoir des réplications réussies"
    Assert.IsTrue conflicts > 0, "Devrait détecter des conflits"
    
    'Rapport
    helper.Logger.LogInfo "Test réplication terminé:" & vbNewLine & _
                         "- Succès: " & successes & vbNewLine & _
                         "- Erreurs: " & errors & vbNewLine & _
                         "- Conflits: " & conflicts
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Integration")
'@Description("Simule un scénario de failover et récupération")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestFailoverRecoveryScenario()
    On Error GoTo TestFail
    
    'Configuration
    With helper
        'DB primaire instable
        .Database.SetLatency 1000
        .Database.SetErrorRate 0.5
        
        'DB secondaire plus stable
        .Database.ConfigureSecondary 200, 0.1
        
        'Métriques
        .ConfigurePerformanceTest True, "failover_test", 1, 1
    End With
    
    'Paramètres
    Const NUM_OPERATIONS As Long = 20
    Dim primaryFailures As Long
    Dim switchovers As Long
    
    'Simuler opérations avec failover
    Dim i As Long
    For i = 1 To NUM_OPERATIONS
        On Error Resume Next
        
        'Tenter opération sur primaire
        helper.ExecuteWithMetrics "db_primary", "SELECT * FROM Data WHERE ID = " & i
        
        If Err.Number <> 0 Then
            'Échec primaire
            helper.Logger.LogError "Échec primaire #" & i & ": " & Err.Description
            helper.Metrics.RecordMetric "primary_failure", 1, "failover_test"
            primaryFailures = primaryFailures + 1
            
            'Basculer sur secondaire
            helper.ExecuteWithMetrics "db_switch", "FAILOVER TO SECONDARY"
            If Err.Number = 0 Then
                switchovers = switchovers + 1
                helper.Metrics.RecordMetric "switchover", 1, "failover_test"
                
                'Réessayer sur secondaire
                helper.ExecuteWithMetrics "db_secondary", "SELECT * FROM Data WHERE ID = " & i
                If Err.Number = 0 Then
                    helper.Metrics.RecordMetric "secondary_success", 1, "failover_test"
                Else
                    helper.Metrics.RecordMetric "complete_failure", 1, "failover_test"
                End If
            End If
        Else
            helper.Metrics.RecordMetric "primary_success", 1, "failover_test"
        End If
        
        'Tenter récupération périodique du primaire
        If primaryFailures > 0 And i Mod 5 = 0 Then
            helper.ExecuteWithMetrics "db_recovery", "RECOVER PRIMARY"
            If Err.Number = 0 Then
                helper.Logger.LogInfo "Primaire récupéré après " & primaryFailures & " échecs"
                primaryFailures = 0
                helper.Metrics.RecordMetric "primary_recovery", 1, "failover_test"
            End If
        End If
        
        On Error GoTo TestFail
    Next i
    
    'Analyse résultats
    Dim primarySuccess As Double
    Dim secondarySuccess As Double
    Dim completeFailures As Double
    Dim recoveries As Double
    
    primarySuccess = helper.Metrics.GetTotalMetric("primary_success", "failover_test")
    secondarySuccess = helper.Metrics.GetTotalMetric("secondary_success", "failover_test")
    completeFailures = helper.Metrics.GetTotalMetric("complete_failure", "failover_test")
    recoveries = helper.Metrics.GetTotalMetric("primary_recovery", "failover_test")
    
    'Assertions
    Assert.IsTrue primarySuccess + secondarySuccess + completeFailures = NUM_OPERATIONS, _
                 "Nombre total d'opérations incorrect"
    Assert.IsTrue switchovers > 0, "Devrait avoir des basculements"
    Assert.IsTrue recoveries > 0, "Devrait avoir des récupérations"
    
    'Rapport
    helper.Logger.LogInfo "Test failover terminé:" & vbNewLine & _
                         "- Succès primaire: " & primarySuccess & vbNewLine & _
                         "- Succès secondaire: " & secondarySuccess & vbNewLine & _
                         "- Échecs complets: " & completeFailures & vbNewLine & _
                         "- Basculements: " & switchovers & vbNewLine & _
                         "- Récupérations: " & recoveries
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Integration")
'@Description("Simule un scénario de migration de données avec validation")
'@Description: 
'@Param: 
'@Returns: 

Public Sub TestDataMigrationScenario()
    On Error GoTo TestFail
    
    'Configuration
    With helper
        'Source et destination
        .Database.ConfigureSource 100, 0.1
        .Database.ConfigureDestination 150, 0.2
        
        'Validation stricte
        .Database.EnableStrictValidation True
        
        'Métriques
        .ConfigurePerformanceTest True, "migration_test", 1, 1
    End With
    
    'Structure pour le suivi
    Dim migrated As New Collection
    Dim validated As New Collection
    
    'Paramètres
    Const BATCH_SIZE As Long = 50
    Const NUM_RECORDS As Long = 200
    Dim currentBatch As Long
    
    'Fonction de validation
    '@Description: 
'@Param: 
'@Returns: 

Private Function ValidateRecord(ByVal recordId As Long) As Boolean
        On Error Resume Next
        
        'Lecture source
        Dim sourceData As String
        helper.ExecuteWithMetrics "db_source", "SELECT * FROM Source WHERE ID = " & recordId
        If Err.Number = 0 Then
            sourceData = "DATA_" & recordId
            
            'Lecture destination
            helper.ExecuteWithMetrics "db_dest", "SELECT * FROM Destination WHERE ID = " & recordId
            If Err.Number = 0 Then
                'Comparer données
                Dim destData As String
                destData = "DATA_" & recordId
                
                If sourceData = destData Then
                    helper.Metrics.RecordMetric "validation_success", 1, "migration_test", _
                                              "record=" & recordId
                    ValidateRecord = True
                Else
                    helper.Metrics.RecordMetric "validation_mismatch", 1, "migration_test", _
                                              "record=" & recordId
                End If
            End If
        End If
        
        If Err.Number <> 0 Then
            helper.Metrics.RecordMetric "validation_error", 1, "migration_test", _
                                      "record=" & recordId
        End If
        
        On Error GoTo TestFail
    End Function
    
    'Migration par lots
    For currentBatch = 1 To (NUM_RECORDS \ BATCH_SIZE)
        Dim startId As Long
        Dim endId As Long
        startId = ((currentBatch - 1) * BATCH_SIZE) + 1
        endId = currentBatch * BATCH_SIZE
        
        'Extraction
        On Error Resume Next
        helper.ExecuteWithMetrics "db_source", _
            "SELECT * FROM Source WHERE ID BETWEEN " & startId & " AND " & endId
            
        If Err.Number = 0 Then
            'Transformation
            helper.ExecuteWithMetrics "transform", "TRANSFORM_BATCH " & currentBatch
            
            If Err.Number = 0 Then
                'Chargement
                helper.ExecuteWithMetrics "db_dest", _
                    "INSERT INTO Destination SELECT * FROM Transformed WHERE BatchID = " & currentBatch
                
                If Err.Number = 0 Then
                    'Marquer comme migrés
                    Dim i As Long
                    For i = startId To endId
                        migrated.Add i
                        helper.Metrics.RecordMetric "migration_success", 1, "migration_test", _
                                                  "record=" & i
                    Next i
                Else
                    helper.Metrics.RecordMetric "load_error", 1, "migration_test", _
                                              "batch=" & currentBatch
                End If
            Else
                helper.Metrics.RecordMetric "transform_error", 1, "migration_test", _
                                          "batch=" & currentBatch
            End If
        Else
            helper.Metrics.RecordMetric "extract_error", 1, "migration_test", _
                                      "batch=" & currentBatch
        End If
        
        On Error GoTo TestFail
        
        'Validation du lot
        If migrated.Count > 0 Then
            Dim recordId As Long
            For Each recordId In migrated
                If ValidateRecord(recordId) Then
                    validated.Add recordId
                End If
            Next recordId
        End If
    Next currentBatch
    
    'Analyse résultats
    Dim successCount As Double
    Dim extractErrors As Double
    Dim transformErrors As Double
    Dim loadErrors As Double
    Dim validationSuccess As Double
    Dim validationErrors As Double
    
    successCount = helper.Metrics.GetTotalMetric("migration_success", "migration_test")
    extractErrors = helper.Metrics.GetTotalMetric("extract_error", "migration_test")
    transformErrors = helper.Metrics.GetTotalMetric("transform_error", "migration_test")
    loadErrors = helper.Metrics.GetTotalMetric("load_error", "migration_test")
    validationSuccess = helper.Metrics.GetTotalMetric("validation_success", "migration_test")
    validationErrors = helper.Metrics.GetTotalMetric("validation_error", "migration_test")
    
    'Assertions
    Assert.IsTrue successCount > 0, "Devrait avoir des migrations réussies"
    Assert.IsTrue validationSuccess > 0, "Devrait avoir des validations réussies"
    Assert.AreEqual migrated.Count, validated.Count, "Tous les enregistrements migrés devraient être validés"
    
    'Rapport
    helper.Logger.LogInfo "Test migration terminé:" & vbNewLine & _
                         "- Migrations réussies: " & successCount & vbNewLine & _
                         "- Erreurs extraction: " & extractErrors & vbNewLine & _
                         "- Erreurs transformation: " & transformErrors & vbNewLine & _
                         "- Erreurs chargement: " & loadErrors & vbNewLine & _
                         "- Validations réussies: " & validationSuccess & vbNewLine & _
                         "- Erreurs validation: " & validationErrors
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub 