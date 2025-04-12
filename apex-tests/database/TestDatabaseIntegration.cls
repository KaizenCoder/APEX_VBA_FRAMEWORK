VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestDatabaseIntegration"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@TestModule
'@Folder("APEX.Tests.Database.Integration")
Option Explicit

Private Assert As Object
Private Fakes As Object

Private securityManager As clsSecurityManager
Private cacheManager As clsCacheManager
Private transactionManager As clsTransactionManager
Private migrationManager As clsMigrationManager
Private mockLogger As ILoggerBase
Private mockMetrics As clsMetricsCollector

'@ModuleInitialize
'@Description: 
'@Param: 
'@Returns: 

Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
'@Description: 
'@Param: 
'@Returns: 

Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestInitialize()
    'This method runs before every test in the module..
    Set securityManager = New clsSecurityManager
    Set cacheManager = New clsCacheManager
    Set transactionManager = New clsTransactionManager
    Set migrationManager = New clsMigrationManager
    Set mockLogger = New MockLogger
    Set mockMetrics = New clsMetricsCollector
    
    ' Initialiser les composants
    securityManager.Initialize mockLogger
    cacheManager.Initialize mockLogger, mockMetrics
    transactionManager.Initialize mockLogger, mockMetrics
    
    ' Créer les dossiers de test pour la migration
    CreateTestFolders
End Sub

'@TestCleanup
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set securityManager = Nothing
    Set cacheManager = Nothing
    Set transactionManager = Nothing
    Set migrationManager = Nothing
    Set mockLogger = Nothing
    Set mockMetrics = Nothing
    
    ' Nettoyer les dossiers de test
    CleanupTestFolders
End Sub

'@TestMethod("Integration")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestSecureTransactionWithCache()
    On Error GoTo TestFail
    
    'Arrange
    Dim connString As String
    connString = "Server=TestServer;Database=TestDB;User=TestUser;Password=TestPass;"
    
    'Act - Chiffrer la chaîne de connexion
    Dim encryptedConn As String
    encryptedConn = securityManager.EncryptConnectionString(connString)
    
    'Act - Stocker dans le cache
    cacheManager.SetValue "conn_string", encryptedConn
    
    'Act - Démarrer une transaction
    transactionManager.BeginTransaction
    
    'Act - Récupérer et déchiffrer
    Dim found As Boolean
    Dim cachedConn As String
    cachedConn = cacheManager.GetValue("conn_string", found)
    
    Dim decryptedConn As String
    decryptedConn = securityManager.DecryptConnectionString(cachedConn)
    
    'Assert
    Assert.IsTrue found, "La chaîne devrait être dans le cache"
    Assert.AreEqual connString, decryptedConn, "La chaîne déchiffrée devrait correspondre"
    
    'Act - Valider la transaction
    Dim result As Boolean
    result = transactionManager.CommitTransaction
    
    'Assert
    Assert.IsTrue result, "La transaction devrait être validée"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Integration")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestMigrationWithSecurityAndCache()
    On Error GoTo TestFail
    
    'Arrange
    CreateTestMigrationScript "V1.0.0__Initial.sql", "CREATE TABLE Users (ID INT, Username VARCHAR(50), Password VARCHAR(100))"
    
    'Act - Initialiser la migration avec cache et sécurité
    migrationManager.Initialize _
        scriptsPath:=GetTestScriptsPath(), _
        backupPath:=GetTestBackupPath(), _
        logger:=mockLogger, _
        metricsCollector:=mockMetrics, _
        transactionManager:=transactionManager
    
    'Act - Stocker un mot de passe dans le cache
    Dim testPassword As String
    testPassword = "P@ssw0rd123!"
    
    'Valider et hasher le mot de passe
    Assert.IsTrue securityManager.ValidatePassword(testPassword), "Le mot de passe devrait être valide"
    
    Dim hashedPassword As String
    hashedPassword = securityManager.HashPassword(testPassword)
    cacheManager.SetValue "admin_password", hashedPassword
    
    'Act - Exécuter la migration
    Dim result As Boolean
    result = migrationManager.MigrateTo("1.0.0")
    
    'Assert
    Assert.IsTrue result, "La migration devrait réussir"
    
    'Vérifier que le mot de passe est toujours dans le cache
    Dim found As Boolean
    Dim cachedHash As String
    cachedHash = cacheManager.GetValue("admin_password", found)
    
    Assert.IsTrue found, "Le hash devrait être dans le cache"
    Assert.AreEqual hashedPassword, cachedHash, "Le hash devrait correspondre"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Integration")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestSecurityAuditWithCache()
    On Error GoTo TestFail
    
    'Arrange
    Dim username As String
    username = "testUser"
    
    'Act - Simuler des tentatives de connexion
    For i = 1 To 3
        securityManager.RecordFailedLogin username
        
        'Stocker le nombre de tentatives dans le cache
        cacheManager.SetValue "login_attempts_" & username, i
    Next i
    
    'Assert - Vérifier le cache et le statut
    Dim found As Boolean
    Dim attempts As Long
    attempts = cacheManager.GetValue("login_attempts_" & username, found)
    
    Assert.IsTrue found, "Les tentatives devraient être dans le cache"
    Assert.AreEqual 3, attempts, "Devrait avoir 3 tentatives"
    
    'Act - Vérifier les logs d'audit
    Dim logs As String
    logs = securityManager.GetAuditLogs(username)
    
    'Assert
    Assert.IsTrue InStr(logs, "Failed login attempt") > 0, "Devrait avoir des échecs de connexion dans les logs"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Integration")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestTransactionRollbackWithCache()
    On Error GoTo TestFail
    
    'Arrange
    cacheManager.SetValue "important_data", "original_value"
    
    'Act - Démarrer une transaction
    transactionManager.BeginTransaction
    
    'Modifier la donnée dans le cache
    cacheManager.SetValue "important_data", "new_value"
    
    'Rollback
    transactionManager.RollbackTransaction
    
    'Assert - La donnée devrait revenir à sa valeur originale
    Dim found As Boolean
    Dim value As String
    value = cacheManager.GetValue("important_data", found)
    
    Assert.IsTrue found, "La donnée devrait être dans le cache"
    Assert.AreEqual "original_value", value, "La valeur devrait être restaurée"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Integration")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestConcurrentTransactionsWithCache()
    On Error GoTo TestFail
    
    'Arrange
    Dim conn1 As New ADODB.Connection
    Dim conn2 As New ADODB.Connection
    
    'Act - Transaction 1
    transactionManager.BeginTransaction
    transactionManager.AddConnection conn1
    cacheManager.SetValue "shared_data", "transaction1_value"
    
    'Act - Transaction 2 (devrait attendre)
    Dim tx2Started As Boolean
    tx2Started = transactionManager.BeginTransaction
    
    'Assert
    Assert.IsFalse tx2Started, "La deuxième transaction ne devrait pas démarrer"
    
    'Act - Valider Transaction 1
    transactionManager.CommitTransaction
    
    'Maintenant Transaction 2 peut démarrer
    tx2Started = transactionManager.BeginTransaction
    Assert.IsTrue tx2Started, "La deuxième transaction devrait démarrer"
    
    'Vérifier les métriques
    Dim stats As TMetricAggregation
    stats = mockMetrics.GetMetricStats("TransactionWaitTime")
    Assert.IsTrue stats.Count > 0, "Devrait avoir des métriques de temps d'attente"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Integration")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestFailedMigrationRecovery()
    On Error GoTo TestFail
    
    'Arrange
    CreateTestMigrationScript "V1.0.0__Initial.sql", "CREATE TABLE Test (ID INT)"
    CreateTestMigrationScript "V1.0.1__Invalid.sql", "INVALID SQL COMMAND"
    
    'Act - Démarrer la migration avec backup
    migrationManager.Initialize GetTestScriptsPath(), GetTestBackupPath(), mockLogger, mockMetrics, transactionManager
    
    'Première migration réussie
    Dim result As Boolean
    result = migrationManager.MigrateTo("1.0.0")
    Assert.IsTrue result, "La première migration devrait réussir"
    
    'Deuxième migration échoue
    result = migrationManager.MigrateTo("1.0.1")
    Assert.IsFalse result, "La migration invalide devrait échouer"
    
    'Vérifier les logs
    Assert.IsTrue mockLogger.ContainsMessage("Migration failed"), "Devrait logger l'échec"
    
    'Vérifier le rollback automatique
    Assert.IsTrue mockLogger.ContainsMessage("Rollback completed"), "Devrait logger le rollback"
    
    'Vérifier les métriques
    Assert.IsTrue mockMetrics.HasMetric("MigrationFailure"), "Devrait avoir une métrique d'échec"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Integration")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestSecurityBreachDetection()
    On Error GoTo TestFail
    
    'Arrange
    Dim username As String
    username = "testUser"
    
    'Act - Simuler des attaques
    Dim maliciousInputs(4) As String
    maliciousInputs(0) = "' OR '1'='1"
    maliciousInputs(1) = "UNION SELECT * FROM Users"
    maliciousInputs(2) = "'; DROP TABLE Users; --"
    maliciousInputs(3) = "admin'--"
    maliciousInputs(4) = "' UNION SELECT NULL,NULL,NULL,NULL--"
    
    Dim i As Long
    For i = 0 To 4
        'Tentative d'injection
        On Error Resume Next
        securityManager.SanitizeSqlInput maliciousInputs(i)
        
        'Enregistrer la tentative
        If Err.Number <> 0 Then
            securityManager.RecordFailedLogin username
            cacheManager.SetValue "breach_attempt_" & i, maliciousInputs(i)
        End If
        On Error GoTo TestFail
    Next i
    
    'Assert
    'Vérifier le verrouillage du compte
    Assert.IsTrue securityManager.IsAccountLocked(username), "Le compte devrait être verrouillé"
    
    'Vérifier les logs d'audit
    Dim logs As String
    logs = securityManager.GetAuditLogs(username)
    Assert.IsTrue InStr(logs, "SQL Injection attempt") > 0, "Devrait détecter l'injection SQL"
    
    'Vérifier les métriques
    Dim stats As TMetricAggregation
    stats = mockMetrics.GetMetricStats("SecurityBreach")
    Assert.IsTrue stats.Count >= 5, "Devrait avoir enregistré toutes les tentatives"
    
    'Vérifier le cache des tentatives
    Dim found As Boolean
    Dim attempts As String
    attempts = cacheManager.GetValue("breach_attempt_0", found)
    Assert.IsTrue found, "Devrait avoir enregistré la tentative dans le cache"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Integration")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestNetworkResilienceWithCache()
    On Error GoTo TestFail
    
    'Arrange
    Dim mockNetwork As New MockNetworkConnection
    mockNetwork.Initialize mockLogger, mockMetrics
    
    'Configurer pour simuler des problèmes réseau
    mockNetwork.Configure _
        latencyMs:=1000, _
        packetLossRate:=0.3, _
        errorRate:=0.2, _
        timeoutMs:=500
    
    'Act - Tenter des opérations avec retry
    Dim i As Long
    For i = 1 To 10
        'Simuler une requête
        Dim response As String
        response = mockNetwork.SendRequest("SELECT * FROM Users")
        
        'Stocker dans le cache si succès
        If Left$(response, 7) = "SUCCESS" Then
            cacheManager.SetValue "last_successful_query", response
        End If
    Next i
    
    'Assert
    'Vérifier les statistiques réseau
    Dim stats As String
    stats = mockNetwork.GetStats()
    
    Assert.IsTrue InStr(stats, "Timeouts:") > 0, "Devrait avoir des timeouts"
    Assert.IsTrue InStr(stats, "Échouées:") > 0, "Devrait avoir des échecs"
    
    'Vérifier que le cache contient au moins une réponse
    Dim found As Boolean
    Dim cachedResponse As String
    cachedResponse = cacheManager.GetValue("last_successful_query", found)
    
    Assert.IsTrue found, "Devrait avoir au moins une requête en cache"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Integration")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestStressTestWithMetrics()
    On Error GoTo TestFail
    
    'Arrange
    Const ITERATIONS As Long = 100
    Const CACHE_SIZE As Long = 1000
    
    'Act - Remplir le cache
    Dim i As Long
    For i = 1 To CACHE_SIZE
        cacheManager.SetValue "stress_key_" & i, "stress_value_" & i
    Next i
    
    'Simuler des opérations intensives
    For i = 1 To ITERATIONS
        'Transaction
        transactionManager.BeginTransaction
        
        'Modifications aléatoires
        Dim randomKey As Long
        randomKey = Int(Rnd * CACHE_SIZE) + 1
        
        'Lecture
        Dim found As Boolean
        cacheManager.GetValue "stress_key_" & randomKey, found
        
        'Écriture
        cacheManager.SetValue "stress_key_" & randomKey, "updated_" & Now
        
        'Validation
        transactionManager.CommitTransaction
    Next i
    
    'Assert
    'Vérifier les métriques de performance
    Dim cacheStats As String
    cacheStats = cacheManager.GetStats()
    
    Assert.IsTrue InStr(cacheStats, "Hits:") > 0, "Devrait avoir des hits"
    Assert.IsTrue InStr(cacheStats, "Misses:") > 0, "Devrait avoir des misses"
    
    'Vérifier les métriques de transaction
    Dim txStats As String
    txStats = transactionManager.GetStats()
    
    Assert.IsTrue InStr(txStats, "Total: " & ITERATIONS) > 0, "Devrait avoir le bon nombre de transactions"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Integration")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestConcurrencyAndDeadlockDetection()
    On Error GoTo TestFail
    
    'Arrange
    Dim resource1 As String
    Dim resource2 As String
    resource1 = "Resource1"
    resource2 = "Resource2"
    
    'Act - Transaction 1
    transactionManager.BeginTransaction
    cacheManager.SetValue resource1, "locked_by_tx1"
    
    'Act - Transaction 2 (simulée)
    Dim tx2Started As Boolean
    tx2Started = transactionManager.BeginTransaction
    
    If tx2Started Then
        'Tentative de verrouiller resource2
        cacheManager.SetValue resource2, "locked_by_tx2"
        
        'Tentative de verrouiller resource1 (devrait détecter le deadlock)
        Dim canLockResource1 As Boolean
        canLockResource1 = cacheManager.TrySetValue(resource1, "tx2_attempt")
        
        'Assert
        Assert.IsFalse canLockResource1, "Devrait détecter le deadlock potentiel"
        
        'Rollback transaction 2
        transactionManager.RollbackTransaction
    End If
    
    'Compléter transaction 1
    cacheManager.SetValue resource2, "tx1_complete"
    transactionManager.CommitTransaction
    
    'Assert
    'Vérifier les métriques
    Assert.IsTrue mockMetrics.HasMetric("DeadlockDetected"), "Devrait avoir détecté un deadlock"
    
    'Vérifier les logs
    Assert.IsTrue mockLogger.ContainsMessage("Deadlock avoided"), "Devrait logger la détection du deadlock"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'Helper Methods
'@Description: 
'@Param: 
'@Returns: 

Private Function GetTestScriptsPath() As String
    GetTestScriptsPath = ThisWorkbook.Path & "\TestScripts"
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function GetTestBackupPath() As String
    GetTestBackupPath = ThisWorkbook.Path & "\TestBackups"
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Sub CreateTestFolders()
    On Error Resume Next
    MkDir GetTestScriptsPath()
    MkDir GetTestBackupPath()
    On Error GoTo 0
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub CleanupTestFolders()
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(GetTestScriptsPath()) Then
        fso.DeleteFolder GetTestScriptsPath()
    End If
    
    If fso.FolderExists(GetTestBackupPath()) Then
        fso.DeleteFolder GetTestBackupPath()
    End If
    
    On Error GoTo 0
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub CreateTestMigrationScript(ByVal fileName As String, ByVal content As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ts As Object
    Set ts = fso.CreateTextFile(GetTestScriptsPath() & "\" & fileName, True)
    ts.Write content
    ts.Close
End Sub 