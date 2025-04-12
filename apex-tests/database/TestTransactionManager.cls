VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestTransactionManager"

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
'@Folder("APEX.Tests.Database")
Option Explicit

Private Assert As Object
Private Fakes As Object
Private txManager As clsTransactionManager
Private mockConnection As ADODB.Connection

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
    Set txManager = New clsTransactionManager
    Set mockConnection = New ADODB.Connection
    txManager.Initialize
End Sub

'@TestCleanup
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set txManager = Nothing
    Set mockConnection = Nothing
End Sub

'@TestMethod("Transaction")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestBeginTransaction()
    On Error GoTo TestFail
    
    'Act
    Dim result As Boolean
    result = txManager.BeginTransaction
    
    'Assert
    Assert.IsTrue result, "La transaction devrait démarrer"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Transaction")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestCommitTransaction()
    On Error GoTo TestFail
    
    'Arrange
    txManager.BeginTransaction
    txManager.AddConnection mockConnection
    
    'Act
    Dim result As Boolean
    result = txManager.CommitTransaction
    
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

'@TestMethod("Transaction")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestRollbackTransaction()
    On Error GoTo TestFail
    
    'Arrange
    txManager.BeginTransaction
    txManager.AddConnection mockConnection
    
    'Act
    Dim result As Boolean
    result = txManager.RollbackTransaction
    
    'Assert
    Assert.IsTrue result, "La transaction devrait être annulée"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Transaction")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestSavepoint()
    On Error GoTo TestFail
    
    'Arrange
    txManager.BeginTransaction
    txManager.AddConnection mockConnection
    
    'Act
    Dim savepointName As String
    savepointName = txManager.CreateSavepoint
    
    'Assert
    Assert.IsTrue Len(savepointName) > 0, "Un nom de savepoint devrait être généré"
    
    'Act - Rollback
    Dim result As Boolean
    result = txManager.RollbackToSavepoint(savepointName)
    
    'Assert
    Assert.IsTrue result, "Le rollback au savepoint devrait réussir"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Transaction")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestMultipleConnections()
    On Error GoTo TestFail
    
    'Arrange
    txManager.BeginTransaction
    
    Dim conn1 As New ADODB.Connection
    Dim conn2 As New ADODB.Connection
    
    'Act
    Dim result1 As Boolean
    Dim result2 As Boolean
    result1 = txManager.AddConnection(conn1)
    result2 = txManager.AddConnection(conn2)
    
    'Assert
    Assert.IsTrue result1, "La première connexion devrait être ajoutée"
    Assert.IsTrue result2, "La deuxième connexion devrait être ajoutée"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Transaction")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestStats()
    On Error GoTo TestFail
    
    'Arrange
    txManager.BeginTransaction
    txManager.AddConnection mockConnection
    txManager.CommitTransaction
    
    txManager.BeginTransaction
    txManager.AddConnection mockConnection
    txManager.RollbackTransaction
    
    'Act
    Dim stats As String
    stats = txManager.GetStats()
    
    'Assert
    Assert.IsTrue InStr(stats, "Total: 2") > 0, "Devrait avoir 2 transactions"
    Assert.IsTrue InStr(stats, "Réussies: 1") > 0, "Devrait avoir 1 succès"
    Assert.IsTrue InStr(stats, "Échouées: 1") > 0, "Devrait avoir 1 échec"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub 