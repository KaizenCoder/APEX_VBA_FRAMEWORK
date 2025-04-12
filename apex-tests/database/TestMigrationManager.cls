VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestMigrationManager"

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
Private migrationManager As clsMigrationManager
Private mockTransactionManager As clsTransactionManager

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
    Set migrationManager = New clsMigrationManager
    Set mockTransactionManager = New clsTransactionManager
    
    ' Créer les dossiers de test
    CreateTestFolders
End Sub

'@TestCleanup
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set migrationManager = Nothing
    Set mockTransactionManager = Nothing
    
    ' Nettoyer les dossiers de test
    CleanupTestFolders
End Sub

'@TestMethod("Migration")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestInitialize()
    On Error GoTo TestFail
    
    'Act
    Dim result As Boolean
    result = migrationManager.Initialize( _
        scriptsPath:=GetTestScriptsPath(), _
        backupPath:=GetTestBackupPath(), _
        transactionManager:=mockTransactionManager _
    )
    
    'Assert
    Assert.IsTrue result, "L'initialisation devrait réussir"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Migration")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestMigrationUpgrade()
    On Error GoTo TestFail
    
    'Arrange
    CreateTestMigrationScript "V1.0.0__Initial.sql", "CREATE TABLE Test (ID INT)"
    CreateTestMigrationScript "V1.0.1__AddColumn.sql", "ALTER TABLE Test ADD Column2 VARCHAR(50)"
    
    migrationManager.Initialize GetTestScriptsPath(), GetTestBackupPath(), Nothing, Nothing, mockTransactionManager
    
    'Act
    Dim result As Boolean
    result = migrationManager.MigrateTo("1.0.1")
    
    'Assert
    Assert.IsTrue result, "La migration devrait réussir"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Migration")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestMigrationDowngrade()
    On Error GoTo TestFail
    
    'Arrange
    CreateTestMigrationScript "V1.0.0__Initial.sql", "CREATE TABLE Test (ID INT)"
    CreateTestMigrationScript "V1.0.1__AddColumn.sql", "ALTER TABLE Test ADD Column2 VARCHAR(50)"
    
    migrationManager.Initialize GetTestScriptsPath(), GetTestBackupPath(), Nothing, Nothing, mockTransactionManager
    
    'Upgrade first
    migrationManager.MigrateTo "1.0.1"
    
    'Act - Downgrade
    Dim result As Boolean
    result = migrationManager.MigrateTo("1.0.0")
    
    'Assert
    Assert.IsTrue result, "Le downgrade devrait réussir"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Migration")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestBackup()
    On Error GoTo TestFail
    
    'Arrange
    migrationManager.Initialize GetTestScriptsPath(), GetTestBackupPath(), Nothing, Nothing, mockTransactionManager
    
    'Act
    Dim result As Boolean
    result = migrationManager.MigrateTo("1.0.0", backup:=True)
    
    'Assert
    Assert.IsTrue result, "La migration avec backup devrait réussir"
    Assert.IsTrue FolderHasFiles(GetTestBackupPath()), "Un fichier de backup devrait être créé"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Migration")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestStats()
    On Error GoTo TestFail
    
    'Arrange
    CreateTestMigrationScript "V1.0.0__Initial.sql", "CREATE TABLE Test (ID INT)"
    migrationManager.Initialize GetTestScriptsPath(), GetTestBackupPath(), Nothing, Nothing, mockTransactionManager
    
    'Act
    migrationManager.MigrateTo "1.0.0" ' Succès
    migrationManager.MigrateTo "1.0.1" ' Échec (script n'existe pas)
    
    'Assert
    Dim stats As String
    stats = migrationManager.GetStats()
    
    Assert.IsTrue InStr(stats, "Total: 2") > 0, "Devrait avoir 2 migrations"
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
    
    ' Supprimer les fichiers
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
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Function FolderHasFiles(ByVal folderPath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    
    FolderHasFiles = folder.Files.Count > 0
End Function 