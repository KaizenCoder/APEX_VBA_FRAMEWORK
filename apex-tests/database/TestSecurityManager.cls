VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestSecurityManager"

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
Private securityManager As clsSecurityManager

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
    securityManager.Initialize
End Sub

'@TestCleanup
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set securityManager = Nothing
End Sub

'@TestMethod("Security")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestConnectionStringEncryption()
    On Error GoTo TestFail
    
    'Arrange
    Dim connString As String
    connString = "Server=MyServer;Database=MyDB;User=MyUser;Password=MyPassword;"
    
    'Act
    Dim encrypted As String
    encrypted = securityManager.EncryptConnectionString(connString)
    
    Dim decrypted As String
    decrypted = securityManager.DecryptConnectionString(encrypted)
    
    'Assert
    Assert.AreNotEqual connString, encrypted, "La chaîne chiffrée devrait être différente"
    Assert.AreEqual connString, decrypted, "La chaîne déchiffrée devrait correspondre"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Security")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestSqlInjectionPrevention()
    On Error GoTo TestFail
    
    'Arrange
    Dim maliciousInput As String
    maliciousInput = "'; DROP TABLE Users; --"
    
    'Act
    Dim sanitized As String
    sanitized = securityManager.SanitizeSqlInput(maliciousInput)
    
    'Assert
    Assert.AreNotEqual maliciousInput, sanitized, "L'entrée devrait être nettoyée"
    Assert.IsFalse InStr(sanitized, ";") > 0, "Les points-virgules devraient être supprimés"
    Assert.IsFalse InStr(sanitized, "--") > 0, "Les commentaires devraient être supprimés"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Security")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestSqlInjectionDetection()
    On Error GoTo TestFail
    
    'Arrange
    Dim inputs(3) As String
    inputs(0) = "UNION SELECT * FROM Users"
    inputs(1) = "INSERT INTO Users"
    inputs(2) = "DROP TABLE"
    inputs(3) = "Normal input"
    
    'Act & Assert
    Dim i As Long
    For i = 0 To 2 ' Les 3 premiers devraient échouer
        On Error Resume Next
        securityManager.SanitizeSqlInput inputs(i)
        
        Assert.IsTrue Err.Number <> 0, "L'injection SQL devrait être détectée: " & inputs(i)
        On Error GoTo TestFail
    Next i
    
    ' Le dernier devrait réussir
    Dim result As String
    result = securityManager.SanitizeSqlInput(inputs(3))
    Assert.AreEqual inputs(3), result, "L'entrée normale devrait être préservée"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Security")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestPasswordPolicy()
    On Error GoTo TestFail
    
    'Arrange
    Dim weakPasswords(2) As String
    weakPasswords(0) = "password"
    weakPasswords(1) = "12345678"
    weakPasswords(2) = "qwerty"
    
    Dim strongPasswords(2) As String
    strongPasswords(0) = "P@ssw0rd123!"
    strongPasswords(1) = "Str0ng#P@ss"
    strongPasswords(2) = "C0mpl3x!P@ss"
    
    'Act & Assert
    Dim i As Long
    For i = 0 To 2
        ' Les mots de passe faibles devraient échouer
        On Error Resume Next
        securityManager.ValidatePassword weakPasswords(i)
        Assert.IsTrue Err.Number <> 0, "Le mot de passe faible devrait être rejeté: " & weakPasswords(i)
        On Error GoTo TestFail
        
        ' Les mots de passe forts devraient réussir
        Dim result As Boolean
        result = securityManager.ValidatePassword(strongPasswords(i))
        Assert.IsTrue result, "Le mot de passe fort devrait être accepté: " & strongPasswords(i)
    Next i
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Security")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestLockoutMechanism()
    On Error GoTo TestFail
    
    'Arrange
    Dim username As String
    username = "testUser"
    
    'Act
    Dim i As Long
    For i = 1 To 5 ' Simuler 5 échecs de connexion
        securityManager.RecordFailedLogin username
    Next i
    
    'Assert
    Assert.IsTrue securityManager.IsAccountLocked(username), "Le compte devrait être verrouillé"
    
    'Act - Attendre le délai de verrouillage
    Sleep 5000 ' 5 secondes
    
    'Assert
    Assert.IsFalse securityManager.IsAccountLocked(username), "Le compte devrait être déverrouillé"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Security")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestAuditLogging()
    On Error GoTo TestFail
    
    'Arrange
    Dim username As String
    Dim action As String
    username = "testUser"
    action = "LOGIN"
    
    'Act
    securityManager.LogAuditEvent username, action
    
    'Assert
    Dim logs As String
    logs = securityManager.GetAuditLogs(username)
    
    Assert.IsTrue InStr(logs, username) > 0, "Le journal devrait contenir l'utilisateur"
    Assert.IsTrue InStr(logs, action) > 0, "Le journal devrait contenir l'action"
    Assert.IsTrue InStr(logs, Format(Date, "yyyy-mm-dd")) > 0, "Le journal devrait contenir la date"
    
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

Private Sub Sleep(ByVal milliseconds As Long)
    Dim endTime As Date
    endTime = DateAdd("s", milliseconds / 1000#, Now())
    Do While Now() < endTime
        DoEvents
    Loop
End Sub 