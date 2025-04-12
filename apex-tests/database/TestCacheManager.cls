VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCacheManager"

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
Private cache As clsCacheManager

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
    Set cache = New clsCacheManager
    cache.Initialize
End Sub

'@TestCleanup
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set cache = Nothing
End Sub

'@TestMethod("Cache")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestSetAndGetValue()
    On Error GoTo TestFail
    
    'Arrange
    Dim key As String
    Dim value As String
    key = "test_key"
    value = "test_value"
    
    'Act
    cache.SetValue key, value
    
    'Assert
    Dim found As Boolean
    Dim result As Variant
    result = cache.GetValue(key, found)
    
    Assert.IsTrue found, "La valeur devrait être trouvée"
    Assert.AreEqual value, result, "La valeur récupérée devrait correspondre"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Cache")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestExpiration()
    On Error GoTo TestFail
    
    'Arrange
    Dim key As String
    key = "expiring_key"
    
    'Act
    cache.Configure 100, 1, 1 ' TTL de 1 seconde
    cache.SetValue key, "value"
    
    'Attendre l'expiration
    Sleep 1100 ' 1.1 secondes
    
    'Assert
    Dim found As Boolean
    cache.GetValue key, found
    Assert.IsFalse found, "La valeur devrait avoir expiré"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Cache")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestEviction()
    On Error GoTo TestFail
    
    'Arrange
    cache.Configure 2, 3600, 3600 ' Taille max de 2 entrées
    
    'Act
    cache.SetValue "key1", "value1"
    cache.SetValue "key2", "value2"
    cache.SetValue "key3", "value3" ' Devrait évincer la plus ancienne
    
    'Assert
    Dim found As Boolean
    cache.GetValue "key1", found
    Assert.IsFalse found, "key1 devrait avoir été évincée"
    
    cache.GetValue "key2", found
    Assert.IsTrue found, "key2 devrait toujours être présente"
    
    cache.GetValue "key3", found
    Assert.IsTrue found, "key3 devrait être présente"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Cache")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestClear()
    On Error GoTo TestFail
    
    'Arrange
    cache.SetValue "key1", "value1"
    cache.SetValue "key2", "value2"
    
    'Act
    cache.Clear
    
    'Assert
    Dim found As Boolean
    cache.GetValue "key1", found
    Assert.IsFalse found, "key1 devrait être supprimée"
    
    cache.GetValue "key2", found
    Assert.IsFalse found, "key2 devrait être supprimée"
    
TestExit:
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Cache")
'@Description: 
'@Param: 
'@Returns: 

Private Sub TestStats()
    On Error GoTo TestFail
    
    'Arrange
    cache.SetValue "key1", "value1"
    
    'Act - Succès
    Dim found As Boolean
    cache.GetValue "key1", found
    
    'Act - Échec
    cache.GetValue "nonexistent", found
    
    'Assert
    Dim stats As String
    stats = cache.GetStats()
    
    Assert.IsTrue InStr(stats, "Succès: 1") > 0, "Devrait avoir 1 succès"
    Assert.IsTrue InStr(stats, "Échecs: 1") > 0, "Devrait avoir 1 échec"
    
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