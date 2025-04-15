' Migrated to apex-core/testing - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestSuite"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ==========================================================================
' Class   : clsTestSuite
' Version : 2.0
' Purpose : Classe représentant une suite de tests unitaires avec fonctionnalités étendues
' Date    : 10/04/2025
' ==========================================================================

' --- Constantes ---
Private Const CLASS_NAME As String = "clsTestSuite"

' --- Types personnalisés ---
Private Type TestInfo
    TestName As String
    ProcedureName As String
    ModuleName As String
    Description As String
    Tags As String
    Priority As Long
    Disabled As Boolean
    Dependencies As String
    Timeout As Long
    ExpectedExceptions As String
End Type

' --- Propriétés ---
Private m_Name As String
Private m_Description As String
Private m_Author As String
Private m_Tests() As TestInfo
Private m_TestCount As Long
Private m_Results As Collection
Private m_PassedCount As Long
private m_SkippedCount As Long
Private m_FailedCount As Long
Private m_SetupProcedure As String
Private m_TeardownProcedure As String
Private m_SetupSuiteProcedure As String
Private m_TeardownSuiteProcedure As String
Private m_Category As String
Private m_ConfigManager As Object

' --- Initialisation ---
Private Sub Class_Initialize()
    m_Name = "TestSuite" & Format(Now(), "yyyymmddhhnnss")
    ReDim m_Tests(0)
    Set m_Results = New Collection
    m_TestCount = 0
    m_PassedCount = 0
    m_FailedCount = 0
    m_SkippedCount = 0
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub Class_Terminate()
    Set m_Results = Nothing
    Set m_ConfigManager = Nothing
End Sub

' --- Propriétés publiques ---
Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(ByVal value As String)
    m_Name = value
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Description(ByVal value As String)
    m_Description = value
End Property

Public Property Get Author() As String
    Author = m_Author
End Property

Public Property Let Author(ByVal value As String)
    m_Author = value
End Property

Public Property Get Category() As String
    Category = m_Category
End Property

Public Property Let Category(ByVal value As String)
    m_Category = value
End Property

Public Property Get TestCount() As Long
    TestCount = m_TestCount
End Property

Public Property Get PassedCount() As Long
    PassedCount = m_PassedCount
End Property

Public Property Get FailedCount() As Long
    FailedCount = m_FailedCount
End Property

Public Property Get SkippedCount() As Long
    SkippedCount = m_SkippedCount
End Property

Public Property Get Results() As Collection
    Set Results = m_Results
End Property

Public Property Let SetupProcedure(ByVal value As String)
    m_SetupProcedure = value
End Property

Public Property Let TeardownProcedure(ByVal value As String)
    m_TeardownProcedure = value
End Property

Public Property Let SetupSuiteProcedure(ByVal value As String)
    m_SetupSuiteProcedure = value
End Property

Public Property Let TeardownSuiteProcedure(ByVal value As String)
    m_TeardownSuiteProcedure = value
End Property

' --- Méthodes publiques ---
'@Description: 
'@Param: 
'@Returns: 

Public Sub AddTest(ByVal testName As String, ByVal testProc As String, _
                  Optional ByVal moduleName As String = "", Optional ByVal description As String = "")
    ' Méthode de compatibilité avec l'ancienne version
    AddTestEx testName, testProc, moduleName, description, "", 1, False, "", 0, ""
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Public Sub AddTestEx(ByVal testName As String, ByVal testProc As String, _
                    Optional ByVal moduleName As String = "", _
                    Optional ByVal description As String = "", _
                    Optional ByVal tags As String = "", _
                    Optional ByVal priority As Long = 1, _
                    Optional ByVal disabled As Boolean = False, _
                    Optional ByVal dependencies As String = "", _
                    Optional ByVal timeout As Long = 0, _
                    Optional ByVal expectedExceptions As String = "")
    
    ' Redimensionner le tableau si nécessaire
    If m_TestCount > 0 Then
        ReDim Preserve m_Tests(m_TestCount)
    End If
    
    ' Ajouter les informations du test
    With m_Tests(m_TestCount)
        .TestName = testName
        .ProcedureName = testProc
        .ModuleName = moduleName
        .Description = description
        .Tags = tags
        .Priority = priority
        .Disabled = disabled
        .Dependencies = dependencies
        .Timeout = IIf(timeout = 0, GetDefaultTimeout(), timeout)
        .ExpectedExceptions = expectedExceptions
    End With
    
    ' Incrémenter le compteur
    m_TestCount = m_TestCount + 1
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Public Sub RunAllTests(Optional ByVal stopOnFailure As Boolean = False)
    If m_TestCount = 0 Then Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    
    Dim i As Long
    Dim result As Object ' clsTestResult
    Dim logger As Object
    
    ' Réinitialiser les compteurs
    m_PassedCount = 0
    m_FailedCount = 0
    m_SkippedCount = 0
    Set m_Results = New Collection
    
    ' Essayer de récupérer un logger si disponible
    On Error Resume Next
    Set logger = CreateObject("APEX.Logger")
    If Err.Number <> 0 Then Set logger = Nothing
    On Error GoTo 0
    
    ' Exécuter la procédure de setup de la suite si définie
    If m_SetupSuiteProcedure <> "" Then
        If logger Is Nothing Then
            Debug.Print "  Exécution du setup de la suite: " & m_SetupSuiteProcedure
        Else
            logger.LogInfo CLASS_NAME, "Exécution du setup de la suite: " & m_SetupSuiteProcedure
        End If
        
        Application.Run m_SetupSuiteProcedure
    End If
    
    ' Trier les tests par priorité si nécessaire
    SortTestsByPriority
    
    ' Exécuter chaque test
    For i = 0 To m_TestCount - 1
        With m_Tests(i)
            ' Vérifier si le test est désactivé
            If .Disabled Then
                If logger Is Nothing Then
                    Debug.Print "  Test ignoré (désactivé): " & .TestName
                Else
                    logger.LogInfo CLASS_NAME, "Test ignoré (désactivé): " & .TestName
                End If
                
                m_SkippedCount = m_SkippedCount + 1
                Set result = New clsTestResult
                result.TestName = .TestName
                result.Skipped = True
                result.Success = True ' Un test ignoré n'est pas considéré comme un échec
                m_Results.Add result
                
                GoTo NextTest
            End If
            
            ' Vérifier les dépendances
            If Not CheckDependencies(.Dependencies) Then
                If logger Is Nothing Then
                    Debug.Print "  Test ignoré (dépendances non satisfaites): " & .TestName
                Else
                    logger.LogInfo CLASS_NAME, "Test ignoré (dépendances non satisfaites): " & .TestName
                End If
                
                m_SkippedCount = m_SkippedCount + 1
                Set result = New clsTestResult
                result.TestName = .TestName
                result.Skipped = True
                result.SkipReason = "Dépendances non satisfaites"
                result.Success = True ' Un test ignoré n'est pas considéré comme un échec
                m_Results.Add result
                
                GoTo NextTest
            End If
            
            ' Exécuter la procédure de setup si définie
            If m_SetupProcedure <> "" Then
                Application.Run m_SetupProcedure
            End If
            
            ' Log de début de test
            If logger Is Nothing Then
                Debug.Print "  Exécution du test: " & .TestName
            Else
                logger.LogInfo CLASS_NAME, "Exécution du test: " & .TestName
            End If
            
            ' Exécuter le test
            Set result = modTestRunner.RunTest(.TestName, .ProcedureName, .ModuleName, .Timeout, .ExpectedExceptions)
            
            ' Exécuter la procédure de teardown si définie
            If m_TeardownProcedure <> "" Then
                Application.Run m_TeardownProcedure
            End If
            
            ' Ajouter le résultat à la collection
            m_Results.Add result
            
            ' Mettre à jour les compteurs
            If result.Success Then
                m_PassedCount = m_PassedCount + 1
                If logger Is Nothing Then
                    Debug.Print "    RÉUSSI en " & Format(result.ExecutionTime, "0.000") & " sec"
                Else
                    logger.LogInfo CLASS_NAME, "Test RÉUSSI: " & .TestName & " en " & Format(result.ExecutionTime, "0.000") & " sec"
                End If
            Else
                m_FailedCount = m_FailedCount + 1
                If logger Is Nothing Then
                    Debug.Print "    ÉCHEC en " & Format(result.ExecutionTime, "0.000") & " sec: " & result.ErrorMessage
                Else
                    logger.LogError CLASS_NAME, "Test ÉCHEC: " & .TestName & " en " & Format(result.ExecutionTime, "0.000") & " sec: " & result.ErrorMessage
                End If
                
                ' Arrêter si nécessaire
                If stopOnFailure Then Exit For
            End If
        End With
NextTest:
    Next i
    
    ' Exécuter la procédure de teardown de la suite si définie
    If m_TeardownSuiteProcedure <> "" Then
        If logger Is Nothing Then
            Debug.Print "  Exécution du teardown de la suite: " & m_TeardownSuiteProcedure
        Else
            logger.LogInfo CLASS_NAME, "Exécution du teardown de la suite: " & m_TeardownSuiteProcedure
        End If
        
        Application.Run m_TeardownSuiteProcedure
    End If
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Public Function GetFailureDetails() As String
    Dim result As String
    Dim testResult As Object ' clsTestResult
    Dim i As Long
    
    i = 0
    For Each testResult In m_Results
        If Not testResult.Success And Not testResult.Skipped Then
            i = i + 1
            result = result & i & ". " & testResult.TestName & vbCrLf & _
                     "   Source: " & testResult.ErrorSource & vbCrLf & _
                     "   Message: " & testResult.ErrorMessage & vbCrLf & vbCrLf
        End If
    Next testResult
    
    GetFailureDetails = result
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetTestInfo(ByVal testIndex As Long) As Variant
    Dim result As Object
    
    If testIndex < 0 Or testIndex >= m_TestCount Then
        GetTestInfo = Null
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    Set result = CreateObject("Scripting.Dictionary")
    With m_Tests(testIndex)
        result.Add "TestName", .TestName
        result.Add "ProcedureName", .ProcedureName
        result.Add "ModuleName", .ModuleName
        result.Add "Description", .Description
        result.Add "Tags", .Tags
        result.Add "Priority", .Priority
        result.Add "Disabled", .Disabled
        result.Add "Dependencies", .Dependencies
        result.Add "Timeout", .Timeout
        result.Add "ExpectedExceptions", .ExpectedExceptions
    End With
    
    GetTestInfo = result
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetDetailedReport() As String
    Dim report As String
    Dim testResult As Object ' clsTestResult
    Dim i As Long
    
    report = "RAPPORT DÉTAILLÉ DE LA SUITE: " & m_Name & vbCrLf & _
             "Description: " & m_Description & vbCrLf & _
             "Auteur: " & m_Author & vbCrLf & _
             "Catégorie: " & m_Category & vbCrLf & _
             "Date d'exécution: " & Format(Now, "yyyy-mm-dd hh:mm:ss") & vbCrLf & vbCrLf & _
             "RÉSUMÉ:" & vbCrLf & _
             "  Tests: " & m_TestCount & vbCrLf & _
             "  Réussis: " & m_PassedCount & " (" & Format(IIf(m_TestCount > 0, m_PassedCount / m_TestCount * 100, 0), "0.00") & "%)" & vbCrLf & _
             "  Échoués: " & m_FailedCount & vbCrLf & _
             "  Ignorés: " & m_SkippedCount & vbCrLf & vbCrLf
    
    report = report & "DÉTAILS DES TESTS:" & vbCrLf
    
    i = 0
    For Each testResult In m_Results
        i = i + 1
        report = report & i & ". " & testResult.TestName & ": "
        
        If testResult.Skipped Then
            report = report & "IGNORÉ" & IIf(testResult.SkipReason <> "", " (" & testResult.SkipReason & ")", "") & vbCrLf
        ElseIf testResult.Success Then
            report = report & "RÉUSSI en " & Format(testResult.ExecutionTime, "0.000") & " sec" & vbCrLf
        Else
            report = report & "ÉCHEC en " & Format(testResult.ExecutionTime, "0.000") & " sec" & vbCrLf & _
                     "   Source: " & testResult.ErrorSource & vbCrLf & _
                     "   Message: " & testResult.ErrorMessage & vbCrLf
        End If
        
        report = report & vbCrLf
    Next testResult
    
    GetDetailedReport = report
End Function

' --- Méthodes privées ---
'@Description: 
'@Param: 
'@Returns: 

Private Function GetDefaultTimeout() As Long
    On Error Resume Next
    
    ' Essayer de lire depuis la configuration
    If m_ConfigManager Is Nothing Then
        ' Essayer de récupérer le gestionnaire de configuration
        Set m_ConfigManager = CreateObject("APEX.ConfigManager")
        If Err.Number <> 0 Then Set m_ConfigManager = Nothing
    End If
    
    If Not m_ConfigManager Is Nothing Then
        GetDefaultTimeout = m_ConfigManager.GetSetting("General", "TimeoutSeconds", 30) * 1000
    Else
        GetDefaultTimeout = 30000 ' 30 secondes par défaut
    End If
    
    On Error GoTo 0
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Sub SortTestsByPriority()
    Dim i As Long, j As Long
    Dim tempTest As TestInfo
    
    ' Tri à bulles simple par priorité
    For i = 0 To m_TestCount - 2
        For j = i + 1 To m_TestCount - 1
            If m_Tests(i).Priority > m_Tests(j).Priority Then
                tempTest = m_Tests(i)
                m_Tests(i) = m_Tests(j)
                m_Tests(j) = tempTest
            End If
        Next j
    Next i
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Function CheckDependencies(ByVal dependencies As String) As Boolean
    If Trim(dependencies) = "" Then
        CheckDependencies = True
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    Dim depArray() As String
    Dim i As Long
    
    depArray = Split(dependencies, ",")
    
    For i = 0 To UBound(depArray)
        Dim depName As String
        depName = Trim(depArray(i))
        
        ' Vérifier si le test existe dans les résultats
        Dim found As Boolean
        Dim testPassed As Boolean
        Dim testResult As Object ' clsTestResult
        
        found = False
        testPassed = False
        
        For Each testResult In m_Results
            If testResult.TestName = depName Then
                found = True
                testPassed = testResult.Success And Not testResult.Skipped
                Exit For
            End If
        Next testResult
        
        ' Si dépendance non trouvée ou échouée
        If Not found Or Not testPassed Then
            CheckDependencies = False
            Exit'@Description: 
'@Param: 
'@Returns: 

 Function
        End If
    Next i
    
    CheckDependencies = True
End Function 