' Migrated to apex-core/testing - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Attribute VB_Name = "modTestRegistry"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Option Explicit
' ==========================================================================
' Module : modTestRegistry
' Version : 2.0
' Purpose : Module d'enregistrement et de découverte automatique des tests
' Date    : 10/04/2025
' ==========================================================================

' --- Constantes ---
Private Const MODULE_NAME As String = "modTestRegistry"
Private Const DEFAULT_SUITE_PREFIX As String = "Test"

' --- Variables globales ---
Private m_testSuites As Collection
Private m_initialized As Boolean
Private m_logger As Object
Private m_configManager As Object

' --- Initialisation ---
'@Description: 
'@Param: 
'@Returns: 

Public Sub Initialize()
    If m_initialized Then Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    
    Set m_testSuites = New Collection
    
    ' Initialiser les dépendances
    InitializeDependencies
    
    ' Marquer comme initialisé
    m_initialized = True
    
    ' Log
    LogMessage "Registre de tests initialisé", "info"
End Sub

' --- Enregistrement des suites ---
'@Description: 
'@Param: 
'@Returns: 

Public Sub RegisterSuite(suite As clsTestSuite)
    ' Initialiser si ce n'est pas déjà fait
    If Not m_initialized Then Initialize
    
    ' Vérifier si la suite existe déjà
    On Error Resume Next
    Dim existingSuite As clsTestSuite
    Set existingSuite = m_testSuites(suite.Name)
    
    If Err.Number = 0 Then
        ' La suite existe déjà
        LogMessage "Suite de test '" & suite.Name & "' déjà enregistrée", "warning"
        Err.Clear
    Else
        ' Ajouter la suite
        m_testSuites.Add suite, suite.Name
        
        ' Enregistrer également auprès du TestRunner
        modTestRunner.RegisterTestSuite suite
        
        LogMessage "Suite de test '" & suite.Name & "' enregistrée avec " & suite.TestCount & " test(s)", "info"
        Err.Clear
    End If
    
    On Error GoTo 0
End Sub

' --- Récupération des suites ---
'@Description: 
'@Param: 
'@Returns: 

Public Function GetSuite(ByVal suiteName As String) As clsTestSuite
    ' Initialiser si ce n'est pas déjà fait
    If Not m_initialized Then Initialize
    
    ' Vérifier si la suite existe
    On Error Resume Next
    Set GetSuite = m_testSuites(suiteName)
    
    If Err.Number <> 0 Then
        Set GetSuite = Nothing
        Err.Clear
    End If
    
    On Error GoTo 0
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetAllSuites() As Collection
    ' Initialiser si ce n'est pas déjà fait
    If Not m_initialized Then Initialize
    
    Set GetAllSuites = m_testSuites
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetSuiteCount() As Long
    ' Initialiser si ce n'est pas déjà fait
    If Not m_initialized Then Initialize
    
    GetSuiteCount = m_testSuites.Count
End Function

' --- Découverte automatique ---
'@Description: 
'@Param: 
'@Returns: 

Public Sub DiscoverTests()
    ' Initialiser si ce n'est pas déjà fait
    If Not m_initialized Then Initialize
    
    LogMessage "Début de la découverte automatique des tests...", "info"
    
    ' Obtenir les modules VBA
    Dim vbProj As Object
    Dim vbComp As Object
    Dim prefixToUse As String
    
    ' Obtenir le préfixe à utiliser pour les modules de test
    If Not m_configManager Is Nothing Then
        prefixToUse = m_configManager.GetSetting("General", "DefaultSuitePrefix", DEFAULT_SUITE_PREFIX)
    Else
        prefixToUse = DEFAULT_SUITE_PREFIX
    End If
    
    On Error Resume Next
    Set vbProj = Application.VBE.ActiveVBProject
    If Err.Number <> 0 Then
        LogMessage "Erreur lors de l'accès au projet VBA. Vérifiez que la sécurité du modèle d'objet est activée.", "error"
        Err.Clear
        Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    End If
    On Error GoTo 0
    
    ' Scanner les modules commençant par le préfixe
    Dim modulesFound As Long
    modulesFound = 0
    
    For Each vbComp In vbProj.VBComponents
        If Left(vbComp.Name, Len(prefixToUse)) = prefixToUse Then
            ' Traiter ce module comme une suite de test potentielle
            If ProcessTestModule(vbComp) Then
                modulesFound = modulesFound + 1
            End If
        End If
    Next vbComp
    
    LogMessage "Découverte automatique terminée. " & modulesFound & " module(s) de test identifié(s).", "info"
End Sub

' --- Exécution des tests ---
'@Description: 
'@Param: 
'@Returns: 

Public Sub RunAllDiscoveredTests(Optional ByVal outputReport As Boolean = True)
    ' Initialiser si ce n'est pas déjà fait
    If Not m_initialized Then Initialize
    
    ' Si aucune suite n'a été découverte, lancer la découverte
    If m_testSuites.Count = 0 Then
        DiscoverTests
    End If
    
    ' Vérifier si des suites ont été trouvées
    If m_testSuites.Count = 0 Then
        LogMessage "Aucune suite de test trouvée. Impossible d'exécuter les tests.", "warning"
        Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    End If
    
    ' Obtenir le paramètre d'arrêt sur échec
    Dim stopOnFailure As Boolean
    If Not m_configManager Is Nothing Then
        stopOnFailure = (m_configManager.GetSetting("General", "StopOnFailure", "False") = "True")
    Else
        stopOnFailure = False
    End If
    
    ' Exécuter tous les tests via TestRunner
    modTestRunner.RunAllTests outputReport, stopOnFailure
End Sub

' --- Création de suite ---
'@Description: 
'@Param: 
'@Returns: 

Public Function CreateSuite(ByVal suiteName As String) As clsTestSuite
    ' Initialiser si ce n'est pas déjà fait
    If Not m_initialized Then Initialize
    
    ' Vérifier si la suite existe déjà
    Dim existingSuite As clsTestSuite
    Set existingSuite = GetSuite(suiteName)
    
    If Not existingSuite Is Nothing Then
        Set CreateSuite = existingSuite
    Else
        ' Créer une nouvelle suite
        Dim newSuite As New clsTestSuite
        newSuite.Name = suiteName
        
        ' Enregistrer la suite
        RegisterSuite newSuite
        
        Set CreateSuite = newSuite
    End If
End Function

' --- Fonctions utilitaires ---
'@Description: 
'@Param: 
'@Returns: 

Private Function ProcessTestModule(vbComp As Object) As Boolean
    Dim procKind As Long
    Dim procName As String
    Dim lineNum As Long
    Dim suiteCreated As Boolean
    Dim suite As clsTestSuite
    Dim moduleName As String
    
    moduleName = vbComp.Name
    suiteCreated = False
    
    ' Parcourir les procédures du module
    For i = 1 To vbComp.CodeModule.CountOfLines
        lineNum = 1
        On Error Resume Next
        procName = vbComp.CodeModule.ProcOfLine(i, procKind)
        
        If Err.Number = 0 And procName <> "" Then
            ' Si le nom de la procédure commence par "Test", c'est un test
            If Left(procName, 4) = "Test" Then
                ' Créer la suite si ce n'est pas déjà fait
                If Not suiteCreated Then
                    Set suite = CreateSuite(moduleName)
                    suiteCreated = True
                End If
                
                ' Ajouter le test à la suite
                suite.AddTest procName, procName, moduleName
                
                ' Avancer au-delà de cette procédure
                i = i + vbComp.CodeModule.ProcCountLines(procName, procKind)
            End If
        End If
        
        Err.Clear
        On Error GoTo 0
    Next i
    
    ProcessTestModule = suiteCreated
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Sub InitializeDependencies()
    ' Initialiser le logger si disponible
    On Error Resume Next
    Set m_logger = CreateObject("APEX.Logger")
    If Err.Number <> 0 Then Set m_logger = Nothing
    
    ' Initialiser le gestionnaire de configuration si disponible
    Set m_configManager = CreateObject("APEX.ConfigManager")
    If Err.Number <> 0 Then Set m_configManager = Nothing
    
    Err.Clear
    On Error GoTo 0
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub LogMessage(message As String, logLevel As String)
    ' Écrire dans le journal si disponible
    On Error Resume Next
    If Not m_logger Is Nothing Then
        Select Case LCase(logLevel)
            Case "debug"
                m_logger.LogDebug MODULE_NAME, message
            Case "info"
                m_logger.LogInfo MODULE_NAME, message
            Case "warning"
                m_logger.LogWarning MODULE_NAME, message
            Case "error"
                m_logger.LogError MODULE_NAME, message
            Case Else
                m_logger.LogInfo MODULE_NAME, message
        End Select
    Else
        ' Écrire dans la fenêtre de débogage
        Debug.Print message
    End If
    On Error GoTo 0
End Sub

' --- Cleanup ---
Private Sub Class_Terminate()
    Set m_testSuites = Nothing
    Set m_logger = Nothing
    Set m_configManager = Nothing
End Sub 