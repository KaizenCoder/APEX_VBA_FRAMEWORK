VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsTestRunner.cls
' Classe d'exécution des tests pour le framework de test
'
' Cette classe gère l'exécution des suites de tests, y compris la configuration,
' l'exécution et la génération de rapports.
'
' @class clsTestRunner
' @author APEX Framework Team
' @version 1.0
' @date 2024-07-27
Option Explicit

' États internes
Private Type TTestInfo
    Name As String               ' Nom du test
    Source As String             ' Nom du module/classe source
    Procedure As String          ' Nom de la procédure
    Level As String              ' Niveau de test (UNIT, INTEGRATION, PERFORMANCE)
    Tags As Collection           ' Tags associés au test
    Status As String             ' Statut du test (PENDING, PASSED, FAILED, SKIPPED)
    Duration As Double           ' Durée d'exécution en millisecondes
    ErrorMessage As String       ' Message d'erreur (si échec)
    SetupProcedure As String     ' Procédure de setup (optionnelle)
    TeardownProcedure As String  ' Procédure de teardown (optionnelle)
    Timeout As Long              ' Timeout en millisecondes
    IsSlow As Boolean            ' Indique si le test est lent
End Type

Private Type TTestRunnerState
    Tests As Collection          ' Collection de tests
    SuiteName As String          ' Nom de la suite de tests
    TotalTests As Long           ' Nombre total de tests
    PassedTests As Long          ' Nombre de tests réussis
    FailedTests As Long          ' Nombre de tests échoués
    SkippedTests As Long         ' Nombre de tests ignorés
    CurrentTestIndex As Long     ' Index du test en cours d'exécution
    IsConfigured As Boolean      ' Indicateur de configuration
    VerboseMode As Boolean       ' Mode verbeux pour les logs
    StopOnFirstFailure As Boolean ' Arrêter l'exécution après le premier échec
    TestLevel As String          ' Niveau de test à exécuter
    TestFilter As String         ' Filtre de tests à exécuter
    SkipSlowTests As Boolean     ' Ignorer les tests lents
    DefaultTimeoutMs As Long     ' Timeout par défaut en millisecondes
    BeforeSuiteProcedure As String ' Procédure exécutée avant la suite
    AfterSuiteProcedure As String  ' Procédure exécutée après la suite
End Type

Private this As TTestRunnerState

'------------------------------------------------------------------------------
' Initialise l'objet TestRunner
'
' @param suiteName Nom de la suite de tests
'------------------------------------------------------------------------------
Public Sub Initialize(ByVal suiteName As String)
    Set this.Tests = New Collection
    this.SuiteName = suiteName
    this.TotalTests = 0
    this.PassedTests = 0
    this.FailedTests = 0
    this.SkippedTests = 0
    this.CurrentTestIndex = 0
    this.IsConfigured = False
    this.VerboseMode = True
    this.StopOnFirstFailure = False
    this.TestLevel = TEST_LEVEL_UNIT
    this.TestFilter = "*"
    this.SkipSlowTests = False
    this.DefaultTimeoutMs = 30000 ' 30 secondes
    this.BeforeSuiteProcedure = ""
    this.AfterSuiteProcedure = ""
    
    LogMessage "Suite de tests '" & suiteName & "' initialisée."
End Sub

'------------------------------------------------------------------------------
' Configure le TestRunner
'
' @param verboseMode Mode verbeux pour les logs
' @param stopOnFirstFailure Arrêter l'exécution après le premier échec
' @param testLevel Niveau de test à exécuter
' @param testFilter Filtre de tests à exécuter
' @param skipSlowTests Ignorer les tests lents
' @param defaultTimeoutMs Timeout par défaut en millisecondes
'------------------------------------------------------------------------------
Public Sub Configure(Optional ByVal verboseMode As Boolean = True, _
                    Optional ByVal stopOnFirstFailure As Boolean = False, _
                    Optional ByVal testLevel As String = TEST_LEVEL_UNIT, _
                    Optional ByVal testFilter As String = "*", _
                    Optional ByVal skipSlowTests As Boolean = False, _
                    Optional ByVal defaultTimeoutMs As Long = 30000)
    this.VerboseMode = verboseMode
    this.StopOnFirstFailure = stopOnFirstFailure
    this.TestLevel = testLevel
    this.TestFilter = testFilter
    this.SkipSlowTests = skipSlowTests
    this.DefaultTimeoutMs = defaultTimeoutMs
    this.IsConfigured = True
    
    LogMessage "Suite de tests '" & this.SuiteName & "' configurée."
End Sub

'------------------------------------------------------------------------------
' Accesseur pour le nom de la suite
'
' @return String Nom de la suite de tests
'------------------------------------------------------------------------------
Public Property Get SuiteName() As String
    SuiteName = this.SuiteName
End Property

'------------------------------------------------------------------------------
' Mutateur pour le nom de la suite
'
' @param value Nouveau nom de la suite
'------------------------------------------------------------------------------
Public Property Let SuiteName(ByVal value As String)
    this.SuiteName = value
End Property

'------------------------------------------------------------------------------
' Accesseur pour le nombre total de tests
'
' @return Long Nombre total de tests
'------------------------------------------------------------------------------
Public Property Get TotalTests() As Long
    TotalTests = this.TotalTests
End Property

'------------------------------------------------------------------------------
' Accesseur pour le nombre de tests réussis
'
' @return Long Nombre de tests réussis
'------------------------------------------------------------------------------
Public Property Get PassedTests() As Long
    PassedTests = this.PassedTests
End Property

'------------------------------------------------------------------------------
' Accesseur pour le nombre de tests échoués
'
' @return Long Nombre de tests échoués
'------------------------------------------------------------------------------
Public Property Get FailedTests() As Long
    FailedTests = this.FailedTests
End Property

'------------------------------------------------------------------------------
' Accesseur pour le nombre de tests ignorés
'
' @return Long Nombre de tests ignorés
'------------------------------------------------------------------------------
Public Property Get SkippedTests() As Long
    SkippedTests = this.SkippedTests
End Property

'------------------------------------------------------------------------------
' Accesseur pour le nombre de tests
'
' @return Long Nombre de tests dans la suite
'------------------------------------------------------------------------------
Public Property Get TestCount() As Long
    If this.Tests Is Nothing Then
        TestCount = 0
    Else
        TestCount = this.Tests.Count
    End If
End Property

'------------------------------------------------------------------------------
' Définit une procédure à exécuter avant la suite de tests
'
' @param procedureName Nom de la procédure
'------------------------------------------------------------------------------
Public Sub SetBeforeSuiteProcedure(ByVal procedureName As String)
    this.BeforeSuiteProcedure = procedureName
End Sub

'------------------------------------------------------------------------------
' Définit une procédure à exécuter après la suite de tests
'
' @param procedureName Nom de la procédure
'------------------------------------------------------------------------------
Public Sub SetAfterSuiteProcedure(ByVal procedureName As String)
    this.AfterSuiteProcedure = procedureName
End Sub

'------------------------------------------------------------------------------
' Ajoute un test à la suite
'
' @param testName Nom du test
' @param sourceName Nom du module ou de la classe source
' @param procedureName Nom de la procédure de test
' @param level Niveau du test (UNIT, INTEGRATION, PERFORMANCE)
' @param setupProc Nom de la procédure de setup (optionnel)
' @param teardownProc Nom de la procédure de teardown (optionnel)
' @param timeoutMs Timeout en millisecondes (optionnel)
' @param isSlow Indique si le test est lent (optionnel)
' @return Long Index du test ajouté
'------------------------------------------------------------------------------
Public Function AddTest(ByVal testName As String, _
                       ByVal sourceName As String, _
                       ByVal procedureName As String, _
                       Optional ByVal level As String = TEST_LEVEL_UNIT, _
                       Optional ByVal setupProc As String = "", _
                       Optional ByVal teardownProc As String = "", _
                       Optional ByVal timeoutMs As Long = 0, _
                       Optional ByVal isSlow As Boolean = False) As Long
    
    Dim testInfo As TTestInfo
    
    ' Initialiser le test
    testInfo.Name = testName
    testInfo.Source = sourceName
    testInfo.Procedure = procedureName
    testInfo.Level = level
    Set testInfo.Tags = New Collection
    testInfo.Status = "PENDING"
    testInfo.Duration = 0
    testInfo.ErrorMessage = ""
    testInfo.SetupProcedure = setupProc
    testInfo.TeardownProcedure = teardownProc
    
    ' Utiliser le timeout par défaut si non spécifié
    If timeoutMs <= 0 Then
        testInfo.Timeout = this.DefaultTimeoutMs
    Else
        testInfo.Timeout = timeoutMs
    End If
    
    testInfo.IsSlow = isSlow
    
    ' Ajouter à la collection
    Dim testIndex As Long
    testIndex = this.Tests.Count + 1
    
    ' Stocker dans une variable pour contourner la limitation de VBA
    Dim tempTest As TTestInfo
    tempTest = testInfo
    
    ' Ajouter à la collection en utilisant une clé unique
    this.Tests.Add tempTest, testName & "_" & testIndex
    
    this.TotalTests = this.TotalTests + 1
    
    LogMessage "Test ajouté: " & testName & " (" & level & ")"
    
    AddTest = testIndex
End Function

'------------------------------------------------------------------------------
' Ajoute un tag à un test spécifique
'
' @param testIndex Index du test
' @param tag Tag à ajouter
' @return Boolean Succès de l'opération
'------------------------------------------------------------------------------
Public Function AddTagToTest(ByVal testIndex As Long, ByVal tag As String) As Boolean
    If testIndex <= 0 Or testIndex > this.Tests.Count Then
        AddTagToTest = False
        Exit Function
    End If
    
    ' Récupérer le test
    Dim testInfo As TTestInfo
    testInfo = this.Tests(testIndex)
    
    ' Ajouter le tag (éviter les duplications)
    On Error Resume Next
    testInfo.Tags.Add tag, tag
    On Error GoTo 0
    
    ' Stocker à nouveau dans la collection
    this.Tests.Remove testIndex
    this.Tests.Add testInfo, testInfo.Name & "_" & testIndex, testIndex
    
    AddTagToTest = True
End Function

'------------------------------------------------------------------------------
' Exécute tous les tests de la suite
'
' @return String Résumé des résultats de la suite
'------------------------------------------------------------------------------
Public Function RunTests() As String
    If Not this.IsConfigured Then Configure
    
    ' Réinitialiser les compteurs
    this.PassedTests = 0
    this.FailedTests = 0
    this.SkippedTests = 0
    this.CurrentTestIndex = 0
    
    ' Logs de début de suite
    LogMessage "=== EXÉCUTION DE LA SUITE: " & this.SuiteName & " ==="
    LogMessage "Tests: " & this.Tests.Count & ", Mode: " & this.TestLevel & ", Filtre: " & this.TestFilter
    
    Dim startTime As Double
    startTime = Timer
    
    ' Exécuter BeforeSuite si spécifié
    If Len(this.BeforeSuiteProcedure) > 0 Then
        On Error Resume Next
        LogMessage "Exécution de la procédure BeforeSuite: " & this.BeforeSuiteProcedure
        Application.Run this.BeforeSuiteProcedure
        
        If Err.Number <> 0 Then
            LogMessage "ERREUR dans BeforeSuite: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If
    
    ' Exécuter les tests
    Dim i As Long
    Dim testInfo As TTestInfo
    
    For i = 1 To this.Tests.Count
        this.CurrentTestIndex = i
        
        ' Récupérer les infos du test
        testInfo = this.Tests(i)
        
        ' Vérifier si le test doit être exécuté selon le niveau et le filtre
        If ShouldRunTest(testInfo) Then
            ' Exécuter le test
            ExecuteTest i
        Else
            ' Marquer comme ignoré
            testInfo = this.Tests(i)
            testInfo.Status = "SKIPPED"
            this.Tests.Remove i
            this.Tests.Add testInfo, testInfo.Name & "_" & i, i
            this.SkippedTests = this.SkippedTests + 1
            
            LogMessage "IGNORÉ: " & testInfo.Name & " (niveau/filtre/lent)"
        End If
        
        ' Arrêter l'exécution après le premier échec si demandé
        If this.StopOnFirstFailure And this.FailedTests > 0 Then
            LogMessage "Arrêt de l'exécution après échec (StopOnFirstFailure=True)"
            Exit For
        End If
    Next i
    
    ' Exécuter AfterSuite si spécifié
    If Len(this.AfterSuiteProcedure) > 0 Then
        On Error Resume Next
        LogMessage "Exécution de la procédure AfterSuite: " & this.AfterSuiteProcedure
        Application.Run this.AfterSuiteProcedure
        
        If Err.Number <> 0 Then
            LogMessage "ERREUR dans AfterSuite: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If
    
    ' Durée totale
    Dim duration As Double
    duration = Timer - startTime
    
    ' Générer le résumé
    Dim summary As String
    summary = GenerateTestSummary(duration)
    
    RunTests = summary
End Function

'------------------------------------------------------------------------------
' Vérifie si un test doit être exécuté selon les critères de filtrage
'
' @param testInfo Informations sur le test
' @return Boolean True si le test doit être exécuté
'------------------------------------------------------------------------------
Private Function ShouldRunTest(ByVal testInfo As TTestInfo) As Boolean
    ' Vérifier le niveau de test
    If this.TestLevel <> testInfo.Level And this.TestLevel <> "*" Then
        ShouldRunTest = False
        Exit Function
    End If
    
    ' Vérifier si les tests lents doivent être ignorés
    If this.SkipSlowTests And testInfo.IsSlow Then
        ShouldRunTest = False
        Exit Function
    End If
    
    ' Vérifier le filtre de nom (wildcard)
    If this.TestFilter <> "*" Then
        If Not MatchesWildcard(testInfo.Name, this.TestFilter) Then
            ShouldRunTest = False
            Exit Function
        End If
    End If
    
    ShouldRunTest = True
End Function

'------------------------------------------------------------------------------
' Exécute un test spécifique
'
' @param testIndex Index du test à exécuter
'------------------------------------------------------------------------------
Private Sub ExecuteTest(ByVal testIndex As Long)
    Dim testInfo As TTestInfo
    testInfo = this.Tests(testIndex)
    
    ' Log de début de test
    LogMessage "TEST: " & testInfo.Name & " (" & testInfo.Level & ")"
    
    ' Exécuter le setup si spécifié
    If Len(testInfo.SetupProcedure) > 0 Then
        On Error Resume Next
        LogVerbose "Setup: " & testInfo.SetupProcedure
        Application.Run testInfo.SetupProcedure
        
        If Err.Number <> 0 Then
            LogMessage "ERREUR dans Setup: " & Err.Description
            UpdateTestStatus testIndex, "FAILED", "Erreur dans Setup: " & Err.Description, 0
            Err.Clear
            Exit Sub
        End If
        On Error GoTo 0
    End If
    
    ' Exécuter le test avec timeout
    Dim startTime As Double
    startTime = Timer * 1000 ' Millisecondes
    
    On Error Resume Next
    
    ' Exécuter le test
    Application.Run testInfo.Procedure
    
    ' Calculer la durée
    Dim duration As Double
    duration = (Timer * 1000) - startTime
    
    ' Vérifier si une erreur s'est produite
    If Err.Number <> 0 Then
        LogMessage "ÉCHEC: " & testInfo.Name & " - " & Err.Description
        UpdateTestStatus testIndex, "FAILED", Err.Description, duration
        Err.Clear
    Else
        LogMessage "RÉUSSI: " & testInfo.Name & " (" & Format(duration, "0.00") & " ms)"
        UpdateTestStatus testIndex, "PASSED", "", duration
    End If
    
    On Error GoTo 0
    
    ' Exécuter le teardown si spécifié
    If Len(testInfo.TeardownProcedure) > 0 Then
        On Error Resume Next
        LogVerbose "Teardown: " & testInfo.TeardownProcedure
        Application.Run testInfo.TeardownProcedure
        
        If Err.Number <> 0 Then
            LogMessage "AVERTISSEMENT: Erreur dans Teardown: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Sub

'------------------------------------------------------------------------------
' Met à jour le statut d'un test
'
' @param testIndex Index du test
' @param status Nouveau statut (PASSED, FAILED, SKIPPED)
' @param errorMessage Message d'erreur (si échec)
' @param duration Durée d'exécution en millisecondes
'------------------------------------------------------------------------------
Private Sub UpdateTestStatus(ByVal testIndex As Long, ByVal status As String, _
                          ByVal errorMessage As String, ByVal duration As Double)
    Dim testInfo As TTestInfo
    testInfo = this.Tests(testIndex)
    
    ' Mettre à jour les informations
    testInfo.Status = status
    testInfo.ErrorMessage = errorMessage
    testInfo.Duration = duration
    
    ' Mettre à jour la collection
    this.Tests.Remove testIndex
    this.Tests.Add testInfo, testInfo.Name & "_" & testIndex, testIndex
    
    ' Mettre à jour les compteurs
    Select Case status
        Case "PASSED"
            this.PassedTests = this.PassedTests + 1
        Case "FAILED"
            this.FailedTests = this.FailedTests + 1
        Case "SKIPPED"
            this.SkippedTests = this.SkippedTests + 1
    End Select
End Sub

'------------------------------------------------------------------------------
' Génère un résumé des résultats de la suite de tests
'
' @param duration Durée totale d'exécution en secondes
' @return String Résumé formaté
'------------------------------------------------------------------------------
Private Function GenerateTestSummary(ByVal duration As Double) As String
    Dim summary As String
    
    summary = String(50, "-") & vbCrLf & _
              "RÉSUMÉ: " & this.SuiteName & vbCrLf & _
              String(50, "-") & vbCrLf & _
              "Total: " & this.Tests.Count & vbCrLf & _
              "Réussis: " & this.PassedTests & " (" & Format(SafeDiv(this.PassedTests, this.Tests.Count) * 100, "0.00") & "%)" & vbCrLf & _
              "Échoués: " & this.FailedTests & " (" & Format(SafeDiv(this.FailedTests, this.Tests.Count) * 100, "0.00") & "%)" & vbCrLf & _
              "Ignorés: " & this.SkippedTests & " (" & Format(SafeDiv(this.SkippedTests, this.Tests.Count) * 100, "0.00") & "%)" & vbCrLf & _
              "Durée: " & Format(duration, "0.00") & " secondes" & vbCrLf & _
              String(50, "-")
              
    GenerateTestSummary = summary
End Function

'------------------------------------------------------------------------------
' Récupère le nom d'un test à l'index spécifié
'
' @param testIndex Index du test
' @return String Nom du test
'------------------------------------------------------------------------------
Public Function GetTestName(ByVal testIndex As Long) As String
    If testIndex <= 0 Or testIndex > this.Tests.Count Then
        GetTestName = ""
        Exit Function
    End If
    
    Dim testInfo As TTestInfo
    testInfo = this.Tests(testIndex)
    
    GetTestName = testInfo.Name
End Function

'------------------------------------------------------------------------------
' Récupère le statut d'un test à l'index spécifié
'
' @param testIndex Index du test
' @return String Statut du test
'------------------------------------------------------------------------------
Public Function GetTestStatus(ByVal testIndex As Long) As String
    If testIndex <= 0 Or testIndex > this.Tests.Count Then
        GetTestStatus = ""
        Exit Function
    End If
    
    Dim testInfo As TTestInfo
    testInfo = this.Tests(testIndex)
    
    GetTestStatus = testInfo.Status
End Function

'------------------------------------------------------------------------------
' Récupère la durée d'exécution d'un test à l'index spécifié
'
' @param testIndex Index du test
' @return Double Durée d'exécution en millisecondes
'------------------------------------------------------------------------------
Public Function GetTestDuration(ByVal testIndex As Long) As Double
    If testIndex <= 0 Or testIndex > this.Tests.Count Then
        GetTestDuration = 0
        Exit Function
    End If
    
    Dim testInfo As TTestInfo
    testInfo = this.Tests(testIndex)
    
    GetTestDuration = testInfo.Duration
End Function

'------------------------------------------------------------------------------
' Récupère le message d'erreur d'un test à l'index spécifié
'
' @param testIndex Index du test
' @return String Message d'erreur
'------------------------------------------------------------------------------
Public Function GetTestErrorMessage(ByVal testIndex As Long) As String
    If testIndex <= 0 Or testIndex > this.Tests.Count Then
        GetTestErrorMessage = ""
        Exit Function
    End If
    
    Dim testInfo As TTestInfo
    testInfo = this.Tests(testIndex)
    
    GetTestErrorMessage = testInfo.ErrorMessage
End Function

'------------------------------------------------------------------------------
' Journalise un message si le mode verbeux est activé
'
' @param message Message à journaliser
'------------------------------------------------------------------------------
Private Sub LogVerbose(ByVal message As String)
    If this.VerboseMode Then
        LogMessage message
    End If
End Sub

'------------------------------------------------------------------------------
' Journalise un message
'
' @param message Message à journaliser
'------------------------------------------------------------------------------
Private Sub LogMessage(ByVal message As String)
    Debug.Print "[" & this.SuiteName & "] " & message
End Sub

'------------------------------------------------------------------------------
' Vérifie si une chaîne correspond à un motif wildcard
'
' @param str Chaîne à vérifier
' @param pattern Motif wildcard (* = plusieurs caractères, ? = un caractère)
' @return Boolean True si la chaîne correspond au motif
'------------------------------------------------------------------------------
Private Function MatchesWildcard(ByVal str As String, ByVal pattern As String) As Boolean
    ' Fonction simplifiée - à améliorer avec une implémentation plus robuste si nécessaire
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Convertir le motif wildcard en expression régulière
    Dim regexPattern As String
    regexPattern = "^" & Replace(Replace(pattern, "\", "\\"), ".", "\.") & "$"
    regexPattern = Replace(regexPattern, "?", ".")
    regexPattern = Replace(regexPattern, "*", ".*")
    
    regex.Pattern = regexPattern
    regex.IgnoreCase = True
    
    MatchesWildcard = regex.Test(str)
End Function

'------------------------------------------------------------------------------
' Fonction utilitaire pour la division sécurisée (évite division par zéro)
'
' @param numerator Numérateur
' @param denominator Dénominateur
' @return Double Résultat de la division ou 0 si dénominateur = 0
'------------------------------------------------------------------------------
Private Function SafeDiv(ByVal numerator As Double, ByVal denominator As Double) As Double
    If denominator = 0 Then
        SafeDiv = 0
    Else
        SafeDiv = numerator / denominator
    End If
End Function 