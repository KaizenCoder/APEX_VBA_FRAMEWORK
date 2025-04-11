' Migrated to apex-core/testing - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Attribute VB_Name = "modTestRunner"
Option Explicit
' ==========================================================================
' Module : modTestRunner
' Version : 2.0
' Purpose : Module principal d'automatisation des tests
' Date    : 10/04/2025
' ==========================================================================

' --- Constantes ---
Private Const MODULE_NAME As String = "modTestRunner"
Private Const REPORT_FOLDER As String = "TestReports\"
Private Const MAX_EXECUTE_TIME_MS As Long = 30000 ' 30 secondes max par test
Private Const CONFIG_FILE As String = "config\test_config.ini"

' --- Déclarations globales ---
Private m_suites As Collection
Private m_runningTest As Boolean
Private m_stopTests As Boolean
Private m_currentTestName As String
Private m_perfResults As Collection
Private m_configManager As Object
Private m_logger As Object

' --- Enregistrement des suites de test ---
Public Sub RegisterTestSuite(suite As clsTestSuite)
    If m_suites Is Nothing Then Set m_suites = New Collection
    On Error Resume Next
    m_suites.Add suite, suite.Name
    If Err.Number <> 0 Then
        LogMessage "AVERTISSEMENT: La suite '" & suite.Name & "' est déjà enregistrée", "warning"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' --- Exécution des tests ---
Public Sub RunAllTests(Optional ByVal outputReport As Boolean = True, Optional ByVal stopOnFailure As Boolean = False)
    If m_runningTest Then
        MsgBox "Des tests sont déjà en cours d'exécution.", vbExclamation, "Tests en cours"
        Exit Sub
    End If
    
    m_runningTest = True
    m_stopTests = False
    
    ' Initialiser la configuration et le logger
    InitializeConfig
    
    ' Vérifier si les tests sont activés
    If Not GetConfigBool("General", "EnableTests", True) Then
        LogMessage "Les tests sont désactivés dans la configuration. Modifiez le fichier " & CONFIG_FILE & " pour les activer.", "warning"
        m_runningTest = False
        Exit Sub
    End If
    
    ' Initialiser la collection des résultats de performance
    Set m_perfResults = New Collection
    
    Dim startTime As Double
    startTime = Timer
    
    ' Assurer l'existence du répertoire de rapports si nécessaire
    If outputReport Then EnsureReportFolder
    
    Dim suite As clsTestSuite
    Dim totalSuites As Long
    Dim totalTests As Long
    Dim totalPassed As Long
    Dim totalFailed As Long
    Dim totalSkipped As Long
    Dim suiteResults As String
    Dim allResults As String
    
    ' Log d'en-tête
    LogMessage "==========================================================", "info"
    LogMessage "DÉBUT DE L'EXÉCUTION DE TOUS LES TESTS: " & Format(Now, "yyyy-mm-dd hh:mm:ss"), "info"
    LogMessage "==========================================================", "info"
    
    If m_suites Is Nothing Or m_suites.Count = 0 Then
        LogMessage "AUCUNE SUITE DE TEST ENREGISTRÉE", "warning"
        If outputReport Then
            WriteReportFile "AUCUNE SUITE DE TEST ENREGISTRÉE", "AllTests"
        End If
        m_runningTest = False
        Exit Sub
    End If
    
    ' Filtrer les suites à exécuter si nécessaire
    Dim suitesToRun As Collection
    Set suitesToRun = GetSuitesToRun(m_suites)
    
    ' Exécuter chaque suite
    For Each suite In suitesToRun
        suiteResults = RunSuite(suite, stopOnFailure)
        allResults = allResults & suiteResults & vbCrLf & vbCrLf
        
        totalSuites = totalSuites + 1
        totalTests = totalTests + suite.TestCount
        totalPassed = totalPassed + suite.PassedCount
        totalFailed = totalFailed + suite.FailedCount
        totalSkipped = totalSkipped + suite.SkippedCount
        
        If m_stopTests Then Exit For
    Next suite
    
    ' Log de résumé
    LogMessage "==========================================================", "info"
    LogMessage "RÉSUMÉ DES TESTS:", "info"
    LogMessage "  Suites: " & totalSuites, "info"
    LogMessage "  Tests: " & totalTests, "info"
    LogMessage "  Réussis: " & totalPassed & " (" & Format(IIf(totalTests > 0, totalPassed / totalTests * 100, 0), "0.00") & "%)", "info"
    LogMessage "  Échoués: " & totalFailed, "info"
    LogMessage "  Ignorés: " & totalSkipped, "info"
    LogMessage "  Temps total: " & Format(Timer - startTime, "0.000") & " secondes", "info"
    LogMessage "==========================================================", "info"
    
    ' Afficher les résultats de performance
    If m_perfResults.Count > 0 And GetConfigBool("Performance", "LogPerformance", True) Then
        LogMessage vbCrLf & "TOP 5 DES TESTS LES PLUS LENTS:", "info"
        OutputTopSlowTests GetConfigInt("Reporting", "Top5SlowTests", 5)
    End If
    
    ' Générer rapport si demandé
    If outputReport And GetConfigBool("Reporting", "GenerateReports", True) Then
        Dim summary As String
        summary = "RÉSUMÉ DES TESTS:" & vbCrLf & _
                 "  Suites: " & totalSuites & vbCrLf & _
                 "  Tests: " & totalTests & vbCrLf & _
                 "  Réussis: " & totalPassed & " (" & Format(IIf(totalTests > 0, totalPassed / totalTests * 100, 0), "0.00") & "%)" & vbCrLf & _
                 "  Échoués: " & totalFailed & vbCrLf & _
                 "  Ignorés: " & totalSkipped & vbCrLf & _
                 "  Temps total: " & Format(Timer - startTime, "0.000") & " secondes"
                 
        ' Ajouter les tests les plus lents
        If m_perfResults.Count > 0 And GetConfigBool("Reporting", "Top5SlowTests", True) Then
            summary = summary & vbCrLf & vbCrLf & "TOP 5 DES TESTS LES PLUS LENTS:" & vbCrLf
            summary = summary & GetTopSlowTestsText(5)
        End If
        
        allResults = allResults & vbCrLf & summary
        WriteReportFile allResults, "AllTests"
        
        ' Envoi par email si configuré
        If GetConfigBool("Email", "SendReportByEmail", False) Then
            SendReportByEmail summary
        End If
    End If
    
    ' Terminer
    m_runningTest = False
    Set m_perfResults = Nothing
End Sub

' --- Exécution d'une suite de tests ---
Public Function RunSuite(suite As clsTestSuite, Optional ByVal stopOnFailure As Boolean = False) As String
    Dim result As String
    Dim startTime As Double
    startTime = Timer
    
    LogMessage "DÉBUT DE LA SUITE: " & suite.Name, "info"
    result = "SUITE DE TEST: " & suite.Name & vbCrLf
    
    ' Exécuter les tests de la suite
    suite.RunAllTests stopOnFailure
    
    ' Afficher le résumé de la suite
    LogMessage "  Tests: " & suite.TestCount, "info"
    LogMessage "  Réussis: " & suite.PassedCount & " (" & Format(IIf(suite.TestCount > 0, suite.PassedCount / suite.TestCount * 100, 0), "0.00") & "%)", "info"
    LogMessage "  Échoués: " & suite.FailedCount, "info"
    LogMessage "  Ignorés: " & suite.SkippedCount, "info"
    LogMessage "  Temps: " & Format(Timer - startTime, "0.000") & " secondes", "info"
    
    ' Ajouter au résultat textuel
    result = result & "  Tests: " & suite.TestCount & vbCrLf & _
             "  Réussis: " & suite.PassedCount & " (" & Format(IIf(suite.TestCount > 0, suite.PassedCount / suite.TestCount * 100, 0), "0.00") & "%)" & vbCrLf & _
             "  Échoués: " & suite.FailedCount & vbCrLf & _
             "  Ignorés: " & suite.SkippedCount & vbCrLf & _
             "  Temps: " & Format(Timer - startTime, "0.000") & " secondes" & vbCrLf
    
    ' Ajouter les détails des tests échoués
    If suite.FailedCount > 0 Then
        result = result & vbCrLf & "DÉTAILS DES ÉCHECS:" & vbCrLf & suite.GetFailureDetails
    End If
    
    ' Indiquer si on doit arrêter les tests
    If suite.FailedCount > 0 And stopOnFailure Then
        LogMessage "ARRÊT DES TESTS DEMANDÉ EN RAISON D'UN ÉCHEC", "warning"
        m_stopTests = True
    End If
    
    RunSuite = result
End Function

' --- Exécution d'un test unitaire ---
Public Function RunTest(testName As String, testProc As String, _
                        Optional callingModule As String = "", _
                        Optional timeout As Long = 0, _
                        Optional expectedExceptions As String = "") As clsTestResult
    Dim result As New clsTestResult
    Dim startTime As Double
    Dim elapsedTime As Double
    Dim perfEntry As String
    Dim actualTimeout As Long
    
    ' Initialiser le résultat
    result.TestName = testName
    result.Success = False ' Par défaut, considérer que le test échoue
    
    ' Capturer le nom du test en cours
    m_currentTestName = testName
    
    ' Déterminer le timeout à utiliser
    If timeout <= 0 Then
        actualTimeout = GetConfigInt("Performance", "MaxTestDurationSeconds", 30) * 1000
    Else
        actualTimeout = timeout
    End If
    
    ' Log de début de test si en mode debug
    If GetConfigBool("Debug", "VerboseLogging", False) Then
        LogMessage "Début d'exécution du test: " & testName, "debug"
    End If
    
    On Error Resume Next
    startTime = Timer
    
    ' Exécuter le test avec protection contre les boucles infinies
    Application.OnTime Now + TimeSerial(0, 0, actualTimeout / 1000), "TestTimeoutHandler"
    
    If callingModule <> "" Then
        Application.Run callingModule & "." & testProc
    Else
        Application.Run testProc
    End If
    
    ' Vérifier si une erreur s'est produite
    Dim errorOccurred As Boolean
    Dim errorWasExpected As Boolean
    
    errorOccurred = (Err.Number <> 0)
    errorWasExpected = IsErrorExpected(Err.Number, expectedExceptions)
    
    ' Annuler le délai d'attente si le test s'est terminé normalement
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, actualTimeout / 1000), "TestTimeoutHandler", , False
    
    If errorOccurred Then
        ' Vérifier si l'erreur était attendue
        If errorWasExpected Then
            result.Success = True
            result.ErrorMessage = "Exception attendue: " & Err.Description & " (Code: " & Err.Number & ")"
        Else
            result.Success = False
            result.ErrorMessage = "Erreur d'exécution: " & Err.Description & " (Code: " & Err.Number & ")"
            result.ErrorSource = IIf(Err.Source <> "", Err.Source, callingModule & "." & testProc)
        End If
        Err.Clear
    Else
        ' Si aucune erreur ne s'est produite mais qu'on en attendait une
        If expectedExceptions <> "" Then
            result.Success = False
            result.ErrorMessage = "Exception attendue non déclenchée: " & expectedExceptions
            result.ErrorSource = callingModule & "." & testProc
        Else
            ' Si pas d'erreur, le test est réussi
            result.Success = True
        End If
    End If
    
    ' Calculer le temps d'exécution
    elapsedTime = Timer - startTime
    result.ExecutionTime = elapsedTime
    
    ' Vérifier si le test est lent
    Dim perfThreshold As Long
    perfThreshold = GetConfigInt("Performance", "PerformanceThresholdMs", 1000) / 1000
    
    If elapsedTime > perfThreshold Then
        result.IsPerformanceIssue = True
        If GetConfigBool("Debug", "VerboseLogging", False) Then
            LogMessage "Test lent détecté: " & testName & " (" & Format(elapsedTime, "0.000") & " sec)", "warning"
        End If
    End If
    
    ' Enregistrer les performances
    perfEntry = testName & "|" & Format(elapsedTime, "0.000") & "|" & IIf(result.Success, "Réussi", "Échoué")
    m_perfResults.Add perfEntry
    
    ' Log de fin de test si en mode debug
    If GetConfigBool("Debug", "VerboseLogging", False) Then
        LogMessage "Fin d'exécution du test: " & testName & " - " & IIf(result.Success, "RÉUSSI", "ÉCHEC"), _
                  IIf(result.Success, "info", "error")
    End If
    
    ' Réinitialiser le nom du test en cours
    m_currentTestName = ""
    
    Set RunTest = result
    On Error GoTo 0
End Function

' --- Gestionnaire de timeout ---
Public Sub TestTimeoutHandler()
    If m_currentTestName <> "" Then
        LogMessage "TIMEOUT: Le test '" & m_currentTestName & "' a dépassé le temps maximum d'exécution", "error"
    End If
End Sub

' --- Initialisation de la configuration ---
Private Sub InitializeConfig()
    On Error Resume Next
    
    ' Essayer de récupérer le gestionnaire de configuration
    Set m_configManager = CreateObject("APEX.ConfigManager")
    If Err.Number <> 0 Then
        Set m_configManager = Nothing
        Err.Clear
    End If
    
    ' Essayer de récupérer le logger
    Set m_logger = CreateObject("APEX.Logger")
    If Err.Number <> 0 Then
        Set m_logger = Nothing
        Err.Clear
    End If
    
    On Error GoTo 0
End Sub

' --- Lecture de la configuration ---
Private Function GetConfigString(section As String, key As String, defaultValue As String) As String
    On Error Resume Next
    Dim result As String
    
    If Not m_configManager Is Nothing Then
        result = m_configManager.GetSetting(section, key, defaultValue)
    Else
        ' Lecture directe du fichier INI si le ConfigManager n'est pas disponible
        result = GetINISetting(CONFIG_FILE, section, key, defaultValue)
    End If
    
    If Err.Number <> 0 Then
        result = defaultValue
        Err.Clear
    End If
    
    GetConfigString = result
    On Error GoTo 0
End Function

Private Function GetConfigInt(section As String, key As String, defaultValue As Long) As Long
    On Error Resume Next
    Dim result As Long
    
    If Not m_configManager Is Nothing Then
        result = CLng(m_configManager.GetSetting(section, key, CStr(defaultValue)))
    Else
        ' Lecture directe du fichier INI si le ConfigManager n'est pas disponible
        result = CLng(GetINISetting(CONFIG_FILE, section, key, CStr(defaultValue)))
    End If
    
    If Err.Number <> 0 Then
        result = defaultValue
        Err.Clear
    End If
    
    GetConfigInt = result
    On Error GoTo 0
End Function

Private Function GetConfigBool(section As String, key As String, defaultValue As Boolean) As Boolean
    On Error Resume Next
    Dim result As String
    
    If Not m_configManager Is Nothing Then
        result = m_configManager.GetSetting(section, key, IIf(defaultValue, "True", "False"))
    Else
        ' Lecture directe du fichier INI si le ConfigManager n'est pas disponible
        result = GetINISetting(CONFIG_FILE, section, key, IIf(defaultValue, "True", "False"))
    End If
    
    If Err.Number <> 0 Then
        GetConfigBool = defaultValue
        Err.Clear
        Exit Function
    End If
    
    ' Conversion en booléen
    result = UCase(Trim(result))
    GetConfigBool = (result = "TRUE" Or result = "YES" Or result = "1")
    
    On Error GoTo 0
End Function

' --- Utilitaires ---
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

Private Sub EnsureReportFolder()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folderPath As String
    folderPath = GetConfigString("General", "ReportFolder", REPORT_FOLDER)
    
    If Not fso.FolderExists(folderPath) Then
        On Error Resume Next
        fso.CreateFolder folderPath
        If Err.Number <> 0 Then
            LogMessage "AVERTISSEMENT: Impossible de créer le dossier de rapports: " & Err.Description, "warning"
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Sub

Private Sub WriteReportFile(content As String, reportName As String)
    Dim fso As Object
    Dim ts As Object
    Dim fileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Créer un nom de fichier unique avec horodatage
    Dim folderPath As String
    folderPath = GetConfigString("General", "ReportFolder", REPORT_FOLDER)
    
    fileName = folderPath & "\" & reportName & "_" & Format(Now, "yyyymmdd_hhmmss") & "." & _
               LCase(GetConfigString("Reporting", "ReportFormat", "TXT"))
    
    On Error Resume Next
    Set ts = fso.CreateTextFile(fileName, True)
    If Err.Number <> 0 Then
        LogMessage "ERREUR: Impossible de créer le fichier de rapport: " & Err.Description, "error"
        Err.Clear
        Exit Sub
    End If
    
    ' Écrire l'en-tête
    ts.WriteLine "RAPPORT DE TESTS APEX VBA FRAMEWORK"
    ts.WriteLine "Date: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    ts.WriteLine "=============================================="
    ts.WriteLine ""
    
    ' Écrire le contenu
    ts.Write content
    
    ' Fermer le fichier
    ts.Close
    LogMessage "Rapport de test enregistré: " & fileName, "info"
    On Error GoTo 0
End Sub

Private Function GetSuitesToRun(allSuites As Collection) As Collection
    Dim result As New Collection
    
    ' Vérifier si on doit exécuter toutes les suites
    If GetConfigBool("TestSelection", "RunAllTests", True) Then
        Set GetSuitesToRun = allSuites
        Exit Function
    End If
    
    ' Sinon, filtrer en fonction des suites sélectionnées
    Dim selectedSuites As String
    selectedSuites = GetConfigString("TestSelection", "SelectedSuites", "")
    
    If Trim(selectedSuites) = "" Then
        ' Si aucune suite n'est spécifiée, exécuter toutes les suites quand même
        Set GetSuitesToRun = allSuites
        Exit Function
    End If
    
    ' Parcourir les suites sélectionnées
    Dim suiteArray() As String
    Dim i As Long
    Dim suite As clsTestSuite
    
    suiteArray = Split(selectedSuites, ",")
    
    For i = 0 To UBound(suiteArray)
        Dim suiteName As String
        suiteName = Trim(suiteArray(i))
        
        On Error Resume Next
        Set suite = allSuites(suiteName)
        If Err.Number = 0 Then
            result.Add suite, suiteName
        Else
            LogMessage "Suite de test introuvable: " & suiteName, "warning"
            Err.Clear
        End If
        On Error GoTo 0
    Next i
    
    If result.Count = 0 Then
        LogMessage "Aucune suite de test valide trouvée dans la sélection. Exécution de toutes les suites.", "warning"
        Set GetSuitesToRun = allSuites
    Else
        Set GetSuitesToRun = result
    End If
End Function

Private Function IsErrorExpected(errNumber As Long, expectedExceptions As String) As Boolean
    If Trim(expectedExceptions) = "" Then
        IsErrorExpected = False
        Exit Function
    End If
    
    Dim expArray() As String
    Dim i As Long
    
    expArray = Split(expectedExceptions, ",")
    
    For i = 0 To UBound(expArray)
        Dim expErr As Long
        
        On Error Resume Next
        expErr = CLng(Trim(expArray(i)))
        If Err.Number = 0 And expErr = errNumber Then
            IsErrorExpected = True
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    Next i
    
    IsErrorExpected = False
End Function

' --- Fonctions d'analyse de performance ---
Private Sub OutputTopSlowTests(ByVal topCount As Long)
    Dim i As Long
    Dim perfArray() As String
    Dim parts() As String
    Dim temp As String
    Dim j As Long, k As Long
    
    ' Convertir la collection en tableau pour le tri
    ReDim perfArray(1 To m_perfResults.Count)
    For i = 1 To m_perfResults.Count
        perfArray(i) = m_perfResults(i)
    Next i
    
    ' Tri à bulles simple par temps d'exécution (décroissant)
    For j = 1 To UBound(perfArray) - 1
        For k = j + 1 To UBound(perfArray)
            parts = Split(perfArray(j), "|")
            Dim time1 As Double
            time1 = CDbl(parts(1))
            
            parts = Split(perfArray(k), "|")
            Dim time2 As Double
            time2 = CDbl(parts(1))
            
            If time1 < time2 Then
                temp = perfArray(j)
                perfArray(j) = perfArray(k)
                perfArray(k) = temp
            End If
        Next k
    Next j
    
    ' Afficher les N plus lents
    For i = 1 To IIf(topCount < UBound(perfArray), topCount, UBound(perfArray))
        parts = Split(perfArray(i), "|")
        LogMessage "  " & i & ". " & parts(0) & " - " & parts(1) & " sec (" & parts(2) & ")", "info"
    Next i
End Sub

Private Function GetTopSlowTestsText(ByVal topCount As Long) As String
    Dim i As Long
    Dim perfArray() As String
    Dim parts() As String
    Dim temp As String
    Dim j As Long, k As Long
    Dim result As String
    
    ' Convertir la collection en tableau pour le tri
    ReDim perfArray(1 To m_perfResults.Count)
    For i = 1 To m_perfResults.Count
        perfArray(i) = m_perfResults(i)
    Next i
    
    ' Tri à bulles simple par temps d'exécution (décroissant)
    For j = 1 To UBound(perfArray) - 1
        For k = j + 1 To UBound(perfArray)
            parts = Split(perfArray(j), "|")
            Dim time1 As Double
            time1 = CDbl(parts(1))
            
            parts = Split(perfArray(k), "|")
            Dim time2 As Double
            time2 = CDbl(parts(1))
            
            If time1 < time2 Then
                temp = perfArray(j)
                perfArray(j) = perfArray(k)
                perfArray(k) = temp
            End If
        Next k
    Next j
    
    ' Formater le texte
    For i = 1 To IIf(topCount < UBound(perfArray), topCount, UBound(perfArray))
        parts = Split(perfArray(i), "|")
        result = result & "  " & i & ". " & parts(0) & " - " & parts(1) & " sec (" & parts(2) & ")" & vbCrLf
    Next i
    
    GetTopSlowTestsText = result
End Function

Private Sub SendReportByEmail(ByVal reportContent As String)
    On Error Resume Next
    
    Dim olApp As Object
    Dim olMail As Object
    Dim recipients As String
    Dim subject As String
    
    recipients = GetConfigString("Email", "EmailRecipients", "")
    subject = GetConfigString("Email", "EmailSubject", "Rapport de tests APEX Framework")
    
    If Trim(recipients) = "" Then
        LogMessage "Pas de destinataires configurés pour l'envoi du rapport par email", "warning"
        Exit Sub
    End If
    
    ' Créer l'objet Outlook
    Set olApp = CreateObject("Outlook.Application")
    If Err.Number <> 0 Then
        LogMessage "Impossible de démarrer Outlook pour l'envoi du rapport: " & Err.Description, "error"
        Err.Clear
        Exit Sub
    End If
    
    ' Créer le message
    Set olMail = olApp.CreateItem(0) ' olMailItem = 0
    
    With olMail
        .To = recipients
        .Subject = subject & " - " & Format(Now, "yyyy-mm-dd")
        .Body = reportContent
        .Send
    End With
    
    If Err.Number <> 0 Then
        LogMessage "Erreur lors de l'envoi du rapport par email: " & Err.Description, "error"
        Err.Clear
    Else
        LogMessage "Rapport envoyé par email à: " & recipients, "info"
    End If
    
    Set olMail = Nothing
    Set olApp = Nothing
    
    On Error GoTo 0
End Sub

' --- Lecture directe des fichiers INI ---
Private Function GetINISetting(ByVal filePath As String, ByVal section As String, ByVal key As String, ByVal defaultValue As String) As String
    Dim result As String
    result = Space(255)
    
    ' Appel de l'API Windows pour lire le fichier INI
    Dim length As Long
    length = GetPrivateProfileString(section, key, defaultValue, result, Len(result), filePath)
    
    If length = 0 Then
        GetINISetting = defaultValue
    Else
        GetINISetting = Left(result, length)
    End If
End Function

' --- Déclarations API ---
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

' --- Initialisation et nettoyage ---
Private Sub Class_Initialize()
    Set m_suites = New Collection
    m_runningTest = False
    m_stopTests = False
    InitializeConfig
End Sub

Private Sub Class_Terminate()
    Set m_suites = Nothing
    Set m_perfResults = Nothing
    Set m_configManager = Nothing
    Set m_logger = Nothing
End Sub 