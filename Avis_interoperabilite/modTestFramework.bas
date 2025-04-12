' modTestFramework.bas
' Module principal du framework de test pour l'architecture d'interopérabilité Apex-Excel
'
' Ce module fournit les fonctions principales pour configurer et exécuter des tests,
' ainsi que des utilitaires pour la gestion des environnements de test.
'
' @module modTestFramework
' @author APEX Framework Team
' @version 1.0
' @date 2024-07-27
Option Explicit

' ==========================================================================
' Module           : modTestFramework
' Version          : 1.0
' Description      : Module principal du framework de test APEX
' Auteur           : Équipe APEX
' Date             : 2024-07-27
' ==========================================================================

' ----------------------------------------
' Constantes publiques du framework de test
' ----------------------------------------
Public Const TEST_LEVEL_UNIT As String = "UNIT"
Public Const TEST_LEVEL_INTEGRATION As String = "INTEGRATION"
Public Const TEST_LEVEL_SYSTEM As String = "SYSTEM"
Public Const TEST_LEVEL_PERFORMANCE As String = "PERFORMANCE"

Public Const TEST_RESULT_PASS As String = "PASS"
Public Const TEST_RESULT_FAIL As String = "FAIL"
Public Const TEST_RESULT_SKIP As String = "SKIP"
Public Const TEST_RESULT_ERROR As String = "ERROR"

Public Const FORMAT_TEXT As Integer = 0
Public Const FORMAT_CSV As Integer = 1
Public Const FORMAT_MARKDOWN As Integer = 2
Public Const FORMAT_HTML As Integer = 3

' ----------------------------------------
' Constantes privées
' ----------------------------------------
Private Const MODULE_NAME As String = "modTestFramework"
Private Const ERR_PROC_NOT_FOUND As Long = 1000
Private Const ERR_INVALID_ARGUMENT As Long = 1001
Private Const ERR_INTERNAL_ERROR As Long = 1002

' ----------------------------------------
' Types de données pour test
' ----------------------------------------
Public Type TTestResult
    TestName As String
    TestModule As String
    Description As String
    Level As String
    StartTime As Double
    EndTime As Double
    Duration As Double
    Result As String
    ErrorMessage As String
    Tags As String
End Type

' ----------------------------------------
' Configuration globale pour le framework
' ----------------------------------------
Private mConfig As Object ' Dictionary
Private mInitialized As Boolean

' ----------------------------------------
' Méthodes d'initialisation
' ----------------------------------------

' Initialise le framework de test avec les paramètres par défaut
Public Sub InitializeTestFramework()
    On Error GoTo ErrorHandler
    
    ' Vérifier si déjà initialisé
    If mInitialized Then Exit Sub
    
    ' Créer le dictionnaire de configuration
    Set mConfig = CreateObject("Scripting.Dictionary")
    
    ' Paramètres par défaut
    mConfig.Add "Verbose", False
    mConfig.Add "StopOnFailure", False
    mConfig.Add "LogToFile", False
    mConfig.Add "LogPath", ""
    mConfig.Add "DefaultTimeout", 30 ' secondes
    mConfig.Add "TestFilter", ""
    mConfig.Add "TestLevel", TEST_LEVEL_UNIT
    mConfig.Add "OutputFormat", FORMAT_TEXT
    
    mInitialized = True
    
    Exit Sub
ErrorHandler:
    Err.Raise ERR_INTERNAL_ERROR, MODULE_NAME & ".InitializeTestFramework", _
              "Erreur lors de l'initialisation du framework de test: " & Err.Description
End Sub

' Configure un paramètre du framework
Public Sub SetTestConfig(ByVal configName As String, ByVal configValue As Variant)
    On Error GoTo ErrorHandler
    
    ' S'assurer que le framework est initialisé
    If Not mInitialized Then InitializeTestFramework
    
    ' Vérifier si la clé existe déjà
    If mConfig.Exists(configName) Then
        mConfig(configName) = configValue
    Else
        mConfig.Add configName, configValue
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise ERR_INTERNAL_ERROR, MODULE_NAME & ".SetTestConfig", _
              "Erreur lors de la configuration du paramètre '" & configName & "': " & Err.Description
End Sub

' Récupère un paramètre de configuration du framework
Public Function GetTestConfig(ByVal configName As String) As Variant
    On Error GoTo ErrorHandler
    
    ' S'assurer que le framework est initialisé
    If Not mInitialized Then InitializeTestFramework
    
    ' Vérifier si la clé existe
    If mConfig.Exists(configName) Then
        GetTestConfig = mConfig(configName)
    Else
        Err.Raise ERR_INVALID_ARGUMENT, , "Le paramètre de configuration '" & configName & "' n'existe pas"
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise ERR_INTERNAL_ERROR, MODULE_NAME & ".GetTestConfig", _
              "Erreur lors de la récupération du paramètre '" & configName & "': " & Err.Description
End Function

' ----------------------------------------
' Exécution dynamique de procédures
' ----------------------------------------

' Exécute une procédure par son nom et son module
Public Function RunProcedureByName(ByVal procName As String, ByVal moduleName As String) As Variant
    On Error GoTo ErrorHandler
    
    ' Variables pour manipuler l'objet de code
    Dim vbComp As Object
    Dim vbProj As Object
    Dim procFound As Boolean
    
    procFound = False
    
    ' Obtenir le projet VBA courant
    Set vbProj = Application.VBE.ActiveVBProject
    
    ' Parcourir les composants du projet
    For Each vbComp In vbProj.VBComponents
        If vbComp.Name = moduleName Then
            ' Le module a été trouvé
            Application.Run moduleName & "." & procName
            procFound = True
            Exit For
        End If
    Next vbComp
    
    ' Vérifier si la procédure a été trouvée
    If Not procFound Then
        Err.Raise ERR_PROC_NOT_FOUND, , "Procédure ou module non trouvé: " & moduleName & "." & procName
    End If
    
    Exit Function
ErrorHandler:
    If Err.Number = ERR_PROC_NOT_FOUND Then
        Err.Raise ERR_PROC_NOT_FOUND, MODULE_NAME & ".RunProcedureByName", _
                 "La procédure '" & procName & "' dans le module '" & moduleName & "' n'a pas été trouvée"
    Else
        Err.Raise ERR_INTERNAL_ERROR, MODULE_NAME & ".RunProcedureByName", _
                 "Erreur lors de l'exécution de '" & procName & "' dans '" & moduleName & "': " & Err.Description
    End If
End Function

' Vérifie si une procédure existe dans un module donné
Public Function ProcedureExists(ByVal procName As String, ByVal moduleName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Variables pour manipuler l'objet de code
    Dim vbComp As Object
    Dim vbProj As Object
    Dim procFound As Boolean
    Dim codeModule As Object
    Dim procKind As Integer
    Dim lineNum As Long
    
    procFound = False
    
    ' Obtenir le projet VBA courant
    Set vbProj = Application.VBE.ActiveVBProject
    
    ' Parcourir les composants du projet
    For Each vbComp In vbProj.VBComponents
        If vbComp.Name = moduleName Then
            ' Le module a été trouvé, vérifier si la procédure existe
            Set codeModule = vbComp.CodeModule
            
            ' Tenter de trouver la procédure
            On Error Resume Next
            lineNum = codeModule.ProcStartLine(procName, 0) ' 0 = vbext_pk_Proc
            If Err.Number = 0 And lineNum > 0 Then
                procFound = True
            End If
            On Error GoTo ErrorHandler
            
            Exit For
        End If
    Next vbComp
    
    ProcedureExists = procFound
    Exit Function
    
ErrorHandler:
    ProcedureExists = False
End Function

' ----------------------------------------
' Méthodes auxiliaires pour les tests
' ----------------------------------------

' Écrit dans le journal des tests si le mode verbose est activé
Public Sub TestLog(ByVal message As String, Optional ByVal level As String = "INFO")
    On Error Resume Next
    
    ' S'assurer que le framework est initialisé
    If Not mInitialized Then InitializeTestFramework
    
    ' Obtenir le mode verbose
    Dim isVerbose As Boolean
    isVerbose = mConfig("Verbose")
    
    ' Écrire le message si en mode verbose ou si c'est un message d'erreur
    If isVerbose Or level = "ERROR" Then
        Debug.Print Format(Now, "yyyy-mm-dd hh:nn:ss") & " [" & level & "] " & message
    End If
    
    ' Si la journalisation dans un fichier est activée
    If mConfig("LogToFile") Then
        LogToFile level & ": " & message
    End If
End Sub

' Journalise un message dans un fichier
Private Sub LogToFile(ByVal message As String)
    On Error Resume Next
    
    Dim logPath As String
    logPath = mConfig("LogPath")
    
    ' Vérifier si un chemin de fichier journal est défini
    If logPath = "" Then
        logPath = ThisWorkbook.Path & "\test_log.txt"
    End If
    
    ' Écrire dans le fichier journal
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open logPath For Append As #fileNum
    Print #fileNum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & message
    Close #fileNum
End Sub

' Vérifie si un test correspond aux filtres actuels
Public Function TestMatchesFilter(ByVal testName As String, ByVal testModule As String, _
                               ByVal level As String, ByVal tags As String) As Boolean
    On Error Resume Next
    
    ' S'assurer que le framework est initialisé
    If Not mInitialized Then InitializeTestFramework
    
    ' Obtenir les filtres actuels
    Dim testFilter As String
    Dim testLevel As String
    
    testFilter = mConfig("TestFilter")
    testLevel = mConfig("TestLevel")
    
    ' Si pas de filtre, accepter tous les tests
    If testFilter = "" And (testLevel = "" Or level = "") Then
        TestMatchesFilter = True
        Exit Function
    End If
    
    ' Vérifier le niveau de test
    If testLevel <> "" And level <> "" Then
        If testLevel <> level Then
            TestMatchesFilter = False
            Exit Function
        End If
    End If
    
    ' Vérifier les filtres de tag ou de nom
    If testFilter <> "" Then
        ' Vérifier si le filtre correspond au nom du test
        If InStr(1, testName, testFilter, vbTextCompare) > 0 Then
            TestMatchesFilter = True
            Exit Function
        End If
        
        ' Vérifier si le filtre correspond à un tag
        If InStr(1, tags, testFilter, vbTextCompare) > 0 Then
            TestMatchesFilter = True
            Exit Function
        End If
        
        ' Le filtre ne correspond pas
        TestMatchesFilter = False
    Else
        ' Pas de filtre de tag/nom, match si le niveau correspond
        TestMatchesFilter = True
    End If
End Function

' Formate un résultat de test selon le format de sortie configuré
Public Function FormatTestResult(ByVal result As TTestResult) As String
    On Error Resume Next
    
    ' S'assurer que le framework est initialisé
    If Not mInitialized Then InitializeTestFramework
    
    ' Obtenir le format de sortie
    Dim outputFormat As Integer
    outputFormat = mConfig("OutputFormat")
    
    ' Formater selon le format demandé
    Select Case outputFormat
        Case FORMAT_TEXT
            FormatTestResult = FormatTestResultAsText(result)
        Case FORMAT_CSV
            FormatTestResult = FormatTestResultAsCSV(result)
        Case FORMAT_MARKDOWN
            FormatTestResult = FormatTestResultAsMarkdown(result)
        Case FORMAT_HTML
            FormatTestResult = FormatTestResultAsHTML(result)
        Case Else
            FormatTestResult = FormatTestResultAsText(result)
    End Select
End Function

' Formate un résultat au format texte
Private Function FormatTestResultAsText(ByVal result As TTestResult) As String
    Dim output As String
    
    output = "Test: " & result.TestName & " [" & result.TestModule & "]" & vbCrLf & _
             "Description: " & result.Description & vbCrLf & _
             "Niveau: " & result.Level & vbCrLf & _
             "Résultat: " & result.Result & vbCrLf & _
             "Durée: " & Format(result.Duration, "0.000") & " secondes" & vbCrLf
            
    If result.ErrorMessage <> "" Then
        output = output & "Erreur: " & result.ErrorMessage & vbCrLf
    End If
    
    If result.Tags <> "" Then
        output = output & "Tags: " & result.Tags & vbCrLf
    End If
    
    output = output & String(40, "-") & vbCrLf
    
    FormatTestResultAsText = output
End Function

' Formate un résultat au format CSV
Private Function FormatTestResultAsCSV(ByVal result As TTestResult) As String
    FormatTestResultAsCSV = _
        """" & result.TestName & """," & _
        """" & result.TestModule & """," & _
        """" & result.Description & """," & _
        """" & result.Level & """," & _
        Format(result.StartTime, "yyyy-mm-dd hh:nn:ss") & "," & _
        Format(result.Duration, "0.000") & "," & _
        """" & result.Result & """," & _
        """" & Replace(result.ErrorMessage, """", """""") & """," & _
        """" & result.Tags & """"
End Function

' Formate un résultat au format Markdown
Private Function FormatTestResultAsMarkdown(ByVal result As TTestResult) As String
    Dim statusEmoji As String
    Dim output As String
    
    ' Déterminer l'emoji en fonction du résultat
    Select Case result.Result
        Case TEST_RESULT_PASS
            statusEmoji = "?"
        Case TEST_RESULT_FAIL
            statusEmoji = "?"
        Case TEST_RESULT_SKIP
            statusEmoji = "??"
        Case TEST_RESULT_ERROR
            statusEmoji = "??"
        Case Else
            statusEmoji = "?"
    End Select
    
    output = "### " & statusEmoji & " " & result.TestName & vbCrLf & _
             "**Module:** `" & result.TestModule & "`" & vbCrLf & _
             "**Description:** " & result.Description & vbCrLf & _
             "**Niveau:** " & result.Level & vbCrLf & _
             "**Durée:** " & Format(result.Duration, "0.000") & " secondes" & vbCrLf
            
    If result.ErrorMessage <> "" Then
        output = output & "**Erreur:** ```" & vbCrLf & result.ErrorMessage & vbCrLf & "```" & vbCrLf
    End If
    
    If result.Tags <> "" Then
        output = output & "**Tags:** `" & Replace(result.Tags, ",", "`, `") & "`" & vbCrLf
    End If
    
    output = output & "---" & vbCrLf
    
    FormatTestResultAsMarkdown = output
End Function

' Formate un résultat au format HTML
Private Function FormatTestResultAsHTML(ByVal result As TTestResult) As String
    Dim statusClass As String
    Dim output As String
    
    ' Déterminer la classe CSS en fonction du résultat
    Select Case result.Result
        Case TEST_RESULT_PASS
            statusClass = "success"
        Case TEST_RESULT_FAIL
            statusClass = "danger"
        Case TEST_RESULT_SKIP
            statusClass = "warning"
        Case TEST_RESULT_ERROR
            statusClass = "danger"
        Case Else
            statusClass = "secondary"
    End Select
    
    output = "<div class=""test-result " & statusClass & """>" & vbCrLf & _
             "  <h3>" & result.TestName & "</h3>" & vbCrLf & _
             "  <div class=""test-details"">" & vbCrLf & _
             "    <p><strong>Module:</strong> " & result.TestModule & "</p>" & vbCrLf & _
             "    <p><strong>Description:</strong> " & result.Description & "</p>" & vbCrLf & _
             "    <p><strong>Niveau:</strong> " & result.Level & "</p>" & vbCrLf & _
             "    <p><strong>Durée:</strong> " & Format(result.Duration, "0.000") & " secondes</p>" & vbCrLf
            
    If result.ErrorMessage <> "" Then
        output = output & "    <div class=""error-message""><strong>Erreur:</strong><pre>" & _
                 result.ErrorMessage & "</pre></div>" & vbCrLf
    End If
    
    If result.Tags <> "" Then
        Dim tags As Variant
        Dim tag As Variant
        Dim tagsHTML As String
        
        tags = Split(result.Tags, ",")
        tagsHTML = ""
        
        For Each tag In tags
            tagsHTML = tagsHTML & "<span class=""tag"">" & Trim(tag) & "</span> "
        Next tag
        
        output = output & "    <p><strong>Tags:</strong> " & tagsHTML & "</p>" & vbCrLf
    End If
    
    output = output & "  </div>" & vbCrLf & "</div>" & vbCrLf
    
    FormatTestResultAsHTML = output
End Function

' ----------------------------------------
' Méthodes utilitaires pour l'exécution
' ----------------------------------------

' Obtient le temps actuel en secondes (pour mesurer les durées)
Public Function GetTimeInSeconds() As Double
    GetTimeInSeconds = Timer
End Function

' Génère un nom de fichier unique basé sur la date et l'heure
Public Function GenerateUniqueFileName(ByVal basePath As String, _
                                    ByVal baseName As String, _
                                    ByVal extension As String) As String
    Dim fileName As String
    fileName = baseName & "_" & Format(Now, "yyyymmdd_hhnnss")
    
    ' S'assurer que l'extension commence par un point
    If Left(extension, 1) <> "." Then
        extension = "." & extension
    End If
    
    ' S'assurer que le chemin se termine par un séparateur
    If Right(basePath, 1) <> "\" Then
        basePath = basePath & "\"
    End If
    
    GenerateUniqueFileName = basePath & fileName & extension
End Function

' Crée un rapport de résultats de test
Public Function GenerateTestReport(ByVal results As Collection, _
                                ByVal title As String, _
                                Optional ByVal format As Integer = -1) As String
    On Error GoTo ErrorHandler
    
    ' S'assurer que le framework est initialisé
    If Not mInitialized Then InitializeTestFramework
    
    ' Si format non spécifié, utiliser celui par défaut
    If format = -1 Then
        format = mConfig("OutputFormat")
    End If
    
    ' Statistiques
    Dim totalTests As Long
    Dim passedTests As Long
    Dim failedTests As Long
    Dim skippedTests As Long
    Dim errorTests As Long
    Dim totalDuration As Double
    
    totalTests = results.Count
    passedTests = 0
    failedTests = 0
    skippedTests = 0
    errorTests = 0
    totalDuration = 0
    
    ' Calculer les statistiques
    Dim i As Long
    Dim result As TTestResult
    
    For i = 1 To totalTests
        result = results(i)
        
        totalDuration = totalDuration + result.Duration
        
        Select Case result.Result
            Case TEST_RESULT_PASS
                passedTests = passedTests + 1
            Case TEST_RESULT_FAIL
                failedTests = failedTests + 1
            Case TEST_RESULT_SKIP
                skippedTests = skippedTests + 1
            Case TEST_RESULT_ERROR
                errorTests = errorTests + 1
        End Select
    Next i
    
    ' Générer le rapport selon le format
    Dim report As String
    
    Select Case format
        Case FORMAT_TEXT
            report = GenerateTextReport(results, title, totalTests, passedTests, _
                    failedTests, skippedTests, errorTests, totalDuration)
        Case FORMAT_CSV
            report = GenerateCSVReport(results, title)
        Case FORMAT_MARKDOWN
            report = GenerateMarkdownReport(results, title, totalTests, passedTests, _
                    failedTests, skippedTests, errorTests, totalDuration)
        Case FORMAT_HTML
            report = GenerateHTMLReport(results, title, totalTests, passedTests, _
                    failedTests, skippedTests, errorTests, totalDuration)
        Case Else
            report = GenerateTextReport(results, title, totalTests, passedTests, _
                    failedTests, skippedTests, errorTests, totalDuration)
    End Select
    
    GenerateTestReport = report
    Exit Function
    
ErrorHandler:
    Err.Raise ERR_INTERNAL_ERROR, MODULE_NAME & ".GenerateTestReport", _
              "Erreur lors de la génération du rapport: " & Err.Description
End Function

' Génère un rapport au format texte
Private Function GenerateTextReport(ByVal results As Collection, _
                                 ByVal title As String, _
                                 ByVal totalTests As Long, _
                                 ByVal passedTests As Long, _
                                 ByVal failedTests As Long, _
                                 ByVal skippedTests As Long, _
                                 ByVal errorTests As Long, _
                                 ByVal totalDuration As Double) As String
    Dim report As String
    Dim i As Long
    Dim result As TTestResult
    
    ' En-tête
    report = String(60, "=") & vbCrLf & _
             "RAPPORT DE TESTS: " & title & vbCrLf & _
             "Date: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
             String(60, "=") & vbCrLf & vbCrLf
    
    ' Résumé
    report = report & "RÉSUMÉ" & vbCrLf & _
                      String(20, "-") & vbCrLf & _
                      "Total des tests:    " & totalTests & vbCrLf & _
                      "Tests réussis:      " & passedTests & " (" & Format(passedTests / totalTests, "0.0%") & ")" & vbCrLf & _
                      "Tests échoués:      " & failedTests & " (" & Format(failedTests / totalTests, "0.0%") & ")" & vbCrLf & _
                      "Tests ignorés:      " & skippedTests & " (" & Format(skippedTests / totalTests, "0.0%") & ")" & vbCrLf & _
                      "Tests en erreur:    " & errorTests & " (" & Format(errorTests / totalTests, "0.0%") & ")" & vbCrLf & _
                      "Durée totale:       " & Format(totalDuration, "0.000") & " secondes" & vbCrLf & vbCrLf
    
    ' Détails
    report = report & "DÉTAILS DES TESTS" & vbCrLf & _
                      String(60, "-") & vbCrLf
    
    For i = 1 To results.Count
        result = results(i)
        report = report & FormatTestResultAsText(result)
    Next i
    
    GenerateTextReport = report
End Function

' Génère un rapport au format CSV
Private Function GenerateCSVReport(ByVal results As Collection, ByVal title As String) As String
    Dim report As String
    Dim i As Long
    Dim result As TTestResult
    
    ' En-tête
    report = "Nom du test,Module,Description,Niveau,Heure de début,Durée,Résultat,Message d'erreur,Tags" & vbCrLf
    
    ' Détails
    For i = 1 To results.Count
        result = results(i)
        report = report & FormatTestResultAsCSV(result) & vbCrLf
    Next i
    
    GenerateCSVReport = report
End Function

' Génère un rapport au format Markdown
Private Function GenerateMarkdownReport(ByVal results As Collection, _
                                     ByVal title As String, _
                                     ByVal totalTests As Long, _
                                     ByVal passedTests As Long, _
                                     ByVal failedTests As Long, _
                                     ByVal skippedTests As Long, _
                                     ByVal errorTests As Long, _
                                     ByVal totalDuration As Double) As String
    Dim report As String
    Dim i As Long
    Dim result As TTestResult
    
    ' En-tête
    report = "# Rapport de Tests: " & title & vbCrLf & _
             "Date: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & vbCrLf
    
    ' Résumé
    report = report & "## Résumé" & vbCrLf & vbCrLf & _
                      "| Métrique | Valeur | Pourcentage |" & vbCrLf & _
                      "|----------|--------|------------|" & vbCrLf & _
                      "| **Tests totaux** | " & totalTests & " | 100% |" & vbCrLf & _
                      "| **Tests réussis** | " & passedTests & " | " & Format(passedTests / totalTests, "0.0%") & " |" & vbCrLf & _
                      "| **Tests échoués** | " & failedTests & " | " & Format(failedTests / totalTests, "0.0%") & " |" & vbCrLf & _
                      "| **Tests ignorés** | " & skippedTests & " | " & Format(skippedTests / totalTests, "0.0%") & " |" & vbCrLf & _
                      "| **Tests en erreur** | " & errorTests & " | " & Format(errorTests / totalTests, "0.0%") & " |" & vbCrLf & _
                      "| **Durée totale** | " & Format(totalDuration, "0.000") & " sec | |" & vbCrLf & vbCrLf
    
    ' Détails
    report = report & "## Détails des Tests" & vbCrLf & vbCrLf
    
    For i = 1 To results.Count
        result = results(i)
        report = report & FormatTestResultAsMarkdown(result)
    Next i
    
    GenerateMarkdownReport = report
End Function

' Génère un rapport au format HTML
Private Function GenerateHTMLReport(ByVal results As Collection, _
                                 ByVal title As String, _
                                 ByVal totalTests As Long, _
                                 ByVal passedTests As Long, _
                                 ByVal failedTests As Long, _
                                 ByVal skippedTests As Long, _
                                 ByVal errorTests As Long, _
                                 ByVal totalDuration As Double) As String
    Dim report As String
    Dim i As Long
    Dim result As TTestResult
    
    ' En-tête HTML et styles CSS
    report = "<!DOCTYPE html>" & vbCrLf & _
             "<html lang=""fr"">" & vbCrLf & _
             "<head>" & vbCrLf & _
             "  <meta charset=""UTF-8"">" & vbCrLf & _
             "  <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & vbCrLf & _
             "  <title>Rapport de Tests: " & title & "</title>" & vbCrLf & _
             "  <style>" & vbCrLf & _
             "    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; }" & vbCrLf & _
             "    h1, h2 { color: #333; }" & vbCrLf & _
             "    .summary { background-color: #f5f5f5; padding: 15px; border-radius: 5px; margin-bottom: 20px; }" & vbCrLf & _
             "    .summary-table { width: 100%; border-collapse: collapse; }" & vbCrLf & _
             "    .summary-table th, .summary-table td { padding: 8px; text-align: left; border-bottom: 1px solid #ddd; }" & vbCrLf & _
             "    .success { background-color: #dff0d8; }" & vbCrLf & _
             "    .danger { background-color: #f2dede; }" & vbCrLf & _
             "    .warning { background-color: #fcf8e3; }" & vbCrLf & _
             "    .secondary { background-color: #e7e7e7; }" & vbCrLf & _
             "    .test-result { padding: 10px; margin-bottom: 15px; border-radius: 5px; }" & vbCrLf & _
             "    .test-details { margin-left: 15px; }" & vbCrLf & _
             "    .error-message { background-color: #f8f8f8; padding: 10px; border-left: 3px solid #f2dede; }" & vbCrLf & _
             "    .tag { display: inline-block; padding: 2px 8px; background-color: #007bff; color: white; border-radius: 12px; margin-right: 5px; font-size: 0.8em; }" & vbCrLf & _
             "    pre { background-color: #f8f8f8; padding: 10px; border-radius: 3px; overflow-x: auto; }" & vbCrLf & _
             "  </style>" & vbCrLf & _
             "</head>" & vbCrLf & _
             "<body>" & vbCrLf & _
             "  <h1>Rapport de Tests: " & title & "</h1>" & vbCrLf & _
             "  <p>Date: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & "</p>" & vbCrLf
    
    ' Résumé
    report = report & "  <div class=""summary"">" & vbCrLf & _
                      "    <h2>Résumé</h2>" & vbCrLf & _
                      "    <table class=""summary-table"">" & vbCrLf & _
                      "      <tr><th>Métrique</th><th>Valeur</th><th>Pourcentage</th></tr>" & vbCrLf & _
                      "      <tr><td>Tests totaux</td><td>" & totalTests & "</td><td>100%</td></tr>" & vbCrLf & _
                      "      <tr class=""" & IIf(failedTests = 0 And errorTests = 0, "success", "danger") & """>" & _
                      "<td>Tests réussis</td><td>" & passedTests & "</td><td>" & Format(passedTests / totalTests, "0.0%") & "</td></tr>" & vbCrLf & _
                      "      <tr class=""" & IIf(failedTests > 0, "danger", "success") & """>" & _
                      "<td>Tests échoués</td><td>" & failedTests & "</td><td>" & Format(failedTests / totalTests, "0.0%") & "</td></tr>" & vbCrLf & _
                      "      <tr class=""" & IIf(skippedTests > 0, "warning", "") & """>" & _
                      "<td>Tests ignorés</td><td>" & skippedTests & "</td><td>" & Format(skippedTests / totalTests, "0.0%") & "</td></tr>" & vbCrLf & _
                      "      <tr class=""" & IIf(errorTests > 0, "danger", "") & """>" & _
                      "<td>Tests en erreur</td><td>" & errorTests & "</td><td>" & Format(errorTests / totalTests, "0.0%") & "</td></tr>" & vbCrLf & _
                      "      <tr><td>Durée totale</td><td>" & Format(totalDuration, "0.000") & " sec</td><td></td></tr>" & vbCrLf & _
                      "    </table>" & vbCrLf & _
                      "  </div>" & vbCrLf
    
    ' Détails
    report = report & "  <h2>Détails des Tests</h2>" & vbCrLf
    
    For i = 1 To results.Count
        result = results(i)
        report = report & "  " & FormatTestResultAsHTML(result)
    Next i
    
    ' Pied de page HTML
    report = report & "</body>" & vbCrLf & "</html>"
    
    GenerateHTMLReport = report
End Function

' Sauvegarde un rapport dans un fichier
Public Sub SaveReportToFile(ByVal report As String, _
                         ByVal filePath As String, _
                         Optional ByVal overwrite As Boolean = False)
    On Error GoTo ErrorHandler
    
    ' Vérifier si le fichier existe déjà
    If Dir(filePath) <> "" And Not overwrite Then
        Err.Raise ERR_INVALID_ARGUMENT, , "Le fichier existe déjà. Utilisez overwrite=True pour l'écraser."
    End If
    
    ' Écrire le rapport dans un fichier
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Output As #fileNum
    Print #fileNum, report
    Close #fileNum
    
    Exit Sub
    
ErrorHandler:
    Err.Raise ERR_INTERNAL_ERROR, MODULE_NAME & ".SaveReportToFile", _
              "Erreur lors de la sauvegarde du rapport: " & Err.Description
End Sub 