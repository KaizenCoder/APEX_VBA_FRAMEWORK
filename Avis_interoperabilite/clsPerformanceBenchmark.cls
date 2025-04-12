VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPerformanceBenchmark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsPerformanceBenchmark.cls
' Classe de benchmarking pour l'évaluation des performances
'
' Cette classe fournit des outils pour mesurer les performances des opérations,
' comparer différentes implémentations et générer des rapports.
'
' @class clsPerformanceBenchmark
' @author APEX Framework Team
' @version 1.0
' @date 2024-07-27
Option Explicit

' Structure pour les informations d'une exécution de benchmark
Private Type TBenchmarkRun
    Name As String               ' Nom de l'exécution
    Description As String        ' Description de l'exécution
    StartTime As Double          ' Heure de début (millisecondes)
    EndTime As Double            ' Heure de fin (millisecondes)
    Duration As Double           ' Durée en millisecondes
    MemoryUsageBefore As Double  ' Utilisation mémoire avant (si disponible)
    MemoryUsageAfter As Double   ' Utilisation mémoire après (si disponible)
    Tags As Collection           ' Tags associés à l'exécution
    Metrics As Collection        ' Métriques additionnelles
    Parameters As Collection     ' Paramètres de l'exécution
    Category As String           ' Catégorie de benchmark
    Iterations As Long           ' Nombre d'itérations
End Type

' Structure pour les résultats de comparaison
Private Type TComparisonResult
    BaselineName As String       ' Nom de la référence
    ComparedName As String       ' Nom de la comparaison
    BaselineDuration As Double   ' Durée de la référence
    ComparedDuration As Double   ' Durée de la comparaison
    PercentageDifference As Double ' Différence en pourcentage
    AbsoluteDifference As Double ' Différence absolue en ms
    Category As String           ' Catégorie
    Description As String        ' Description de la comparaison
End Type

' État interne du benchmark
Private Type TBenchmarkState
    Name As String               ' Nom du benchmark
    Runs As Collection           ' Collection des exécutions
    ComparisonResults As Collection ' Résultats des comparaisons
    CurrentRun As String         ' Nom de l'exécution en cours
    IsRunning As Boolean         ' Indique si un benchmark est en cours
    HighPrecisionAvailable As Boolean ' Indique si la mesure haute précision est disponible
    ReportFormat As String       ' Format de rapport (TEXT/MARKDOWN)
    BaselineName As String       ' Nom de la référence pour les comparaisons
    DefaultIterations As Long    ' Nombre d'itérations par défaut
End Type

Private this As TBenchmarkState

' Constantes
Private Const DEFAULT_ITERATIONS As Long = 1
Private Const FORMAT_TEXT As String = "TEXT"
Private Const FORMAT_MARKDOWN As String = "MARKDOWN"

'------------------------------------------------------------------------------
' Initialise l'objet de benchmark
'
' @param name Nom du benchmark
'------------------------------------------------------------------------------
Public Sub Initialize(Optional ByVal name As String = "Default")
    this.Name = name
    Set this.Runs = New Collection
    Set this.ComparisonResults = New Collection
    this.CurrentRun = ""
    this.IsRunning = False
    this.ReportFormat = FORMAT_MARKDOWN
    this.BaselineName = ""
    this.DefaultIterations = DEFAULT_ITERATIONS
    
    ' Vérifier si les fonctions de haute précision sont disponibles
    On Error Resume Next
    ' TODO: Ajouter la vérification de disponibilité de QueryPerformanceCounter
    this.HighPrecisionAvailable = True
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Configure le benchmark
'
' @param reportFormat Format de rapport (TEXT/MARKDOWN)
' @param defaultIterations Nombre d'itérations par défaut
'------------------------------------------------------------------------------
Public Sub Configure(Optional ByVal reportFormat As String = FORMAT_MARKDOWN, _
                    Optional ByVal defaultIterations As Long = DEFAULT_ITERATIONS)
    If UCase(reportFormat) = FORMAT_TEXT Or UCase(reportFormat) = FORMAT_MARKDOWN Then
        this.ReportFormat = UCase(reportFormat)
    End If
    
    If defaultIterations > 0 Then
        this.DefaultIterations = defaultIterations
    End If
End Sub

'------------------------------------------------------------------------------
' Démarre une mesure de performance
'
' @param name Nom de l'exécution
' @param description Description de l'exécution
' @param category Catégorie de benchmark
' @param iterations Nombre d'itérations
'------------------------------------------------------------------------------
Public Sub StartMeasurement(ByVal name As String, _
                          Optional ByVal description As String = "", _
                          Optional ByVal category As String = "", _
                          Optional ByVal iterations As Long = 0)
    If this.IsRunning Then
        Debug.Print "Avertissement: Un benchmark est déjà en cours. Il sera automatiquement terminé."
        EndMeasurement
    End If
    
    ' Créer une nouvelle exécution
    Dim newRun As TBenchmarkRun
    newRun.Name = name
    newRun.Description = description
    newRun.Category = category
    Set newRun.Tags = New Collection
    Set newRun.Metrics = New Collection
    Set newRun.Parameters = New Collection
    newRun.StartTime = GetTimestamp()
    
    ' Définir les itérations
    If iterations <= 0 Then
        newRun.Iterations = this.DefaultIterations
    Else
        newRun.Iterations = iterations
    End If
    
    ' Essayer de capturer l'utilisation mémoire (si possible)
    On Error Resume Next
    ' TODO: Ajouter la mesure de mémoire si nécessaire
    newRun.MemoryUsageBefore = 0
    On Error GoTo 0
    
    ' Stocker l'exécution courante
    Dim tempRun As TBenchmarkRun
    tempRun = newRun
    
    ' Déterminer un nom unique
    Dim uniqueName As String
    uniqueName = name
    Dim counter As Long
    counter = 1
    
    ' Éviter les noms en double
    On Error Resume Next
    Do While Not this.Runs(uniqueName) Is Nothing
        uniqueName = name & "_" & counter
        counter = counter + 1
    Loop
    On Error GoTo 0
    
    this.CurrentRun = uniqueName
    this.IsRunning = True
    
    Debug.Print "Benchmark '" & uniqueName & "' démarré"
End Sub

'------------------------------------------------------------------------------
' Termine une mesure de performance
'
' @return Double Durée en millisecondes
'------------------------------------------------------------------------------
Public Function EndMeasurement() As Double
    If Not this.IsRunning Then
        Debug.Print "Erreur: Aucun benchmark en cours."
        EndMeasurement = 0
        Exit Function
    End If
    
    ' Récupérer l'exécution courante
    Dim currentRun As TBenchmarkRun
    
    ' Vérifier si l'exécution existe déjà
    On Error Resume Next
    Dim existingRun As Boolean
    existingRun = False
    
    Dim i As Long
    For i = 1 To this.Runs.Count
        If this.Runs(i).Name = this.CurrentRun Then
            currentRun = this.Runs(i)
            existingRun = True
            Exit For
        End If
    Next i
    On Error GoTo 0
    
    If Not existingRun Then
        ' Si on arrive ici, c'est que quelque chose s'est mal passé
        Debug.Print "Erreur: Exécution courante non trouvée."
        this.IsRunning = False
        EndMeasurement = 0
        Exit Function
    End If
    
    ' Mettre à jour les temps
    currentRun.EndTime = GetTimestamp()
    currentRun.Duration = currentRun.EndTime - currentRun.StartTime
    
    ' Essayer de capturer l'utilisation mémoire finale (si possible)
    On Error Resume Next
    ' TODO: Ajouter la mesure de mémoire si nécessaire
    currentRun.MemoryUsageAfter = 0
    On Error GoTo 0
    
    ' Mettre à jour l'exécution dans la collection
    Dim runAdded As Boolean
    runAdded = False
    
    ' Ajouter seulement si elle n'existe pas encore
    If Not existingRun Then
        On Error Resume Next
        this.Runs.Add currentRun, this.CurrentRun
        If Err.Number = 0 Then
            runAdded = True
        End If
        On Error GoTo 0
    Else
        ' Remplacer l'existante
        On Error Resume Next
        this.Runs.Remove this.CurrentRun
        this.Runs.Add currentRun, this.CurrentRun
        runAdded = True
        On Error GoTo 0
    End If
    
    If Not runAdded Then
        Debug.Print "Avertissement: Impossible d'ajouter l'exécution à la collection."
    End If
    
    Debug.Print "Benchmark '" & this.CurrentRun & "' terminé: " & _
                Format(currentRun.Duration, "0.000") & " ms"
    
    this.IsRunning = False
    
    EndMeasurement = currentRun.Duration
End Function

'------------------------------------------------------------------------------
' Exécute une fonction et mesure sa performance
'
' @param name Nom du benchmark
' @param functionName Nom de la fonction à exécuter
' @param description Description du benchmark
' @param category Catégorie
' @param iterations Nombre d'itérations
' @return Double Durée moyenne d'exécution en millisecondes
'------------------------------------------------------------------------------
Public Function Measure(ByVal name As String, _
                       ByVal functionName As String, _
                       Optional ByVal description As String = "", _
                       Optional ByVal category As String = "", _
                       Optional ByVal iterations As Long = 0) As Double
    ' Utiliser le nombre d'itérations par défaut si non spécifié
    If iterations <= 0 Then
        iterations = this.DefaultIterations
    End If
    
    StartMeasurement name, description, category, iterations
    
    ' Exécuter la fonction le nombre d'itérations spécifié
    Dim i As Long
    On Error Resume Next
    
    For i = 1 To iterations
        Application.Run functionName
        
        ' Si une erreur survient, enregistrer et sortir
        If Err.Number <> 0 Then
            EndMeasurement
            Debug.Print "Erreur lors de l'exécution de " & functionName & ": " & Err.Description
            Measure = 0
            Exit Function
        End If
    Next i
    
    On Error GoTo 0
    
    ' Calculer la durée moyenne par itération
    Dim totalDuration As Double
    totalDuration = EndMeasurement()
    
    Dim avgDuration As Double
    avgDuration = totalDuration / iterations
    
    Measure = avgDuration
End Function

'------------------------------------------------------------------------------
' Définit une exécution comme référence pour les comparaisons
'
' @param name Nom de l'exécution de référence
' @return Boolean Succès de l'opération
'------------------------------------------------------------------------------
Public Function SetBaseline(ByVal name As String) As Boolean
    ' Vérifier si l'exécution existe
    On Error Resume Next
    Dim baselineRun As TBenchmarkRun
    baselineRun = this.Runs(name)
    
    If Err.Number <> 0 Then
        Debug.Print "Erreur: Exécution '" & name & "' non trouvée."
        Err.Clear
        SetBaseline = False
        Exit Function
    End If
    On Error GoTo 0
    
    this.BaselineName = name
    SetBaseline = True
    
    Debug.Print "Référence définie: " & name
End Function

'------------------------------------------------------------------------------
' Compare une exécution à la référence
'
' @param name Nom de l'exécution à comparer
' @param description Description de la comparaison
' @return TComparisonResult Résultat de la comparaison
'------------------------------------------------------------------------------
Public Function CompareTo(ByVal name As String, _
                        Optional ByVal description As String = "") As Boolean
    ' Vérifier si une référence est définie
    If Len(this.BaselineName) = 0 Then
        Debug.Print "Erreur: Aucune référence définie. Utilisez SetBaseline d'abord."
        CompareTo = False
        Exit Function
    End If
    
    ' Vérifier si les exécutions existent
    On Error Resume Next
    Dim baselineRun As TBenchmarkRun
    baselineRun = this.Runs(this.BaselineName)
    
    If Err.Number <> 0 Then
        Debug.Print "Erreur: Exécution de référence '" & this.BaselineName & "' non trouvée."
        Err.Clear
        CompareTo = False
        Exit Function
    End If
    
    Dim comparedRun As TBenchmarkRun
    comparedRun = this.Runs(name)
    
    If Err.Number <> 0 Then
        Debug.Print "Erreur: Exécution à comparer '" & name & "' non trouvée."
        Err.Clear
        CompareTo = False
        Exit Function
    End If
    On Error GoTo 0
    
    ' Créer le résultat de comparaison
    Dim result As TComparisonResult
    result.BaselineName = this.BaselineName
    result.ComparedName = name
    result.BaselineDuration = baselineRun.Duration / baselineRun.Iterations
    result.ComparedDuration = comparedRun.Duration / comparedRun.Iterations
    result.AbsoluteDifference = result.ComparedDuration - result.BaselineDuration
    
    ' Éviter division par zéro
    If result.BaselineDuration > 0 Then
        result.PercentageDifference = (result.AbsoluteDifference / result.BaselineDuration) * 100
    Else
        result.PercentageDifference = 0
    End If
    
    result.Category = comparedRun.Category
    result.Description = description
    
    ' Ajouter aux résultats
    On Error Resume Next
    this.ComparisonResults.Add result, this.BaselineName & "_vs_" & name
    If Err.Number <> 0 Then
        ' Si déjà présent, remplacer
        this.ComparisonResults.Remove this.BaselineName & "_vs_" & name
        this.ComparisonResults.Add result, this.BaselineName & "_vs_" & name
    End If
    On Error GoTo 0
    
    ' Afficher un résumé de la comparaison
    Dim faster As Boolean
    faster = result.PercentageDifference < 0
    
    Debug.Print "Comparaison: " & name & " vs " & this.BaselineName & " = " & _
                IIf(faster, "Plus rapide de ", "Plus lent de ") & _
                Format(Abs(result.PercentageDifference), "0.00") & "% " & _
                "(" & Format(result.ComparedDuration, "0.000") & " ms vs " & _
                Format(result.BaselineDuration, "0.000") & " ms)"
    
    CompareTo = True
End Function

'------------------------------------------------------------------------------
' Génère un rapport détaillé des benchmarks
'
' @param format Format du rapport (TEXT/MARKDOWN)
' @return String Rapport formaté
'------------------------------------------------------------------------------
Public Function GenerateReport(Optional ByVal format As String = "") As String
    ' Utiliser le format par défaut si non spécifié
    If Len(format) = 0 Then
        format = this.ReportFormat
    End If
    
    Select Case UCase(format)
        Case FORMAT_MARKDOWN
            GenerateReport = GenerateMarkdownReport()
        Case Else ' FORMAT_TEXT par défaut
            GenerateReport = GenerateTextReport()
    End Select
End Function

'------------------------------------------------------------------------------
' Génère un rapport au format texte brut
'
' @return String Rapport en texte brut
'------------------------------------------------------------------------------
Private Function GenerateTextReport() As String
    Dim report As String
    
    ' En-tête
    report = String(70, "=") & vbCrLf & _
             "RAPPORT DE PERFORMANCE: " & this.Name & vbCrLf & _
             "Date: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
             String(70, "=") & vbCrLf & vbCrLf
    
    ' Section des exécutions
    report = report & "EXÉCUTIONS DE BENCHMARK" & vbCrLf & _
                     String(30, "-") & vbCrLf
    
    Dim run As Variant
    Dim i As Long
    
    If this.Runs.Count = 0 Then
        report = report & "Aucune exécution enregistrée." & vbCrLf
    Else
        ' En-tête du tableau
        report = report & "Nom                    | Catégorie     | Durée (ms)  | Itérations | Durée/itér. (ms)" & vbCrLf & _
                         String(80, "-") & vbCrLf
        
        ' Lignes du tableau
        For i = 1 To this.Runs.Count
            Dim currentRun As TBenchmarkRun
            currentRun = this.Runs(i)
            
            report = report & PadRight(Left(currentRun.Name, 22), 23) & "| " & _
                             PadRight(Left(currentRun.Category, 12), 13) & "| " & _
                             PadLeft(Format(currentRun.Duration, "0.000"), 11) & " | " & _
                             PadLeft(CStr(currentRun.Iterations), 10) & " | " & _
                             PadLeft(Format(currentRun.Duration / currentRun.Iterations, "0.000"), 15) & vbCrLf
        Next i
    End If
    
    report = report & vbCrLf
    
    ' Section des comparaisons
    report = report & "COMPARAISONS DE PERFORMANCE" & vbCrLf & _
                     String(30, "-") & vbCrLf
    
    If this.ComparisonResults.Count = 0 Then
        report = report & "Aucune comparaison enregistrée." & vbCrLf
    Else
        ' En-tête du tableau
        report = report & "Comparaison              | Différence  | Comparé (ms) | Référence (ms) | Différence (%)" & vbCrLf & _
                         String(90, "-") & vbCrLf
        
        ' Lignes du tableau
        For i = 1 To this.ComparisonResults.Count
            Dim currentComparison As TComparisonResult
            currentComparison = this.ComparisonResults(i)
            
            Dim comparisonName As String
            comparisonName = currentComparison.ComparedName & " vs " & currentComparison.BaselineName
            
            Dim differenceText As String
            If currentComparison.PercentageDifference < 0 Then
                differenceText = "Plus rapide"
            ElseIf currentComparison.PercentageDifference > 0 Then
                differenceText = "Plus lent"
            Else
                differenceText = "Identique"
            End If
            
            report = report & PadRight(Left(comparisonName, 24), 25) & "| " & _
                             PadRight(differenceText, 11) & "| " & _
                             PadLeft(Format(currentComparison.ComparedDuration, "0.000"), 12) & " | " & _
                             PadLeft(Format(currentComparison.BaselineDuration, "0.000"), 14) & " | " & _
                             PadLeft(Format(currentComparison.PercentageDifference, "+0.00;-0.00;0.00"), 14) & vbCrLf
        Next i
    End If
    
    GenerateTextReport = report
End Function

'------------------------------------------------------------------------------
' Génère un rapport au format Markdown
'
' @return String Rapport en Markdown
'------------------------------------------------------------------------------
Private Function GenerateMarkdownReport() As String
    Dim report As String
    
    ' En-tête
    report = "# Rapport de Performance: " & this.Name & vbCrLf & vbCrLf & _
             "Date: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & vbCrLf
    
    ' Section des exécutions
    report = report & "## Exécutions de Benchmark" & vbCrLf & vbCrLf
    
    Dim run As Variant
    Dim i As Long
    
    If this.Runs.Count = 0 Then
        report = report & "*Aucune exécution enregistrée.*" & vbCrLf
    Else
        ' En-tête du tableau
        report = report & "| Nom | Catégorie | Durée (ms) | Itérations | Durée/itér. (ms) |" & vbCrLf & _
                         "|-----|-----------|------------|------------|------------------|" & vbCrLf
        
        ' Lignes du tableau
        For i = 1 To this.Runs.Count
            Dim currentRun As TBenchmarkRun
            currentRun = this.Runs(i)
            
            report = report & "| " & currentRun.Name & " | " & _
                             currentRun.Category & " | " & _
                             Format(currentRun.Duration, "0.000") & " | " & _
                             currentRun.Iterations & " | " & _
                             Format(currentRun.Duration / currentRun.Iterations, "0.000") & " |" & vbCrLf
        Next i
    End If
    
    report = report & vbCrLf
    
    ' Section des comparaisons
    report = report & "## Comparaisons de Performance" & vbCrLf & vbCrLf
    
    If this.ComparisonResults.Count = 0 Then
        report = report & "*Aucune comparaison enregistrée.*" & vbCrLf
    Else
        ' En-tête du tableau
        report = report & "| Comparaison | Résultat | Comparé (ms) | Référence (ms) | Différence (%) |" & vbCrLf & _
                         "|-------------|----------|--------------|----------------|----------------|" & vbCrLf
        
        ' Lignes du tableau
        For i = 1 To this.ComparisonResults.Count
            Dim currentComparison As TComparisonResult
            currentComparison = this.ComparisonResults(i)
            
            Dim comparisonName As String
            comparisonName = currentComparison.ComparedName & " vs " & currentComparison.BaselineName
            
            Dim differenceText As String
            Dim differenceEmoji As String
            
            If currentComparison.PercentageDifference < 0 Then
                differenceText = "Plus rapide"
                differenceEmoji = "?"
            ElseIf currentComparison.PercentageDifference > 0 Then
                differenceText = "Plus lent"
                differenceEmoji = "??"
            Else
                differenceText = "Identique"
                differenceEmoji = "??"
            End If
            
            report = report & "| " & comparisonName & " | " & _
                             differenceEmoji & " " & differenceText & " | " & _
                             Format(currentComparison.ComparedDuration, "0.000") & " | " & _
                             Format(currentComparison.BaselineDuration, "0.000") & " | " & _
                             Format(currentComparison.PercentageDifference, "+0.00;-0.00;0.00") & "% |" & vbCrLf
        Next i
    End If
    
    ' Ajouter des visualisations si beaucoup de données
    If this.Runs.Count > 5 Or this.ComparisonResults.Count > 5 Then
        report = report & vbCrLf & "## Recommandations" & vbCrLf & vbCrLf & _
                         "Pour une analyse visuelle complète, il est recommandé de:" & vbCrLf & vbCrLf & _
                         "1. Exporter ces données vers un outil de visualisation" & vbCrLf & _
                         "2. Créer des graphiques de comparaison de performance" & vbCrLf & _
                         "3. Analyser les tendances sur plusieurs séries de tests" & vbCrLf
    End If
    
    GenerateMarkdownReport = report
End Function

'------------------------------------------------------------------------------
' Exporte les données de benchmark au format CSV
'
' @param filePath Chemin du fichier CSV
' @return Boolean Succès de l'opération
'------------------------------------------------------------------------------
Public Function ExportToCSV(ByVal filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Output As #fileNum
    
    ' En-tête du fichier CSV
    Print #fileNum, "Type,Name,Category,Duration,Iterations,DurationPerIter,Description"
    
    ' Données des exécutions
    Dim i As Long
    For i = 1 To this.Runs.Count
        Dim currentRun As TBenchmarkRun
        currentRun = this.Runs(i)
        
        Dim runLine As String
        runLine = "Run," & _
                  Chr(34) & currentRun.Name & Chr(34) & "," & _
                  Chr(34) & currentRun.Category & Chr(34) & "," & _
                  Format(currentRun.Duration, "0.000") & "," & _
                  currentRun.Iterations & "," & _
                  Format(currentRun.Duration / currentRun.Iterations, "0.000") & "," & _
                  Chr(34) & currentRun.Description & Chr(34)
        
        Print #fileNum, runLine
    Next i
    
    ' Données des comparaisons
    For i = 1 To this.ComparisonResults.Count
        Dim currentComparison As TComparisonResult
        currentComparison = this.ComparisonResults(i)
        
        Dim comparisonLine As String
        comparisonLine = "Comparison," & _
                        Chr(34) & currentComparison.ComparedName & " vs " & currentComparison.BaselineName & Chr(34) & "," & _
                        Chr(34) & currentComparison.Category & Chr(34) & "," & _
                        Format(currentComparison.AbsoluteDifference, "0.000") & ",0," & _
                        Format(currentComparison.PercentageDifference, "0.000") & "," & _
                        Chr(34) & currentComparison.Description & Chr(34)
        
        Print #fileNum, comparisonLine
    Next i
    
    Close #fileNum
    
    ExportToCSV = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur lors de l'exportation CSV: " & Err.Description
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
    ExportToCSV = False
End Function

'------------------------------------------------------------------------------
' Ajoute un paramètre à l'exécution en cours
'
' @param name Nom du paramètre
' @param value Valeur du paramètre
' @return Boolean Succès de l'opération
'------------------------------------------------------------------------------
Public Function AddParameter(ByVal name As String, ByVal value As Variant) As Boolean
    If Not this.IsRunning Then
        Debug.Print "Erreur: Aucun benchmark en cours."
        AddParameter = False
        Exit Function
    End If
    
    ' Récupérer l'exécution courante
    On Error Resume Next
    Dim currentRun As TBenchmarkRun
    
    Dim i As Long
    For i = 1 To this.Runs.Count
        If this.Runs(i).Name = this.CurrentRun Then
            currentRun = this.Runs(i)
            Exit For
        End If
    Next i
    
    If Err.Number <> 0 Then
        Err.Clear
        AddParameter = False
        Exit Function
    End If
    On Error GoTo 0
    
    ' Ajouter le paramètre
    On Error Resume Next
    currentRun.Parameters.Add value, name
    If Err.Number <> 0 Then
        ' Remplacer si existe déjà
        currentRun.Parameters.Remove name
        currentRun.Parameters.Add value, name
    End If
    On Error GoTo 0
    
    ' Mettre à jour l'exécution
    Dim j As Long
    For j = 1 To this.Runs.Count
        If this.Runs(j).Name = this.CurrentRun Then
            this.Runs.Remove j
            this.Runs.Add currentRun, this.CurrentRun
            Exit For
        End If
    Next j
    
    AddParameter = True
End Function

'------------------------------------------------------------------------------
' Retourne la durée d'exécution d'un benchmark spécifique
'
' @param name Nom du benchmark
' @return Double Durée en millisecondes ou 0 si non trouvé
'------------------------------------------------------------------------------
Public Function GetDuration(ByVal name As String) As Double
    On Error Resume Next
    Dim run As TBenchmarkRun
    run = this.Runs(name)
    
    If Err.Number <> 0 Then
        GetDuration = 0
        Err.Clear
    Else
        GetDuration = run.Duration
    End If
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Utilitaire pour remplir une chaîne à droite jusqu'à une certaine longueur
'
' @param str Chaîne à remplir
' @param length Longueur désirée
' @param fillChar Caractère de remplissage (espace par défaut)
' @return String Chaîne remplie
'------------------------------------------------------------------------------
Private Function PadRight(ByVal str As String, ByVal length As Long, _
                         Optional ByVal fillChar As String = " ") As String
    If Len(str) >= length Then
        PadRight = str
    Else
        PadRight = str & String(length - Len(str), fillChar)
    End If
End Function

'------------------------------------------------------------------------------
' Utilitaire pour remplir une chaîne à gauche jusqu'à une certaine longueur
'
' @param str Chaîne à remplir
' @param length Longueur désirée
' @param fillChar Caractère de remplissage (espace par défaut)
' @return String Chaîne remplie
'------------------------------------------------------------------------------
Private Function PadLeft(ByVal str As String, ByVal length As Long, _
                        Optional ByVal fillChar As String = " ") As String
    If Len(str) >= length Then
        PadLeft = str
    Else
        PadLeft = String(length - Len(str), fillChar) & str
    End If
End Function

'------------------------------------------------------------------------------
' Obtient un timestamp haute précision en millisecondes
'
' @return Double Timestamp en millisecondes
'------------------------------------------------------------------------------
Private Function GetTimestamp() As Double
    ' Pour l'instant, utiliser Timer (moins précis) - en attendant implémentation QueryPerformanceCounter
    GetTimestamp = Timer * 1000
End Function 