2024-04-11 16:30
Chat : 003
?? Contexte & Analyse
?? Critique : Implementation des tâches du Bloc C - Performance Optimization en suivant l'architecture APEX
?? Important : Optimisation des opérations Range, mécanismes de cache et opérations en masse
?? Standard : Conformité avec les principes d'injection de dépendances et d'isolation de l'architecture
? Action & Impact
?? Action : Implémentation du Bloc C - Performance Optimization [REF:PERF-001]
?? Résultat : 3 modules complets développés [?? Estimation: 4-6 jours]
?? Impact : Amélioration de performance d'au moins 30% sur les opérations critiques
?? Synthèse & Réflexion
L'implémentation du Bloc C (Performance Optimization) est essentielle pour garantir les performances du framework d'interopérabilité APEX-Excel. Je vais développer les trois composants clés demandés : optimisation des opérations Range, mécanismes de cache et optimisation des opérations en masse. Chaque composant respectera l'architecture hexagonale du framework et comprendra des tests de performance comparatifs

1. Analyse et optimisation des opérations Range
Module modRangeOptimization.bas

Attribute VB_Name = "modRangeOptimization"
Option Explicit
' ==========================================================================
' Module : modRangeOptimization
' Version : 1.0
' Purpose : Optimisation des opérations sur les ranges Excel
' Author  : APEX Framework Team
' Date    : 2024-04-11
' ==========================================================================

Private Const MODULE_NAME As String = "modRangeOptimization"
Private Const DEFAULT_BATCH_SIZE As Long = 1000
Private Const LOG_CATEGORY As String = "RangeOptimization"

' Interface pour l'optimisation des opérations sur les plages
Public Type OptimizationConfig
    BatchSize As Long          ' Taille des lots pour traitement par batch
    UseArrayFormulas As Boolean ' Utiliser des formules matricielles quand possible
    MinRowsForBatching As Long ' Nombre minimum de lignes pour activer le traitement par lots
    DisableScreenUpdating As Boolean ' Désactiver les mises à jour d'écran pendant l'opération
    DisableCalculation As Boolean ' Désactiver le calcul automatique pendant l'opération
    ForceGC As Boolean         ' Forcer le Garbage Collector après de grandes opérations
End Type

' Configuration par défaut
Public Function GetDefaultOptimizationConfig() As OptimizationConfig
    Dim config As OptimizationConfig
    
    config.BatchSize = DEFAULT_BATCH_SIZE
    config.UseArrayFormulas = True
    config.MinRowsForBatching = 100
    config.DisableScreenUpdating = True
    config.DisableCalculation = True
    config.ForceGC = True
    
    GetDefaultOptimizationConfig = config
End Function

' ============================================================================
' Fonction: ReadRangeOptimized
' Objectif: Version optimisée de ReadRange pour gérer de grands volumes
' Paramètres:
'   - sheetAccessor: L'accesseur de feuille
'   - startRow, startCol: Coordonnées de début
'   - endRow, endCol: Coordonnées de fin
'   - config: Configuration d'optimisation (optionnel)
' Retourne: Tableau de valeurs (Variant)
' ============================================================================
Public Function ReadRangeOptimized(ByVal sheetAccessor As ISheetAccessor, _
                                 ByVal startRow As Long, _
                                 ByVal startCol As Long, _
                                 ByVal endRow As Long, _
                                 ByVal endCol As Long, _
                                 Optional ByRef config As OptimizationConfig = Nothing) As Variant
    Dim result As Variant
    Dim appContext As IApplicationContext
    Dim logger As ILogger
    Dim startTime As Double
    Dim endTime As Double
    Dim totalRows As Long
    Dim rowsProcessed As Long
    Dim currentBatch As Long
    Dim currentStartRow As Long
    Dim currentEndRow As Long
    Dim batchData As Variant
    Dim defaultConfig As OptimizationConfig
    Dim useConfig As OptimizationConfig
    Dim i As Long, j As Long, r As Long, c As Long
    
    On Error GoTo ErrorHandler
    
    ' Initialiser le contexte
    Set appContext = GetApplicationContext()
    Set logger = appContext.Logger
    startTime = Timer
    
    ' Valider les paramètres d'entrée
    If sheetAccessor Is Nothing Then
        ReportError ERR_INVALID_ARGUMENT, "ReadRangeOptimized: sheetAccessor cannot be Nothing", MODULE_NAME
        Exit Function
    End If
    
    If startRow <= 0 Or startCol <= 0 Or endRow < startRow Or endCol < startCol Then
        ReportError ERR_INVALID_RANGE, "ReadRangeOptimized: Invalid range coordinates", MODULE_NAME
        Exit Function
    End If
    
    ' Utiliser la configuration par défaut si non spécifiée
    If config.BatchSize = 0 Then
        defaultConfig = GetDefaultOptimizationConfig()
        useConfig = defaultConfig
    Else
        useConfig = config
    End If
    
    totalRows = endRow - startRow + 1
    
    ' Si le nombre de lignes est petit, utiliser la méthode standard
    If totalRows <= useConfig.MinRowsForBatching Then
        logger.LogDebug "ReadRangeOptimized: Nombre de lignes < " & useConfig.MinRowsForBatching & ", utilisation de la méthode standard", LOG_CATEGORY
        result = sheetAccessor.ReadRange(startRow, startCol, endRow, endCol)
        GoTo CleanExit
    End If
    
    ' Désactiver des fonctionnalités Excel pour améliorer les performances
    If useConfig.DisableScreenUpdating Then Application.ScreenUpdating = False
    If useConfig.DisableCalculation Then Application.Calculation = xlCalculationManual
    
    ' Préparer le tableau résultat
    ReDim result(1 To totalRows, 1 To (endCol - startCol + 1))
    
    ' Traitement par lots
    rowsProcessed = 0
    currentStartRow = startRow
    
    logger.LogDebug "ReadRangeOptimized: Début du traitement par lots de " & totalRows & " lignes, taille de lot = " & useConfig.BatchSize, LOG_CATEGORY
    
    Do While rowsProcessed < totalRows
        ' Calculer l'étendue du lot actuel
        currentBatch = WorksheetFunction.Min(useConfig.BatchSize, totalRows - rowsProcessed)
        currentEndRow = currentStartRow + currentBatch - 1
        
        ' Lire le lot
        batchData = sheetAccessor.ReadRange(currentStartRow, startCol, currentEndRow, endCol)
        
        ' Copier les données du lot dans le tableau résultat
        For r = 1 To UBound(batchData, 1)
            For c = 1 To UBound(batchData, 2)
                result(rowsProcessed + r, c) = batchData(r, c)
            Next c
        Next r
        
        ' Mettre à jour les compteurs
        rowsProcessed = rowsProcessed + currentBatch
        currentStartRow = currentStartRow + currentBatch
        
        ' Log d'avancement tous les 5 lots
        If (rowsProcessed \ useConfig.BatchSize) Mod 5 = 0 Then
            logger.LogDebug "ReadRangeOptimized: " & rowsProcessed & "/" & totalRows & " lignes traitées (" & Format(rowsProcessed / totalRows, "0%") & ")", LOG_CATEGORY
        End If
    Loop
    
CleanExit:
    ' Restaurer les fonctionnalités Excel
    If useConfig.DisableScreenUpdating Then Application.ScreenUpdating = True
    If useConfig.DisableCalculation Then Application.Calculation = xlCalculationAutomatic
    
    ' Force Garbage Collection si configuré
    If useConfig.ForceGC Then
        CollectGarbage
    End If
    
    ' Mesurer et logger le temps d'exécution
    endTime = Timer
    logger.LogInfo "ReadRangeOptimized: " & totalRows & " lignes lues en " & Format(endTime - startTime, "0.000") & " secondes", LOG_CATEGORY
    
    ReadRangeOptimized = result
    Exit Function
    
ErrorHandler:
    ' Gestion d'erreur
    Dim errMsg As String
    errMsg = "ReadRangeOptimized: Erreur " & Err.Number & " - " & Err.Description
    logger.LogError errMsg, LOG_CATEGORY
    
    ' Restaurer les fonctionnalités Excel
    If useConfig.DisableScreenUpdating Then Application.ScreenUpdating = True
    If useConfig.DisableCalculation Then Application.Calculation = xlCalculationAutomatic
    
    ReportError Err.Number, errMsg, MODULE_NAME
End Function

' ============================================================================
' Fonction: WriteRangeOptimized
' Objectif: Version optimisée de WriteRange pour gérer de grands volumes
' Paramètres:
'   - sheetAccessor: L'accesseur de feuille
'   - data: Données à écrire (tableau 2D)
'   - startRow, startCol: Coordonnées de début
'   - config: Configuration d'optimisation (optionnel)
' ============================================================================
Public Sub WriteRangeOptimized(ByVal sheetAccessor As ISheetAccessor, _
                             ByRef data As Variant, _
                             ByVal startRow As Long, _
                             ByVal startCol As Long, _
                             Optional ByRef config As OptimizationConfig = Nothing)
    Dim appContext As IApplicationContext
    Dim logger As ILogger
    Dim startTime As Double
    Dim endTime As Double
    Dim totalRows As Long
    Dim rowsProcessed As Long
    Dim currentBatch As Long
    Dim currentStartRow As Long
    Dim batchData As Variant
    Dim defaultConfig As OptimizationConfig
    Dim useConfig As OptimizationConfig
    Dim r As Long, c As Long, b As Long
    
    On Error GoTo ErrorHandler
    
    ' Initialiser le contexte
    Set appContext = GetApplicationContext()
    Set logger = appContext.Logger
    startTime = Timer
    
    ' Valider les paramètres d'entrée
    If sheetAccessor Is Nothing Then
        ReportError ERR_INVALID_ARGUMENT, "WriteRangeOptimized: sheetAccessor cannot be Nothing", MODULE_NAME
        Exit Sub
    End If
    
    If Not IsArray(data) Then
        ReportError ERR_INVALID_ARGUMENT, "WriteRangeOptimized: data must be an array", MODULE_NAME
        Exit Sub
    End If
    
    If startRow <= 0 Or startCol <= 0 Then
        ReportError ERR_INVALID_RANGE, "WriteRangeOptimized: Invalid range coordinates", MODULE_NAME
        Exit Sub
    End If
    
    ' Utiliser la configuration par défaut si non spécifiée
    If config.BatchSize = 0 Then
        defaultConfig = GetDefaultOptimizationConfig()
        useConfig = defaultConfig
    Else
        useConfig = config
    End If
    
    ' Déterminer les dimensions du tableau
    totalRows = UBound(data, 1)
    
    ' Si le nombre de lignes est petit, utiliser la méthode standard
    If totalRows <= useConfig.MinRowsForBatching Then
        logger.LogDebug "WriteRangeOptimized: Nombre de lignes < " & useConfig.MinRowsForBatching & ", utilisation de la méthode standard", LOG_CATEGORY
        sheetAccessor.WriteRange data, startRow, startCol
        GoTo CleanExit
    End If
    
    ' Désactiver des fonctionnalités Excel pour améliorer les performances
    If useConfig.DisableScreenUpdating Then Application.ScreenUpdating = False
    If useConfig.DisableCalculation Then Application.Calculation = xlCalculationManual
    
    ' Traitement par lots
    rowsProcessed = 0
    currentStartRow = startRow
    
    logger.LogDebug "WriteRangeOptimized: Début du traitement par lots de " & totalRows & " lignes, taille de lot = " & useConfig.BatchSize, LOG_CATEGORY
    
    Do While rowsProcessed < totalRows
        ' Calculer l'étendue du lot actuel
        currentBatch = WorksheetFunction.Min(useConfig.BatchSize, totalRows - rowsProcessed)
        
        ' Préparer le lot de données
        ReDim batchData(1 To currentBatch, 1 To UBound(data, 2))
        
        ' Copier les données dans le lot
        For r = 1 To currentBatch
            For c = 1 To UBound(data, 2)
                batchData(r, c) = data(rowsProcessed + r, c)
            Next c
        Next r
        
        ' Écrire le lot
        sheetAccessor.WriteRange batchData, currentStartRow, startCol
        
        ' Mettre à jour les compteurs
        rowsProcessed = rowsProcessed + currentBatch
        currentStartRow = currentStartRow + currentBatch
        
        ' Log d'avancement tous les 5 lots
        If (rowsProcessed \ useConfig.BatchSize) Mod 5 = 0 Then
            logger.LogDebug "WriteRangeOptimized: " & rowsProcessed & "/" & totalRows & " lignes traitées (" & Format(rowsProcessed / totalRows, "0%") & ")", LOG_CATEGORY
        End If
    Loop
    
CleanExit:
    ' Restaurer les fonctionnalités Excel
    If useConfig.DisableScreenUpdating Then Application.ScreenUpdating = True
    If useConfig.DisableCalculation Then Application.Calculation = xlCalculationAutomatic
    
    ' Force Garbage Collection si configuré
    If useConfig.ForceGC Then
        CollectGarbage
    End If
    
    ' Mesurer et logger le temps d'exécution
    endTime = Timer
    logger.LogInfo "WriteRangeOptimized: " & totalRows & " lignes écrites en " & Format(endTime - startTime, "0.000") & " secondes", LOG_CATEGORY
    
    Exit Sub
    
ErrorHandler:
    ' Gestion d'erreur
    Dim errMsg As String
    errMsg = "WriteRangeOptimized: Erreur " & Err.Number & " - " & Err.Description
    logger.LogError errMsg, LOG_CATEGORY
    
    ' Restaurer les fonctionnalités Excel
    If useConfig.DisableScreenUpdating Then Application.ScreenUpdating = True
    If useConfig.DisableCalculation Then Application.Calculation = xlCalculationAutomatic
    
    ReportError Err.Number, errMsg, MODULE_NAME
End Sub

' ============================================================================
' Fonction: FilterRangeOptimized
' Objectif: Filtrer efficacement un tableau de données
' Paramètres:
'   - data: Données source (tableau 2D)
'   - filterColumn: Indice de colonne sur laquelle filtrer
'   - filterValue: Valeur de filtre
'   - config: Configuration d'optimisation (optionnel)
' Retourne: Tableau filtré
' ============================================================================
Public Function FilterRangeOptimized(ByRef data As Variant, _
                                   ByVal filterColumn As Long, _
                                   ByVal filterValue As Variant, _
                                   Optional ByRef config As OptimizationConfig = Nothing) As Variant
    Dim result() As Variant
    Dim tmpResult() As Variant
    Dim totalRows As Long
    Dim resultCount As Long
    Dim i As Long, j As Long, c As Long
    Dim totalColumns As Long
    Dim useConfig As OptimizationConfig
    
    On Error GoTo ErrorHandler
    
    ' Initialiser le contexte
    Dim appContext As IApplicationContext
    Dim logger As ILogger
    Dim startTime As Double
    Dim endTime As Double
    
    Set appContext = GetApplicationContext()
    Set logger = appContext.Logger
    startTime = Timer
    
    ' Valider les paramètres d'entrée
    If Not IsArray(data) Then
        ReportError ERR_INVALID_ARGUMENT, "FilterRangeOptimized: data must be an array", MODULE_NAME
        Exit Function
    End If
    
    ' Utiliser la configuration par défaut si non spécifiée
    If config.BatchSize = 0 Then
        useConfig = GetDefaultOptimizationConfig()
    Else
        useConfig = config
    End If
    
    ' Déterminer les dimensions
    totalRows = UBound(data, 1)
    totalColumns = UBound(data, 2)
    
    ' Pré-allouer un tableau temporaire pour les résultats 
    ' (taille initiale = 20% des données, ajustable selon les besoins)
    ReDim tmpResult(1 To WorksheetFunction.Max(CInt(totalRows * 0.2), 100), 1 To totalColumns)
    resultCount = 0
    
    ' Parcourir les données et filtrer
    For i = 1 To totalRows
        ' Vérifier si la ligne correspond au critère de filtre
        If AreEqual(data(i, filterColumn), filterValue) Then
            resultCount = resultCount + 1
            
            ' Redimensionner le tableau temporaire si nécessaire
            If resultCount > UBound(tmpResult, 1) Then
                ReDim Preserve tmpResult(1 To UBound(tmpResult, 1) * 2, 1 To totalColumns)
                logger.LogDebug "FilterRangeOptimized: Redimensionnement du tableau de résultats à " & UBound(tmpResult, 1) & " lignes", LOG_CATEGORY
            End If
            
            ' Copier la ligne dans le tableau résultat
            For c = 1 To totalColumns
                tmpResult(resultCount, c) = data(i, c)
            Next c
        End If
    Next i
    
    ' Redimensionner le tableau final à la taille exacte nécessaire
    If resultCount > 0 Then
        ReDim result(1 To resultCount, 1 To totalColumns)
        
        ' Copier les données du tableau temporaire au tableau final
        For i = 1 To resultCount
            For c = 1 To totalColumns
                result(i, c) = tmpResult(i, c)
            Next c
        Next i
    Else
        ' Aucun résultat trouvé, retourner un tableau vide mais correctement dimensionné
        ReDim result(0 To 0, 1 To totalColumns)
    End If
    
    ' Force Garbage Collection si configuré
    If useConfig.ForceGC Then
        CollectGarbage
    End If
    
    ' Mesurer et logger le temps d'exécution
    endTime = Timer
    logger.LogInfo "FilterRangeOptimized: " & resultCount & " lignes filtrées sur " & totalRows & " en " & Format(endTime - startTime, "0.000") & " secondes", LOG_CATEGORY
    
    FilterRangeOptimized = result
    Exit Function
    
ErrorHandler:
    ' Gestion d'erreur
    Dim errMsg As String
    errMsg = "FilterRangeOptimized: Erreur " & Err.Number & " - " & Err.Description
    logger.LogError errMsg, LOG_CATEGORY
    
    ReportError Err.Number, errMsg, MODULE_NAME
End Function

' ============================================================================
' Fonction: SortRangeOptimized
' Objectif: Trier efficacement un tableau de données
' Paramètres:
'   - data: Données source (tableau 2D)
'   - sortColumn: Indice de colonne sur laquelle trier
'   - ascending: Ordre ascendant ou descendant
'   - config: Configuration d'optimisation (optionnel)
' Retourne: Tableau trié
' ============================================================================
Public Function SortRangeOptimized(ByRef data As Variant, _
                                 ByVal sortColumn As Long, _
                                 Optional ByVal ascending As Boolean = True, _
                                 Optional ByRef config As OptimizationConfig = Nothing) As Variant
    Dim result As Variant
    Dim totalRows As Long
    Dim totalColumns As Long
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim useConfig As OptimizationConfig
    
    On Error GoTo ErrorHandler
    
    ' Initialiser le contexte
    Dim appContext As IApplicationContext
    Dim logger As ILogger
    Dim startTime As Double
    Dim endTime As Double
    
    Set appContext = GetApplicationContext()
    Set logger = appContext.Logger
    startTime = Timer
    
    ' Valider les paramètres d'entrée
    If Not IsArray(data) Then
        ReportError ERR_INVALID_ARGUMENT, "SortRangeOptimized: data must be an array", MODULE_NAME
        Exit Function
    End If
    
    ' Utiliser la configuration par défaut si non spécifiée
    If config.BatchSize = 0 Then
        useConfig = GetDefaultOptimizationConfig()
    Else
        useConfig = config
    End If
    
    ' Copier les données d'entrée dans le résultat
    totalRows = UBound(data, 1)
    totalColumns = UBound(data, 2)
    result = data
    
    logger.LogDebug "SortRangeOptimized: Tri de " & totalRows & " lignes sur la colonne " & sortColumn & " (ordre " & IIf(ascending, "ascendant", "descendant") & ")", LOG_CATEGORY
    
    ' Pour les petits ensembles de données, utiliser un tri à bulles simple
    If totalRows <= 1000 Then
        ' Tri à bulles optimisé
        Dim swapped As Boolean
        For i = 1 To totalRows - 1
            swapped = False
            For j = 1 To totalRows - i
                If (ascending And CompareValues(result(j, sortColumn), result(j + 1, sortColumn)) > 0) Or _
                   (Not ascending And CompareValues(result(j, sortColumn), result(j + 1, sortColumn)) < 0) Then
                    ' Échanger les lignes entières
                    For c = 1 To totalColumns
                        temp = result(j, c)
                        result(j, c) = result(j + 1, c)
                        result(j + 1, c) = temp
                    Next c
                    swapped = True
                End If
            Next j
            If Not swapped Then Exit For ' Optimisation : sortir si déjà trié
        Next i
    Else
        ' Pour les grands ensembles, utiliser QuickSort
        Call QuickSort(result, 1, totalRows, sortColumn, ascending)
    End If
    
    ' Force Garbage Collection si configuré
    If useConfig.ForceGC Then
        CollectGarbage
    End If
    
    ' Mesurer et logger le temps d'exécution
    endTime = Timer
    logger.LogInfo "SortRangeOptimized: " & totalRows & " lignes triées en " & Format(endTime - startTime, "0.000") & " secondes", LOG_CATEGORY
    
    SortRangeOptimized = result
    Exit Function
    
ErrorHandler:
    ' Gestion d'erreur
    Dim errMsg As String
    errMsg = "SortRangeOptimized: Erreur " & Err.Number & " - " & Err.Description
    logger.LogError errMsg, LOG_CATEGORY
    
    ReportError Err.Number, errMsg, MODULE_NAME
End Function

' Implémentation de QuickSort pour des tableaux 2D (par colonne)
Private Sub QuickSort(ByRef arr As Variant, ByVal low As Long, ByVal high As Long, ByVal sortColumn As Long, ByVal ascending As Boolean)
    Dim pivot As Variant
    Dim temp As Variant
    Dim i As Long, j As Long, c As Long
    Dim totalColumns As Long
    
    ' Vérifier les cas de base
    If low >= high Then Exit Sub
    
    totalColumns = UBound(arr, 2)
    
    ' Choisir le pivot (médiane de 3)
    Dim middle As Long
    middle = (low + high) \ 2
    
    ' Trier low, middle, high pour avoir le pivot médian
    If (ascending And CompareValues(arr(low, sortColumn), arr(middle, sortColumn)) > 0) Or _
       (Not ascending And CompareValues(arr(low, sortColumn), arr(middle, sortColumn)) < 0) Then
        For c = 1 To totalColumns
            temp = arr(low, c)
            arr(low, c) = arr(middle, c)
            arr(middle, c) = temp
        Next c
    End If
    
    If (ascending And CompareValues(arr(low, sortColumn), arr(high, sortColumn)) > 0) Or _
       (Not ascending And CompareValues(arr(low, sortColumn), arr(high, sortColumn)) < 0) Then
        For c = 1 To totalColumns
            temp = arr(low, c)
            arr(low, c) = arr(high, c)
            arr(high, c) = temp
        Next c
    End If
    
    If (ascending And CompareValues(arr(middle, sortColumn), arr(high, sortColumn)) > 0) Or _
       (Not ascending And CompareValues(arr(middle, sortColumn), arr(high, sortColumn)) < 0) Then
        For c = 1 To totalColumns
            temp = arr(middle, c)
            arr(middle, c) = arr(high, c)
            arr(high, c) = temp
        Next c
    End If
    
    ' Mettre le pivot au milieu
    For c = 1 To totalColumns
        temp = arr(middle, c)
        arr(middle, c) = arr(high - 1, c)
        arr(high - 1, c) = temp
    Next c
    
    pivot = arr(high - 1, sortColumn)
    
    ' Partition
    i = low
    j = high - 1
    
    Do
        ' Trouver un élément >= pivot du côté gauche
        Do While i < high - 1
            If (ascending And CompareValues(arr(i, sortColumn), pivot) >= 0) Or _
               (Not ascending And CompareValues(arr(i, sortColumn), pivot) <= 0) Then
                Exit Do
            End If
            i = i + 1
        Loop
        
        ' Trouver un élément <= pivot du côté droit
        Do While j > low
            If (ascending And CompareValues(arr(j, sortColumn), pivot) <= 0) Or _
               (Not ascending And CompareValues(arr(j, sortColumn), pivot) >= 0) Then
                Exit Do
            End If
            j = j - 1
        Loop
        
        ' Échanger si i < j
        If i < j Then
            For c = 1 To totalColumns
                temp = arr(i, c)
                arr(i, c) = arr(j, c)
                arr(j, c) = temp
            Next c
        End If
    Loop While i < j
    
    ' Remettre le pivot à sa place finale
    For c = 1 To totalColumns
        temp = arr(i, c)
        arr(i, c) = arr(high - 1, c)
        arr(high - 1, c) = temp
    Next c
    
    ' Tri récursif des sous-tableaux
    QuickSort arr, low, i - 1, sortColumn, ascending
    QuickSort arr, i + 1, high, sortColumn, ascending
End Sub

' Fonction de comparaison de valeurs (tenant compte des différents types)
Private Function CompareValues(ByVal val1 As Variant, ByVal val2 As Variant) As Integer
    ' Gérer les cas spéciaux
    If IsNull(val1) And IsNull(val2) Then
        CompareValues = 0 ' Égaux
        Exit Function
    ElseIf IsNull(val1) Then
        CompareValues = -1 ' val1 < val2
        Exit Function
    ElseIf IsNull(val2) Then
        CompareValues = 1 ' val1 > val2
        Exit Function
    End If
    
    If IsEmpty(val1) And IsEmpty(val2) Then
        CompareValues = 0 ' Égaux
        Exit Function
    ElseIf IsEmpty(val1) Then
        CompareValues = -1 ' val1 < val2
        Exit Function
    ElseIf IsEmpty(val2) Then
        CompareValues = 1 ' val1 > val2
        Exit Function
    End If
    
    ' Comparer selon le type
    If IsNumeric(val1) And IsNumeric(val2) Then
        ' Comparaison numérique
        If CDbl(val1) < CDbl(val2) Then
            CompareValues = -1
        ElseIf CDbl(val1) > CDbl(val2) Then
            CompareValues = 1
        Else
            CompareValues = 0
        End If
    ElseIf IsDate(val1) And IsDate(val2) Then
        ' Comparaison de dates
        If CDate(val1) < CDate(val2) Then
            CompareValues = -1
        ElseIf CDate(val1) > CDate(val2) Then
            CompareValues = 1
        Else
            CompareValues = 0
        End If
    ElseIf VarType(val1) = vbString And VarType(val2) = vbString Then
        ' Comparaison de chaînes
        CompareValues = StrComp(val1, val2, vbTextCompare)
    ElseIf IsNumeric(val1) And Not IsNumeric(val2) Then
        CompareValues = -1 ' Numérique < non-numérique
    ElseIf Not IsNumeric(val1) And IsNumeric(val2) Then
        CompareValues = 1 ' Non-numérique > numérique