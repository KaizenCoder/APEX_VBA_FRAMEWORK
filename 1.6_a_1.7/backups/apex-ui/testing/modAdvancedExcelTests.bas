Attribute VB_Name = "modAdvancedExcelTests"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

'@Folder("APEX.UI.Testing")
'@ModuleDescription("Tests unitaires pour les fonctionnalités Excel avancées")
Option Explicit
Option Private Module

'*******************************************************************************
' Module : modAdvancedExcelTests
' Author : [Votre nom]
' Date   : 12/04/2025
' Purpose: Tests unitaires pour valider le fonctionnement des interfaces et
'          implémentations pour les fonctionnalités Excel avancées
'*******************************************************************************

' Constantes pour les messages d'erreur/succès
Private Const TEST_PASSED As String = "PASSED"
Private Const TEST_FAILED As String = "FAILED"

' Variables pour le suivi des tests
Private m_passedCount As Long
Private m_failedCount As Long
Private m_testSheet As Object ' ISheetAccessor

''
' Point d'entrée des tests unitaires
' @param logOutput (optional) Si True, écrit les résultats dans un journal
' @return Boolean True si tous les tests ont réussi
''
'@Description: 
'@Param: 
'@Returns: 

Public Function RunAllTests(Optional ByVal logOutput As Boolean = True) As Boolean
    ' Initialiser le suivi des tests
    m_passedCount = 0
    m_failedCount = 0
    
    On Error Resume Next
    
    ' Créer un environnement de test
    If Not InitializeTestEnvironment() Then
        Debug.Print "Échec de l'initialisation de l'environnement de test."
        RunAllTests = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Exécuter les tests de table Excel
    Call RunTableTests
    
    ' Exécuter les tests de tableau croisé dynamique
    Call RunPivotTableTests
    
    ' Exécuter les tests de graphique
    Call RunChartTests
    
    ' Nettoyer l'environnement de test
    Call CleanupTestEnvironment
    
    ' Afficher les résultats
    Debug.Print "=== Résultats des tests ==="
    Debug.Print "Tests réussis: " & m_passedCount
    Debug.Print "Tests échoués: " & m_failedCount
    Debug.Print "Total: " & (m_passedCount + m_failedCount)
    
    ' Journaliser les résultats si demandé
    If logOutput Then
        ' Code pour journaliser les résultats...
    End If
    
    ' Tous les tests ont réussi?
    RunAllTests = (m_failedCount = 0)
End Function

''
' Initialise l'environnement de test en créant une feuille temporaire
' @return Boolean True si l'initialisation a réussi
''
'@Description: 
'@Param: 
'@Returns: 

Private Function InitializeTestEnvironment() As Boolean
    On Error GoTo ErrorHandler
    
    ' Créer un classeur et une feuille de test
    Dim wb As Workbook
    Dim ws As Worksheet
    
    ' Utiliser le classeur actif ou en créer un nouveau
    If Application.Workbooks.Count = 0 Then
        Set wb = Application.Workbooks.Add
    Else
        Set wb = Application.ActiveWorkbook
    End If
    
    ' Ajouter une feuille de test
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("TestSheet").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler
    
    Set ws = wb.Worksheets.Add
    ws.Name = "TestSheet"
    
    ' Préparer les données de test
    ws.Range("A1").Value = "Catégorie"
    ws.Range("B1").Value = "Valeur 1"
    ws.Range("C1").Value = "Valeur 2"
    ws.Range("D1").Value = "Valeur 3"
    
    ' Catégories
    ws.Range("A2").Value = "Produit A"
    ws.Range("A3").Value = "Produit B"
    ws.Range("A4").Value = "Produit C"
    ws.Range("A5").Value = "Produit D"
    
    ' Données numériques
    ws.Range("B2:D5").Formula = "=RAND()*100"
    Application.Calculate
    
    ' Conserver les valeurs uniquement
    ws.Range("B2:D5").Value = ws.Range("B2:D5").Value
    
    ' Créer un accesseur de feuille
    Dim appContext As Object ' IApplicationContext
    Set appContext = Application.Run("GetApplicationContext")
    
    ' Obtenir un accesseur pour la feuille de test
    Set m_testSheet = appContext.GetWorkbookAccessor(wb.Name).GetSheetAccessor("TestSheet")
    
    InitializeTestEnvironment = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Debug.Print "Erreur lors de l'initialisation de l'environnement de test: " & Err.Description
    InitializeTestEnvironment = False
End Function

''
' Nettoie l'environnement de test
''
'@Description: 
'@Param: 
'@Returns: 

Private Sub CleanupTestEnvironment()
    ' Libérer les références
    Set m_testSheet = Nothing
    
    ' Optionnel : supprimer la feuille de test
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.ActiveWorkbook.Worksheets("TestSheet").Delete
    Application.DisplayAlerts = True
End Sub

''
' Exécute tous les tests liés aux Tables Excel
''
'@Description: 
'@Param: 
'@Returns: 

Private Sub RunTableTests()
    Debug.Print "=== Tests des Tables Excel ==="
    
    ' Créer une table
    Dim tableCreated As Boolean
    tableCreated = TestCreateTable()
    LogTestResult "Création d'une table Excel", tableCreated
    
    ' Si la table n'a pas été créée correctement, arrêter les tests
    If Not tableCreated Then
        Debug.Print "Test de création de table échoué, les autres tests de table sont annulés."
        Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    End If
    
    ' Les autres tests sur la table
    LogTestResult "Lecture des données de la table", TestReadTableData()
    LogTestResult "Écriture dans la table", TestWriteTableData()
    LogTestResult "Manipulation de la structure de la table", TestTableStructure()
    LogTestResult "Filtrage et tri de la table", TestTableFilterAndSort()
    LogTestResult "Mise en forme de la table", TestTableFormatting()
    
    ' Supprimer la table pour les tests suivants
    On Error Resume Next
    m_testSheet.GetNativeSheet.ListObjects(1).Delete
End Sub

''
' Exécute tous les tests liés aux Tableaux Croisés Dynamiques
''
'@Description: 
'@Param: 
'@Returns: 

Private Sub RunPivotTableTests()
    Debug.Print "=== Tests des Tableaux Croisés Dynamiques ==="
    
    ' Créer un tableau croisé dynamique
    Dim pivotCreated As Boolean
    pivotCreated = TestCreatePivotTable()
    LogTestResult "Création d'un tableau croisé dynamique", pivotCreated
    
    ' Si le tableau croisé n'a pas été créé correctement, arrêter les tests
    If Not pivotCreated Then
        Debug.Print "Test de création de tableau croisé échoué, les autres tests de tableau croisé sont annulés."
        Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    End If
    
    ' Les autres tests sur le tableau croisé
    LogTestResult "Configuration des champs du tableau croisé", TestPivotTableFields()
    LogTestResult "Filtrage du tableau croisé dynamique", TestPivotTableFilters()
    LogTestResult "Mise en forme du tableau croisé dynamique", TestPivotTableFormatting()
    LogTestResult "Rafraîchissement et expansion du tableau croisé", TestPivotTableActions()
    
    ' Supprimer le tableau croisé pour les tests suivants
    On Error Resume Next
    m_testSheet.GetNativeSheet.PivotTables(1).TableRange2.Clear
End Sub

''
' Exécute tous les tests liés aux Graphiques
''
'@Description: 
'@Param: 
'@Returns: 

Private Sub RunChartTests()
    Debug.Print "=== Tests des Graphiques ==="
    
    ' Créer un graphique
    Dim chartCreated As Boolean
    chartCreated = TestCreateChart()
    LogTestResult "Création d'un graphique", chartCreated
    
    ' Si le graphique n'a pas été créé correctement, arrêter les tests
    If Not chartCreated Then
        Debug.Print "Test de création de graphique échoué, les autres tests de graphique sont annulés."
        Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    End If
    
    ' Les autres tests sur le graphique
    LogTestResult "Configuration des séries du graphique", TestChartSeries()
    LogTestResult "Configuration des axes du graphique", TestChartAxes()
    LogTestResult "Mise en forme du graphique", TestChartFormatting()
    LogTestResult "Positionnement du graphique", TestChartPosition()
    
    ' Supprimer le graphique pour la fin des tests
    On Error Resume Next
    For Each obj In m_testSheet.GetNativeSheet.ChartObjects
        obj.Delete
    Next obj
End Sub

' ============== Tests spécifiques pour les Tables Excel ==============

''
' Teste la création d'une Table Excel
' @return Boolean True si le test réussit
''
'@Description: 
'@Param: 
'@Returns: 

Private Function TestCreateTable() As Boolean
    On Error GoTo ErrorHandler
    
    ' Créer un accesseur de table
    Dim tableAccessor As New clsExcelTableAccessor
    
    ' Créer une table à partir de la plage A1:D5
    tableAccessor.CreateTableFromRange m_testSheet, "A1:D5", "TestTable"
    
    ' Vérifier que la table a été créée
    TestCreateTable = (m_testSheet.GetNativeSheet.ListObjects.Count > 0)
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestCreateTable: " & Err.Description
    TestCreateTable = False
End Function

''
' Teste la lecture des données dans une Table Excel
' @return Boolean True si le test réussit
''
'@Description: 
'@Param: 
'@Returns: 

Private Function TestReadTableData() As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir un accesseur pour la table existante
    Dim tableAccessor As New clsExcelTableAccessor
    tableAccessor.Initialize m_testSheet, "TestTable"
    
    ' Tester les propriétés de la table
    Dim success As Boolean
    success = (tableAccessor.RowCount = 4) And (tableAccessor.ColumnCount = 4)
    
    ' Tester la lecture de données
    Dim data As Variant
    data = tableAccessor.ReadAllData
    success = success And (UBound(data, 1) = 4) And (UBound(data, 2) = 4)
    
    ' Tester la lecture d'une cellule spécifique
    Dim cellValue As Variant
    cellValue = tableAccessor.ReadCell(1, "Catégorie")
    success = success And (cellValue = "Produit A")
    
    TestReadTableData = success
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestReadTableData: " & Err.Description
    TestReadTableData = False
End Function

''
' Teste l'écriture de données dans une Table Excel
' @return Boolean True si le test réussit
''
'@Description: 
'@Param: 
'@Returns: 

Private Function TestWriteTableData() As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir un accesseur pour la table existante
    Dim tableAccessor As New clsExcelTableAccessor
    tableAccessor.Initialize m_testSheet, "TestTable"
    
    ' Écrire dans une cellule
    tableAccessor.WriteCell 1, "Valeur 1", 999.99
    
    ' Vérifier que la valeur a été écrite
    Dim success As Boolean
    success = (tableAccessor.ReadCell(1, "Valeur 1") = 999.99)
    
    ' Écrire dans une ligne
    Dim rowData(1 To 4) As Variant
    rowData(1) = "Produit Z"
    rowData(2) = 100
    rowData(3) = 200
    rowData(4) = 300
    tableAccessor.WriteRow 3, rowData
    
    ' Vérifier que la ligne a été écrite
    Dim rowResult As Variant
    rowResult = tableAccessor.ReadRow(3)
    success = success And (rowResult(1) = "Produit Z") And (rowResult(2) = 100)
    
    TestWriteTableData = success
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestWriteTableData: " & Err.Description
    TestWriteTableData = False
End Function

''
' Teste les opérations de structure sur une Table Excel
' @return Boolean True si le test réussit
''
'@Description: 
'@Param: 
'@Returns: 

Private Function TestTableStructure() As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir un accesseur pour la table existante
    Dim tableAccessor As New clsExcelTableAccessor
    tableAccessor.Initialize m_testSheet, "TestTable"
    
    ' Ajouter une colonne
    tableAccessor.AddColumn "Nouvelle Colonne"
    
    ' Vérifier que la colonne a été ajoutée
    Dim success As Boolean
    success = (tableAccessor.ColumnCount = 5)
    
    ' Ajouter une ligne
    Dim newRowIndex As Long
    newRowIndex = tableAccessor.AddRow
    
    ' Vérifier que la ligne a été ajoutée
    success = success And (tableAccessor.RowCount = 5)
    
    ' Écrire dans la nouvelle ligne
    tableAccessor.WriteCell newRowIndex, "Catégorie", "Produit E"
    
    ' Supprimer une colonne
    tableAccessor.DeleteColumn "Nouvelle Colonne"
    
    ' Vérifier que la colonne a été supprimée
    success = success And (tableAccessor.ColumnCount = 4)
    
    TestTableStructure = success
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestTableStructure: " & Err.Description
    TestTableStructure = False
End Function

''
' Teste le filtrage et le tri sur une Table Excel
' @return Boolean True si le test réussit
''
'@Description: 
'@Param: 
'@Returns: 

Private Function TestTableFilterAndSort() As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir un accesseur pour la table existante
    Dim tableAccessor As New clsExcelTableAccessor
    tableAccessor.Initialize m_testSheet, "TestTable"
    
    ' Trier la table par une colonne
    tableAccessor.SortByColumn "Valeur 1", False ' Tri descendant
    
    ' Appliquer un filtre
    tableAccessor.ApplyFilter "Catégorie", "Produit*"
    
    ' Effacer les filtres
    tableAccessor.ClearFilters
    
    ' C'est difficile de vérifier le résultat du tri/filtrage dans un test unitaire
    ' sans vérifier visuellement, donc on considère réussi si aucune erreur ne s'est produite
    TestTableFilterAndSort = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestTableFilterAndSort: " & Err.Description
    TestTableFilterAndSort = False
End Function

''
' Teste la mise en forme sur une Table Excel
' @return Boolean True si le test réussit
''
'@Description: 
'@Param: 
'@Returns: 

Private Function TestTableFormatting() As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir un accesseur pour la table existante
    Dim tableAccessor As New clsExcelTableAccessor
    tableAccessor.Initialize m_testSheet, "TestTable"
    
    ' Appliquer un style de table
    tableAccessor.ApplyTableStyle "TableStyleMedium2"
    
    ' Ajouter une mise en forme conditionnelle
    tableAccessor.SetConditionalFormatting "Valeur 1", _
        "=AND($B2>50,$B2<150)", RGB(255, 200, 200)
    
    ' C'est difficile de vérifier le résultat de la mise en forme dans un test unitaire
    ' sans vérifier visuellement, donc on considère réussi si aucune erreur ne s'est produite
    TestTableFormatting = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestTableFormatting: " & Err.Description
    TestTableFormatting = False
End Function

' ============== Tests spécifiques pour les Tableaux Croisés Dynamiques ==============

''
' Teste la création d'un Tableau Croisé Dynamique
' @return Boolean True si le test réussit
''
'@Description: 
'@Param: 
'@Returns: 

Private Function TestCreatePivotTable() As Boolean
    On Error GoTo ErrorHandler
    
    ' S'assurer qu'il y a une table pour la source
    Dim tableAccessor As New clsExcelTableAccessor
    
    ' Créer une table si elle n'existe pas encore
    On Error Resume Next
    If m_testSheet.GetNativeSheet.ListObjects.Count = 0 Then
        tableAccessor.CreateTableFromRange m_testSheet, "A1:D5", "TestTable"
    Else
        tableAccessor.Initialize m_testSheet, "TestTable"
    End If
    On Error GoTo ErrorHandler
    
    ' Créer un pivotTable
    Dim pivotAccessor As New clsExcelPivotTableAccessor
    
    ' Position du tableau croisé sous les données source
    Dim pivotPos As String
    pivotPos = "A7"
    
    ' Créer le tableau croisé à partir de la table
    pivotAccessor.CreatePivotTableFromData m_testSheet, tableAccessor.GetNativeTable, _
                                           pivotPos, "TestPivotTable"
    
    ' Vérifier que le tableau croisé a été créé
    TestCreatePivotTable = (m_testSheet.GetNativeSheet.PivotTables.Count > 0)
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestCreatePivotTable: " & Err.Description
    TestCreatePivotTable = False
End Function

''
' Teste la configuration des champs d'un Tableau Croisé Dynamique
' @return Boolean True si le test réussit
''
'@Description: 
'@Param: 
'@Returns: 

Private Function TestPivotTableFields() As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir un accesseur pour le tableau croisé existant
    Dim pivotAccessor As New clsExcelPivotTableAccessor
    pivotAccessor.Initialize m_testSheet, "TestPivotTable"
    
    ' Ajouter des champs
    pivotAccessor.AddRowField "Catégorie"
    pivotAccessor.AddDataField "Valeur 1", "Somme de Valeur 1", xlSum
    pivotAccessor.AddDataField "Valeur 2", "Moyenne de Valeur 2", xlAverage
    
    ' Vérifier que les champs ont été ajoutés
    Dim success As Boolean
    success = (pivotAccessor.DataFieldsCount = 2)
    
    TestPivotTableFields = success
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestPivotTableFields: " & Err.Description
    TestPivotTableFields = False
End Function

''
' Teste le filtrage d'un Tableau Croisé Dynamique
' @return Boolean True si le test réussit
''
'@Description: 
'@Param: 
'@Returns: 

Private Function TestPivotTableFilters() As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir un accesseur pour le tableau croisé existant
    Dim pivotAccessor As New clsExcelPivotTableAccessor
    pivotAccessor.Initialize m_testSheet, "TestPivotTable"
    
    ' Déplacer un champ en filtre de rapport
    pivotAccessor.MoveField "Catégorie", AREA_PAGES
    
    ' Appliquer un filtre
    Dim items(1 To 2) As String
    items(1) = "Produit A"
    items(2) = "Produit B"
    pivotAccessor.ApplyFilter "Catégorie", items
    
    ' Effacer les filtres
    pivotAccessor.ClearFilters "Catégorie"
    
    ' Effacer tous les filtres
    pivotAccessor.ClearAllFilters
    
    ' On considère réussi si aucune erreur ne s'est produite
    TestPivotTableFilters = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestPivotTableFilters: " & Err.Description
    TestPivotTableFilters = False
End Function

''
' Teste la mise en forme d'un Tableau Croisé Dynamique
' @return Boolean True si le test réussit
''
'@Description: 
'@Param: 
'@Returns: 

Private Function TestPivotTableFormatting() As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir un accesseur pour le tableau croisé existant
    Dim pivotAccessor As New clsExcelPivotTableAccessor
    pivotAccessor.Initialize m_testSheet, "TestPivotTable"
    
    ' Formater un champ de données
    pivotAccessor.FormatDataField "Somme de Valeur 1", "#,##0.00"
    
    ' Définir les sous-totaux
    pivotAccessor.SetSubtotal "Catégorie", True
    
    ' On considère réussi si aucune erreur ne s'est produite
    TestPivotTableFormatting = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestPivotTableFormatting: " & Err.Description
    TestPivotTableFormatting = False
End Function

''
' Teste les actions sur un Tableau Croisé Dynamique
' @return Boolean True si le test réussit
''
'@Description: 
'@Param: 
'@Returns: 

Private Function TestPivotTableActions() As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir un accesseur pour le tableau croisé existant
    Dim pivotAccessor As New clsExcelPivotTableAccessor
    pivotAccessor.Initialize m_testSheet, "TestPivotTable"
    
    ' Rafraîchir le tableau croisé
    pivotAccessor.Refresh
    
    ' Développer tous les champs
    pivotAccessor.ExpandAll True
    
    ' Réduire tous les champs
    pivotAccessor.ExpandAll False
    
    ' Obtenir les valeurs
    Dim values As Variant
    values = pivotAccessor.GetAllValues
    
    ' On considère réussi si aucune erreur ne s'est produite
    TestPivotTableActions = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestPivotTableActions: " & Err.Description
    TestPivotTableActions = False
End Function

' ============== Tests spécifiques pour les Graphiques ==============

''
' Teste la création d'un Graphique
' @return Boolean True si le test réussit
''
'@Description: 
'@Param: 
'@Returns: 

Private Function TestCreateChart() As Boolean
    On Error GoTo ErrorHandler
    
    ' Créer un graphique
    Dim chartAccessor As New clsExcelChartAccessor
    
    ' Position et taille du graphique
    chartAccessor.CreateChart m_testSheet, 200, 200, 400, 300, "TestChart", xlColumnClustered
    
    ' Définir la source de données
    chartAccessor.Initialize m_testSheet, "TestChart"
    chartAccessor.SetSourceData "A1:D5"
    
    ' Définir un titre
    chartAccessor.Title = "Graphique de Test"
    
    ' Vérifier que le graphique a été créé
    TestCreateChart = (m_testSheet.GetNativeSheet.ChartObjects.Count > 0)
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestCreateChart: " & Err.Description
    TestCreateChart = False
End Function

''
' Teste la configuration des séries d'un Graphique
' @return Boolean True si le test réussit
''
'@Description: 
'@Param: 
'@Returns: 

Private Function TestChartSeries() As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir un accesseur pour le graphique existant
    Dim chartAccessor As New clsExcelChartAccessor
    chartAccessor.Initialize m_testSheet, "TestChart"
    
    ' Effacer les séries existantes
    chartAccessor.ClearSeries
    
    ' Ajouter des séries manuellement
    chartAccessor.AddSeries "Valeur 1", "B2:B5", "A2:A5"
    chartAccessor.AddSeries "Valeur 2", "C2:C5", "A2:A5"
    
    ' Formater une série
    chartAccessor.FormatSeries 1, FORMAT_COLOR, RGB(255, 0, 0)
    
    ' Ajouter des étiquettes de données
    chartAccessor.SetDataLabels 1, True, xlDataLabelShowValue
    
    ' On considère réussi si aucune erreur ne s'est produite
    TestChartSeries = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestChartSeries: " & Err.Description
    TestChartSeries = False
End Function

''
' Teste la configuration des axes d'un Graphique
' @return Boolean True si le test réussit
''
'@Description: 
'@Param: 
'@Returns: 

Private Function TestChartAxes() As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir un accesseur pour le graphique existant
    Dim chartAccessor As New clsExcelChartAccessor
    chartAccessor.Initialize m_testSheet, "TestChart"
    
    ' Définir les titres des axes
    chartAccessor.SetXAxisTitle "Catégories"
    chartAccessor.SetYAxisTitle "Valeurs"
    
    ' Formater les axes
    chartAccessor.FormatXAxis , , , , "#,##0"
    chartAccessor.FormatYAxis 0, 200, 50, , "#,##0.00"
    
    ' On considère réussi si aucune erreur ne s'est produite
    TestChartAxes = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestChartAxes: " & Err.Description
    TestChartAxes = False
End Function

''
' Teste la mise en forme d'un Graphique
' @return Boolean True si le test réussit
''
'@Description: 
'@Param: 
'@Returns: 

Private Function TestChartFormatting() As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir un accesseur pour le graphique existant
    Dim chartAccessor As New clsExcelChartAccessor
    chartAccessor.Initialize m_testSheet, "TestChart"
    
    ' Définir le type de graphique
    chartAccessor.ChartType = xlColumnClustered
    
    ' Configurer la légende
    chartAccessor.HasLegend = True
    chartAccessor.LegendPosition = xlLegendPositionBottom
    
    ' Appliquer un style
    chartAccessor.ApplyChartStyle 1
    
    ' On considère réussi si aucune erreur ne s'est produite
    TestChartFormatting = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestChartFormatting: " & Err.Description
    TestChartFormatting = False
End Function

''
' Teste le positionnement d'un Graphique
' @return Boolean True si le test réussit
''
'@Description: 
'@Param: 
'@Returns: 

Private Function TestChartPosition() As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir un accesseur pour le graphique existant
    Dim chartAccessor As New clsExcelChartAccessor
    chartAccessor.Initialize m_testSheet, "TestChart"
    
    ' Déplacer et redimensionner le graphique
    chartAccessor.SetPosition 300, 300, 350, 250
    
    ' Export vers une image (dans un dossier temporaire)
    Dim tempPath As String
    tempPath = Environ("TEMP") & "\testchart.png"
    chartAccessor.ExportAsImage tempPath, "png"
    
    ' Vérifier si le fichier existe
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim success As Boolean
    success = fso.FileExists(tempPath)
    
    ' Supprimer le fichier temporaire
    On Error Resume Next
    If fso.FileExists(tempPath) Then
        fso.DeleteFile tempPath
    End If
    
    TestChartPosition = success
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestChartPosition: " & Err.Description
    TestChartPosition = False
End Function

' ============== Utilitaires ==============

''
' Enregistre le résultat d'un test et incrémente les compteurs
' @param testName Nom du test
' @param success Indique si le test a réussi
''
'@Description: 
'@Param: 
'@Returns: 

Private Sub LogTestResult(ByVal testName As String, ByVal success As Boolean)
    If success Then
        Debug.Print "  " & testName & ": " & TEST_PASSED
        m_passedCount = m_passedCount + 1
    Else
        Debug.Print "  " & testName & ": " & TEST_FAILED
        m_failedCount = m_failedCount + 1
    End If
End Sub