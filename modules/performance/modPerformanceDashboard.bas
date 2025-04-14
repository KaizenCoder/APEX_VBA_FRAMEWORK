Attribute VB_Name = "modPerformanceDashboard"
'------------------------------------------------------------------------------
' Module       : modPerformanceDashboard
' Description  : Module pour générer des tableaux de bord de performance Excel
' Date         : 14/04/2025
' Auteur       : APEX Framework Team
' Version      : 1.0
'------------------------------------------------------------------------------
Option Explicit

' Constantes pour la configuration du tableau de bord
Private Const DASHBOARD_SHEET_NAME As String = "Performance_Dashboard"
Private Const RESULTS_SHEET_NAME As String = "Performance_Results"
Private Const CHART_WIDTH As Double = 400
Private Const CHART_HEIGHT As Double = 250
Private Const CHART_TOP_MARGIN As Double = 20
Private Const CHART_LEFT_MARGIN As Double = 50
Private Const SPACE_BETWEEN_CHARTS As Double = 30

' Énumération pour les types de graphiques
Public Enum ChartTypePerformance
    CT_BarChart = 1
    CT_LineChart = 2
    CT_PieChart = 3
    CT_ScatterChart = 4
End Enum

' Structure pour configurer un graphique
Private Type ChartConfig
    Title As String
    ChartType As ChartTypePerformance
    DataRange As Range
    XAxisTitle As String
    YAxisTitle As String
    ShowLegend As Boolean
    Position As Range
End Type

'------------------------------------------------------------------------------
' Procédure     : GeneratePerformanceDashboard
' Description   : Génère un tableau de bord complet à partir des résultats de test
' Paramètres    : 
'   - resultFilePath: Chemin du fichier CSV contenant les résultats des tests
'   - targetWorkbook: Classeur cible (Optional, utilise ActiveWorkbook si absent)
'------------------------------------------------------------------------------
Public Sub GeneratePerformanceDashboard(ByVal resultFilePath As String, _
                                      Optional ByVal targetWorkbook As Workbook = Nothing)
    On Error GoTo ErrorHandler
    
    ' Valider les paramètres d'entrée
    If Len(Dir(resultFilePath)) = 0 Then
        MsgBox "Le fichier de résultats n'existe pas: " & resultFilePath, vbExclamation
        Exit Sub
    End If
    
    ' Utiliser ActiveWorkbook si targetWorkbook n'est pas fourni
    If targetWorkbook Is Nothing Then
        Set targetWorkbook = ActiveWorkbook
    End If
    
    ' Configurer Excel pour les performances
    OptimizeForPerformance
    
    ' Importer les données
    Dim resultsSheet As Worksheet
    Set resultsSheet = ImportPerformanceData(resultFilePath, targetWorkbook)
    
    ' Créer le tableau de bord
    Dim dashboardSheet As Worksheet
    Set dashboardSheet = CreateDashboardSheet(targetWorkbook)
    
    ' Générer les graphiques
    GenerateCharts resultsSheet, dashboardSheet
    
    ' Ajouter un tableau récapitulatif
    AddSummaryTable resultsSheet, dashboardSheet
    
    ' Ajouter des filtres interactifs
    AddInteractiveFilters dashboardSheet
    
    ' Mise en forme finale
    FormatDashboard dashboardSheet
    
    ' Activer la feuille de tableau de bord
    dashboardSheet.Activate
    
    ' Restaurer les paramètres Excel
    RestoreExcelState
    
    MsgBox "Tableau de bord de performance généré avec succès!", vbInformation
    Exit Sub
    
ErrorHandler:
    RestoreExcelState
    MsgBox "Erreur lors de la génération du tableau de bord: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' Fonction      : ImportPerformanceData
' Description   : Importe les données de performance depuis le fichier CSV
' Paramètres    : 
'   - filePath: Chemin du fichier CSV
'   - targetWb: Classeur cible
' Retour        : Feuille contenant les données importées
'------------------------------------------------------------------------------
Private Function ImportPerformanceData(ByVal filePath As String, ByVal targetWb As Workbook) As Worksheet
    On Error GoTo ErrorHandler
    
    ' Supprimer la feuille si elle existe déjà
    On Error Resume Next
    Application.DisplayAlerts = False
    targetWb.Worksheets(RESULTS_SHEET_NAME).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler
    
    ' Créer une nouvelle feuille
    Dim ws As Worksheet
    Set ws = targetWb.Worksheets.Add
    ws.Name = RESULTS_SHEET_NAME
    
    ' Ouvrir le fichier CSV
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    ' Lire l'en-tête
    Dim headerLine As String
    Line Input #fileNum, headerLine
    
    Dim headers() As String
    headers = Split(headerLine, ",")
    
    ' Écrire les en-têtes
    Dim col As Integer
    For col = 0 To UBound(headers)
        ws.Cells(1, col + 1).Value = headers(col)
    Next col
    
    ' Lire et écrire les données
    Dim rowNum As Long
    rowNum = 2 ' Commencer à la ligne 2 après les en-têtes
    
    Dim dataLine As String
    Dim dataValues() As String
    
    Do Until EOF(fileNum)
        Line Input #fileNum, dataLine
        dataValues = Split(dataLine, ",")
        
        For col = 0 To UBound(dataValues)
            ws.Cells(rowNum, col + 1).Value = dataValues(col)
        Next col
        
        rowNum = rowNum + 1
    Loop
    
    Close #fileNum
    
    ' Formater en tant que tableau
    Dim dataRange As Range
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowNum - 1, UBound(headers) + 1))
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
    tbl.Name = "PerformanceData"
    
    ' Appliquer un format
    With tbl
        .TableStyle = "TableStyleMedium2"
        .Range.Columns.AutoFit
    End With
    
    ' Ajouter des filtres avancés
    ws.Range("A1").AutoFilter
    
    Set ImportPerformanceData = ws
    Exit Function
    
ErrorHandler:
    Close #fileNum
    MsgBox "Erreur lors de l'importation des données: " & Err.Description, vbCritical
    Set ImportPerformanceData = Nothing
End Function

'------------------------------------------------------------------------------
' Fonction      : CreateDashboardSheet
' Description   : Crée la feuille du tableau de bord
' Paramètres    : 
'   - targetWb: Classeur cible
' Retour        : Feuille du tableau de bord
'------------------------------------------------------------------------------
Private Function CreateDashboardSheet(ByVal targetWb As Workbook) As Worksheet
    On Error Resume Next
    
    ' Supprimer la feuille si elle existe déjà
    Application.DisplayAlerts = False
    targetWb.Worksheets(DASHBOARD_SHEET_NAME).Delete
    Application.DisplayAlerts = True
    
    ' Créer une nouvelle feuille
    Dim ws As Worksheet
    Set ws = targetWb.Worksheets.Add(Before:=targetWb.Worksheets(1))
    ws.Name = DASHBOARD_SHEET_NAME
    
    ' Configurer la feuille
    ws.Range("A1").Value = "APEX Framework - Tableau de Bord de Performance"
    ws.Range("A2").Value = "Généré le: " & Format(Now, "dd/mm/yyyy à hh:mm:ss")
    
    ' Mise en forme
    With ws.Range("A1")
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(44, 62, 80)
    End With
    
    With ws.Range("A2")
        .Font.Size = 10
        .Font.Italic = True
    End With
    
    ' Configurer la mise en page
    With ws.PageSetup
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .CenterVertically = False
        .PrintGridlines = False
    End With
    
    Set CreateDashboardSheet = ws
End Function

'------------------------------------------------------------------------------
' Procédure     : GenerateCharts
' Description   : Génère les graphiques sur le tableau de bord
' Paramètres    : 
'   - dataSheet: Feuille contenant les données
'   - dashboardSheet: Feuille du tableau de bord
'------------------------------------------------------------------------------
Private Sub GenerateCharts(ByVal dataSheet As Worksheet, ByVal dashboardSheet As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Position de départ pour les graphiques
    Dim topPosition As Double
    topPosition = CHART_TOP_MARGIN + 60 ' Laisser de la place pour le titre
    
    ' Nettoyage des graphiques existants
    Dim cht As ChartObject
    For Each cht In dashboardSheet.ChartObjects
        cht.Delete
    Next cht
    
    ' 1. Graphique des temps d'exécution par type de test
    Dim chartConfig1 As ChartConfig
    With chartConfig1
        .Title = "Temps d'Exécution par Type de Test"
        .ChartType = CT_BarChart
        .XAxisTitle = "Type de Test"
        .YAxisTitle = "Temps (secondes)"
        .ShowLegend = False
        Set .Position = dashboardSheet.Range("B5")
    End With
    
    ' Créer une requête PowerPivot/tableau croisé dynamique pour agréger les données
    Dim ptSheet As Worksheet
    Dim pt As PivotTable
    Dim pc As PivotCache
    
    On Error Resume Next
    Set ptSheet = dataSheet.Parent.Worksheets.Add
    ptSheet.Name = "TempPivot"
    ptSheet.Visible = xlSheetVeryHidden
    
    ' Créer le cache du tableau croisé
    Set pc = dataSheet.Parent.PivotCaches.Create(xlDatabase, dataSheet.ListObjects("PerformanceData").Range)
    Set pt = pc.CreatePivotTable(ptSheet.Range("A3"), "PerfPivot")
    
    ' Configurer le pivot
    With pt
        .PivotFields("TestType").Orientation = xlRowField
        .PivotFields("ExecutionTime").Orientation = xlDataField
        .PivotFields("Sum of ExecutionTime").Function = xlAverage
        .PivotFields("Sum of ExecutionTime").Name = "Temps Moyen (s)"
    End With
    
    ' Créer le graphique
    AddPivotChart dashboardSheet, pt.TableRange2, chartConfig1
    
    ' 2. Graphique des performances de mémoire
    Dim chartConfig2 As ChartConfig
    With chartConfig2
        .Title = "Utilisation Mémoire par Opération"
        .ChartType = CT_LineChart
        .XAxisTitle = "Type d'Opération"
        .YAxisTitle = "Mémoire (MB)"
        .ShowLegend = False
        Set .Position = dashboardSheet.Range("J5")
    End With
    
    ' Configurer le nouveau pivot pour la mémoire
    Set pt = pc.CreatePivotTable(ptSheet.Range("A20"), "MemPivot")
    
    With pt
        .PivotFields("OperationType").Orientation = xlRowField
        .PivotFields("MemoryDelta").Orientation = xlDataField
        .PivotFields("Sum of MemoryDelta").Function = xlAverage
        .PivotFields("Sum of MemoryDelta").Name = "Mémoire Moyenne (MB)"
    End With
    
    AddPivotChart dashboardSheet, pt.TableRange2, chartConfig2
    
    ' 3. Graphique de comparaison des méthodes d'accès
    Dim chartConfig3 As ChartConfig
    With chartConfig3
        .Title = "Comparaison des Méthodes d'Accès"
        .ChartType = CT_BarChart
        .XAxisTitle = "Méthode"
        .YAxisTitle = "Temps (secondes)"
        .ShowLegend = False
        Set .Position = dashboardSheet.Range("B18")
    End With
    
    ' Filtrer pour avoir uniquement les tests de méthodes d'accès
    Set pt = pc.CreatePivotTable(ptSheet.Range("A40"), "AccessMethodsPivot")
    
    With pt
        .PivotFields("TestName").Orientation = xlPageField
        .PivotFields("TestName").CurrentPage = "TestPerformance_CompareAccessMethods"
        .PivotFields("AccessMethod").Orientation = xlRowField
        .PivotFields("ExecutionTime").Orientation = xlDataField
        .PivotFields("Sum of ExecutionTime").Name = "Temps (s)"
    End With
    
    AddPivotChart dashboardSheet, pt.TableRange2, chartConfig3
    
    ' 4. Graphique d'évolution des performances au fil du temps
    Dim chartConfig4 As ChartConfig
    With chartConfig4
        .Title = "Analyse des Performances par Volume"
        .ChartType = CT_LineChart
        .XAxisTitle = "Volume (nombre de cellules)"
        .YAxisTitle = "Temps (secondes)"
        .ShowLegend = True
        Set .Position = dashboardSheet.Range("J18")
    End With
    
    ' Configurer le pivot pour les différents volumes
    Set pt = pc.CreatePivotTable(ptSheet.Range("A60"), "VolumesPivot")
    
    With pt
        .PivotFields("CellCount").Orientation = xlRowField
        .PivotFields("OperationType").Orientation = xlColumnField
        .PivotFields("ExecutionTime").Orientation = xlDataField
        .PivotFields("Sum of ExecutionTime").Name = "Temps (s)"
    End With
    
    AddPivotChart dashboardSheet, pt.TableRange2, chartConfig4
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Erreur dans GenerateCharts: " & Err.Description
    ' Continue l'exécution
End Sub

'------------------------------------------------------------------------------
' Procédure     : AddPivotChart
' Description   : Ajoute un graphique basé sur un pivot au tableau de bord
' Paramètres    : 
'   - dashboardSheet: Feuille du tableau de bord
'   - dataRange: Plage de données source
'   - config: Configuration du graphique
'------------------------------------------------------------------------------
Private Sub AddPivotChart(ByVal dashboardSheet As Worksheet, _
                        ByVal dataRange As Range, _
                        ByVal config As ChartConfig)
    
    On Error GoTo ErrorHandler
    
    ' Déterminer la position du graphique
    Dim leftPosition As Double
    Dim topPosition As Double
    
    leftPosition = config.Position.Left
    topPosition = config.Position.Top
    
    ' Créer le graphique
    Dim chartObj As ChartObject
    Set chartObj = dashboardSheet.ChartObjects.Add(leftPosition, topPosition, CHART_WIDTH, CHART_HEIGHT)
    
    ' Configurer le graphique
    With chartObj.Chart
        ' Définir le type de graphique
        Select Case config.ChartType
            Case CT_BarChart
                .ChartType = xlColumnClustered
            Case CT_LineChart
                .ChartType = xlLine
            Case CT_PieChart
                .ChartType = xlPie
            Case CT_ScatterChart
                .ChartType = xlXYScatterLines
            Case Else
                .ChartType = xlColumnClustered
        End Select
        
        ' Définir la source de données
        .SetSourceData dataRange
        
        ' Configurer le titre
        .HasTitle = True
        .ChartTitle.Text = config.Title
        
        ' Configurer les axes
        On Error Resume Next ' Certains types de graphiques n'ont pas ces axes
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = config.XAxisTitle
        
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = config.YAxisTitle
        On Error GoTo ErrorHandler
        
        ' Configurer la légende
        .HasLegend = config.ShowLegend
        
        ' Appliquer un style
        .ChartStyle = 201
        
        ' Mise en forme avancée
        .ApplyLayout 3
        
        ' Configurer la couleur de fond
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(248, 248, 248)
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Erreur dans AddPivotChart: " & Err.Description
    ' Continue l'exécution
End Sub

'------------------------------------------------------------------------------
' Procédure     : AddSummaryTable
' Description   : Ajoute un tableau récapitulatif au tableau de bord
' Paramètres    : 
'   - dataSheet: Feuille contenant les données
'   - dashboardSheet: Feuille du tableau de bord
'------------------------------------------------------------------------------
Private Sub AddSummaryTable(ByVal dataSheet As Worksheet, ByVal dashboardSheet As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Position du tableau récapitulatif
    Dim tableRange As Range
    Set tableRange = dashboardSheet.Range("B31:F38")
    
    ' Entêtes du tableau
    tableRange.Cells(1, 1).Value = "Mesure"
    tableRange.Cells(1, 2).Value = "Min"
    tableRange.Cells(1, 3).Value = "Max"
    tableRange.Cells(1, 4).Value = "Moyenne"
    tableRange.Cells(1, 5).Value = "Écart type"
    
    ' Lignes du tableau
    tableRange.Cells(2, 1).Value = "Temps d'exécution (s)"
    tableRange.Cells(3, 1).Value = "Mémoire utilisée (MB)"
    tableRange.Cells(4, 1).Value = "Cellules traitées/sec"
    tableRange.Cells(5, 1).Value = "Efficacité mémoire (cellules/MB)"
    tableRange.Cells(6, 1).Value = "Opérations/sec"
    tableRange.Cells(7, 1).Value = "Gain vs méthode standard"
    tableRange.Cells(8, 1).Value = "Score de performance"
    
    ' Calculs des statistiques (simulation - dans un cas réel, ces valeurs seraient calculées)
    ' Temps d'exécution
    tableRange.Cells(2, 2).Value = Application.WorksheetFunction.Min( _
        dataSheet.ListObjects("PerformanceData").ListColumns("ExecutionTime").DataBodyRange)
    tableRange.Cells(2, 3).Value = Application.WorksheetFunction.Max( _
        dataSheet.ListObjects("PerformanceData").ListColumns("ExecutionTime").DataBodyRange)
    tableRange.Cells(2, 4).Value = Application.WorksheetFunction.Average( _
        dataSheet.ListObjects("PerformanceData").ListColumns("ExecutionTime").DataBodyRange)
    tableRange.Cells(2, 5).Value = Application.WorksheetFunction.StDev( _
        dataSheet.ListObjects("PerformanceData").ListColumns("ExecutionTime").DataBodyRange)
    
    ' Mémoire utilisée (avec gestion des valeurs manquantes)
    On Error Resume Next
    tableRange.Cells(3, 2).Value = Application.WorksheetFunction.Min( _
        dataSheet.ListObjects("PerformanceData").ListColumns("MemoryDelta").DataBodyRange)
    tableRange.Cells(3, 3).Value = Application.WorksheetFunction.Max( _
        dataSheet.ListObjects("PerformanceData").ListColumns("MemoryDelta").DataBodyRange)
    tableRange.Cells(3, 4).Value = Application.WorksheetFunction.Average( _
        dataSheet.ListObjects("PerformanceData").ListColumns("MemoryDelta").DataBodyRange)
    tableRange.Cells(3, 5).Value = Application.WorksheetFunction.StDev( _
        dataSheet.ListObjects("PerformanceData").ListColumns("MemoryDelta").DataBodyRange)
    On Error GoTo ErrorHandler
    
    ' Format du tableau
    With tableRange
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        
        ' En-têtes
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(220, 230, 241)
        
        ' Alternance des lignes
        For i = 2 To 8 Step 2
            .Rows(i).Interior.Color = RGB(240, 240, 240)
        Next i
        
        ' Alignement
        .Columns(1).HorizontalAlignment = xlLeft
        .Columns(2).HorizontalAlignment = xlRight
        .Columns(3).HorizontalAlignment = xlRight
        .Columns(4).HorizontalAlignment = xlRight
        .Columns(5).HorizontalAlignment = xlRight
        
        ' Format des nombres
        .Columns(2).NumberFormat = "0.000"
        .Columns(3).NumberFormat = "0.000"
        .Columns(4).NumberFormat = "0.000"
        .Columns(5).NumberFormat = "0.000"
        
        ' Ajuster la largeur des colonnes
        .Columns.AutoFit
    End With
    
    ' Titre du tableau
    dashboardSheet.Range("B30").Value = "Synthèse des Résultats de Performance"
    With dashboardSheet.Range("B30")
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Erreur dans AddSummaryTable: " & Err.Description
    ' Continue l'exécution
End Sub

'------------------------------------------------------------------------------
' Procédure     : AddInteractiveFilters
' Description   : Ajoute des filtres interactifs au tableau de bord
' Paramètres    : 
'   - dashboardSheet: Feuille du tableau de bord
'------------------------------------------------------------------------------
Private Sub AddInteractiveFilters(ByVal dashboardSheet As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Position des contrôles
    Dim controlsRange As Range
    Set controlsRange = dashboardSheet.Range("J31:N38")
    
    ' Titre
    dashboardSheet.Range("J30").Value = "Filtres et Options"
    With dashboardSheet.Range("J30")
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    ' Noms des filtres
    controlsRange.Cells(1, 1).Value = "Type de test:"
    controlsRange.Cells(2, 1).Value = "Volume de données:"
    controlsRange.Cells(3, 1).Value = "Méthode d'accès:"
    controlsRange.Cells(4, 1).Value = "Opération:"
    controlsRange.Cells(5, 1).Value = "Affichage:"
    controlsRange.Cells(6, 1).Value = "Échelle:"
    
    ' Format des étiquettes
    With controlsRange.Columns(1)
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    
    ' Création simulée des contrôles (normalement on utiliserait ActiveX ou des boutons)
    With controlsRange.Cells(1, 2)
        .Value = "Tous"
    End With
    
    With controlsRange.Cells(2, 2)
        .Value = "Tous"
    End With
    
    With controlsRange.Cells(3, 2)
        .Value = "Toutes"
    End With
    
    With controlsRange.Cells(4, 2)
        .Value = "Toutes"
    End With
    
    With controlsRange.Cells(5, 2)
        .Value = "Temps+Mémoire"
    End With
    
    With controlsRange.Cells(6, 2)
        .Value = "Linéaire"
    End With
    
    ' Ajout de boutons pour rafraîchir et exporter
    dashboardSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        controlsRange.Cells(7, 1).Left, controlsRange.Cells(7, 1).Top, 80, 25).TextFrame.Characters.Text = "Rafraîchir"
    
    dashboardSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        controlsRange.Cells(7, 3).Left, controlsRange.Cells(7, 3).Top, 80, 25).TextFrame.Characters.Text = "Exporter PDF"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Erreur dans AddInteractiveFilters: " & Err.Description
    ' Continue l'exécution
End Sub

'------------------------------------------------------------------------------
' Procédure     : FormatDashboard
' Description   : Applique la mise en forme finale au tableau de bord
' Paramètres    : 
'   - dashboardSheet: Feuille du tableau de bord
'------------------------------------------------------------------------------
Private Sub FormatDashboard(ByVal dashboardSheet As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Colorer l'en-tête
    With dashboardSheet.Range("A1:Z3")
        .Interior.Color = RGB(230, 242, 250)
    End With
    
    ' Ajouter un pied de page
    dashboardSheet.Range("A42").Value = "© APEX Framework " & Year(Now) & " - Généré par modPerformanceDashboard"
    With dashboardSheet.Range("A42")
        .Font.Size = 8
        .Font.Italic = True
        .Font.Color = RGB(128, 128, 128)
    End With
    
    ' Optimiser la zone affichée
    dashboardSheet.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Erreur dans FormatDashboard: " & Err.Description
    ' Continue l'exécution
End Sub

'------------------------------------------------------------------------------
' Fonction      : GeneratePerformanceReport
' Description   : Génère un rapport de performance complet (dashboard + export PDF)
' Paramètres    : 
'   - resultFilePath: Chemin du fichier CSV contenant les résultats des tests
'   - reportFilePath: Chemin du fichier de sortie (PDF)
'   - includeRawData: Inclure les données brutes dans le rapport
' Retour        : Booléen indiquant si l'opération a réussi
'------------------------------------------------------------------------------
Public Function GeneratePerformanceReport(ByVal resultFilePath As String, _
                                        ByVal reportFilePath As String, _
                                        Optional ByVal includeRawData As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    ' Créer un classeur temporaire
    Dim tempWb As Workbook
    Set tempWb = Workbooks.Add
    
    ' Générer le tableau de bord
    GeneratePerformanceDashboard resultFilePath, tempWb
    
    ' Exporter en PDF
    tempWb.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=reportFilePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    
    ' Fermer le classeur temporaire
    tempWb.Close SaveChanges:=False
    
    GeneratePerformanceReport = True
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If Not tempWb Is Nothing Then
        tempWb.Close SaveChanges:=False
    End If
    
    MsgBox "Erreur lors de la génération du rapport: " & Err.Description, vbCritical
    GeneratePerformanceReport = False
End Function

'------------------------------------------------------------------------------
' Procédure     : OptimizeForPerformance
' Description   : Configure Excel pour maximiser les performances
'------------------------------------------------------------------------------
Private Sub OptimizeForPerformance()
    ' Désactiver les mises à jour d'écran
    Application.ScreenUpdating = False
    
    ' Désactiver les événements
    Application.EnableEvents = False
    
    ' Désactiver les calculs automatiques
    Application.Calculation = xlCalculationManual
    
    ' Désactiver les alertes
    Application.DisplayAlerts = False
    
    ' Désactiver la barre d'état
    Application.DisplayStatusBar = False
End Sub

'------------------------------------------------------------------------------
' Procédure     : RestoreExcelState
' Description   : Restaure l'état normal d'Excel
'------------------------------------------------------------------------------
Private Sub RestoreExcelState()
    ' Réactiver les mises à jour d'écran
    Application.ScreenUpdating = True
    
    ' Réactiver les événements
    Application.EnableEvents = True
    
    ' Réactiver les calculs automatiques
    Application.Calculation = xlCalculationAutomatic
    
    ' Réactiver les alertes
    Application.DisplayAlerts = True
    
    ' Réactiver la barre d'état
    Application.DisplayStatusBar = True
End Sub