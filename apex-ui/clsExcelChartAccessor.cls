VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelChartAccessor"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("APEX.UI.Excel")
'@ModuleDescription("Implémentation de l'interface IChartAccessor pour manipuler les graphiques Excel")
'@Implements IChartAccessor
Option Explicit

Implements IChartAccessor

' Variables privées
Private mobjSheet As Object ' ISheetAccessor
Private mobjNativeChart As Object ' Chart
Private mstrChartName As String

' Constantes d'erreur
Private Const ERR_INVALID_CHART As Long = 3001
Private Const ERR_INVALID_SERIES As Long = 3002
Private Const ERR_INVALID_RANGE As Long = 3003
Private Const ERR_INVALID_FORMAT As Long = 3004
Private Const ERR_EXPORT_FAILED As Long = 3005

' Constantes pour les types de format
Private Const FORMAT_COLOR As Long = 1
Private Const FORMAT_PATTERN As Long = 2
Private Const FORMAT_MARKER As Long = 3
Private Const FORMAT_LINE As Long = 4

'*******************************************************************************
' Module : clsExcelChartAccessor
' Author : [Votre nom]
' Date   : 12/04/2025
' Purpose: Implémentation de l'interface IChartAccessor pour manipuler les
'          graphiques Excel avec opérations complètes de personnalisation
'*******************************************************************************

''
' Initialise un nouvel accesseur de graphique Excel
' @param sheetAccessor L'accesseur de feuille qui contient le graphique
' @param chartName Le nom du graphique à manipuler
' @return Boolean True si l'initialisation a réussi
''
'@Description: 
'@Param: 
'@Returns: 

Public Function Initialize(ByVal sheetAccessor As Object, ByVal chartName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier et stocker l'accesseur de feuille
    Set mobjSheet = sheetAccessor
    
    ' Récupérer la feuille native
    Dim nativeSheet As Object
    Set nativeSheet = mobjSheet.GetNativeSheet
    
    ' Rechercher le graphique par son nom
    Dim cht As Object
    For Each cht In nativeSheet.ChartObjects
        If cht.Name = chartName Then
            Set mobjNativeChart = cht.Chart
            mstrChartName = chartName
            Initialize = True
            Exit'@Description: 
'@Param: 
'@Returns: 

 Function
        End If
    Next cht
    
    ' Graphique non trouvé
    Err.Raise ERR_INVALID_CHART, "clsExcelChartAccessor", "Le graphique '" & chartName & "' n'existe pas dans la feuille"
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
ErrorHandler:
    Initialize = False
    If Err.Number <> ERR_INVALID_CHART Then
        Err.Raise Err.Number, "clsExcelChartAccessor.Initialize", Err.Description
    End If
End Function

''
' Crée un nouveau graphique Excel dans la feuille spécifiée
' @param sheetAccessor L'accesseur de la feuille cible
' @param left Position horizontale (points)
' @param top Position verticale (points)
' @param width Largeur (points)
' @param height Hauteur (points)
' @param chartName Nom à donner au nouveau graphique (optionnel)
' @param chartType Type de graphique (constante XlChartType, optionnel)
' @return clsExcelChartAccessor L'accesseur pour le graphique nouvellement créé
''
'@Description: 
'@Param: 
'@Returns: 

Public Function CreateChart(ByVal sheetAccessor As Object, ByVal left As Single, ByVal top As Single, _
                          ByVal width As Single, ByVal height As Single, _
                          Optional ByVal chartName As String = "", _
                          Optional ByVal chartType As Long = xlColumnClustered) As clsExcelChartAccessor
    On Error GoTo ErrorHandler
    
    ' Obtenir la feuille native
    Dim nativeSheet As Object
    Set nativeSheet = sheetAccessor.GetNativeSheet
    
    ' Créer un nouvel objet graphique
    Dim newChartObject As Object
    Set newChartObject = nativeSheet.ChartObjects.Add(left, top, width, height)
    
    ' Définir le type de graphique
    newChartObject.Chart.ChartType = chartType
    
    ' Définir le nom du graphique si fourni
    If chartName <> "" Then
        newChartObject.Name = chartName
    End If
    
    ' Créer et initialiser un nouvel accesseur pour ce graphique
    Dim chartAccessor As New clsExcelChartAccessor
    chartAccessor.Initialize sheetAccessor, newChartObject.Name
    
    Set CreateChart = chartAccessor
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsExcelChartAccessor.CreateChart", Err.Description
End Function

' ==================== Implémentation de IChartAccessor ====================

' ----- Propriétés -----

Private Property Get IChartAccessor_ChartName() As String
    IChartAccessor_ChartName = mstrChartName
End Property

Private Property Get IChartAccessor_ParentSheet() As Object
    Set IChartAccessor_ParentSheet = mobjSheet
End Property

Private Property Get IChartAccessor_ChartType() As Long
    IChartAccessor_ChartType = mobjNativeChart.ChartType
End Property

Private Property Let IChartAccessor_ChartType(ByVal value As Long)
    mobjNativeChart.ChartType = value
End Property

Private Property Get IChartAccessor_Title() As String
    On Error Resume Next
    If mobjNativeChart.HasTitle Then
        IChartAccessor_Title = mobjNativeChart.ChartTitle.Text
    Else
        IChartAccessor_Title = ""
    End If
End Property

Private Property Let IChartAccessor_Title(ByVal value As String)
    mobjNativeChart.HasTitle = (value <> "")
    If value <> "" Then
        mobjNativeChart.ChartTitle.Text = value
    End If
End Property

Private Property Get IChartAccessor_HasLegend() As Boolean
    IChartAccessor_HasLegend = mobjNativeChart.HasLegend
End Property

Private Property Let IChartAccessor_HasLegend(ByVal value As Boolean)
    mobjNativeChart.HasLegend = value
End Property

Private Property Get IChartAccessor_LegendPosition() As Long
    If mobjNativeChart.HasLegend Then
        IChartAccessor_LegendPosition = mobjNativeChart.Legend.Position
    Else
        IChartAccessor_LegendPosition = -1
    End If
End Property

Private Property Let IChartAccessor_LegendPosition(ByVal value As Long)
    If Not mobjNativeChart.HasLegend Then
        mobjNativeChart.HasLegend = True
    End If
    mobjNativeChart.Legend.Position = value
End Property

' ----- Méthodes de création et manipulation -----

'@Description: 
'@Param: 
'@Returns: 

Private Function IChartAccessor_SetSourceData(ByVal sourceRange As String, Optional ByVal seriesInRows As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir la feuille native
    Dim nativeSheet As Object
    Set nativeSheet = mobjSheet.GetNativeSheet
    
    ' Obtenir la plage source
    Dim srcRange As Object
    Set srcRange = nativeSheet.Range(sourceRange)
    
    ' Définir la source de données du graphique
    mobjNativeChart.SetSourceData Source:=srcRange, PlotBy:=IIf(seriesInRows, xlRows, xlColumns)
    
    IChartAccessor_SetSourceData = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    IChartAccessor_SetSourceData = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IChartAccessor_AddSeries(ByVal name As String, ByVal values As String, Optional ByVal xValues As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir la feuille native
    Dim nativeSheet As Object
    Set nativeSheet = mobjSheet.GetNativeSheet
    
    ' Ajouter une nouvelle série
    mobjNativeChart.SeriesCollection.NewSeries
    
    ' Obtenir l'index de la dernière série ajoutée
    Dim seriesIndex As Long
    seriesIndex = mobjNativeChart.SeriesCollection.Count
    
    ' Définir les propriétés de la série
    With mobjNativeChart.SeriesCollection(seriesIndex)
        ' Définir le nom de la série (peut être une référence de cellule)
        If Left(name, 1) = "=" Then
            .Name = name
        Else
            .Name = name
        End If
        
        ' Définir les valeurs Y
        If Left(values, 1) = "=" Then
            .Values = values
        Else
            .Values = nativeSheet.Range(values).Value
        End If
        
        ' Définir les valeurs X si fournies
        If xValues <> "" Then
            If Left(xValues, 1) = "=" Then
                .XValues = xValues
            Else
                .XValues = nativeSheet.Range(xValues).Value
            End If
        End If
    End With
    
    IChartAccessor_AddSeries = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    IChartAccessor_AddSeries = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IChartAccessor_RemoveSeries(ByVal seriesIndex As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que l'index est valide
    If seriesIndex < 1 Or seriesIndex > mobjNativeChart.SeriesCollection.Count Then
        Err.Raise ERR_INVALID_SERIES, "clsExcelChartAccessor.RemoveSeries", _
                "Index de série invalide : " & seriesIndex
    End If
    
    ' Supprimer la série
    mobjNativeChart.SeriesCollection(seriesIndex).Delete
    
    IChartAccessor_RemoveSeries = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    IChartAccessor_RemoveSeries = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IChartAccessor_ClearSeries() As Boolean
    On Error GoTo ErrorHandler
    
    ' Supprimer toutes les séries une par une (de la dernière à la première)
    Do While mobjNativeChart.SeriesCollection.Count > 0
        mobjNativeChart.SeriesCollection(mobjNativeChart.SeriesCollection.Count).Delete
    Loop
    
    IChartAccessor_ClearSeries = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    IChartAccessor_ClearSeries = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IChartAccessor_SetXAxisTitle(ByVal title As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Mettre à jour le titre de l'axe X
    With mobjNativeChart.Axes(xlCategory, xlPrimary)
        .HasTitle = (title <> "")
        If title <> "" Then
            .AxisTitle.Text = title
        End If
    End With
    
    IChartAccessor_SetXAxisTitle = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    IChartAccessor_SetXAxisTitle = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IChartAccessor_SetYAxisTitle(ByVal title As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Mettre à jour le titre de l'axe Y
    With mobjNativeChart.Axes(xlValue, xlPrimary)
        .HasTitle = (title <> "")
        If title <> "" Then
            .AxisTitle.Text = title
        End If
    End With
    
    IChartAccessor_SetYAxisTitle = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    IChartAccessor_SetYAxisTitle = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IChartAccessor_FormatXAxis(Optional ByVal min As Variant, Optional ByVal max As Variant, _
                                         Optional ByVal majorUnit As Variant, Optional ByVal minorUnit As Variant, _
                                         Optional ByVal numberFormat As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Formater l'axe X
    With mobjNativeChart.Axes(xlCategory, xlPrimary)
        ' Définir les limites si spécifiées
        If Not IsMissing(min) Then
            .MinimumScale = min
        End If
        
        If Not IsMissing(max) Then
            .MaximumScale = max
        End If
        
        ' Définir les intervalles si spécifiés
        If Not IsMissing(majorUnit) Then
            .MajorUnit = majorUnit
        End If
        
        If Not IsMissing(minorUnit) Then
            .MinorUnit = minorUnit
        End If
        
        ' Définir le format numérique si spécifié
        If numberFormat <> "" Then
            .TickLabels.NumberFormat = numberFormat
        End If
    End With
    
    IChartAccessor_FormatXAxis = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    IChartAccessor_FormatXAxis = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IChartAccessor_FormatYAxis(Optional ByVal min As Variant, Optional ByVal max As Variant, _
                                         Optional ByVal majorUnit As Variant, Optional ByVal minorUnit As Variant, _
                                         Optional ByVal numberFormat As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Formater l'axe Y
    With mobjNativeChart.Axes(xlValue, xlPrimary)
        ' Définir les limites si spécifiées
        If Not IsMissing(min) Then
            .MinimumScale = min
        End If
        
        If Not IsMissing(max) Then
            .MaximumScale = max
        End If
        
        ' Définir les intervalles si spécifiés
        If Not IsMissing(majorUnit) Then
            .MajorUnit = majorUnit
        End If
        
        If Not IsMissing(minorUnit) Then
            .MinorUnit = minorUnit
        End If
        
        ' Définir le format numérique si spécifié
        If numberFormat <> "" Then
            .TickLabels.NumberFormat = numberFormat
        End If
    End With
    
    IChartAccessor_FormatYAxis = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    IChartAccessor_FormatYAxis = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IChartAccessor_ApplyChartStyle(ByVal styleIndex As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' Appliquer un style prédéfini
    mobjNativeChart.ApplyChartTemplate ("Chart Style " & styleIndex)
    
    IChartAccessor_ApplyChartStyle = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    IChartAccessor_ApplyChartStyle = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IChartAccessor_FormatSeries(ByVal seriesIndex As Long, ByVal formatType As Long, ByVal formatValue As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que l'index est valide
    If seriesIndex < 1 Or seriesIndex > mobjNativeChart.SeriesCollection.Count Then
        Err.Raise ERR_INVALID_SERIES, "clsExcelChartAccessor.FormatSeries", _
                "Index de série invalide : " & seriesIndex
    End If
    
    ' Appliquer le format selon le type
    Select Case formatType
        Case FORMAT_COLOR
            ' formatValue doit être une couleur (valeur RGB ou nom)
            mobjNativeChart.SeriesCollection(seriesIndex).Format.Fill.ForeColor.RGB = formatValue
            
        Case FORMAT_PATTERN
            ' formatValue doit être un index de motif
            mobjNativeChart.SeriesCollection(seriesIndex).Interior.Pattern = formatValue
            
        Case FORMAT_MARKER
            ' formatValue doit être un style de marqueur (ex: xlMarkerStyleDiamond)
            With mobjNativeChart.SeriesCollection(seriesIndex)
                .MarkerStyle = formatValue
                .MarkerSize = 6  ' Taille par défaut
            End With
            
        Case FORMAT_LINE
            ' formatValue doit être un style de ligne (ex: xlContinuous)
            With mobjNativeChart.SeriesCollection(seriesIndex).Format.Line
                .DashStyle = formatValue
                .Weight = 2  ' Épaisseur par défaut
            End With
            
        Case Else
            Err.Raise ERR_INVALID_FORMAT, "clsExcelChartAccessor.FormatSeries", _
                    "Type de format non supporté : " & formatType
    End Select
    
    IChartAccessor_FormatSeries = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    IChartAccessor_FormatSeries = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IChartAccessor_SetDataLabels(ByVal seriesIndex As Long, ByVal showLabels As Boolean, _
                                           Optional ByVal labelContent As Long = 0) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que l'index est valide
    If seriesIndex < 1 Or seriesIndex > mobjNativeChart.SeriesCollection.Count Then
        Err.Raise ERR_INVALID_SERIES, "clsExcelChartAccessor.SetDataLabels", _
                "Index de série invalide : " & seriesIndex
    End If
    
    ' Activer/désactiver les étiquettes de données
    With mobjNativeChart.SeriesCollection(seriesIndex)
        .HasDataLabels = showLabels
        
        ' Si les étiquettes sont activées, définir leur contenu
        If showLabels And labelContent > 0 Then
            With .DataLabels
                ' Valeur par défaut = Valeur uniquement
                If labelContent = 0 Then
                    labelContent = xlDataLabelShowValue
                End If
                
                ' Configurer le contenu des étiquettes
                .ShowSeriesName = (labelContent And xlDataLabelShowSeriesName) > 0
                .ShowCategoryName = (labelContent And xlDataLabelShowCategoryName) > 0
                .ShowValue = (labelContent And xlDataLabelShowValue) > 0
                .ShowPercentage = (labelContent And xlDataLabelShowPercentage) > 0
                .ShowBubbleSize = (labelContent And xlDataLabelShowBubbleSizes) > 0
                .ShowLegendKey = False
            End With
        End If
    End With
    
    IChartAccessor_SetDataLabels = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    IChartAccessor_SetDataLabels = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IChartAccessor_SetPosition(ByVal left As Single, ByVal top As Single, _
                                         ByVal width As Single, ByVal height As Single) As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir l'objet ChartObject qui contient le graphique
    Dim chartObj As Object
    Set chartObj = mobjNativeChart.Parent
    
    ' Définir la position et la taille
    chartObj.left = left
    chartObj.top = top
    chartObj.width = width
    chartObj.height = height
    
    IChartAccessor_SetPosition = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    IChartAccessor_SetPosition = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IChartAccessor_ExportAsImage(ByVal filePath As String, ByVal formatType As String, _
                                           Optional ByVal resolution As Long = 96) As Boolean
    On Error GoTo ErrorHandler
    
    ' Déterminer le type de fichier à partir de formatType
    Dim fileFormat As Long
    
    Select Case LCase(formatType)
        Case "png"
            fileFormat = xlPicture
        Case "jpg", "jpeg"
            fileFormat = xlPicture
        Case "gif"
            fileFormat = xlPicture
        Case "bmp"
            fileFormat = xlPicture
        Case Else
            fileFormat = xlPicture ' Par défaut
    End Select
    
    ' Exporter le graphique comme image
    mobjNativeChart.Export Filename:=filePath, Filtername:=formatType, Interactive:=False
    
    IChartAccessor_ExportAsImage = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    IChartAccessor_ExportAsImage = False
End Function

' ==================== Méthodes publiques additionnelles ====================

''
' Retourne l'objet Chart sous-jacent pour un accès direct par d'autres composants
' @return Object L'objet Chart natif d'Excel
''
'@Description: 
'@Param: 
'@Returns: 

Public Function GetNativeChart() As Object
    Set GetNativeChart = mobjNativeChart
End Function