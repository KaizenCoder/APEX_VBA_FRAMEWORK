VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelChartAccessor"

'@Module: clsExcelChartAccessor
'@Description: Impl�mentation de l'interface IChartAccessor pour manipuler les graphiques Excel
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("APEX.UI.Excel")
'@ModuleDescription("Impl�mentation de l'interface IChartAccessor pour manipuler les graphiques Excel")
'@Implements IChartAccessor
Option Explicit

Implements IChartAccessor

' Variables priv�es
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
' Author : APEX Framework Team
' Date   : 12/04/2025
' Purpose: Impl�mentation de l'interface IChartAccessor pour manipuler les
'          graphiques Excel avec op�rations compl�tes de personnalisation
'*******************************************************************************

''
' Initialise un nouvel accesseur de graphique Excel
' @param sheetAccessor L'accesseur de feuille qui contient le graphique
' @param chartName Le nom du graphique � manipuler
' @return Boolean True si l'initialisation a r�ussi
''
Public Function Initialize(ByVal sheetAccessor As Object, ByVal chartName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' V�rifier et stocker l'accesseur de feuille
    Set mobjSheet = sheetAccessor
    
    ' R�cup�rer la feuille native
    Dim nativeSheet As Object
    Set nativeSheet = mobjSheet.GetNativeSheet
    
    ' Rechercher le graphique par son nom
    Dim cht As Object
    For Each cht In nativeSheet.ChartObjects
        If cht.Name = chartName Then
            Set mobjNativeChart = cht.Chart
            mstrChartName = chartName
            Initialize = True
            Exit Function
        End If
    Next cht
    
    ' Graphique non trouv�
    Err.Raise ERR_INVALID_CHART, "clsExcelChartAccessor", "Le graphique '" & chartName & "' n'existe pas dans la feuille"
    
    Exit Function
ErrorHandler:
    Initialize = False
    If Err.Number <> ERR_INVALID_CHART Then
        Err.Raise Err.Number, "clsExcelChartAccessor.Initialize", Err.Description
    End If
End Function

''
' Cr�e un nouveau graphique Excel dans la feuille sp�cifi�e
' @param sheetAccessor L'accesseur de la feuille cible
' @param left Position horizontale (points)
' @param top Position verticale (points)
' @param width Largeur (points)
' @param height Hauteur (points)
' @param chartName Nom � donner au nouveau graphique (optionnel)
' @param chartType Type de graphique (constante XlChartType, optionnel)
' @return clsExcelChartAccessor L'accesseur pour le graphique nouvellement cr��
''
Public Function CreateChart(ByVal sheetAccessor As Object, ByVal left As Single, ByVal top As Single, _
                          ByVal width As Single, ByVal height As Single, _
                          Optional ByVal chartName As String = "", _
                          Optional ByVal chartType As Long = xlColumnClustered) As clsExcelChartAccessor
    On Error GoTo ErrorHandler
    
    ' Obtenir la feuille native
    Dim nativeSheet As Object
    Set nativeSheet = sheetAccessor.GetNativeSheet
    
    ' Cr�er un nouvel objet graphique
    Dim newChartObject As Object
    Set newChartObject = nativeSheet.ChartObjects.Add(left, top, width, height)
    
    ' D�finir le type de graphique
    newChartObject.Chart.ChartType = chartType
    
    ' D�finir le nom du graphique si fourni
    If chartName <> "" Then
        newChartObject.Name = chartName
    End If
    
    ' Cr�er et initialiser un nouvel accesseur pour ce graphique
    Dim chartAccessor As New clsExcelChartAccessor
    chartAccessor.Initialize sheetAccessor, newChartObject.Name
    
    Set CreateChart = chartAccessor
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsExcelChartAccessor.CreateChart", Err.Description
End Function

' ==================== Impl�mentation de IChartAccessor ====================

' ----- Propri�t�s -----

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

' ----- M�thodes de cr�ation et manipulation -----

Private Function IChartAccessor_SetSourceData(ByVal sourceRange As String, Optional ByVal seriesInRows As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir la feuille native
    Dim nativeSheet As Object
    Set nativeSheet = mobjSheet.GetNativeSheet
    
    ' Obtenir la plage source
    Dim srcRange As Object
    Set srcRange = nativeSheet.Range(sourceRange)
    
    ' D�finir la source de donn�es du graphique
    mobjNativeChart.SetSourceData Source:=srcRange, PlotBy:=IIf(seriesInRows, xlRows, xlColumns)
    
    IChartAccessor_SetSourceData = True
    Exit Function
    
ErrorHandler:
    IChartAccessor_SetSourceData = False
    Err.Raise Err.Number, "IChartAccessor_SetSourceData", Err.Description
End Function

Private Function IChartAccessor_AddSeries(ByVal name As String, ByVal values As String, Optional ByVal xValues As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir la feuille native
    Dim nativeSheet As Object
    Set nativeSheet = mobjSheet.GetNativeSheet
    
    ' Ajouter une nouvelle s�rie
    mobjNativeChart.SeriesCollection.NewSeries
    
    ' Obtenir l'index de la derni�re s�rie ajout�e
    Dim seriesIndex As Long
    seriesIndex = mobjNativeChart.SeriesCollection.Count
    
    ' D�finir les propri�t�s de la s�rie
    With mobjNativeChart.SeriesCollection(seriesIndex)
        ' D�finir le nom de la s�rie (peut �tre une r�f�rence de cellule)
        If Left(name, 1) = "=" Then
            .Name = name
        Else
            .Name = name
        End If
        
        ' D�finir les valeurs Y
        If Left(values, 1) = "=" Then
            .Values = values
        Else
            .Values = nativeSheet.Range(values).Value
        End If
        
        ' D�finir les valeurs X si fournies
        If xValues <> "" Then
            If Left(xValues, 1) = "=" Then
                .XValues = xValues
            Else
                .XValues = nativeSheet.Range(xValues).Value
            End If
        End If
    End With
    
    IChartAccessor_AddSeries = True
    Exit Function
    
ErrorHandler:
    IChartAccessor_AddSeries = False
    Err.Raise Err.Number, "IChartAccessor_AddSeries", Err.Description
End Function

Private Function IChartAccessor_RemoveSeries(ByVal seriesIndex As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' V�rifier que l'index est valide
    If seriesIndex < 1 Or seriesIndex > mobjNativeChart.SeriesCollection.Count Then
        Err.Raise ERR_INVALID_SERIES, "clsExcelChartAccessor.RemoveSeries", _
                "Index de s�rie invalide : " & seriesIndex
    End If
    
    ' Supprimer la s�rie
    mobjNativeChart.SeriesCollection(seriesIndex).Delete
    
    IChartAccessor_RemoveSeries = True
    Exit Function
    
ErrorHandler:
    IChartAccessor_RemoveSeries = False
    Err.Raise Err.Number, "IChartAccessor_RemoveSeries", Err.Description
End Function

Private Function IChartAccessor_ClearSeries() As Boolean
    On Error GoTo ErrorHandler
    
    ' Supprimer toutes les s�ries une par une (de la derni�re � la premi�re)
    Do While mobjNativeChart.SeriesCollection.Count > 0
        mobjNativeChart.SeriesCollection(mobjNativeChart.SeriesCollection.Count).Delete
    Loop
    
    IChartAccessor_ClearSeries = True
    Exit Function
    
ErrorHandler:
    IChartAccessor_ClearSeries = False
    Err.Raise Err.Number, "IChartAccessor_ClearSeries", Err.Description
End Function

Private Function IChartAccessor_SetXAxisTitle(ByVal title As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Mettre � jour le titre de l'axe X
    With mobjNativeChart.Axes(xlCategory, xlPrimary)
        .HasTitle = (title <> "")
        If title <> "" Then
            .AxisTitle.Text = title
        End If
    End With
    
    IChartAccessor_SetXAxisTitle = True
    Exit Function
    
ErrorHandler:
    IChartAccessor_SetXAxisTitle = False
    Err.Raise Err.Number, "IChartAccessor_SetXAxisTitle", Err.Description
End Function

Private Function IChartAccessor_SetYAxisTitle(ByVal title As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Mettre � jour le titre de l'axe Y
    With mobjNativeChart.Axes(xlValue, xlPrimary)
        .HasTitle = (title <> "")
        If title <> "" Then
            .AxisTitle.Text = title
        End If
    End With
    
    IChartAccessor_SetYAxisTitle = True
    Exit Function
    
ErrorHandler:
    IChartAccessor_SetYAxisTitle = False
    Err.Raise Err.Number, "IChartAccessor_SetYAxisTitle", Err.Description
End Function

Private Function IChartAccessor_FormatXAxis(Optional ByVal min As Variant, Optional ByVal max As Variant, _
                                         Optional ByVal majorUnit As Variant, Optional ByVal minorUnit As Variant, _
                                         Optional ByVal numberFormat As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Formater l'axe X
    With mobjNativeChart.Axes(xlCategory, xlPrimary)
        ' D�finir les limites si sp�cifi�es
        If Not IsMissing(min) Then
            .MinimumScale = min
        End If
        
        If Not IsMissing(max) Then
            .MaximumScale = max
        End If
        
        ' D�finir les intervalles si sp�cifi�s
        If Not IsMissing(majorUnit) Then
            .MajorUnit = majorUnit
        End If
        
        If Not IsMissing(minorUnit) Then
            .MinorUnit = minorUnit
        End If
        
        ' D�finir le format num�rique si sp�cifi�
        If numberFormat <> "" Then
            .TickLabels.NumberFormat = numberFormat
        End If
    End With
    
    IChartAccessor_FormatXAxis = True
    Exit Function
    
ErrorHandler:
    IChartAccessor_FormatXAxis = False
    Err.Raise Err.Number, "IChartAccessor_FormatXAxis", Err.Description
End Function

Private Function IChartAccessor_FormatYAxis(Optional ByVal min As Variant, Optional ByVal max As Variant, _
                                         Optional ByVal majorUnit As Variant, Optional ByVal minorUnit As Variant, _
                                         Optional ByVal numberFormat As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Formater l'axe Y
    With mobjNativeChart.Axes(xlValue, xlPrimary)
        ' D�finir les limites si sp�cifi�es
        If Not IsMissing(min) Then
            .MinimumScale = min
        End If
        
        If Not IsMissing(max) Then
            .MaximumScale = max
        End If
        
        ' D�finir les intervalles si sp�cifi�s
        If Not IsMissing(majorUnit) Then
            .MajorUnit = majorUnit
        End If
        
        If Not IsMissing(minorUnit) Then
            .MinorUnit = minorUnit
        End If
        
        ' D�finir le format num�rique si sp�cifi�
        If numberFormat <> "" Then
            .TickLabels.NumberFormat = numberFormat
        End If
    End With
    
    IChartAccessor_FormatYAxis = True
    Exit Function
    
ErrorHandler:
    IChartAccessor_FormatYAxis = False
    Err.Raise Err.Number, "IChartAccessor_FormatYAxis", Err.Description
End Function

Private Function IChartAccessor_ApplyChartStyle(ByVal styleIndex As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' Appliquer un style pr�d�fini
    mobjNativeChart.ApplyChartTemplate ("Chart Style " & styleIndex)
    
    IChartAccessor_ApplyChartStyle = True
    Exit Function
    
ErrorHandler:
    IChartAccessor_ApplyChartStyle = False
    Err.Raise Err.Number, "IChartAccessor_ApplyChartStyle", Err.Description
End Function

Private Function IChartAccessor_FormatSeries(ByVal seriesIndex As Long, ByVal formatType As Long, ByVal formatValue As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' V�rifier que l'index est valide
    If seriesIndex < 1 Or seriesIndex > mobjNativeChart.SeriesCollection.Count Then
        Err.Raise ERR_INVALID_SERIES, "clsExcelChartAccessor.FormatSeries", _
                "Index de s�rie invalide : " & seriesIndex
    End If
    
    ' Appliquer le format selon le type
    Select Case formatType
        Case FORMAT_COLOR
            ' formatValue doit �tre une couleur (valeur RGB ou nom)
            mobjNativeChart.SeriesCollection(seriesIndex).Format.Fill.ForeColor.RGB = formatValue
            
        Case FORMAT_PATTERN
            ' formatValue doit �tre un index de motif
            mobjNativeChart.SeriesCollection(seriesIndex).Interior.Pattern = formatValue
            
        Case FORMAT_MARKER
            ' formatValue doit �tre un style de marqueur (ex: xlMarkerStyleDiamond)
            With mobjNativeChart.SeriesCollection(seriesIndex)
                .MarkerStyle = formatValue
                .MarkerSize = 6  ' Taille par d�faut
            End With
            
        Case FORMAT_LINE
            ' formatValue doit �tre un style de ligne (ex: xlContinuous)
            With mobjNativeChart.SeriesCollection(seriesIndex).Format.Line
                .DashStyle = formatValue
                .Weight = 2  ' �paisseur par d�faut
            End With
            
        Case Else
            Err.Raise ERR_INVALID_FORMAT, "clsExcelChartAccessor.FormatSeries", _
                    "Type de format non support� : " & formatType
    End Select
    
    IChartAccessor_FormatSeries = True
    Exit Function
    
ErrorHandler:
    IChartAccessor_FormatSeries = False
    Err.Raise Err.Number, "IChartAccessor_FormatSeries", Err.Description
End Function

Private Function IChartAccessor_SetDataLabels(ByVal seriesIndex As Long, ByVal showLabels As Boolean, _
                                           Optional ByVal labelContent As Long = 0) As Boolean
    On Error GoTo ErrorHandler
    
    ' V�rifier que l'index est valide
    If seriesIndex < 1 Or seriesIndex > mobjNativeChart.SeriesCollection.Count Then
        Err.Raise ERR_INVALID_SERIES, "clsExcelChartAccessor.SetDataLabels", _
                "Index de s�rie invalide : " & seriesIndex
    End If
    
    ' Activer/d�sactiver les �tiquettes de donn�es
    With mobjNativeChart.SeriesCollection(seriesIndex)
        .HasDataLabels = showLabels
        
        ' Si les �tiquettes sont activ�es, d�finir leur contenu
        If showLabels And labelContent > 0 Then
            With .DataLabels
                ' Valeur par d�faut = Valeur uniquement
                If labelContent = 0 Then
                    labelContent = xlDataLabelShowValue
                End If
                
                ' Configurer le contenu des �tiquettes
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
    Exit Function
    
ErrorHandler:
    IChartAccessor_SetDataLabels = False
    Err.Raise Err.Number, "IChartAccessor_SetDataLabels", Err.Description
End Function

Private Function IChartAccessor_SetPosition(ByVal left As Single, ByVal top As Single, _
                                         ByVal width As Single, ByVal height As Single) As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir l'objet ChartObject qui contient le graphique
    Dim chartObj As Object
    Set chartObj = mobjNativeChart.Parent
    
    ' D�finir la position et la taille
    chartObj.left = left
    chartObj.top = top
    chartObj.width = width
    chartObj.height = height
    
    IChartAccessor_SetPosition = True
    Exit Function
    
ErrorHandler:
    IChartAccessor_SetPosition = False
    Err.Raise Err.Number, "IChartAccessor_SetPosition", Err.Description
End Function

Private Function IChartAccessor_ExportAsImage(ByVal filePath As String, ByVal formatType As String, _
                                           Optional ByVal resolution As Long = 96) As Boolean
    On Error GoTo ErrorHandler
    
    ' D�terminer le type de fichier � partir de formatType
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
            fileFormat = xlPicture ' Par d�faut
    End Select
    
    ' Exporter le graphique comme image
    mobjNativeChart.Export Filename:=filePath, Filtername:=formatType, Interactive:=False
    
    IChartAccessor_ExportAsImage = True
    Exit Function
    
ErrorHandler:
    IChartAccessor_ExportAsImage = False
    Err.Raise Err.Number, "IChartAccessor_ExportAsImage", Err.Description
End Function

' ==================== M�thodes publiques additionnelles ====================

''
' Retourne l'objet Chart sous-jacent pour un acc�s direct par d'autres composants
' @return Object L'objet Chart natif d'Excel
''
Public Function GetNativeChart() As Object
    Set GetNativeChart = mobjNativeChart
End Function