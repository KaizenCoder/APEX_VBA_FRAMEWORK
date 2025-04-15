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
'@Folder("APEX.Core.Implementations.Excel")
Option Explicit

Implements IChartAccessor

' ==========================================================================
' Classe    : clsExcelChartAccessor
' Purpose   : Implémentation de IChartAccessor pour les graphiques Excel
' Author    : APEX Framework Team
' Date      : 2024-04-11
' ==========================================================================

Private Type TChartState
    Chart As Chart           ' Référence au graphique
    LastError As String      ' Dernière erreur survenue
    IsInitialized As Boolean ' État d'initialisation
End Type

Private this As TChartState

'@Description("Initialise l'accesseur avec un graphique Excel")
'@Description: 
'@Param: 
'@Returns: 

Public Function Initialize(ByVal chart As Chart) As Boolean
    If chart Is Nothing Then
        this.LastError = "Chart object cannot be Nothing"
        Initialize = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    Set this.Chart = chart
    this.IsInitialized = True
    Initialize = True
End Function

'@Description("Vérifie l'initialisation et retourne une erreur si non initialisé")
'@Description: 
'@Param: 
'@Returns: 

Private Sub CheckInitialized()
    If Not this.IsInitialized Then
        Err.Raise 5, "clsExcelChartAccessor", "Chart accessor not initialized"
    End If
End Sub

' Implémentation de IChartAccessor

Private Property Get IChartAccessor_ChartName() As String
    CheckInitialized
    IChartAccessor_ChartName = this.Chart.name
End Property

Private Property Let IChartAccessor_ChartName(ByVal value As String)
    CheckInitialized
    this.Chart.name = value
End Property

Private Property Get IChartAccessor_ChartTitle() As String
    CheckInitialized
    IChartAccessor_ChartTitle = this.Chart.ChartTitle.Text
End Property

Private Property Let IChartAccessor_ChartTitle(ByVal value As String)
    CheckInitialized
    this.Chart.ChartTitle.Text = value
End Property

Private Property Get IChartAccessor_SourceRange() As Range
    CheckInitialized
    Set IChartAccessor_SourceRange = this.Chart.ChartData.Workbook.Worksheets(1).Range(this.Chart.ChartData.Workbook.Names(1).RefersTo)
End Property

Private Property Set IChartAccessor_SourceRange(ByVal rng As Range)
    CheckInitialized
    this.Chart.SetSourceData Source:=rng
End Property

Private Property Get IChartAccessor_ChartType() As ChartType
    CheckInitialized
    IChartAccessor_ChartType = this.Chart.ChartType
End Property

Private Property Let IChartAccessor_ChartType(ByVal value As ChartType)
    CheckInitialized
    this.Chart.ChartType = value
End Property

'@Description: 
'@Param: 
'@Returns: 

Private Sub IChartAccessor_AddSeries(ByVal name As String, _
                                   ByVal xValues As Range, _
                                   ByVal yValues As Range, _
                                   Optional ByVal plotOrder As Long = -1)
    CheckInitialized
    
    ' Ajouter la série
    Dim series As series
    Set series = this.Chart.SeriesCollection.NewSeries
    
    With series
        .name = name
        .XValues = xValues
        .Values = yValues
        If plotOrder > 0 Then .PlotOrder = plotOrder
    End With
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IChartAccessor_RemoveSeries(ByVal seriesName As String)
    CheckInitialized
    
    ' Rechercher et supprimer la série
    Dim series As series
    For Each series In this.Chart.SeriesCollection
        If series.name = seriesName Then
            series.Delete
            Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
        End If
    Next series
    
    Err.Raise 5, "clsExcelChartAccessor", "Series not found: " & seriesName
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IChartAccessor_ConfigureXAxis(Optional ByVal title As String = "", _
                                        Optional ByVal minValue As Variant = Empty, _
                                        Optional ByVal maxValue As Variant = Empty, _
                                        Optional ByVal majorUnit As Variant = Empty)
    CheckInitialized
    
    With this.Chart.Axes(xlCategory)
        If Len(title) > 0 Then .AxisTitle.Text = title
        If Not IsEmpty(minValue) Then .MinimumScale = minValue
        If Not IsEmpty(maxValue) Then .MaximumScale = maxValue
        If Not IsEmpty(majorUnit) Then .MajorUnit = majorUnit
    End With
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IChartAccessor_ConfigureYAxis(Optional ByVal title As String = "", _
                                        Optional ByVal minValue As Variant = Empty, _
                                        Optional ByVal maxValue As Variant = Empty, _
                                        Optional ByVal majorUnit As Variant = Empty)
    CheckInitialized
    
    With this.Chart.Axes(xlValue)
        If Len(title) > 0 Then .AxisTitle.Text = title
        If Not IsEmpty(minValue) Then .MinimumScale = minValue
        If Not IsEmpty(maxValue) Then .MaximumScale = maxValue
        If Not IsEmpty(majorUnit) Then .MajorUnit = majorUnit
    End With
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IChartAccessor_ConfigureLegend(Optional ByVal position As XlLegendPosition = xlLegendPositionRight, _
                                         Optional ByVal showLegend As Boolean = True)
    CheckInitialized
    
    With this.Chart.Legend
        If showLegend Then
            .position = position
            .Visible = True
        Else
            .Visible = False
        End If
    End With
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IChartAccessor_ApplyChartStyle(ByVal styleName As String)
    CheckInitialized
    this.Chart.ChartStyle = styleName
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IChartAccessor_ConfigureDataLabels(Optional ByVal showValue As Boolean = True, _
                                             Optional ByVal showSeriesName As Boolean = False, _
                                             Optional ByVal showCategoryName As Boolean = False)
    CheckInitialized
    
    With this.Chart
        .HasDataLabels = True
        With .DataLabels
            .ShowValue = showValue
            .ShowSeriesName = showSeriesName
            .ShowCategoryName = showCategoryName
        End With
    End With
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Function IChartAccessor_ExportToNewSheet() As Worksheet
    CheckInitialized
    
    ' Créer une nouvelle feuille
    Dim ws As Worksheet
    Set ws = this.Chart.Parent.Parent.Worksheets.Add
    
    ' Copier le graphique
    this.Chart.ChartArea.Copy
    ws.Paste ws.Range("A1")
    
    Set IChartAccessor_ExportToNewSheet = ws
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Sub IChartAccessor_ExportToImage(ByVal filePath As String, Optional ByVal imageType As XlChartPictureType = xlBitmap)
    CheckInitialized
    this.Chart.Export Filename:=filePath, FilterName:=imageType
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IChartAccessor_Refresh()
    CheckInitialized
    this.Chart.Refresh
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IChartAccessor_Resize(ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single)
    CheckInitialized
    
    With this.Chart.Parent
        .left = left
        .top = top
        .width = width
        .height = height
    End With
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IChartAccessor_Move(ByVal left As Single, ByVal top As Single)
    CheckInitialized
    
    With this.Chart.Parent
        .left = left
        .top = top
    End With
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IChartAccessor_CopyToClipboard()
    CheckInitialized
    this.Chart.ChartArea.Copy
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub Class_Terminate()
    Set this.Chart = Nothing
End Sub 