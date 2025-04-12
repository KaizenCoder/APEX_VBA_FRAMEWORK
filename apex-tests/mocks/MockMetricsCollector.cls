VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockMetricsCollector"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("APEX.Tests.Mocks")
Option Explicit

Private Type TMetricEntry
    Name As String
    Value As Double
    Timestamp As Date
    Category As String
    Tags As String
End Type

Private Type TMetricsState
    Entries As Collection
    TotalCount As Long
    LastUpdate As Date
    Categories As Collection
    EnableAggregation As Boolean
End Type

Private this As TMetricsState

'@Description("Initialise le collecteur de métriques")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Initialize()
    Set this.Entries = New Collection
    Set this.Categories = New Collection
    this.TotalCount = 0
    this.LastUpdate = Now
    this.EnableAggregation = True
End Sub

'@Description("Enregistre une métrique")
'@Description: 
'@Param: 
'@Returns: 

Public Sub RecordMetric(ByVal name As String, _
                       ByVal value As Double, _
                       Optional ByVal category As String = "", _
                       Optional ByVal tags As String = "")
    Dim entry As TMetricEntry
    
    With entry
        .Name = name
        .Value = value
        .Timestamp = Now
        .Category = category
        .Tags = tags
    End With
    
    this.Entries.Add entry
    this.TotalCount = this.TotalCount + 1
    this.LastUpdate = Now
    
    'Ajouter nouvelle catégorie si nécessaire
    If Len(category) > 0 Then
        On Error Resume Next
        this.Categories.Add category, category
        On Error GoTo 0
    End If
End Sub

'@Description("Obtient la valeur moyenne d'une métrique")
'@Description: 
'@Param: 
'@Returns: 

Public Function GetAverageMetric(ByVal name As String, _
                               Optional ByVal category As String = "") As Double
    If Not this.EnableAggregation Then
        Err.Raise 5, "MockMetricsCollector", "L'agrégation est désactivée"
    End If
    
    Dim sum As Double
    Dim count As Long
    Dim entry As TMetricEntry
    Dim i As Long
    
    For i = 1 To this.Entries.Count
        entry = this.Entries(i)
        If entry.Name = name Then
            If Len(category) = 0 Or entry.Category = category Then
                sum = sum + entry.Value
                count = count + 1
            End If
        End If
    Next i
    
    If count > 0 Then
        GetAverageMetric = sum / count
    End If
End Function

'@Description("Obtient le total d'une métrique")
'@Description: 
'@Param: 
'@Returns: 

Public Function GetTotalMetric(ByVal name As String, _
                             Optional ByVal category As String = "") As Double
    If Not this.EnableAggregation Then
        Err.Raise 5, "MockMetricsCollector", "L'agrégation est désactivée"
    End If
    
    Dim total As Double
    Dim entry As TMetricEntry
    Dim i As Long
    
    For i = 1 To this.Entries.Count
        entry = this.Entries(i)
        If entry.Name = name Then
            If Len(category) = 0 Or entry.Category = category Then
                total = total + entry.Value
            End If
        End If
    Next i
    
    GetTotalMetric = total
End Function

'@Description("Obtient toutes les métriques sous forme de texte")
'@Description: 
'@Param: 
'@Returns: 

Public Function GetMetricsReport() As String
    Dim report As String
    Dim entry As TMetricEntry
    Dim i As Long
    
    report = "Rapport des métriques:" & vbNewLine & _
            "Total des entrées: " & this.TotalCount & vbNewLine & _
            "Dernière mise à jour: " & Format(this.LastUpdate, "yyyy-mm-dd hh:nn:ss") & vbNewLine & _
            "Catégories: " & this.Categories.Count & vbNewLine & vbNewLine
            
    report = report & "Détail des métriques:" & vbNewLine
    
    For i = 1 To this.Entries.Count
        entry = this.Entries(i)
        report = report & "- " & entry.Name & ": " & entry.Value & _
                IIf(Len(entry.Category) > 0, " [" & entry.Category & "]", "") & _
                IIf(Len(entry.Tags) > 0, " Tags: " & entry.Tags, "") & _
                " @ " & Format(entry.Timestamp, "hh:nn:ss") & vbNewLine
    Next i
    
    GetMetricsReport = report
End Function

'@Description("Active/désactive l'agrégation des métriques")
'@Description: 
'@Param: 
'@Returns: 

Public Sub EnableMetricAggregation(ByVal enable As Boolean)
    this.EnableAggregation = enable
End Sub

'@Description("Réinitialise toutes les métriques")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Reset()
    Set this.Entries = New Collection
    Set this.Categories = New Collection
    this.TotalCount = 0
    this.LastUpdate = Now
End Sub 