VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelPivotTableAccessor"

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

Implements IPivotTableAccessor

' ==========================================================================
' Classe    : clsExcelPivotTableAccessor
' Purpose   : Implémentation de IPivotTableAccessor pour les tableaux croisés
' Author    : APEX Framework Team
' Date      : 2024-04-11
' ==========================================================================

Private Type TPivotTableState
    PivotTable As PivotTable     ' Référence au tableau croisé
    LastError As String          ' Dernière erreur survenue
    IsInitialized As Boolean     ' État d'initialisation
End Type

Private this As TPivotTableState

'@Description("Initialise l'accesseur avec un tableau croisé dynamique")
'@Description: 
'@Param: 
'@Returns: 

Public Function Initialize(ByVal pivotTable As PivotTable) As Boolean
    If pivotTable Is Nothing Then
        this.LastError = "PivotTable object cannot be Nothing"
        Initialize = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    Set this.PivotTable = pivotTable
    this.IsInitialized = True
    Initialize = True
End Function

'@Description("Vérifie l'initialisation et retourne une erreur si non initialisé")
'@Description: 
'@Param: 
'@Returns: 

Private Sub CheckInitialized()
    If Not this.IsInitialized Then
        Err.Raise 5, "clsExcelPivotTableAccessor", "PivotTable accessor not initialized"
    End If
End Sub

' Implémentation de IPivotTableAccessor

Private Property Get IPivotTableAccessor_PivotTableName() As String
    CheckInitialized
    IPivotTableAccessor_PivotTableName = this.PivotTable.name
End Property

Private Property Get IPivotTableAccessor_PivotRange() As Range
    CheckInitialized
    Set IPivotTableAccessor_PivotRange = this.PivotTable.TableRange2
End Property

Private Property Get IPivotTableAccessor_SourceRange() As Range
    CheckInitialized
    Set IPivotTableAccessor_SourceRange = this.PivotTable.SourceData
End Property

Private Property Set IPivotTableAccessor_SourceRange(ByVal rng As Range)
    CheckInitialized
    this.PivotTable.ChangePivotCache _
        this.PivotTable.Parent.Parent.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rng)
End Property

'@Description: 
'@Param: 
'@Returns: 

Private Sub IPivotTableAccessor_AddRowField(ByVal fieldName As String, Optional ByVal position As Long = -1)
    CheckInitialized
    
    ' Vérifier si le champ existe
    Dim pvtField As PivotField
    On Error Resume Next
    Set pvtField = this.PivotTable.PivotFields(fieldName)
    On Error GoTo 0
    
    If pvtField Is Nothing Then
        Err.Raise 5, "clsExcelPivotTableAccessor", "Field not found: " & fieldName
    End If
    
    ' Ajouter le champ en ligne
    With pvtField
        .Orientation = xlRowField
        If position > 0 Then .Position = position
    End With
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IPivotTableAccessor_AddColumnField(ByVal fieldName As String, Optional ByVal position As Long = -1)
    CheckInitialized
    
    ' Vérifier si le champ existe
    Dim pvtField As PivotField
    On Error Resume Next
    Set pvtField = this.PivotTable.PivotFields(fieldName)
    On Error GoTo 0
    
    If pvtField Is Nothing Then
        Err.Raise 5, "clsExcelPivotTableAccessor", "Field not found: " & fieldName
    End If
    
    ' Ajouter le champ en colonne
    With pvtField
        .Orientation = xlColumnField
        If position > 0 Then .Position = position
    End With
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IPivotTableAccessor_AddValueField(ByVal fieldName As String, _
                                            ByVal aggregationType As PivotAggregationType, _
                                            Optional ByVal customName As String = "")
    CheckInitialized
    
    ' Vérifier si le champ existe
    Dim pvtField As PivotField
    On Error Resume Next
    Set pvtField = this.PivotTable.PivotFields(fieldName)
    On Error GoTo 0
    
    If pvtField Is Nothing Then
        Err.Raise 5, "clsExcelPivotTableAccessor", "Field not found: " & fieldName
    End If
    
    ' Configurer le champ de valeur
    With pvtField
        .Orientation = xlDataField
        
        ' Définir le type d'agrégation
        Select Case aggregationType
            Case pvtSum
                .Function = xlSum
            Case pvtCount
                .Function = xlCount
            Case pvtAverage
                .Function = xlAverage
            Case pvtMax
                .Function = xlMax
            Case pvtMin
                .Function = xlMin
            Case pvtProduct
                .Function = xlProduct
            Case pvtCountNumbers
                .Function = xlCountNums
            Case pvtStdDev
                .Function = xlStDev
            Case pvtStdDevP
                .Function = xlStDevP
            Case pvtVar
                .Function = xlVar
            Case pvtVarP
                .Function = xlVarP
        End Select
        
        ' Définir un nom personnalisé si fourni
        If Len(customName) > 0 Then
            .Caption = customName
        End If
    End With
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IPivotTableAccessor_AddFilterField(ByVal fieldName As String, Optional ByVal position As Long = -1)
    CheckInitialized
    
    ' Vérifier si le champ existe
    Dim pvtField As PivotField
    On Error Resume Next
    Set pvtField = this.PivotTable.PivotFields(fieldName)
    On Error GoTo 0
    
    If pvtField Is Nothing Then
        Err.Raise 5, "clsExcelPivotTableAccessor", "Field not found: " & fieldName
    End If
    
    ' Ajouter le champ en filtre
    With pvtField
        .Orientation = xlPageField
        If position > 0 Then .Position = position
    End With
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IPivotTableAccessor_RemoveField(ByVal fieldName As String)
    CheckInitialized
    
    ' Vérifier si le champ existe
    Dim pvtField As PivotField
    On Error Resume Next
    Set pvtField = this.PivotTable.PivotFields(fieldName)
    On Error GoTo 0
    
    If pvtField Is Nothing Then
        Err.Raise 5, "clsExcelPivotTableAccessor", "Field not found: " & fieldName
    End If
    
    ' Supprimer le champ
    pvtField.Orientation = xlHidden
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IPivotTableAccessor_ApplyFilter(ByVal fieldName As String, ByVal criteria As Variant)
    CheckInitialized
    
    ' Vérifier si le champ existe
    Dim pvtField As PivotField
    On Error Resume Next
    Set pvtField = this.PivotTable.PivotFields(fieldName)
    On Error GoTo 0
    
    If pvtField Is Nothing Then
        Err.Raise 5, "clsExcelPivotTableAccessor", "Field not found: " & fieldName
    End If
    
    ' Appliquer le filtre
    With pvtField
        .ClearAllFilters
        If IsArray(criteria) Then
            .PivotFilters.Add2 Type:=xlCaptionEquals, Value1:=criteria
        Else
            .CurrentPage = criteria
        End If
    End With
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IPivotTableAccessor_ClearFilters()
    CheckInitialized
    
    ' Effacer tous les filtres
    Dim pvtField As PivotField
    For Each pvtField In this.PivotTable.PivotFields
        pvtField.ClearAllFilters
    Next pvtField
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IPivotTableAccessor_ApplyPivotStyle(ByVal styleName As String)
    CheckInitialized
    this.PivotTable.TableStyle2 = styleName
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IPivotTableAccessor_ConfigureDisplay(Optional ByVal showGrandTotals As Boolean = True, _
                                               Optional ByVal showSubTotals As Boolean = True, _
                                               Optional ByVal repeatLabels As Boolean = False)
    CheckInitialized
    
    With this.PivotTable
        .ColumnGrand = showGrandTotals
        .RowGrand = showGrandTotals
        .SubtotalLocation = IIf(showSubTotals, xlAtTop, xlAtBottom)
        .RepeatAllLabels IIf(repeatLabels, xlRepeatLabels, xlDoNotRepeatLabels)
    End With
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IPivotTableAccessor_Refresh()
    CheckInitialized
    this.PivotTable.RefreshTable
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub IPivotTableAccessor_Clear()
    CheckInitialized
    this.PivotTable.ClearTable
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Function IPivotTableAccessor_ExportToNewSheet() As Worksheet
    CheckInitialized
    
    ' Créer une nouvelle feuille
    Dim ws As Worksheet
    Set ws = this.PivotTable.Parent.Parent.Worksheets.Add
    
    ' Copier le tableau croisé
    this.PivotTable.TableRange2.Copy ws.Range("A1")
    
    Set IPivotTableAccessor_ExportToNewSheet = ws
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IPivotTableAccessor_GetPivotData() As Variant
    CheckInitialized
    IPivotTableAccessor_GetPivotData = this.PivotTable.TableRange2.value
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Sub Class_Terminate()
    Set this.PivotTable = Nothing
End Sub 