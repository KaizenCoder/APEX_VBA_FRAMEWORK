VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelTableAccessor"

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

Implements ITableAccessor

' ==========================================================================
' Classe    : clsExcelTableAccessor
' Purpose   : Implémentation de ITableAccessor pour les Tables Excel
' Author    : APEX Framework Team
' Date      : 2024-04-11
' ==========================================================================

Private Type TTableState
    ListObject As ListObject      ' Référence à la Table Excel
    LastError As String          ' Dernière erreur survenue
    IsInitialized As Boolean     ' État d'initialisation
End Type

Private this As TTableState

'@Description("Initialise l'accesseur avec une Table Excel")
'@Description: 
'@Param: 
'@Returns: 

Public Function Initialize(ByVal tableObject As ListObject) As Boolean
    If tableObject Is Nothing Then
        this.LastError = "Table object cannot be Nothing"
        Initialize = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    Set this.ListObject = tableObject
    this.IsInitialized = True
    Initialize = True
End Function

'@Description("Vérifie l'initialisation et retourne une erreur si non initialisé")
'@Description: 
'@Param: 
'@Returns: 

Private Sub CheckInitialized()
    If Not this.IsInitialized Then
        Err.Raise 5, "clsExcelTableAccessor", "Table accessor not initialized"
    End If
End Sub

' Implémentation de ITableAccessor

Private Property Get ITableAccessor_TableName() As String
    CheckInitialized
    ITableAccessor_TableName = this.ListObject.name
End Property

Private Property Get ITableAccessor_DataRange() As Range
    CheckInitialized
    Set ITableAccessor_DataRange = this.ListObject.DataBodyRange
End Property

Private Property Get ITableAccessor_HeaderRange() As Range
    CheckInitialized
    Set ITableAccessor_HeaderRange = this.ListObject.HeaderRowRange
End Property

Private Property Get ITableAccessor_DataRowCount() As Long
    CheckInitialized
    If this.ListObject.DataBodyRange Is Nothing Then
        ITableAccessor_DataRowCount = 0
    Else
        ITableAccessor_DataRowCount = this.ListObject.DataBodyRange.Rows.Count
    End If
End Property

Private Property Get ITableAccessor_ColumnCount() As Long
    CheckInitialized
    ITableAccessor_ColumnCount = this.ListObject.ListColumns.Count
End Property

'@Description: 
'@Param: 
'@Returns: 

Private Function ITableAccessor_ReadData() As Variant
    CheckInitialized
    
    If this.ListObject.DataBodyRange Is Nothing Then
        ' Retourner un tableau vide si pas de données
        Dim emptyArray(1 To 1, 1 To 1) As Variant
        ITableAccessor_ReadData = emptyArray
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Lire les données
    ITableAccessor_ReadData = this.ListObject.DataBodyRange.value
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Sub ITableAccessor_WriteData(ByVal data As Variant, Optional ByVal appendData As Boolean = False)
    CheckInitialized
    
    ' Vérifier que data est un tableau 2D
    If Not IsArray(data) Then
        Err.Raise 5, "clsExcelTableAccessor", "Data must be a 2D array"
    End If
    
    ' Si append est False, effacer d'abord les données existantes
    If Not appendData Then
        If Not this.ListObject.DataBodyRange Is Nothing Then
            this.ListObject.DataBodyRange.Delete
        End If
    End If
    
    ' Écrire les nouvelles données
    If appendData And Not this.ListObject.DataBodyRange Is Nothing Then
        ' Ajouter à la fin
        Dim lastRow As Long
        lastRow = this.ListObject.DataBodyRange.Rows.Count
        this.ListObject.ListRows.Add.Range.Resize(UBound(data, 1), UBound(data, 2)).value = data
    Else
        ' Écrire directement
        this.ListObject.Range.Resize(UBound(data, 1), UBound(data, 2)).Offset(1).value = data
    End If
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ITableAccessor_ApplyFilter(ByVal columnName As String, ByVal criteria As Variant)
    CheckInitialized
    
    ' Trouver la colonne
    Dim col As ListColumn
    Set col = this.ListObject.ListColumns(columnName)
    
    If col Is Nothing Then
        Err.Raise 5, "clsExcelTableAccessor", "Column not found: " & columnName
    End If
    
    ' Appliquer le filtre
    this.ListObject.Range.AutoFilter Field:=col.Index, Criteria1:=criteria
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ITableAccessor_ClearFilters()
    CheckInitialized
    If this.ListObject.AutoFilter Is Nothing Then Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    this.ListObject.AutoFilter.ShowAllData
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ITableAccessor_Sort(ByVal columnName As String, Optional ByVal ascending As Boolean = True)
    CheckInitialized
    
    ' Trouver la colonne
    Dim col As ListColumn
    Set col = this.ListObject.ListColumns(columnName)
    
    If col Is Nothing Then
        Err.Raise 5, "clsExcelTableAccessor", "Column not found: " & columnName
    End If
    
    ' Appliquer le tri
    With this.ListObject.Sort
        .SortFields.Clear
        .SortFields.Add Key:=col.Range, _
                        SortOn:=xlSortOnValues, _
                        Order:=IIf(ascending, xlAscending, xlDescending)
        .Header = xlYes
        .Apply
    End With
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Function ITableAccessor_AddColumn(ByVal columnName As String, Optional ByVal position As Long = -1) As Long
    CheckInitialized
    
    ' Vérifier si la colonne existe déjà
    Dim existingCol As ListColumn
    On Error Resume Next
    Set existingCol = this.ListObject.ListColumns(columnName)
    On Error GoTo 0
    
    If Not existingCol Is Nothing Then
        Err.Raise 5, "clsExcelTableAccessor", "Column already exists: " & columnName
    End If
    
    ' Ajouter la colonne
    Dim newCol As ListColumn
    If position <= 0 Or position > this.ListObject.ListColumns.Count Then
        ' Ajouter à la fin
        Set newCol = this.ListObject.ListColumns.Add
    Else
        ' Ajouter à la position spécifiée
        Set newCol = this.ListObject.ListColumns.Add(position)
    End If
    
    ' Configurer la nouvelle colonne
    newCol.name = columnName
    
    ITableAccessor_AddColumn = newCol.Index
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Sub ITableAccessor_RemoveColumn(ByVal columnName As String)
    CheckInitialized
    
    ' Trouver et supprimer la colonne
    Dim col As ListColumn
    Set col = this.ListObject.ListColumns(columnName)
    
    If col Is Nothing Then
        Err.Raise 5, "clsExcelTableAccessor", "Column not found: " & columnName
    End If
    
    col.Delete
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Function ITableAccessor_GetColumnIndex(ByVal columnName As String) As Long
    CheckInitialized
    
    ' Trouver la colonne
    Dim col As ListColumn
    On Error Resume Next
    Set col = this.ListObject.ListColumns(columnName)
    On Error GoTo 0
    
    If col Is Nothing Then
        ITableAccessor_GetColumnIndex = 0
    Else
        ITableAccessor_GetColumnIndex = col.Index
    End If
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function ITableAccessor_GetColumnName(ByVal columnIndex As Long) As String
    CheckInitialized
    
    ' Vérifier l'index
    If columnIndex <= 0 Or columnIndex > this.ListObject.ListColumns.Count Then
        ITableAccessor_GetColumnName = ""
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ITableAccessor_GetColumnName = this.ListObject.ListColumns(columnIndex).name
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Sub ITableAccessor_ApplyTableStyle(ByVal styleName As String)
    CheckInitialized
    this.ListObject.TableStyle = styleName
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ITableAccessor_ConvertToRange()
    CheckInitialized
    this.ListObject.Unlist
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub ITableAccessor_Refresh()
    CheckInitialized
    this.ListObject.Refresh
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub Class_Terminate()
    Set this.ListObject = Nothing
End Sub 