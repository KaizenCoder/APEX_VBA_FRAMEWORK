VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModExcelFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("APEX.Core.Factories")
Option Explicit

' ==========================================================================
' Module    : ModExcelFactory
' Purpose   : Factory pour créer les accesseurs Excel
' Author    : APEX Framework Team
' Date      : 2024-04-14
' ==========================================================================

'@Description("Crée un accesseur de classeur")
'@Param workbook Le classeur Excel à encapsuler
'@Returns Un accesseur IWorkbookAccessor
Public Function CreateWorkbookAccessor(ByVal workbook As Workbook) As IWorkbookAccessor
    Dim accessor As New clsExcelWorkbookAccessor
    accessor.Initialize workbook
    Set CreateWorkbookAccessor = accessor
End Function

'@Description("Crée un accesseur de feuille")
'@Param worksheet La feuille Excel à encapsuler
'@Returns Un accesseur ISheetAccessor
Public Function CreateSheetAccessor(ByVal worksheet As Worksheet) As ISheetAccessor
    Dim accessor As New clsExcelSheetAccessor
    accessor.Initialize worksheet
    Set CreateSheetAccessor = accessor
End Function

'@Description("Crée un accesseur de plage")
'@Param range La plage Excel à encapsuler
'@Returns Un accesseur IRangeAccessor
Public Function CreateRangeAccessor(ByVal range As Range) As IRangeAccessor
    Dim accessor As New clsExcelRangeAccessor
    accessor.Initialize range
    Set CreateRangeAccessor = accessor
End Function

'@Description("Crée un accesseur de cellule")
'@Param cell La cellule Excel à encapsuler
'@Returns Un accesseur ICellAccessor
Public Function CreateCellAccessor(ByVal cell As Range) As ICellAccessor
    Dim accessor As New clsExcelCellAccessor
    accessor.Initialize cell
    Set CreateCellAccessor = accessor
End Function

'@Description("Crée un accesseur de tableau")
'@Param listObject Le tableau Excel à encapsuler
'@Returns Un accesseur ITableAccessor
Public Function CreateTableAccessor(ByVal listObject As ListObject) As ITableAccessor
    Dim accessor As New clsExcelTableAccessor
    accessor.Initialize listObject
    Set CreateTableAccessor = accessor
End Function

'@Description("Crée un accesseur de tableau croisé dynamique")
'@Param pivotTable Le tableau croisé dynamique à encapsuler
'@Returns Un accesseur IPivotTableAccessor
Public Function CreatePivotTableAccessor(ByVal pivotTable As PivotTable) As IPivotTableAccessor
    Dim accessor As New clsExcelPivotTableAccessor
    accessor.Initialize pivotTable
    Set CreatePivotTableAccessor = accessor
End Function 