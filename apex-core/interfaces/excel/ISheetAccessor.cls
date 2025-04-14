VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISheetAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
'@Folder("APEX.Core.Interfaces.Excel")
Option Explicit

' ==========================================================================
' Interface  : ISheetAccessor
' Purpose    : Interface pour l'accès aux objets Worksheet d'Excel
' Author     : APEX Framework Team
' Date       : 2024-04-11
' ==========================================================================

'@Description("Obtient une cellule par ses coordonnées")
'@Param row Numéro de ligne (1-based)
'@Param col Numéro de colonne (1-based)
'@Returns Une interface ICellAccessor pour manipuler la cellule
Public Function GetCell(ByVal row As Long, ByVal col As Long) As ICellAccessor
End Function

'@Description("Obtient une plage de cellules par son adresse")
'@Param address Adresse de la plage (ex: "A1:C10")
'@Returns Une interface IRangeAccessor pour manipuler la plage
Public Function GetRange(ByVal address As String) As IRangeAccessor
End Function

'@Description("Obtient une plage de cellules par ses coordonnées")
'@Param startRow Ligne de début (1-based)
'@Param startCol Colonne de début (1-based)
'@Param endRow Ligne de fin (1-based)
'@Param endCol Colonne de fin (1-based)
'@Returns Une interface IRangeAccessor pour manipuler la plage
Public Function GetRangeByCoords(ByVal startRow As Long, ByVal startCol As Long, _
                                ByVal endRow As Long, ByVal endCol As Long) As IRangeAccessor
End Function

'@Description("Écrit une valeur dans une cellule")
'@Param row Numéro de ligne (1-based)
'@Param col Numéro de colonne (1-based)
'@Param value Valeur à écrire
'@Returns True si l'écriture a réussi, False sinon
Public Function WriteValue(ByVal row As Long, ByVal col As Long, ByVal value As Variant) As Boolean
End Function

'@Description("Lit la valeur d'une cellule")
'@Param row Numéro de ligne (1-based)
'@Param col Numéro de colonne (1-based)
'@Returns La valeur de la cellule
Public Function ReadValue(ByVal row As Long, ByVal col As Long) As Variant
End Function

'@Description("Obtient une plage utilisée (non vide) de la feuille")
'@Returns Une interface IRangeAccessor pour manipuler la plage utilisée
Public Function GetUsedRange() As IRangeAccessor
End Function

'@Description("Ajoute une ligne")
'@Param position Position de la nouvelle ligne (1-based)
'@Returns True si l'ajout a réussi, False sinon
Public Function InsertRow(ByVal position As Long) As Boolean
End Function

'@Description("Ajoute une colonne")
'@Param position Position de la nouvelle colonne (1-based)
'@Returns True si l'ajout a réussi, False sinon
Public Function InsertColumn(ByVal position As Long) As Boolean
End Function

'@Description("Supprime une ligne")
'@Param position Position de la ligne à supprimer (1-based)
'@Returns True si la suppression a réussi, False sinon
Public Function DeleteRow(ByVal position As Long) As Boolean
End Function

'@Description("Supprime une colonne")
'@Param position Position de la colonne à supprimer (1-based)
'@Returns True si la suppression a réussi, False sinon
Public Function DeleteColumn(ByVal position As Long) As Boolean
End Function

'@Description("Trouve toutes les occurrences d'une valeur dans la feuille")
'@Param value Valeur à rechercher
'@Param matchCase True pour recherche sensible à la casse
'@Param matchEntireCell True pour correspondance exacte de cellule
'@Returns Une collection d'IRangeAccessor contenant les cellules correspondantes
Public Function FindAll(ByVal value As Variant, Optional ByVal matchCase As Boolean = False, _
                        Optional ByVal matchEntireCell As Boolean = False) As Collection
End Function

'@Description("Nom de la feuille")
Public Property Get Name() As String
End Property

'@Description("Définit le nom de la feuille")
Public Property Let Name(ByVal value As String)
End Property

'@Description("Indique si la feuille est visible")
Public Property Get Visible() As Boolean
End Property

'@Description("Définit la visibilité de la feuille")
Public Property Let Visible(ByVal value As Boolean)
End Property

'@Description("Indique si la feuille est en lecture seule")
Public Property Get ReadOnly() As Boolean
End Property

'@Description("Nombre de lignes dans la feuille")
Public Property Get RowCount() As Long
End Property

'@Description("Nombre de colonnes dans la feuille")
Public Property Get ColumnCount() As Long
End Property 