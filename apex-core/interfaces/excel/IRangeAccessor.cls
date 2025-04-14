VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRangeAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
'@Folder("APEX.Core.Interfaces.Excel")
Option Explicit

' ==========================================================================
' Interface  : IRangeAccessor
' Purpose    : Interface pour l'accès aux plages de cellules Excel
' Author     : APEX Framework Team
' Date       : 2024-04-11
' ==========================================================================

'@Description("Obtient une cellule dans la plage par ses coordonnées relatives")
'@Param row Numéro de ligne relatif à la plage (1-based)
'@Param col Numéro de colonne relatif à la plage (1-based)
'@Returns Une interface ICellAccessor pour manipuler la cellule
Public Function GetCell(ByVal row As Long, ByVal col As Long) As ICellAccessor
End Function

'@Description("Valeurs de la plage (tableau 2D)")
Public Property Get Values() As Variant
End Property

'@Description("Définit les valeurs de la plage (tableau 2D)")
Public Property Let Values(ByVal values As Variant)
End Property

'@Description("Formules de la plage (tableau 2D)")
Public Property Get Formulas() As Variant
End Property

'@Description("Définit les formules de la plage (tableau 2D)")
Public Property Let Formulas(ByVal formulas As Variant)
End Property

'@Description("Formules locales de la plage (tableau 2D)")
Public Property Get FormulasLocal() As Variant
End Property

'@Description("Définit les formules locales de la plage (tableau 2D)")
Public Property Let FormulasLocal(ByVal formulas As Variant)
End Property

'@Description("Formules R1C1 de la plage (tableau 2D)")
Public Property Get FormulasR1C1() As Variant
End Property

'@Description("Définit les formules R1C1 de la plage (tableau 2D)")
Public Property Let FormulasR1C1(ByVal formulas As Variant)
End Property

'@Description("Adresse de la plage (ex: 'A1:C10')")
Public Property Get Address() As String
End Property

'@Description("Adresse R1C1 de la plage (ex: 'R1C1:R10C3')")
Public Property Get AddressR1C1() As String
End Property

'@Description("Première ligne de la plage (1-based)")
Public Property Get FirstRow() As Long
End Property

'@Description("Dernière ligne de la plage (1-based)")
Public Property Get LastRow() As Long
End Property

'@Description("Première colonne de la plage (1-based)")
Public Property Get FirstColumn() As Long
End Property

'@Description("Dernière colonne de la plage (1-based)")
Public Property Get LastColumn() As Long
End Property

'@Description("Nombre de lignes dans la plage")
Public Property Get RowCount() As Long
End Property

'@Description("Nombre de colonnes dans la plage")
Public Property Get ColumnCount() As Long
End Property

'@Description("Format numérique de la plage")
Public Property Get NumberFormat() As String
End Property

'@Description("Définit le format numérique de la plage")
Public Property Let NumberFormat(ByVal format As String)
End Property

'@Description("Couleur de fond de la plage (Long RGB)")
Public Property Get BackColor() As Long
End Property

'@Description("Définit la couleur de fond de la plage (Long RGB)")
Public Property Let BackColor(ByVal color As Long)
End Property

'@Description("Couleur du texte de la plage (Long RGB)")
Public Property Get ForeColor() As Long
End Property

'@Description("Définit la couleur du texte de la plage (Long RGB)")
Public Property Let ForeColor(ByVal color As Long)
End Property

'@Description("Style de police de la plage")
Public Property Get FontBold() As Boolean
End Property

'@Description("Définit le style gras de la police de la plage")
Public Property Let FontBold(ByVal bold As Boolean)
End Property

'@Description("Style italique de la police de la plage")
Public Property Get FontItalic() As Boolean
End Property

'@Description("Définit le style italique de la police de la plage")
Public Property Let FontItalic(ByVal italic As Boolean)
End Property

'@Description("Trie la plage")
'@Param keyColumn Numéro de colonne servant de clé de tri, relatif à la plage (1-based)
'@Param ascending True pour tri ascendant, False pour tri descendant
'@Param hasHeader True si la première ligne est un en-tête à ne pas trier
'@Returns True si le tri a réussi, False sinon
Public Function Sort(ByVal keyColumn As Long, Optional ByVal ascending As Boolean = True, _
                     Optional ByVal hasHeader As Boolean = True) As Boolean
End Function

'@Description("Cherche une valeur dans la plage")
'@Param value Valeur à rechercher
'@Param matchCase True pour recherche sensible à la casse
'@Param matchEntireCell True pour correspondance exacte de cellule
'@Returns Une interface ICellAccessor pour la première cellule trouvée, ou Nothing si non trouvée
Public Function Find(ByVal value As Variant, Optional ByVal matchCase As Boolean = False, _
                    Optional ByVal matchEntireCell As Boolean = False) As ICellAccessor
End Function

'@Description("Cherche toutes les occurrences d'une valeur dans la plage")
'@Param value Valeur à rechercher
'@Param matchCase True pour recherche sensible à la casse
'@Param matchEntireCell True pour correspondance exacte de cellule
'@Returns Une collection d'ICellAccessor contenant les cellules correspondantes
Public Function FindAll(ByVal value As Variant, Optional ByVal matchCase As Boolean = False, _
                        Optional ByVal matchEntireCell As Boolean = False) As Collection
End Function

'@Description("Efface le contenu de la plage")
Public Sub Clear()
End Sub

'@Description("Efface le format de la plage")
Public Sub ClearFormat()
End Sub

'@Description("Efface tout (contenu et format) de la plage")
Public Sub ClearAll()
End Sub

'@Description("Copie la plage vers une destination")
'@Param destination Adresse de destination (ex: "D5")
'@Returns True si la copie a réussi, False sinon
Public Function Copy(ByVal destination As String) As Boolean
End Function

'@Description("Crée un tableau Excel à partir de la plage")
'@Param tableName Nom du tableau
'@Param hasHeaders True si la première ligne contient des en-têtes
'@Returns Une interface ITableAccessor pour manipuler le tableau, ou Nothing en cas d'échec
Public Function CreateTable(ByVal tableName As String, Optional ByVal hasHeaders As Boolean = True) As ITableAccessor
End Function

'@Description("Fusionne les cellules de la plage")
'@Returns True si la fusion a réussi, False sinon
Public Function Merge() As Boolean
End Function

'@Description("Sépare les cellules fusionnées de la plage")
'@Returns True si la séparation a réussi, False sinon
Public Function Unmerge() As Boolean
End Function

'@Description("Obtient la feuille contenant cette plage")
'@Returns Une interface ISheetAccessor pour manipuler la feuille
Public Function GetParentSheet() As ISheetAccessor
End Function 