VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICellAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
'@Folder("APEX.Core.Interfaces.Excel")
Option Explicit

' ==========================================================================
' Interface  : ICellAccessor
' Purpose    : Interface pour l'accès aux cellules Excel (Range d'une cellule)
' Author     : APEX Framework Team
' Date       : 2024-04-11
' ==========================================================================

'@Description("Valeur de la cellule")
Public Property Get Value() As Variant
End Property

'@Description("Définit la valeur de la cellule")
Public Property Let Value(ByVal value As Variant)
End Property

'@Description("Formule de la cellule (sans le =)")
Public Property Get Formula() As String
End Property

'@Description("Définit la formule de la cellule (sans le =)")
Public Property Let Formula(ByVal formula As String)
End Property

'@Description("Formule locale de la cellule (sans le =)")
Public Property Get FormulaLocal() As String
End Property

'@Description("Définit la formule locale de la cellule (sans le =)")
Public Property Let FormulaLocal(ByVal formula As String)
End Property

'@Description("Formule R1C1 de la cellule (sans le =)")
Public Property Get FormulaR1C1() As String
End Property

'@Description("Définit la formule R1C1 de la cellule (sans le =)")
Public Property Let FormulaR1C1(ByVal formula As String)
End Property

'@Description("Adresse de la cellule (ex: 'A1')")
Public Property Get Address() As String
End Property

'@Description("Adresse R1C1 de la cellule (ex: 'R1C1')")
Public Property Get AddressR1C1() As String
End Property

'@Description("Numéro de ligne de la cellule (1-based)")
Public Property Get Row() As Long
End Property

'@Description("Numéro de colonne de la cellule (1-based)")
Public Property Get Column() As Long
End Property

'@Description("Texte affiché dans la cellule (tenant compte du format)")
Public Property Get Text() As String
End Property

'@Description("Format numérique de la cellule")
Public Property Get NumberFormat() As String
End Property

'@Description("Définit le format numérique de la cellule")
Public Property Let NumberFormat(ByVal format As String)
End Property

'@Description("Couleur de fond de la cellule (Long RGB)")
Public Property Get BackColor() As Long
End Property

'@Description("Définit la couleur de fond de la cellule (Long RGB)")
Public Property Let BackColor(ByVal color As Long)
End Property

'@Description("Couleur du texte de la cellule (Long RGB)")
Public Property Get ForeColor() As Long
End Property

'@Description("Définit la couleur du texte de la cellule (Long RGB)")
Public Property Let ForeColor(ByVal color As Long)
End Property

'@Description("Style de police de la cellule")
Public Property Get FontBold() As Boolean
End Property

'@Description("Définit le style gras de la police de la cellule")
Public Property Let FontBold(ByVal bold As Boolean)
End Property

'@Description("Style italique de la police de la cellule")
Public Property Get FontItalic() As Boolean
End Property

'@Description("Définit le style italique de la police de la cellule")
Public Property Let FontItalic(ByVal italic As Boolean)
End Property

'@Description("Indique si la cellule contient une formule")
Public Property Get HasFormula() As Boolean
End Property

'@Description("Efface le contenu de la cellule")
Public Sub Clear()
End Sub

'@Description("Efface le format de la cellule")
Public Sub ClearFormat()
End Sub

'@Description("Efface tout (contenu et format) de la cellule")
Public Sub ClearAll()
End Sub

'@Description("Obtient la feuille contenant cette cellule")
'@Returns Une interface ISheetAccessor pour manipuler la feuille
Public Function GetParentSheet() As ISheetAccessor
End Function 