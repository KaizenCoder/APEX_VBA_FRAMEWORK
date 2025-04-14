VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWorkbookAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
'@Folder("APEX.Core.Interfaces.Excel")
Option Explicit

' ==========================================================================
' Interface  : IWorkbookAccessor
' Purpose    : Interface pour l'accès aux objets Workbook d'Excel
' Author     : APEX Framework Team
' Date       : 2024-04-11
' ==========================================================================

'@Description("Obtient une feuille de calcul par son nom")
'@Param name Le nom de la feuille à récupérer
'@Returns Une interface ISheetAccessor pour manipuler la feuille
Public Function GetSheet(ByVal name As String) As ISheetAccessor
End Function

'@Description("Obtient la feuille active")
'@Returns Une interface ISheetAccessor pour manipuler la feuille active
Public Function GetActiveSheet() As ISheetAccessor
End Function

'@Description("Sauvegarde le classeur")
'@Returns True si la sauvegarde a réussi, False sinon
Public Function Save() As Boolean
End Function

'@Description("Sauvegarde le classeur sous un nouveau nom")
'@Param path Le chemin complet où sauvegarder le classeur
'@Returns True si la sauvegarde a réussi, False sinon
Public Function SaveAs(ByVal path As String) As Boolean
End Function

'@Description("Crée une nouvelle feuille")
'@Param name Nom de la nouvelle feuille
'@Param position Position de la nouvelle feuille (1-based), -1 pour la dernière position
'@Returns Une interface ISheetAccessor pour manipuler la nouvelle feuille
Public Function CreateSheet(ByVal name As String, Optional ByVal position As Long = -1) As ISheetAccessor
End Function

'@Description("Supprime une feuille")
'@Param name Nom de la feuille à supprimer
'@Returns True si la suppression a réussi, False sinon
Public Function DeleteSheet(ByVal name As String) As Boolean
End Function

'@Description("Ferme le classeur")
'@Param saveChanges True pour enregistrer les modifications avant la fermeture
'@Returns True si la fermeture a réussi, False sinon
Public Function Close(Optional ByVal saveChanges As Boolean = False) As Boolean
End Function

'@Description("Chemin d'accès au fichier du classeur")
Public Property Get Path() As String
End Property

'@Description("Nom du fichier du classeur (sans le chemin)")
Public Property Get Name() As String
End Property

'@Description("Le classeur est-il en lecture seule")
Public Property Get ReadOnly() As Boolean
End Property

'@Description("Nombre de feuilles dans le classeur")
Public Property Get SheetCount() As Long
End Property

'@Description("Liste des noms de feuilles dans le classeur")
Public Property Get SheetNames() As Variant
End Property 