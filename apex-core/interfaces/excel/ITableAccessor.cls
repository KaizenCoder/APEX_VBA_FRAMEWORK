VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITableAccessor"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
'@Folder("APEX.Core.Interfaces.Excel")
Option Explicit

' ==========================================================================
' Interface  : ITableAccessor
' Purpose   : Interface pour la manipulation des Tables Excel (ListObjects)
' Author    : APEX Framework Team
' Date      : 2024-04-11
' ==========================================================================

'@Description("Obtient le nom de la table")
Public Property Get TableName() As String
End Property

'@Description("Obtient ou définit la plage de données de la table")
Public Property Get DataRange() As Range
End Property

'@Description("Obtient la plage d'en-têtes de la table")
Public Property Get HeaderRange() As Range
End Property

'@Description("Obtient le nombre de lignes de données (sans compter l'en-tête)")
Public Property Get DataRowCount() As Long
End Property

'@Description("Obtient le nombre de colonnes")
Public Property Get ColumnCount() As Long
End Property

'@Description("Lit toutes les données de la table (sans les en-têtes)")
'@Description: 
'@Param: 
'@Returns: 

Public Function ReadData() As Variant
End Function

'@Description("Écrit des données dans la table")
'@Param data Tableau 2D de données à écrire
'@Param appendData Si Vrai, ajoute les données à la fin, sinon remplace les données existantes
'@Description: 
'@Param: 
'@Returns: 

Public Sub WriteData(ByVal data As Variant, Optional ByVal appendData As Boolean = False)
End Sub

'@Description("Applique un filtre sur une colonne")
'@Param columnName Nom de la colonne à filtrer
'@Param criteria Critères de filtrage
'@Description: 
'@Param: 
'@Returns: 

Public Sub ApplyFilter(ByVal columnName As String, ByVal criteria As Variant)
End Sub

'@Description("Efface tous les filtres")
'@Description: 
'@Param: 
'@Returns: 

Public Sub ClearFilters()
End Sub

'@Description("Trie la table selon une colonne")
'@Param columnName Nom de la colonne
'@Param ascending Vrai pour tri ascendant, Faux pour descendant
'@Description: 
'@Param: 
'@Returns: 

Public Sub Sort(ByVal columnName As String, Optional ByVal ascending As Boolean = True)
End Sub

'@Description("Ajoute une colonne à la table")
'@Param columnName Nom de la colonne
'@Param position Position de la colonne (1-based), -1 pour ajouter à la fin
'@Return Index de la nouvelle colonne
'@Description: 
'@Param: 
'@Returns: 

Public Function AddColumn(ByVal columnName As String, Optional ByVal position As Long = -1) As Long
End Function

'@Description("Supprime une colonne de la table")
'@Param columnName Nom de la colonne à supprimer
'@Description: 
'@Param: 
'@Returns: 

Public Sub RemoveColumn(ByVal columnName As String)
End Sub

'@Description("Obtient l'index d'une colonne par son nom")
'@Param columnName Nom de la colonne
'@Return Index de la colonne (1-based), 0 si non trouvée
'@Description: 
'@Param: 
'@Returns: 

Public Function GetColumnIndex(ByVal columnName As String) As Long
End Function

'@Description("Obtient le nom d'une colonne par son index")
'@Param columnIndex Index de la colonne (1-based)
'@Return Nom de la colonne
'@Description: 
'@Param: 
'@Returns: 

Public Function GetColumnName(ByVal columnIndex As Long) As String
End Function

'@Description("Applique un style prédéfini à la table")
'@Param styleName Nom du style de table Excel
'@Description: 
'@Param: 
'@Returns: 

Public Sub ApplyTableStyle(ByVal styleName As String)
End Sub

'@Description("Convertit la table en plage normale")
'@Description: 
'@Param: 
'@Returns: 

Public Sub ConvertToRange()
End Sub

'@Description("Rafraîchit la table et ses calculs")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Refresh()
End Sub 