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
'@Folder("APEX.UI.Interfaces")
'@ModuleDescription("Interface for accessing and manipulating Excel Tables")
Option Explicit

'*******************************************************************************
' Module : ITableAccessor
' Author : [Votre nom]
' Date   : 12/04/2025
' Purpose: Interface définissant les opérations de manipulation des Tables Excel
'*******************************************************************************

' ===== Propriétés =====

''
' Obtient le nom de la table Excel
' @return String Nom de la table
''
Public Property Get TableName() As String
End Property

''
' Obtient la référence à la feuille contenant la table
' @return ISheetAccessor Accesseur de la feuille contenant la table
''
Public Property Get ParentSheet() As Object ' ISheetAccessor
End Property

''
' Obtient le nombre total de lignes dans la table (sans compter l'en-tête)
' @return Long Nombre de lignes de données
''
Public Property Get RowCount() As Long
End Property

''
' Obtient le nombre total de colonnes dans la table
' @return Long Nombre de colonnes
''
Public Property Get ColumnCount() As Long
End Property

''
' Obtient la liste des noms de colonnes de la table
' @return Variant() Tableau de noms de colonnes
''
Public Property Get ColumnNames() As Variant
End Property

' ===== Méthodes de lecture =====

''
' Lit toutes les données de la table (sans l'en-tête)
' @return Variant() Tableau 2D contenant les données
''
'@Description: 
'@Param: 
'@Returns: 

Public Function ReadAllData() As Variant
End Function

''
' Lit l'en-tête de la table (noms des colonnes)
' @return Variant() Tableau 1D contenant les noms des colonnes
''
'@Description: 
'@Param: 
'@Returns: 

Public Function ReadHeader() As Variant
End Function

''
' Lit une ligne spécifique de la table
' @param rowIndex Index de la ligne à lire (1-based)
' @return Variant() Tableau 1D contenant les données de la ligne
''
'@Description: 
'@Param: 
'@Returns: 

Public Function ReadRow(ByVal rowIndex As Long) As Variant
End Function

''
' Lit une colonne spécifique de la table
' @param columnName Nom de la colonne à lire
' @return Variant() Tableau 1D contenant les données de la colonne
''
'@Description: 
'@Param: 
'@Returns: 

Public Function ReadColumn(ByVal columnName As String) As Variant
End Function

''
' Lit une cellule spécifique de la table
' @param rowIndex Index de la ligne (1-based)
' @param columnName Nom de la colonne
' @return Variant Valeur de la cellule
''
'@Description: 
'@Param: 
'@Returns: 

Public Function ReadCell(ByVal rowIndex As Long, ByVal columnName As String) As Variant
End Function

' ===== Méthodes d'écriture =====

''
' Écrit des données dans une ligne spécifique de la table
' @param rowIndex Index de la ligne à modifier (1-based)
' @param rowData Tableau 1D contenant les données à écrire
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function WriteRow(ByVal rowIndex As Long, ByVal rowData As Variant) As Boolean
End Function

''
' Écrit des données dans une colonne spécifique de la table
' @param columnName Nom de la colonne à modifier
' @param columnData Tableau 1D contenant les données à écrire
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function WriteColumn(ByVal columnName As String, ByVal columnData As Variant) As Boolean
End Function

''
' Écrit une valeur dans une cellule spécifique de la table
' @param rowIndex Index de la ligne (1-based)
' @param columnName Nom de la colonne
' @param value Valeur à écrire
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function WriteCell(ByVal rowIndex As Long, ByVal columnName As String, ByVal value As Variant) As Boolean
End Function

' ===== Méthodes de manipulation de la structure =====

''
' Ajoute une nouvelle ligne à la fin de la table
' @param rowData Tableau 1D contenant les données de la nouvelle ligne (optionnel)
' @return Long Index de la nouvelle ligne
''
'@Description: 
'@Param: 
'@Returns: 

Public Function AddRow(Optional ByVal rowData As Variant) As Long
End Function

''
' Supprime une ligne spécifique de la table
' @param rowIndex Index de la ligne à supprimer (1-based)
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function DeleteRow(ByVal rowIndex As Long) As Boolean
End Function

''
' Ajoute une nouvelle colonne à la table
' @param columnName Nom de la nouvelle colonne
' @param columnData Tableau 1D contenant les données de la colonne (optionnel)
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function AddColumn(ByVal columnName As String, Optional ByVal columnData As Variant) As Boolean
End Function

''
' Supprime une colonne spécifique de la table
' @param columnName Nom de la colonne à supprimer
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function DeleteColumn(ByVal columnName As String) As Boolean
End Function

' ===== Méthodes de filtrage et tri =====

''
' Applique un filtre à une colonne de la table
' @param columnName Nom de la colonne à filtrer
' @param criteria Critère de filtrage
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function ApplyFilter(ByVal columnName As String, ByVal criteria As String) As Boolean
End Function

''
' Supprime tous les filtres de la table
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function ClearFilters() As Boolean
End Function

''
' Trie la table selon une colonne
' @param columnName Nom de la colonne de tri
' @param ascending True pour tri ascendant, False pour descendant
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function SortByColumn(ByVal columnName As String, Optional ByVal ascending As Boolean = True) As Boolean
End Function

' ===== Méthodes de mise en forme =====

''
' Applique un style prédéfini à la table
' @param styleName Nom du style à appliquer
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function ApplyTableStyle(ByVal styleName As String) As Boolean
End Function

''
' Définit la mise en forme conditionnelle pour une colonne
' @param columnName Nom de la colonne
' @param formulaString Formule de mise en forme conditionnelle
' @param formatSettings Paramètres de format (couleur, police, etc.)
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function SetConditionalFormatting(ByVal columnName As String, ByVal formulaString As String, ByVal formatSettings As Variant) As Boolean
End Function

'@Interface: ITableAccessor
'@Description: Interface pour manipuler les Tables Excel
'@Version: 1.0
Option Explicit

'@Description: Crée une nouvelle table Excel
'@Param: ws As Worksheet - La feuille où créer la table
'@Param: rangeName As String - La plage pour la table
'@Param: tableName As String - Le nom de la table à créer
'@Returns: Boolean - True si succès, False sinon
Public Function CreateTable(ByVal ws As Worksheet, ByVal rangeName As String, ByVal tableName As String) As Boolean
End Function

'@Description: Supprime une table Excel existante
'@Param: tableName As String - Le nom de la table à supprimer
'@Returns: Boolean - True si succès, False sinon
Public Function DeleteTable(ByVal tableName As String) As Boolean
End Function

'@Description: Ajoute une ligne à la table
'@Param: tableName As String - Le nom de la table
'@Param: rowData As Variant - Les données de la ligne à ajouter
'@Returns: Boolean - True si succès, False sinon
Public Function AddRow(ByVal tableName As String, ByVal rowData As Variant) As Boolean
End Function

'@Description: Applique un filtre à la table
'@Param: tableName As String - Le nom de la table
'@Param: columnName As String - La colonne à filtrer
'@Param: criteria As String - Les critères de filtrage
'@Returns: Boolean - True si succès, False sinon
Public Function ApplyFilter(ByVal tableName As String, ByVal columnName As String, ByVal criteria As String) As Boolean
End Function

'@Description: Trie la table selon une colonne
'@Param: tableName As String - Le nom de la table
'@Param: columnName As String - La colonne de tri
'@Param: ascending As Boolean - True pour tri ascendant
'@Returns: Boolean - True si succès, False sinon
Public Function SortBy(ByVal tableName As String, ByVal columnName As String, Optional ByVal ascending As Boolean = True) As Boolean
End Function

'@Description: Met à jour le style de la table
'@Param: tableName As String - Le nom de la table
'@Param: styleName As String - Le nom du style à appliquer
'@Returns: Boolean - True si succès, False sinon
Public Function ApplyStyle(ByVal tableName As String, ByVal styleName As String) As Boolean
End Function

'@Description: Récupère les données de la table
'@Param: tableName As String - Le nom de la table
'@Returns: Variant - Les données de la table en tableau 2D
Public Function GetTableData(ByVal tableName As String) As Variant
End Function

'@Description: Vérifie si une table existe
'@Param: tableName As String - Le nom de la table à vérifier
'@Returns: Boolean - True si la table existe
Public Function TableExists(ByVal tableName As String) As Boolean
End Function