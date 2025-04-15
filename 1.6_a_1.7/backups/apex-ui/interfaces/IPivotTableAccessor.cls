VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPivotTableAccessor"

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
'@ModuleDescription("Interface for accessing and manipulating Excel PivotTables")
Option Explicit

'*******************************************************************************
' Module : IPivotTableAccessor
' Author : [Votre nom]
' Date   : 12/04/2025
' Purpose: Interface définissant les opérations de manipulation des tableaux
'          croisés dynamiques (PivotTables) dans Excel
'*******************************************************************************

' ===== Propriétés =====

''
' Obtient le nom du tableau croisé dynamique
' @return String Nom du tableau croisé dynamique
''
Public Property Get PivotTableName() As String
End Property

''
' Obtient la référence à la feuille contenant le tableau croisé dynamique
' @return ISheetAccessor Accesseur de la feuille contenant le tableau croisé
''
Public Property Get ParentSheet() As Object ' ISheetAccessor
End Property

''
' Indique si le tableau croisé dynamique a une zone de filtre de rapport
' @return Boolean True si des filtres de rapport existent
''
Public Property Get HasPageFields() As Boolean
End Property

''
' Obtient le nombre de champs de données (mesures) dans le tableau croisé
' @return Long Nombre de champs de données
''
Public Property Get DataFieldsCount() As Long
End Property

''
' Obtient la liste des champs utilisés dans le tableau croisé
' @return Variant() Tableau de noms de champs
''
Public Property Get FieldNames() As Variant
End Property

' ===== Méthodes de structure =====

''
' Ajoute un champ comme ligne dans le tableau croisé dynamique
' @param fieldName Nom du champ à ajouter comme ligne
' @param position Position d'insertion (optionnel)
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function AddRowField(ByVal fieldName As String, Optional ByVal position As Long = -1) As Boolean
End Function

''
' Ajoute un champ comme colonne dans le tableau croisé dynamique
' @param fieldName Nom du champ à ajouter comme colonne
' @param position Position d'insertion (optionnel)
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function AddColumnField(ByVal fieldName As String, Optional ByVal position As Long = -1) As Boolean
End Function

''
' Ajoute un champ comme filtre de rapport dans le tableau croisé dynamique
' @param fieldName Nom du champ à ajouter comme filtre
' @param position Position d'insertion (optionnel)
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function AddPageField(ByVal fieldName As String, Optional ByVal position As Long = -1) As Boolean
End Function

''
' Ajoute un champ comme données (mesure) dans le tableau croisé dynamique
' @param fieldName Nom du champ source pour les données
' @param caption Légende personnalisée pour le champ (optionnel)
' @param function Fonction d'agrégation à utiliser (somme, moyenne, etc.)
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function AddDataField(ByVal fieldName As String, Optional ByVal caption As String = "", _
                            Optional ByVal function As Long = -1) As Boolean
End Function

''
' Supprime un champ du tableau croisé dynamique
' @param fieldName Nom du champ à supprimer
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function RemoveField(ByVal fieldName As String) As Boolean
End Function

''
' Déplace un champ existant vers une autre zone (lignes, colonnes, filtres, données)
' @param fieldName Nom du champ à déplacer
' @param targetArea Zone cible (1=ligne, 2=colonne, 3=filtre, 4=données)
' @param position Position dans la zone cible (optionnel)
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function MoveField(ByVal fieldName As String, ByVal targetArea As Long, _
                         Optional ByVal position As Long = -1) As Boolean
End Function

' ===== Méthodes de filtre et de format =====

''
' Applique un filtre sur un champ spécifique du tableau croisé
' @param fieldName Nom du champ à filtrer
' @param values Valeurs à inclure/exclure dans le filtre
' @param include True pour inclure les valeurs, False pour les exclure
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function ApplyFilter(ByVal fieldName As String, ByVal values As Variant, _
                           Optional ByVal include As Boolean = True) As Boolean
End Function

''
' Efface tous les filtres appliqués à un champ
' @param fieldName Nom du champ dont les filtres doivent être effacés
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function ClearFilters(ByVal fieldName As String) As Boolean
End Function

''
' Efface tous les filtres du tableau croisé
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function ClearAllFilters() As Boolean
End Function

''
' Applique un format numérique à un champ de données
' @param dataFieldName Nom du champ de données à formater
' @param numberFormat Format numérique à appliquer (ex: "#,##0.00")
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function FormatDataField(ByVal dataFieldName As String, ByVal numberFormat As String) As Boolean
End Function

''
' Définit le sous-total pour un champ de ligne ou de colonne
' @param fieldName Nom du champ
' @param showSubtotal True pour afficher les sous-totaux, False pour les masquer
' @param function Fonction d'agrégation pour le sous-total (optionnel)
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function SetSubtotal(ByVal fieldName As String, ByVal showSubtotal As Boolean, _
                           Optional ByVal function As Long = -1) As Boolean
End Function

' ===== Méthodes d'actions =====

''
' Rafraîchit les données du tableau croisé dynamique
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function Refresh() As Boolean
End Function

''
' Développe ou réduit un élément dans le tableau croisé
' @param fieldName Nom du champ
' @param itemName Nom de l'élément à développer/réduire
' @param expand True pour développer, False pour réduire
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function ExpandItem(ByVal fieldName As String, ByVal itemName As String, _
                          ByVal expand As Boolean) As Boolean
End Function

''
' Développe ou réduit tous les éléments d'un champ
' @param fieldName Nom du champ
' @param expand True pour développer, False pour réduire
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function ExpandField(ByVal fieldName As String, ByVal expand As Boolean) As Boolean
End Function

''
' Développe ou réduit tous les champs du tableau croisé
' @param expand True pour développer, False pour réduire
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function ExpandAll(ByVal expand As Boolean) As Boolean
End Function

' ===== Méthodes de données =====

''
' Obtient les valeurs d'un champ de données spécifique sous forme de tableau
' @param dataFieldName Nom du champ de données
' @return Variant() Tableau contenant les valeurs
''
'@Description: 
'@Param: 
'@Returns: 

Public Function GetDataFieldValues(ByVal dataFieldName As String) As Variant
End Function

''
' Obtient la plage contenant le tableau croisé entier
' @return Variant() Tableau 2D contenant toutes les valeurs
''
'@Description: 
'@Param: 
'@Returns: 

Public Function GetAllValues() As Variant
End Function

''
' Obtient la liste des éléments (membres) d'un champ spécifique
' @param fieldName Nom du champ
' @return Variant() Tableau des éléments
''
'@Description: 
'@Param: 
'@Returns: 

Public Function GetFieldItems(ByVal fieldName As String) As Variant
End Function