VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IChartAccessor"

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
'@ModuleDescription("Interface for accessing and manipulating Excel Charts")
Option Explicit

'*******************************************************************************
' Module : IChartAccessor
' Author : [Votre nom]
' Date   : 12/04/2025
' Purpose: Interface définissant les opérations de manipulation des graphiques
'          dans Excel
'*******************************************************************************

' ===== Propriétés =====

''
' Obtient le nom du graphique
' @return String Nom du graphique
''
Public Property Get ChartName() As String
End Property

''
' Obtient la référence à la feuille contenant le graphique
' @return ISheetAccessor Accesseur de la feuille contenant le graphique
''
Public Property Get ParentSheet() As Object ' ISheetAccessor
End Property

''
' Obtient le type de graphique
' @return Long Type de graphique (constante XlChartType)
''
Public Property Get ChartType() As Long
End Property

''
' Définit le type de graphique
' @param value Type de graphique à définir (constante XlChartType)
''
Public Property Let ChartType(ByVal value As Long)
End Property

''
' Obtient le titre du graphique
' @return String Titre du graphique
''
Public Property Get Title() As String
End Property

''
' Définit le titre du graphique
' @param value Titre à définir
''
Public Property Let Title(ByVal value As String)
End Property

''
' Indique si le graphique a une légende
' @return Boolean True si le graphique a une légende
''
Public Property Get HasLegend() As Boolean
End Property

''
' Active ou désactive la légende du graphique
' @param value True pour activer la légende, False pour la désactiver
''
Public Property Let HasLegend(ByVal value As Boolean)
End Property

''
' Obtient la position de la légende
' @return Long Position de la légende (constante XlLegendPosition)
''
Public Property Get LegendPosition() As Long
End Property

''
' Définit la position de la légende
' @param value Position de la légende à définir (constante XlLegendPosition)
''
Public Property Let LegendPosition(ByVal value As Long)
End Property

' ===== Méthodes de création et manipulation =====

''
' Définit la source de données du graphique
' @param sourceRange Plage source contenant les données
' @param seriesInRows True si les séries sont en lignes, False si en colonnes
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function SetSourceData(ByVal sourceRange As String, Optional ByVal seriesInRows As Boolean = False) As Boolean
End Function

''
' Ajoute une série de données au graphique
' @param name Nom de la série (peut être une référence de cellule ou un texte)
' @param values Plage de valeurs pour la série
' @param xValues Plage de valeurs pour l'axe X (optionnel)
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function AddSeries(ByVal name As String, ByVal values As String, Optional ByVal xValues As String = "") As Boolean
End Function

''
' Supprime une série de données du graphique
' @param seriesIndex Index de la série à supprimer (1-based)
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function RemoveSeries(ByVal seriesIndex As Long) As Boolean
End Function

''
' Supprime toutes les séries de données du graphique
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function ClearSeries() As Boolean
End Function

''
' Définit le titre de l'axe X
' @param title Titre à définir pour l'axe X
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function SetXAxisTitle(ByVal title As String) As Boolean
End Function

''
' Définit le titre de l'axe Y
' @param title Titre à définir pour l'axe Y
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function SetYAxisTitle(ByVal title As String) As Boolean
End Function

''
' Définit le format de l'axe X
' @param min Valeur minimale (optionnel)
' @param max Valeur maximale (optionnel)
' @param majorUnit Intervalle entre les graduations principales (optionnel)
' @param minorUnit Intervalle entre les graduations secondaires (optionnel)
' @param numberFormat Format numérique à utiliser (optionnel)
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function FormatXAxis(Optional ByVal min As Variant, Optional ByVal max As Variant, _
                           Optional ByVal majorUnit As Variant, Optional ByVal minorUnit As Variant, _
                           Optional ByVal numberFormat As String = "") As Boolean
End Function

''
' Définit le format de l'axe Y
' @param min Valeur minimale (optionnel)
' @param max Valeur maximale (optionnel)
' @param majorUnit Intervalle entre les graduations principales (optionnel)
' @param minorUnit Intervalle entre les graduations secondaires (optionnel)
' @param numberFormat Format numérique à utiliser (optionnel)
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function FormatYAxis(Optional ByVal min As Variant, Optional ByVal max As Variant, _
                           Optional ByVal majorUnit As Variant, Optional ByVal minorUnit As Variant, _
                           Optional ByVal numberFormat As String = "") As Boolean
End Function

''
' Applique un style prédéfini au graphique
' @param styleIndex Numéro du style à appliquer
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function ApplyChartStyle(ByVal styleIndex As Long) As Boolean
End Function

''
' Formate une série spécifique du graphique
' @param seriesIndex Index de la série à formater (1-based)
' @param formatType Type de format (1=couleur, 2=motif, 3=étiquettes, etc.)
' @param formatValue Valeur du format à appliquer
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function FormatSeries(ByVal seriesIndex As Long, ByVal formatType As Long, ByVal formatValue As Variant) As Boolean
End Function

''
' Active ou désactive les étiquettes de données pour une série
' @param seriesIndex Index de la série (1-based)
' @param showLabels True pour afficher les étiquettes, False pour les masquer
' @param labelContent Types de contenu à afficher (combination de constantes XlDataLabelContent)
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function SetDataLabels(ByVal seriesIndex As Long, ByVal showLabels As Boolean, _
                             Optional ByVal labelContent As Long = 0) As Boolean
End Function

''
' Définit la taille et la position du graphique
' @param left Position horizontale (points)
' @param top Position verticale (points)
' @param width Largeur (points)
' @param height Hauteur (points)
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function SetPosition(ByVal left As Single, ByVal top As Single, _
                           ByVal width As Single, ByVal height As Single) As Boolean
End Function

''
' Exporte le graphique au format image
' @param filePath Chemin complet du fichier d'export
' @param formatType Format d'image (png, jpg, etc.)
' @param resolution Résolution de l'image (dpi)
' @return Boolean True si l'opération réussit
''
'@Description: 
'@Param: 
'@Returns: 

Public Function ExportAsImage(ByVal filePath As String, ByVal formatType As String, _
                             Optional ByVal resolution As Long = 96) As Boolean
End Function