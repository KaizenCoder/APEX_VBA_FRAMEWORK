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
'@Interface
'@Folder("APEX.Core.Interfaces.Excel")
Option Explicit

' ==========================================================================
' Interface  : IChartAccessor
' Purpose   : Interface pour la manipulation des graphiques Excel
' Author    : APEX Framework Team
' Date      : 2024-04-11
' ==========================================================================

'@Description("Types de graphiques supportés")
Public Enum ChartType
    chtColumn = 1        ' Colonnes
    chtBar = 2           ' Barres
    chtLine = 3          ' Lignes
    chtPie = 4           ' Secteurs
    chtXY = 5            ' Nuage de points
    chtArea = 6          ' Aires
    chtDoughnut = 7      ' Anneau
    chtRadar = 8         ' Radar
    chtBubble = 9        ' Bulles
    chtStock = 10        ' Boursier
    chtSurface = 11      ' Surface
End Enum

'@Description("Obtient le nom du graphique")
Public Property Get ChartName() As String
End Property

'@Description("Obtient ou définit le titre du graphique")
Public Property Get Title() As String
End Property
Public Property Let Title(ByVal value As String)
End Property

'@Description("Obtient ou définit la source de données")
Public Property Get SourceRange() As Range
End Property
Public Property Set SourceRange(ByVal rng As Range)
End Property

'@Description("Obtient ou définit le type de graphique")
Public Property Get Type() As ChartType
End Property
Public Property Let Type(ByVal value As ChartType)
End Property

'@Description("Crée une nouvelle série")
'@Param name Nom de la série
'@Param values Plage des valeurs
'@Param xValues Plage des valeurs X (optionnel)
'@Return Index de la nouvelle série
'@Description: 
'@Param: 
'@Returns: 

Public Function AddSeries(ByVal name As String, _
                         ByVal values As Range, _
                         Optional ByVal xValues As Range = Nothing) As Long
End Function

'@Description("Supprime une série")
'@Param seriesIndex Index de la série à supprimer
'@Description: 
'@Param: 
'@Returns: 

Public Sub RemoveSeries(ByVal seriesIndex As Long)
End Sub

'@Description("Configure l'axe des X")
'@Param title Titre de l'axe
'@Param min Valeur minimale (optionnel)
'@Param max Valeur maximale (optionnel)
'@Param majorUnit Unité principale (optionnel)
'@Description: 
'@Param: 
'@Returns: 

Public Sub ConfigureXAxis(ByVal title As String, _
                         Optional ByVal min As Variant = Empty, _
                         Optional ByVal max As Variant = Empty, _
                         Optional ByVal majorUnit As Variant = Empty)
End Sub

'@Description("Configure l'axe des Y")
'@Param title Titre de l'axe
'@Param min Valeur minimale (optionnel)
'@Param max Valeur maximale (optionnel)
'@Param majorUnit Unité principale (optionnel)
'@Description: 
'@Param: 
'@Returns: 

Public Sub ConfigureYAxis(ByVal title As String, _
                         Optional ByVal min As Variant = Empty, _
                         Optional ByVal max As Variant = Empty, _
                         Optional ByVal majorUnit As Variant = Empty)
End Sub

'@Description("Configure la légende")
'@Param position Position de la légende (xlLegendPosition)
'@Param showLegend Afficher/masquer la légende
'@Description: 
'@Param: 
'@Returns: 

Public Sub ConfigureLegend(ByVal position As XlLegendPosition, _
                          Optional ByVal showLegend As Boolean = True)
End Sub

'@Description("Applique un style prédéfini au graphique")
'@Param styleName Nom du style de graphique
'@Description: 
'@Param: 
'@Returns: 

Public Sub ApplyChartStyle(ByVal styleName As String)
End Sub

'@Description("Configure les étiquettes de données")
'@Param showValues Afficher les valeurs
'@Param showSeriesName Afficher le nom de la série
'@Param showPercentage Afficher le pourcentage
'@Param position Position des étiquettes
'@Description: 
'@Param: 
'@Returns: 

Public Sub ConfigureDataLabels(Optional ByVal showValues As Boolean = True, _
                             Optional ByVal showSeriesName As Boolean = False, _
                             Optional ByVal showPercentage As Boolean = False, _
                             Optional ByVal position As XlDataLabelPosition = xlLabelPositionBestFit)
End Sub

'@Description("Exporte le graphique vers une nouvelle feuille")
'@Return Référence vers la nouvelle feuille
'@Description: 
'@Param: 
'@Returns: 

Public Function ExportToNewSheet() As Worksheet
End Function

'@Description("Exporte le graphique vers une image")
'@Param filePath Chemin du fichier image
'@Param format Format de l'image (PNG, JPG, etc.)
'@Description: 
'@Param: 
'@Returns: 

Public Sub ExportToImage(ByVal filePath As String, ByVal format As String)
End Sub

'@Description("Rafraîchit le graphique")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Refresh()
End Sub

'@Description("Redimensionne le graphique")
'@Param width Largeur en points
'@Param height Hauteur en points
'@Description: 
'@Param: 
'@Returns: 

Public Sub Resize(ByVal width As Double, ByVal height As Double)
End Sub

'@Description("Déplace le graphique")
'@Param left Position gauche en points
'@Param top Position haute en points
'@Description: 
'@Param: 
'@Returns: 

Public Sub Move(ByVal left As Double, ByVal top As Double)
End Sub

'@Description("Copie le graphique dans le presse-papiers")
'@Description: 
'@Param: 
'@Returns: 

Public Sub CopyToClipboard()
End Sub 