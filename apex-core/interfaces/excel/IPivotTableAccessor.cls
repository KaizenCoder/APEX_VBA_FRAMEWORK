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
'@Interface
'@Folder("APEX.Core.Interfaces.Excel")
Option Explicit

' ==========================================================================
' Interface  : IPivotTableAccessor
' Purpose   : Interface pour la manipulation des tableaux croisés dynamiques
' Author    : APEX Framework Team
' Date      : 2024-04-11
' ==========================================================================

'@Description("Types d'agrégation supportés")
Public Enum PivotAggregationType
    pvtSum = 1           ' Somme
    pvtCount = 2         ' Nombre
    pvtAverage = 3       ' Moyenne
    pvtMax = 4           ' Maximum
    pvtMin = 5           ' Minimum
    pvtProduct = 6       ' Produit
    pvtCountNumbers = 7  ' Nombre de valeurs
    pvtStdDev = 8       ' Écart-type
    pvtStdDevP = 9      ' Écart-type population
    pvtVar = 10         ' Variance
    pvtVarP = 11        ' Variance population
End Enum

'@Description("Obtient le nom du tableau croisé dynamique")
Public Property Get PivotTableName() As String
End Property

'@Description("Obtient la plage du tableau croisé dynamique")
Public Property Get PivotRange() As Range
End Property

'@Description("Obtient ou définit la source de données")
Public Property Get SourceRange() As Range
End Property
Public Property Set SourceRange(ByVal rng As Range)
End Property

'@Description("Ajoute un champ en ligne")
'@Param fieldName Nom du champ à ajouter
'@Param position Position du champ (1-based), -1 pour ajouter à la fin
'@Description: 
'@Param: 
'@Returns: 

Public Sub AddRowField(ByVal fieldName As String, Optional ByVal position As Long = -1)
End Sub

'@Description("Ajoute un champ en colonne")
'@Param fieldName Nom du champ à ajouter
'@Param position Position du champ (1-based), -1 pour ajouter à la fin
'@Description: 
'@Param: 
'@Returns: 

Public Sub AddColumnField(ByVal fieldName As String, Optional ByVal position As Long = -1)
End Sub

'@Description("Ajoute un champ de valeur avec son agrégation")
'@Param fieldName Nom du champ à ajouter
'@Param aggregationType Type d'agrégation à utiliser
'@Param customName Nom personnalisé pour le champ (optionnel)
'@Description: 
'@Param: 
'@Returns: 

Public Sub AddValueField(ByVal fieldName As String, _
                        ByVal aggregationType As PivotAggregationType, _
                        Optional ByVal customName As String = "")
End Sub

'@Description("Ajoute un champ de filtre")
'@Param fieldName Nom du champ à ajouter
'@Param position Position du champ (1-based), -1 pour ajouter à la fin
'@Description: 
'@Param: 
'@Returns: 

Public Sub AddFilterField(ByVal fieldName As String, Optional ByVal position As Long = -1)
End Sub

'@Description("Supprime un champ")
'@Param fieldName Nom du champ à supprimer
'@Description: 
'@Param: 
'@Returns: 

Public Sub RemoveField(ByVal fieldName As String)
End Sub

'@Description("Applique un filtre sur un champ")
'@Param fieldName Nom du champ à filtrer
'@Param criteria Critères de filtrage
'@Description: 
'@Param: 
'@Returns: 

Public Sub ApplyFilter(ByVal fieldName As String, ByVal criteria As Variant)
End Sub

'@Description("Efface tous les filtres")
'@Description: 
'@Param: 
'@Returns: 

Public Sub ClearFilters()
End Sub

'@Description("Applique un style prédéfini au tableau croisé")
'@Param styleName Nom du style de tableau croisé
'@Description: 
'@Param: 
'@Returns: 

Public Sub ApplyPivotStyle(ByVal styleName As String)
End Sub

'@Description("Configure les options d'affichage")
'@Param showGrandTotals Afficher les totaux généraux
'@Param showSubTotals Afficher les sous-totaux
'@Param repeatLabels Répéter les étiquettes
'@Description: 
'@Param: 
'@Returns: 

Public Sub ConfigureDisplay(Optional ByVal showGrandTotals As Boolean = True, _
                          Optional ByVal showSubTotals As Boolean = True, _
                          Optional ByVal repeatLabels As Boolean = False)
End Sub

'@Description("Rafraîchit les données du tableau croisé")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Refresh()
End Sub

'@Description("Efface le tableau croisé et recommence à zéro")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Clear()
End Sub

'@Description("Exporte le tableau croisé vers une nouvelle feuille")
'@Return Référence vers la nouvelle feuille
'@Description: 
'@Param: 
'@Returns: 

Public Function ExportToNewSheet() As Worksheet
End Function

'@Description("Obtient les données du tableau croisé sous forme de tableau 2D")
'@Return Tableau 2D contenant les données
'@Description: 
'@Param: 
'@Returns: 

Public Function GetPivotData() As Variant
End Function 