VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IConfigManagerBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@ModuleDescription("Interface de base pour les gestionnaires de configuration")
'@Interface
'@Version 1.7.0
'@Author APEX Framework Team
Option Explicit

' =========================================================
' Interface IConfigManagerBase
' Contrat pour tous les gestionnaires de configuration
' =========================================================

' --------------------------------------------------------
' Configuration loading
' --------------------------------------------------------
Public Function LoadConfig(ByVal configPath As String) As Boolean
    ' Charge la configuration � partir du chemin sp�cifi�
    ' @param configPath Chemin vers le fichier de configuration
    ' @return True si charg� avec succ�s, False sinon
End Function

Public Function LoadConfigFromString(ByVal configContent As String) As Boolean
    ' Charge la configuration � partir d'une cha�ne
    ' @param configContent Contenu de la configuration
    ' @return True si charg� avec succ�s, False sinon
End Function

' --------------------------------------------------------
' Value access
' --------------------------------------------------------
Public Function GetValue(ByVal section As String, ByVal key As String, Optional ByVal defaultValue As Variant = Null) As Variant
    ' R�cup�re une valeur de configuration
    ' @param section Section de configuration
    ' @param key Cl� � r�cup�rer
    ' @param defaultValue Valeur par d�faut si la cl� n'existe pas
    ' @return Valeur associ�e � la cl� ou defaultValue
End Function

Public Function SetValue(ByVal section As String, ByVal key As String, ByVal value As Variant) As Boolean
    ' D�finit une valeur de configuration
    ' @param section Section de configuration
    ' @param key Cl� � d�finir
    ' @param value Valeur � associer
    ' @return True si d�fini avec succ�s, False sinon
End Function

' --------------------------------------------------------
' Section management
' --------------------------------------------------------
Public Function SectionExists(ByVal section As String) As Boolean
    ' V�rifie si une section existe
    ' @param section Nom de la section
    ' @return True si la section existe, False sinon
End Function

Public Function GetSections() As Variant
    ' R�cup�re la liste des sections disponibles
    ' @return Tableau des noms de sections
End Function

Public Function GetKeysInSection(ByVal section As String) As Variant
    ' R�cup�re les cl�s dans une section
    ' @param section Nom de la section
    ' @return Tableau des cl�s dans la section
End Function

' --------------------------------------------------------
' Persistence
' --------------------------------------------------------
Public Function SaveConfig(Optional ByVal configPath As String = "") As Boolean
    ' Sauvegarde la configuration actuelle
    ' @param configPath Chemin optionnel pour sauvegarder (si vide, utilise le chemin de chargement)
    ' @return True si sauvegard� avec succ�s, False sinon
End Function

Public Function ReloadConfig() As Boolean
    ' Recharge la configuration � partir de la derni�re source
    ' @return True si recharg� avec succ�s, False sinon
End Function

' --------------------------------------------------------
' Validation
' --------------------------------------------------------
Public Function ValidateConfig() As Boolean
    ' Valide la structure et le contenu de la configuration
    ' @return True si valide, False sinon
End Function

Public Function GetValidationErrors() As Variant
    ' R�cup�re les erreurs de validation
    ' @return Tableau des erreurs de validation
End Function 