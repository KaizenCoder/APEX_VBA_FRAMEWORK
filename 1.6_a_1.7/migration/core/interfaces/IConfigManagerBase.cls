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
    ' Charge la configuration à partir du chemin spécifié
    ' @param configPath Chemin vers le fichier de configuration
    ' @return True si chargé avec succès, False sinon
End Function

Public Function LoadConfigFromString(ByVal configContent As String) As Boolean
    ' Charge la configuration à partir d'une chaîne
    ' @param configContent Contenu de la configuration
    ' @return True si chargé avec succès, False sinon
End Function

' --------------------------------------------------------
' Value access
' --------------------------------------------------------
Public Function GetValue(ByVal section As String, ByVal key As String, Optional ByVal defaultValue As Variant = Null) As Variant
    ' Récupère une valeur de configuration
    ' @param section Section de configuration
    ' @param key Clé à récupérer
    ' @param defaultValue Valeur par défaut si la clé n'existe pas
    ' @return Valeur associée à la clé ou defaultValue
End Function

Public Function SetValue(ByVal section As String, ByVal key As String, ByVal value As Variant) As Boolean
    ' Définit une valeur de configuration
    ' @param section Section de configuration
    ' @param key Clé à définir
    ' @param value Valeur à associer
    ' @return True si défini avec succès, False sinon
End Function

' --------------------------------------------------------
' Section management
' --------------------------------------------------------
Public Function SectionExists(ByVal section As String) As Boolean
    ' Vérifie si une section existe
    ' @param section Nom de la section
    ' @return True si la section existe, False sinon
End Function

Public Function GetSections() As Variant
    ' Récupère la liste des sections disponibles
    ' @return Tableau des noms de sections
End Function

Public Function GetKeysInSection(ByVal section As String) As Variant
    ' Récupère les clés dans une section
    ' @param section Nom de la section
    ' @return Tableau des clés dans la section
End Function

' --------------------------------------------------------
' Persistence
' --------------------------------------------------------
Public Function SaveConfig(Optional ByVal configPath As String = "") As Boolean
    ' Sauvegarde la configuration actuelle
    ' @param configPath Chemin optionnel pour sauvegarder (si vide, utilise le chemin de chargement)
    ' @return True si sauvegardé avec succès, False sinon
End Function

Public Function ReloadConfig() As Boolean
    ' Recharge la configuration à partir de la dernière source
    ' @return True si rechargé avec succès, False sinon
End Function

' --------------------------------------------------------
' Validation
' --------------------------------------------------------
Public Function ValidateConfig() As Boolean
    ' Valide la structure et le contenu de la configuration
    ' @return True si valide, False sinon
End Function

Public Function GetValidationErrors() As Variant
    ' Récupère les erreurs de validation
    ' @return Tableau des erreurs de validation
End Function 