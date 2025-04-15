VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IConfigManagerBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@ModuleDescription("Interface de base pour le gestionnaire de configuration")
'@Folder("APEX.Core.Config")
'@Version 1.7.0
'@Author APEX Framework Team
Option Explicit

'==============================================================================
' IConfigManagerBase
' Interface définissant les méthodes et propriétés du gestionnaire de configuration
'==============================================================================

'------------------------------------------------------------------------------
' Initialisation et configuration
'------------------------------------------------------------------------------

' Initialise le gestionnaire de configuration avec les paramètres spécifiés
' @param configSource Source de configuration (chemin fichier, objet, etc.)
' @param options Options supplémentaires (optionnel)
' @return True si l'initialisation a réussi, False sinon
Public Function Initialize(ByVal configSource As Variant, Optional ByVal options As Object = Nothing) As Boolean
End Function

'------------------------------------------------------------------------------
' Accès aux valeurs
'------------------------------------------------------------------------------

' Récupère une valeur par sa clé
' @param key Clé de la valeur à récupérer
' @param defaultValue Valeur par défaut si la clé n'existe pas
' @return La valeur associée à la clé ou la valeur par défaut
Public Function GetValue(ByVal key As String, Optional ByVal defaultValue As Variant = Null) As Variant
End Function

' Récupère une valeur typée (String) par sa clé
' @param key Clé de la valeur à récupérer
' @param defaultValue Valeur par défaut si la clé n'existe pas
' @return La valeur associée à la clé ou la valeur par défaut
Public Function GetString(ByVal key As String, Optional ByVal defaultValue As String = "") As String
End Function

' Récupère une valeur typée (Integer) par sa clé
' @param key Clé de la valeur à récupérer
' @param defaultValue Valeur par défaut si la clé n'existe pas
' @return La valeur associée à la clé ou la valeur par défaut
Public Function GetInteger(ByVal key As String, Optional ByVal defaultValue As Long = 0) As Long
End Function

' Récupère une valeur typée (Boolean) par sa clé
' @param key Clé de la valeur à récupérer
' @param defaultValue Valeur par défaut si la clé n'existe pas
' @return La valeur associée à la clé ou la valeur par défaut
Public Function GetBoolean(ByVal key As String, Optional ByVal defaultValue As Boolean = False) As Boolean
End Function

' Récupère une valeur typée (Double) par sa clé
' @param key Clé de la valeur à récupérer
' @param defaultValue Valeur par défaut si la clé n'existe pas
' @return La valeur associée à la clé ou la valeur par défaut
Public Function GetDouble(ByVal key As String, Optional ByVal defaultValue As Double = 0#) As Double
End Function

' Récupère une valeur typée (Date) par sa clé
' @param key Clé de la valeur à récupérer
' @param defaultValue Valeur par défaut si la clé n'existe pas
' @return La valeur associée à la clé ou la valeur par défaut
Public Function GetDate(ByVal key As String, Optional ByVal defaultValue As Date = #1/1/1900#) As Date
End Function

' Récupère une valeur typée (Object) par sa clé
' @param key Clé de la valeur à récupérer
' @return L'objet associé à la clé ou Nothing
Public Function GetObject(ByVal key As String) As Object
End Function

' Vérifie si une clé existe
' @param key Clé à vérifier
' @return True si la clé existe, False sinon
Public Function HasKey(ByVal key As String) As Boolean
End Function

'------------------------------------------------------------------------------
' Manipulation des valeurs
'------------------------------------------------------------------------------

' Définit ou met à jour une valeur
' @param key Clé de la valeur à définir
' @param value Valeur à associer à la clé
' @return True si la valeur a été définie avec succès, False sinon
Public Function SetValue(ByVal key As String, ByVal value As Variant) As Boolean
End Function

' Supprime une valeur
' @param key Clé de la valeur à supprimer
' @return True si la valeur a été supprimée avec succès, False sinon
Public Function RemoveValue(ByVal key As String) As Boolean
End Function

' Efface toutes les valeurs
' @return True si les valeurs ont été effacées avec succès, False sinon
Public Function Clear() As Boolean
End Function

'------------------------------------------------------------------------------
' Gestion des sections
'------------------------------------------------------------------------------

' Récupère un dictionnaire contenant toutes les valeurs d'une section
' @param section Nom de la section
' @return Dictionnaire contenant les clés/valeurs de la section
Public Function GetSection(ByVal section As String) As Object
End Function

' Vérifie si une section existe
' @param section Nom de la section à vérifier
' @return True si la section existe, False sinon
Public Function HasSection(ByVal section As String) As Boolean
End Function

' Récupère la liste des sections disponibles
' @return Collection de noms de sections
Public Function GetSectionNames() As Collection
End Function

'------------------------------------------------------------------------------
' Persistance
'------------------------------------------------------------------------------

' Sauvegarde la configuration
' @param destination Destination (optionnel, si différente de la source)
' @return True si la sauvegarde a réussi, False sinon
Public Function Save(Optional ByVal destination As Variant = Null) As Boolean
End Function

' Recharge la configuration depuis la source
' @return True si le rechargement a réussi, False sinon
Public Function Reload() As Boolean
End Function

'------------------------------------------------------------------------------
' Propriétés
'------------------------------------------------------------------------------

' Récupère le chemin ou la source de la configuration
Public Property Get Source() As Variant
End Property

' Récupère l'état d'initialisation
Public Property Get IsInitialized() As Boolean
End Property

' Récupère ou définit si les erreurs doivent être ignorées
Public Property Get IgnoreErrors() As Boolean
End Property
Public Property Let IgnoreErrors(ByVal value As Boolean)
End Property 