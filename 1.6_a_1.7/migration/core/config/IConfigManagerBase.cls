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
' Interface d�finissant les m�thodes et propri�t�s du gestionnaire de configuration
'==============================================================================

'------------------------------------------------------------------------------
' Initialisation et configuration
'------------------------------------------------------------------------------

' Initialise le gestionnaire de configuration avec les param�tres sp�cifi�s
' @param configSource Source de configuration (chemin fichier, objet, etc.)
' @param options Options suppl�mentaires (optionnel)
' @return True si l'initialisation a r�ussi, False sinon
Public Function Initialize(ByVal configSource As Variant, Optional ByVal options As Object = Nothing) As Boolean
End Function

'------------------------------------------------------------------------------
' Acc�s aux valeurs
'------------------------------------------------------------------------------

' R�cup�re une valeur par sa cl�
' @param key Cl� de la valeur � r�cup�rer
' @param defaultValue Valeur par d�faut si la cl� n'existe pas
' @return La valeur associ�e � la cl� ou la valeur par d�faut
Public Function GetValue(ByVal key As String, Optional ByVal defaultValue As Variant = Null) As Variant
End Function

' R�cup�re une valeur typ�e (String) par sa cl�
' @param key Cl� de la valeur � r�cup�rer
' @param defaultValue Valeur par d�faut si la cl� n'existe pas
' @return La valeur associ�e � la cl� ou la valeur par d�faut
Public Function GetString(ByVal key As String, Optional ByVal defaultValue As String = "") As String
End Function

' R�cup�re une valeur typ�e (Integer) par sa cl�
' @param key Cl� de la valeur � r�cup�rer
' @param defaultValue Valeur par d�faut si la cl� n'existe pas
' @return La valeur associ�e � la cl� ou la valeur par d�faut
Public Function GetInteger(ByVal key As String, Optional ByVal defaultValue As Long = 0) As Long
End Function

' R�cup�re une valeur typ�e (Boolean) par sa cl�
' @param key Cl� de la valeur � r�cup�rer
' @param defaultValue Valeur par d�faut si la cl� n'existe pas
' @return La valeur associ�e � la cl� ou la valeur par d�faut
Public Function GetBoolean(ByVal key As String, Optional ByVal defaultValue As Boolean = False) As Boolean
End Function

' R�cup�re une valeur typ�e (Double) par sa cl�
' @param key Cl� de la valeur � r�cup�rer
' @param defaultValue Valeur par d�faut si la cl� n'existe pas
' @return La valeur associ�e � la cl� ou la valeur par d�faut
Public Function GetDouble(ByVal key As String, Optional ByVal defaultValue As Double = 0#) As Double
End Function

' R�cup�re une valeur typ�e (Date) par sa cl�
' @param key Cl� de la valeur � r�cup�rer
' @param defaultValue Valeur par d�faut si la cl� n'existe pas
' @return La valeur associ�e � la cl� ou la valeur par d�faut
Public Function GetDate(ByVal key As String, Optional ByVal defaultValue As Date = #1/1/1900#) As Date
End Function

' R�cup�re une valeur typ�e (Object) par sa cl�
' @param key Cl� de la valeur � r�cup�rer
' @return L'objet associ� � la cl� ou Nothing
Public Function GetObject(ByVal key As String) As Object
End Function

' V�rifie si une cl� existe
' @param key Cl� � v�rifier
' @return True si la cl� existe, False sinon
Public Function HasKey(ByVal key As String) As Boolean
End Function

'------------------------------------------------------------------------------
' Manipulation des valeurs
'------------------------------------------------------------------------------

' D�finit ou met � jour une valeur
' @param key Cl� de la valeur � d�finir
' @param value Valeur � associer � la cl�
' @return True si la valeur a �t� d�finie avec succ�s, False sinon
Public Function SetValue(ByVal key As String, ByVal value As Variant) As Boolean
End Function

' Supprime une valeur
' @param key Cl� de la valeur � supprimer
' @return True si la valeur a �t� supprim�e avec succ�s, False sinon
Public Function RemoveValue(ByVal key As String) As Boolean
End Function

' Efface toutes les valeurs
' @return True si les valeurs ont �t� effac�es avec succ�s, False sinon
Public Function Clear() As Boolean
End Function

'------------------------------------------------------------------------------
' Gestion des sections
'------------------------------------------------------------------------------

' R�cup�re un dictionnaire contenant toutes les valeurs d'une section
' @param section Nom de la section
' @return Dictionnaire contenant les cl�s/valeurs de la section
Public Function GetSection(ByVal section As String) As Object
End Function

' V�rifie si une section existe
' @param section Nom de la section � v�rifier
' @return True si la section existe, False sinon
Public Function HasSection(ByVal section As String) As Boolean
End Function

' R�cup�re la liste des sections disponibles
' @return Collection de noms de sections
Public Function GetSectionNames() As Collection
End Function

'------------------------------------------------------------------------------
' Persistance
'------------------------------------------------------------------------------

' Sauvegarde la configuration
' @param destination Destination (optionnel, si diff�rente de la source)
' @return True si la sauvegarde a r�ussi, False sinon
Public Function Save(Optional ByVal destination As Variant = Null) As Boolean
End Function

' Recharge la configuration depuis la source
' @return True si le rechargement a r�ussi, False sinon
Public Function Reload() As Boolean
End Function

'------------------------------------------------------------------------------
' Propri�t�s
'------------------------------------------------------------------------------

' R�cup�re le chemin ou la source de la configuration
Public Property Get Source() As Variant
End Property

' R�cup�re l'�tat d'initialisation
Public Property Get IsInitialized() As Boolean
End Property

' R�cup�re ou d�finit si les erreurs doivent �tre ignor�es
Public Property Get IgnoreErrors() As Boolean
End Property
Public Property Let IgnoreErrors(ByVal value As Boolean)
End Property 