VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IErrorHandlerBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@ModuleDescription("Interface de base pour le gestionnaire d'erreurs")
'@Folder("APEX.Core.Error")
'@Version 1.7.0
'@Author APEX Framework Team
Option Explicit

'==============================================================================
' IErrorHandlerBase
' Interface définissant les méthodes et propriétés du gestionnaire d'erreurs
'==============================================================================

'------------------------------------------------------------------------------
' Initialisation et configuration
'------------------------------------------------------------------------------

' Initialise le gestionnaire d'erreurs avec les paramètres spécifiés
' @param options Options de configuration (logger, niveau de détail, etc.)
' @return True si l'initialisation a réussi, False sinon
Public Function Initialize(Optional ByVal options As Object = Nothing) As Boolean
End Function

'------------------------------------------------------------------------------
' Gestion des erreurs
'------------------------------------------------------------------------------

' Gère une erreur
' @param errNumber Numéro de l'erreur
' @param errDescription Description de l'erreur
' @param errSource Source de l'erreur
' @param errModuleName Module où l'erreur s'est produite
' @param errProcedureName Procédure où l'erreur s'est produite
' @param errLineNumber Numéro de ligne où l'erreur s'est produite (si disponible)
' @param errAdditionalInfo Informations supplémentaires sur l'erreur
' @param errSeverity Niveau de gravité de l'erreur (1=Critique, 2=Erreur, 3=Avertissement, 4=Information)
' @return Code de résultat du traitement de l'erreur
Public Function HandleError(ByVal errNumber As Long, _
                           ByVal errDescription As String, _
                           Optional ByVal errSource As String = "", _
                           Optional ByVal errModuleName As String = "", _
                           Optional ByVal errProcedureName As String = "", _
                           Optional ByVal errLineNumber As Long = 0, _
                           Optional ByVal errAdditionalInfo As String = "", _
                           Optional ByVal errSeverity As Long = 2) As Long
End Function

' Gère l'erreur en cours (Err object)
' @param errSource Source de l'erreur
' @param errModuleName Module où l'erreur s'est produite
' @param errProcedureName Procédure où l'erreur s'est produite
' @param errAdditionalInfo Informations supplémentaires sur l'erreur
' @param errSeverity Niveau de gravité de l'erreur (1=Critique, 2=Erreur, 3=Avertissement, 4=Information)
' @return Code de résultat du traitement de l'erreur
Public Function HandleCurrentError(Optional ByVal errSource As String = "", _
                                  Optional ByVal errModuleName As String = "", _
                                  Optional ByVal errProcedureName As String = "", _
                                  Optional ByVal errAdditionalInfo As String = "", _
                                  Optional ByVal errSeverity As Long = 2) As Long
End Function

' Crée et lève une nouvelle erreur personnalisée
' @param errNumber Numéro de l'erreur
' @param errDescription Description de l'erreur
' @param errSource Source de l'erreur
' @return Toujours False pour faciliter l'utilisation dans les gestionnaires d'erreurs
Public Function RaiseError(ByVal errNumber As Long, _
                          ByVal errDescription As String, _
                          Optional ByVal errSource As String = "") As Boolean
End Function

' Crée et lève une nouvelle erreur personnalisée avec un numéro d'erreur prédéfini
' @param errDescription Description de l'erreur
' @param errSource Source de l'erreur
' @param errType Type d'erreur (voir constantes ERROR_TYPE_*)
' @return Toujours False pour faciliter l'utilisation dans les gestionnaires d'erreurs
Public Function RaiseTypedError(ByVal errDescription As String, _
                              Optional ByVal errSource As String = "", _
                              Optional ByVal errType As Long = 0) As Boolean
End Function

' Définit si l'erreur doit être propagée après traitement
' @param shouldRethrow True si l'erreur doit être propagée, False sinon
Public Property Let RethrowAfterHandling(ByVal shouldRethrow As Boolean)
End Property

' Indique si l'erreur sera propagée après traitement
' @return True si l'erreur sera propagée, False sinon
Public Property Get RethrowAfterHandling() As Boolean
End Property

'------------------------------------------------------------------------------
' Journalisation des erreurs
'------------------------------------------------------------------------------

' Journalise une erreur sans la gérer
' @param errNumber Numéro de l'erreur
' @param errDescription Description de l'erreur
' @param errSource Source de l'erreur
' @param errSeverity Niveau de gravité de l'erreur (1=Critique, 2=Erreur, 3=Avertissement, 4=Information)
' @return True si la journalisation a réussi, False sinon
Public Function LogError(ByVal errNumber As Long, _
                        ByVal errDescription As String, _
                        Optional ByVal errSource As String = "", _
                        Optional ByVal errSeverity As Long = 2) As Boolean
End Function

' Journalise l'erreur en cours (Err object) sans la gérer
' @param errSource Source de l'erreur
' @param errSeverity Niveau de gravité de l'erreur (1=Critique, 2=Erreur, 3=Avertissement, 4=Information)
' @return True si la journalisation a réussi, False sinon
Public Function LogCurrentError(Optional ByVal errSource As String = "", _
                              Optional ByVal errSeverity As Long = 2) As Boolean
End Function

'------------------------------------------------------------------------------
' Interrogation et statistiques
'------------------------------------------------------------------------------

' Récupère le nombre total d'erreurs traitées
' @return Nombre d'erreurs traitées
Public Property Get ErrorCount() As Long
End Property

' Récupère le dernier code d'erreur traité
' @return Dernier code d'erreur
Public Property Get LastErrorCode() As Long
End Property

' Récupère la dernière description d'erreur traitée
' @return Dernière description d'erreur
Public Property Get LastErrorDescription() As String
End Property

' Récupère la dernière source d'erreur traitée
' @return Dernière source d'erreur
Public Property Get LastErrorSource() As String
End Property

' Récupère les statistiques des erreurs traitées
' @return Dictionnaire contenant les statistiques d'erreurs
Public Function GetErrorStats() As Object
End Function

'------------------------------------------------------------------------------
' Notifications et alertes
'------------------------------------------------------------------------------

' Active ou désactive les notifications pour certains types d'erreurs
' @param errTypeOrSeverity Type ou niveau de gravité de l'erreur
' @param enableNotification True pour activer les notifications, False pour les désactiver
Public Sub SetNotificationEnabled(ByVal errTypeOrSeverity As Long, ByVal enableNotification As Boolean)
End Sub

' Indique si les notifications sont activées pour un type d'erreur
' @param errTypeOrSeverity Type ou niveau de gravité de l'erreur
' @return True si les notifications sont activées, False sinon
Public Function IsNotificationEnabled(ByVal errTypeOrSeverity As Long) As Boolean
End Function 