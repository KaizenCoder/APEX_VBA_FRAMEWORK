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
' Interface d�finissant les m�thodes et propri�t�s du gestionnaire d'erreurs
'==============================================================================

'------------------------------------------------------------------------------
' Initialisation et configuration
'------------------------------------------------------------------------------

' Initialise le gestionnaire d'erreurs avec les param�tres sp�cifi�s
' @param options Options de configuration (logger, niveau de d�tail, etc.)
' @return True si l'initialisation a r�ussi, False sinon
Public Function Initialize(Optional ByVal options As Object = Nothing) As Boolean
End Function

'------------------------------------------------------------------------------
' Gestion des erreurs
'------------------------------------------------------------------------------

' G�re une erreur
' @param errNumber Num�ro de l'erreur
' @param errDescription Description de l'erreur
' @param errSource Source de l'erreur
' @param errModuleName Module o� l'erreur s'est produite
' @param errProcedureName Proc�dure o� l'erreur s'est produite
' @param errLineNumber Num�ro de ligne o� l'erreur s'est produite (si disponible)
' @param errAdditionalInfo Informations suppl�mentaires sur l'erreur
' @param errSeverity Niveau de gravit� de l'erreur (1=Critique, 2=Erreur, 3=Avertissement, 4=Information)
' @return Code de r�sultat du traitement de l'erreur
Public Function HandleError(ByVal errNumber As Long, _
                           ByVal errDescription As String, _
                           Optional ByVal errSource As String = "", _
                           Optional ByVal errModuleName As String = "", _
                           Optional ByVal errProcedureName As String = "", _
                           Optional ByVal errLineNumber As Long = 0, _
                           Optional ByVal errAdditionalInfo As String = "", _
                           Optional ByVal errSeverity As Long = 2) As Long
End Function

' G�re l'erreur en cours (Err object)
' @param errSource Source de l'erreur
' @param errModuleName Module o� l'erreur s'est produite
' @param errProcedureName Proc�dure o� l'erreur s'est produite
' @param errAdditionalInfo Informations suppl�mentaires sur l'erreur
' @param errSeverity Niveau de gravit� de l'erreur (1=Critique, 2=Erreur, 3=Avertissement, 4=Information)
' @return Code de r�sultat du traitement de l'erreur
Public Function HandleCurrentError(Optional ByVal errSource As String = "", _
                                  Optional ByVal errModuleName As String = "", _
                                  Optional ByVal errProcedureName As String = "", _
                                  Optional ByVal errAdditionalInfo As String = "", _
                                  Optional ByVal errSeverity As Long = 2) As Long
End Function

' Cr�e et l�ve une nouvelle erreur personnalis�e
' @param errNumber Num�ro de l'erreur
' @param errDescription Description de l'erreur
' @param errSource Source de l'erreur
' @return Toujours False pour faciliter l'utilisation dans les gestionnaires d'erreurs
Public Function RaiseError(ByVal errNumber As Long, _
                          ByVal errDescription As String, _
                          Optional ByVal errSource As String = "") As Boolean
End Function

' Cr�e et l�ve une nouvelle erreur personnalis�e avec un num�ro d'erreur pr�d�fini
' @param errDescription Description de l'erreur
' @param errSource Source de l'erreur
' @param errType Type d'erreur (voir constantes ERROR_TYPE_*)
' @return Toujours False pour faciliter l'utilisation dans les gestionnaires d'erreurs
Public Function RaiseTypedError(ByVal errDescription As String, _
                              Optional ByVal errSource As String = "", _
                              Optional ByVal errType As Long = 0) As Boolean
End Function

' D�finit si l'erreur doit �tre propag�e apr�s traitement
' @param shouldRethrow True si l'erreur doit �tre propag�e, False sinon
Public Property Let RethrowAfterHandling(ByVal shouldRethrow As Boolean)
End Property

' Indique si l'erreur sera propag�e apr�s traitement
' @return True si l'erreur sera propag�e, False sinon
Public Property Get RethrowAfterHandling() As Boolean
End Property

'------------------------------------------------------------------------------
' Journalisation des erreurs
'------------------------------------------------------------------------------

' Journalise une erreur sans la g�rer
' @param errNumber Num�ro de l'erreur
' @param errDescription Description de l'erreur
' @param errSource Source de l'erreur
' @param errSeverity Niveau de gravit� de l'erreur (1=Critique, 2=Erreur, 3=Avertissement, 4=Information)
' @return True si la journalisation a r�ussi, False sinon
Public Function LogError(ByVal errNumber As Long, _
                        ByVal errDescription As String, _
                        Optional ByVal errSource As String = "", _
                        Optional ByVal errSeverity As Long = 2) As Boolean
End Function

' Journalise l'erreur en cours (Err object) sans la g�rer
' @param errSource Source de l'erreur
' @param errSeverity Niveau de gravit� de l'erreur (1=Critique, 2=Erreur, 3=Avertissement, 4=Information)
' @return True si la journalisation a r�ussi, False sinon
Public Function LogCurrentError(Optional ByVal errSource As String = "", _
                              Optional ByVal errSeverity As Long = 2) As Boolean
End Function

'------------------------------------------------------------------------------
' Interrogation et statistiques
'------------------------------------------------------------------------------

' R�cup�re le nombre total d'erreurs trait�es
' @return Nombre d'erreurs trait�es
Public Property Get ErrorCount() As Long
End Property

' R�cup�re le dernier code d'erreur trait�
' @return Dernier code d'erreur
Public Property Get LastErrorCode() As Long
End Property

' R�cup�re la derni�re description d'erreur trait�e
' @return Derni�re description d'erreur
Public Property Get LastErrorDescription() As String
End Property

' R�cup�re la derni�re source d'erreur trait�e
' @return Derni�re source d'erreur
Public Property Get LastErrorSource() As String
End Property

' R�cup�re les statistiques des erreurs trait�es
' @return Dictionnaire contenant les statistiques d'erreurs
Public Function GetErrorStats() As Object
End Function

'------------------------------------------------------------------------------
' Notifications et alertes
'------------------------------------------------------------------------------

' Active ou d�sactive les notifications pour certains types d'erreurs
' @param errTypeOrSeverity Type ou niveau de gravit� de l'erreur
' @param enableNotification True pour activer les notifications, False pour les d�sactiver
Public Sub SetNotificationEnabled(ByVal errTypeOrSeverity As Long, ByVal enableNotification As Boolean)
End Sub

' Indique si les notifications sont activ�es pour un type d'erreur
' @param errTypeOrSeverity Type ou niveau de gravit� de l'erreur
' @return True si les notifications sont activ�es, False sinon
Public Function IsNotificationEnabled(ByVal errTypeOrSeverity As Long) As Boolean
End Function 