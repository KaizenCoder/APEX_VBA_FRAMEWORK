VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IErrorHandlerBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@ModuleDescription("Interface de base pour les gestionnaires d'erreurs")
'@Interface
'@Version 1.7.0
'@Author APEX Framework Team
Option Explicit

' =========================================================
' Interface IErrorHandlerBase
' Contrat pour tous les gestionnaires d'erreurs
' =========================================================

'@Description("Niveaux de gravit� d'erreur")
Public Enum ErrorSeverity
    SeverityLow = 1       ' Impact minimal, information seulement
    SeverityMedium = 2    ' Avertissement, probl�me potentiel
    SeverityHigh = 3      ' Erreur r�cup�rable
    SeverityCritical = 4  ' Erreur critique, interruption
    SeverityFatal = 5     ' Erreur fatale, terminaison du programme
End Enum

'@Description("Strat�gies de gestion d'erreur")
Public Enum ErrorStrategy
    StrategyContinue = 1  ' Continuer l'ex�cution, ignorer l'erreur
    StrategyRetry = 2     ' R�essayer l'op�ration
    StrategyRollback = 3  ' Annuler les changements
    StrategyTerminate = 4 ' Terminer l'ex�cution
End Enum

' --------------------------------------------------------
' Error handling
' --------------------------------------------------------
Public Function HandleError(ByVal errorNumber As Long, ByVal errorDescription As String, Optional ByVal errorSource As String = "", _
                           Optional ByVal severity As ErrorSeverity = SeverityHigh, Optional ByVal extraData As Variant = Null) As ErrorStrategy
    ' G�re une erreur avec les informations fournies
    ' @param errorNumber Num�ro d'erreur
    ' @param errorDescription Description de l'erreur
    ' @param errorSource Source de l'erreur
    ' @param severity Niveau de gravit�
    ' @param extraData Donn�es suppl�mentaires
    ' @return Strat�gie � adopter
End Function

Public Function HandleException(ByVal ex As Object, Optional ByVal severity As ErrorSeverity = SeverityHigh, _
                               Optional ByVal additionalInfo As String = "") As ErrorStrategy
    ' G�re une exception
    ' @param ex Objet exception
    ' @param severity Niveau de gravit�
    ' @param additionalInfo Information suppl�mentaire
    ' @return Strat�gie � adopter
End Function

' --------------------------------------------------------
' Configuration
' --------------------------------------------------------
Public Function SetDefaultStrategy(ByVal severity As ErrorSeverity, ByVal strategy As ErrorStrategy) As Boolean
    ' D�finit la strat�gie par d�faut pour un niveau de gravit�
    ' @param severity Niveau de gravit�
    ' @param strategy Strat�gie � adopter
    ' @return True si d�fini avec succ�s, False sinon
End Function

Public Function SetLogger(ByVal logger As ILoggerBase) As Boolean
    ' Associe un logger au gestionnaire d'erreurs
    ' @param logger Logger � utiliser
    ' @return True si associ� avec succ�s, False sinon
End Function

' --------------------------------------------------------
' Error information
' --------------------------------------------------------
Public Function GetLastError() As Variant
    ' R�cup�re les informations sur la derni�re erreur
    ' @return Structure avec les informations d'erreur
End Function

Public Function GetErrorHistory() As Variant
    ' R�cup�re l'historique des erreurs
    ' @return Tableau des erreurs pr�c�dentes
End Function

Public Function ClearErrorHistory() As Boolean
    ' Efface l'historique des erreurs
    ' @return True si effac� avec succ�s, False sinon
End Function 