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

'@Description("Niveaux de gravité d'erreur")
Public Enum ErrorSeverity
    SeverityLow = 1       ' Impact minimal, information seulement
    SeverityMedium = 2    ' Avertissement, problème potentiel
    SeverityHigh = 3      ' Erreur récupérable
    SeverityCritical = 4  ' Erreur critique, interruption
    SeverityFatal = 5     ' Erreur fatale, terminaison du programme
End Enum

'@Description("Stratégies de gestion d'erreur")
Public Enum ErrorStrategy
    StrategyContinue = 1  ' Continuer l'exécution, ignorer l'erreur
    StrategyRetry = 2     ' Réessayer l'opération
    StrategyRollback = 3  ' Annuler les changements
    StrategyTerminate = 4 ' Terminer l'exécution
End Enum

' --------------------------------------------------------
' Error handling
' --------------------------------------------------------
Public Function HandleError(ByVal errorNumber As Long, ByVal errorDescription As String, Optional ByVal errorSource As String = "", _
                           Optional ByVal severity As ErrorSeverity = SeverityHigh, Optional ByVal extraData As Variant = Null) As ErrorStrategy
    ' Gère une erreur avec les informations fournies
    ' @param errorNumber Numéro d'erreur
    ' @param errorDescription Description de l'erreur
    ' @param errorSource Source de l'erreur
    ' @param severity Niveau de gravité
    ' @param extraData Données supplémentaires
    ' @return Stratégie à adopter
End Function

Public Function HandleException(ByVal ex As Object, Optional ByVal severity As ErrorSeverity = SeverityHigh, _
                               Optional ByVal additionalInfo As String = "") As ErrorStrategy
    ' Gère une exception
    ' @param ex Objet exception
    ' @param severity Niveau de gravité
    ' @param additionalInfo Information supplémentaire
    ' @return Stratégie à adopter
End Function

' --------------------------------------------------------
' Configuration
' --------------------------------------------------------
Public Function SetDefaultStrategy(ByVal severity As ErrorSeverity, ByVal strategy As ErrorStrategy) As Boolean
    ' Définit la stratégie par défaut pour un niveau de gravité
    ' @param severity Niveau de gravité
    ' @param strategy Stratégie à adopter
    ' @return True si défini avec succès, False sinon
End Function

Public Function SetLogger(ByVal logger As ILoggerBase) As Boolean
    ' Associe un logger au gestionnaire d'erreurs
    ' @param logger Logger à utiliser
    ' @return True si associé avec succès, False sinon
End Function

' --------------------------------------------------------
' Error information
' --------------------------------------------------------
Public Function GetLastError() As Variant
    ' Récupère les informations sur la dernière erreur
    ' @return Structure avec les informations d'erreur
End Function

Public Function GetErrorHistory() As Variant
    ' Récupère l'historique des erreurs
    ' @return Tableau des erreurs précédentes
End Function

Public Function ClearErrorHistory() As Boolean
    ' Efface l'historique des erreurs
    ' @return True si effacé avec succès, False sinon
End Function 