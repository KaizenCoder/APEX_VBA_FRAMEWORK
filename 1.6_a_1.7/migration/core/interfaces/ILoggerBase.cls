VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILoggerBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@ModuleDescription("Interface de base pour tous les loggers du framework APEX")
'@Interface
'@Version 1.7.0
'@Author APEX Framework Team
Option Explicit

' =========================================================
' Interface ILoggerBase
' Contrat pour tous les loggers du framework
' =========================================================

'@Description("Niveaux de log standards")
Public Enum LogLevel
    LogTrace = 0     ' Information détaillée pour le débogage
    LogDebug = 1     ' Information de débogage
    LogInfo = 2      ' Information générale
    LogWarning = 3   ' Avertissement
    LogError = 4     ' Erreur récupérable
    LogFatal = 5     ' Erreur critique
End Enum

' --------------------------------------------------------
' Configuration
' --------------------------------------------------------
Public Function Initialize(Optional ByVal config As Object = Nothing) As Boolean
    ' Initialise le logger avec une configuration optionnelle
    ' @param config Objet de configuration (peut être un IConfigManagerBase)
    ' @return True si initialisé avec succès, False sinon
End Function

Public Function SetLogLevel(ByVal level As LogLevel) As Boolean
    ' Définit le niveau de log minimum pour ce logger
    ' @param level Niveau minimum de log à enregistrer
    ' @return True si défini avec succès, False sinon
End Function

Public Function GetLogLevel() As LogLevel
    ' Récupère le niveau de log actuel
    ' @return Niveau de log actuel
End Function

Public Function GetName() As String
    ' Récupère le nom du logger
    ' @return Nom du logger
End Function

' --------------------------------------------------------
' Logging methods
' --------------------------------------------------------
Public Sub LogMessage(ByVal level As LogLevel, ByVal message As String, Optional ByVal source As String = "", Optional ByVal extraData As Variant = Null)
    ' Enregistre un message avec le niveau spécifié
    ' @param level Niveau de log
    ' @param message Message à logger
    ' @param source Source du message (module, classe, etc.)
    ' @param extraData Données supplémentaires (peut être un objet, un tableau, etc.)
End Sub

Public Sub Trace(ByVal message As String, Optional ByVal source As String = "", Optional ByVal extraData As Variant = Null)
    ' Enregistre un message de niveau trace
    ' @param message Message à logger
    ' @param source Source du message
    ' @param extraData Données supplémentaires
End Sub

Public Sub Debug(ByVal message As String, Optional ByVal source As String = "", Optional ByVal extraData As Variant = Null)
    ' Enregistre un message de niveau debug
    ' @param message Message à logger
    ' @param source Source du message
    ' @param extraData Données supplémentaires
End Sub

Public Sub Info(ByVal message As String, Optional ByVal source As String = "", Optional ByVal extraData As Variant = Null)
    ' Enregistre un message de niveau info
    ' @param message Message à logger
    ' @param source Source du message
    ' @param extraData Données supplémentaires
End Sub

Public Sub Warning(ByVal message As String, Optional ByVal source As String = "", Optional ByVal extraData As Variant = Null)
    ' Enregistre un message de niveau warning
    ' @param message Message à logger
    ' @param source Source du message
    ' @param extraData Données supplémentaires
End Sub

Public Sub Error(ByVal message As String, Optional ByVal source As String = "", Optional ByVal extraData As Variant = Null)
    ' Enregistre un message de niveau error
    ' @param message Message à logger
    ' @param source Source du message
    ' @param extraData Données supplémentaires
End Sub

Public Sub Fatal(ByVal message As String, Optional ByVal source As String = "", Optional ByVal extraData As Variant = Null)
    ' Enregistre un message de niveau fatal
    ' @param message Message à logger
    ' @param source Source du message
    ' @param extraData Données supplémentaires
End Sub

Public Sub LogException(ByVal ex As Object, Optional ByVal level As LogLevel = LogError, Optional ByVal additionalInfo As String = "")
    ' Enregistre une exception
    ' @param ex Objet exception
    ' @param level Niveau de log (par défaut Error)
    ' @param additionalInfo Information supplémentaire
End Sub

' --------------------------------------------------------
' Buffer management
' --------------------------------------------------------
Public Sub Flush()
    ' Vide le buffer de logs si applicable
End Sub

Public Function IsEnabled(ByVal level As LogLevel) As Boolean
    ' Vérifie si un niveau de log est activé
    ' @param level Niveau à vérifier
    ' @return True si le niveau est activé, False sinon
End Function 