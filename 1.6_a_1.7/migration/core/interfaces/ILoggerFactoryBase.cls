VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILoggerFactoryBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@ModuleDescription("Interface de factory pour les loggers")
'@Interface
'@Version 1.7.0
'@Author APEX Framework Team
Option Explicit

' =========================================================
' Interface ILoggerFactoryBase
' Factory pour la création et gestion des loggers
' =========================================================

' --------------------------------------------------------
' Factory methods
' --------------------------------------------------------
Public Function CreateLogger(ByVal name As String, Optional ByVal config As Object = Nothing) As ILoggerBase
    ' Crée une instance de logger avec le nom spécifié
    ' @param name Nom du logger
    ' @param config Configuration optionnelle
    ' @return Instance de ILoggerBase
End Function

Public Function GetLogger(ByVal name As String) As ILoggerBase
    ' Récupère une instance existante de logger ou en crée une nouvelle
    ' @param name Nom du logger
    ' @return Instance de ILoggerBase
End Function

Public Function GetDefaultLogger() As ILoggerBase
    ' Récupère le logger par défaut
    ' @return Instance de ILoggerBase par défaut
End Function

Public Function SetDefaultLogger(ByVal name As String) As Boolean
    ' Définit le logger par défaut
    ' @param name Nom du logger à utiliser par défaut
    ' @return True si défini avec succès, False sinon
End Function

' --------------------------------------------------------
' Configuration
' --------------------------------------------------------
Public Function Configure(ByVal config As Object) As Boolean
    ' Configure la factory avec les paramètres donnés
    ' @param config Objet de configuration
    ' @return True si configuré avec succès, False sinon
End Function

Public Function SetGlobalLogLevel(ByVal level As LogLevel) As Boolean
    ' Définit le niveau de log global pour tous les loggers
    ' @param level Niveau de log à appliquer
    ' @return True si défini avec succès, False sinon
End Function

' --------------------------------------------------------
' Logger management
' --------------------------------------------------------
Public Function GetLoggerNames() As Variant
    ' Récupère la liste des noms de loggers disponibles
    ' @return Tableau des noms de loggers
End Function

Public Sub FlushAll()
    ' Vide tous les buffers de tous les loggers
End Sub

Public Function RemoveLogger(ByVal name As String) As Boolean
    ' Supprime un logger
    ' @param name Nom du logger à supprimer
    ' @return True si supprimé avec succès, False sinon
End Function 