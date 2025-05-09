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
' Factory pour la cr�ation et gestion des loggers
' =========================================================

' --------------------------------------------------------
' Factory methods
' --------------------------------------------------------
Public Function CreateLogger(ByVal name As String, Optional ByVal config As Object = Nothing) As ILoggerBase
    ' Cr�e une instance de logger avec le nom sp�cifi�
    ' @param name Nom du logger
    ' @param config Configuration optionnelle
    ' @return Instance de ILoggerBase
End Function

Public Function GetLogger(ByVal name As String) As ILoggerBase
    ' R�cup�re une instance existante de logger ou en cr�e une nouvelle
    ' @param name Nom du logger
    ' @return Instance de ILoggerBase
End Function

Public Function GetDefaultLogger() As ILoggerBase
    ' R�cup�re le logger par d�faut
    ' @return Instance de ILoggerBase par d�faut
End Function

Public Function SetDefaultLogger(ByVal name As String) As Boolean
    ' D�finit le logger par d�faut
    ' @param name Nom du logger � utiliser par d�faut
    ' @return True si d�fini avec succ�s, False sinon
End Function

' --------------------------------------------------------
' Configuration
' --------------------------------------------------------
Public Function Configure(ByVal config As Object) As Boolean
    ' Configure la factory avec les param�tres donn�s
    ' @param config Objet de configuration
    ' @return True si configur� avec succ�s, False sinon
End Function

Public Function SetGlobalLogLevel(ByVal level As LogLevel) As Boolean
    ' D�finit le niveau de log global pour tous les loggers
    ' @param level Niveau de log � appliquer
    ' @return True si d�fini avec succ�s, False sinon
End Function

' --------------------------------------------------------
' Logger management
' --------------------------------------------------------
Public Function GetLoggerNames() As Variant
    ' R�cup�re la liste des noms de loggers disponibles
    ' @return Tableau des noms de loggers
End Function

Public Sub FlushAll()
    ' Vide tous les buffers de tous les loggers
End Sub

Public Function RemoveLogger(ByVal name As String) As Boolean
    ' Supprime un logger
    ' @param name Nom du logger � supprimer
    ' @return True si supprim� avec succ�s, False sinon
End Function 