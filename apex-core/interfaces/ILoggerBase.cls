VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILoggerBase"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
'@Folder("APEX.Core.Interfaces.Logging")
Option Explicit

' ==========================================================================
' Interface  : ILoggerBase
' Purpose   : Interface de base pour tous les loggers du framework APEX
' Author    : APEX Framework Team
' Date      : 2024-04-11
' ==========================================================================

'@Description("Niveaux de log standards")
Public Enum LogLevel
    LogDebug = 0     ' Messages de débogage détaillés
    LogInfo = 1      ' Informations générales sur le fonctionnement
    LogWarning = 2   ' Avertissements non critiques
    LogError = 3     ' Erreurs récupérables
    LogCritical = 4  ' Erreurs critiques nécessitant une attention immédiate
End Enum

'@Description("Enregistre un message avec un niveau et une catégorie spécifiés")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Log(ByVal level As LogLevel, ByVal message As String, Optional ByVal category As String = "")
End Sub

'@Description("Enregistre un message de niveau Debug")
'@Description: 
'@Param: 
'@Returns: 

Public Sub LogDebug(ByVal message As String, Optional ByVal category As String = "")
End Sub

'@Description("Enregistre un message de niveau Info")
'@Description: 
'@Param: 
'@Returns: 

Public Sub LogInfo(ByVal message As String, Optional ByVal category As String = "")
End Sub

'@Description("Enregistre un message de niveau Warning")
'@Description: 
'@Param: 
'@Returns: 

Public Sub LogWarning(ByVal message As String, Optional ByVal category As String = "")
End Sub

'@Description("Enregistre un message de niveau Error")
'@Description: 
'@Param: 
'@Returns: 

Public Sub LogError(ByVal message As String, Optional ByVal category As String = "")
End Sub

'@Description("Enregistre un message de niveau Critical")
'@Description: 
'@Param: 
'@Returns: 

Public Sub LogCritical(ByVal message As String, Optional ByVal category As String = "")
End Sub

'@Description("Configure le logger avec des options spécifiques")
'@Param options Un dictionnaire contenant les options de configuration
' Options possibles :
' - "minimumLevel" : LogLevel - Niveau minimum des messages à logger
' - "includeTimestamp" : Boolean - Inclure l'horodatage dans les messages
' - "includeCategory" : Boolean - Inclure la catégorie dans les messages
' - "format" : String - Format personnalisé des messages
' - "outputPath" : String - Chemin de sortie pour les loggers de fichier
' - "maxSize" : Long - Taille maximale du fichier/buffer de log
' - "rotateCount" : Long - Nombre de fichiers de rotation à conserver
'@Description: 
'@Param: 
'@Returns: 

Public Sub Configure(ByVal options As Variant)
End Sub

'@Description("Vide le buffer de logs si applicable")
'@Remarks Cette méthode doit être appelée pour s'assurer que tous les logs sont écrits,
' particulièrement important pour les loggers avec buffer ou les loggers de fichier
'@Description: 
'@Param: 
'@Returns: 

Public Sub Flush()
End Sub 