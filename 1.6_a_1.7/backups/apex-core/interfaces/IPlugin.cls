' Migrated to apex-core/interfaces - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

' ==========================================================================
' Interface : IPlugin
' Version : 1.0
' Purpose : Interface de base pour tous les plugins du framework
' ==========================================================================

' --- Informations sur le plugin ---
Public Property Get PluginName() As String: Err.Raise vbObjectError + 1001, "IPlugin": End Property
Public Property Get PluginVersion() As String: Err.Raise vbObjectError + 1001, "IPlugin": End Property
Public Property Get PluginDescription() As String: Err.Raise vbObjectError + 1001, "IPlugin": End Property

' --- Méthodes du cycle de vie ---
'@Description: 
'@Param: 
'@Returns: 

Public Function Initialize(Optional ByVal config As Variant) As Boolean: Err.Raise vbObjectError + 1001, "IPlugin": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Sub Terminate(): Err.Raise vbObjectError + 1001, "IPlugin": End Sub

' --- Gestionnaire d'événements ---
Public Function HandleEvent(ByVal eventName As String, ByVal eventData As Variant) As Variant: Err.Raise vbObjectError + 1001, "IPlugin": End Function 