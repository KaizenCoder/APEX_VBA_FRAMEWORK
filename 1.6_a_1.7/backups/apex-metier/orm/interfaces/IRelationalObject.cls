' Migrated to apex-metier/orm/interfaces - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

' ==========================================================================
' Interface : IRelationalObject
' Version : 1.0
' Purpose : Interface pour les objets ORM supportant les relations
' ==========================================================================

' --- Méthodes de relations ---
'@Description: 
'@Param: 
'@Returns: 

Public Function GetRelatedObjects(ByVal relationName As String) As Collection: Err.Raise vbObjectError + 1501, "IRelationalObject": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function GetRelatedObject(ByVal relationName As String, Optional ByVal relatedId As Variant) As Object: Err.Raise vbObjectError + 1501, "IRelationalObject": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function AddRelatedObject(ByVal relationName As String, ByVal relatedObject As Object) As Boolean: Err.Raise vbObjectError + 1501, "IRelationalObject": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function RemoveRelatedObject(ByVal relationName As String, ByVal relatedObject As Object) As Boolean: Err.Raise vbObjectError + 1501, "IRelationalObject": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function HasRelatedObjects(ByVal relationName As String) As Boolean: Err.Raise vbObjectError + 1501, "IRelationalObject": End Function

' --- Chargement et sauvegarde en cascade ---
'@Description: 
'@Param: 
'@Returns: 

Public Function LoadWithRelations(Optional ByVal depth As Integer = 1) As Boolean: Err.Raise vbObjectError + 1501, "IRelationalObject": End Function
Public Function SaveWithRelations(Optional ByVal cascadingMode As Integer = 1) As Boolean: Err.Raise vbObjectError + 1501, "IRelationalObject": End Function 