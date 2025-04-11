' Migrated to apex-metier/orm/interfaces - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit
' ==========================================================================
' Interface : IRelationalObject
' Version : 1.0
' Purpose : Interface pour les objets ORM supportant les relations
' ==========================================================================

' --- Méthodes de relations ---
Public Function GetRelatedObjects(ByVal relationName As String) As Collection: Err.Raise vbObjectError + 1501, "IRelationalObject": End Function
Public Function GetRelatedObject(ByVal relationName As String, Optional ByVal relatedId As Variant) As Object: Err.Raise vbObjectError + 1501, "IRelationalObject": End Function
Public Function AddRelatedObject(ByVal relationName As String, ByVal relatedObject As Object) As Boolean: Err.Raise vbObjectError + 1501, "IRelationalObject": End Function
Public Function RemoveRelatedObject(ByVal relationName As String, ByVal relatedObject As Object) As Boolean: Err.Raise vbObjectError + 1501, "IRelationalObject": End Function
Public Function HasRelatedObjects(ByVal relationName As String) As Boolean: Err.Raise vbObjectError + 1501, "IRelationalObject": End Function

' --- Chargement et sauvegarde en cascade ---
Public Function LoadWithRelations(Optional ByVal depth As Integer = 1) As Boolean: Err.Raise vbObjectError + 1501, "IRelationalObject": End Function
Public Function SaveWithRelations(Optional ByVal cascadingMode As Integer = 1) As Boolean: Err.Raise vbObjectError + 1501, "IRelationalObject": End Function 