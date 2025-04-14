VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IQueryBuilder"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("Core.Interfaces")
'@ModuleDescription("Interface pour la construction de requêtes SQL")
Option Explicit

' ==========================================================================
' Module    : IQueryBuilder
' Ã‰tat      : Ã€ DÃ‰VELOPPER (Stub)
' Description : Interface pour la construction de requÃªtes SQL.
' Date de crÃ©ation : 11/04/2025
' ==========================================================================

' -- Interface stub --

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Méthodes de construction SELECT
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Select(Optional ByVal columns As String = "*") As IQueryBuilder
End Function

Public Function From(ByVal tableName As String) As IQueryBuilder
End Function

Public Function Join(ByVal tableName As String, ByVal onCondition As String, Optional ByVal joinType As String = "INNER") As IQueryBuilder
End Function

Public Function Where(ByVal condition As String) As IQueryBuilder
End Function

Public Function AndWhere(ByVal condition As String) As IQueryBuilder
End Function

Public Function OrWhere(ByVal condition As String) As IQueryBuilder
End Function

Public Function GroupBy(ByVal columns As String) As IQueryBuilder
End Function

Public Function Having(ByVal condition As String) As IQueryBuilder
End Function

Public Function OrderBy(ByVal columns As String, Optional ByVal direction As String = "ASC") As IQueryBuilder
End Function

Public Function Limit(ByVal count As Long, Optional ByVal offset As Long = 0) As IQueryBuilder
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Méthodes de construction INSERT
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function InsertInto(ByVal tableName As String) As IQueryBuilder
End Function

Public Function Values(ByVal columnValues As Collection) As IQueryBuilder
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Méthodes de construction UPDATE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Update(ByVal tableName As String) As IQueryBuilder
End Function

Public Function Set_(ByVal columnValues As Collection) As IQueryBuilder
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Méthodes de construction DELETE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DeleteFrom(ByVal tableName As String) As IQueryBuilder
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Méthodes utilitaires
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AddParameter(ByVal name As String, ByVal value As Variant, Optional ByVal dataType As Long) As IQueryBuilder
End Function

Public Function GetParameters() As Collection
End Function

Public Function GetSQL() As String
End Function

Public Function Reset() As IQueryBuilder
End Function
