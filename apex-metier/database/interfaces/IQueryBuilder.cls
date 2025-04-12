' Migrated to apex-metier/database/interfaces - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

' ==========================================================================
' Interface : IQueryBuilder
' Version : 4.0
' Purpose : Contract for Advanced SQL Query Builder facade (SELECT focus).
' ==========================================================================

' --- Core Clauses ---
'@Description: 
'@Param: 
'@Returns: 

Public Function SelectColumns(ByVal columns As String) As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function FromTable(ByVal tableName As String) As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function

' --- WHERE clause enhancements ---
'@Description: 
'@Param: 
'@Returns: 

Public Function AddWhere(ByVal field As String, ByVal operator As String, ByVal value As Variant, Optional ByVal paramType As ADODB.DataTypeEnum = adVarWChar, Optional ByVal paramSize As Long = 0) As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function AddWhereClause(ByVal whereClause As String) As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function AddWhereIn(ByVal field As String, ByVal values As Variant) As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function AddWhereBetween(ByVal field As String, ByVal value1 As Variant, ByVal value2 As Variant, Optional ByVal paramType As ADODB.DataTypeEnum = adVarWChar) As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function AddWhereLike(ByVal field As String, ByVal pattern As String, Optional ByVal escapeChar As String = "") As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function AddWhereIsNull(ByVal field As String, Optional ByVal isNotNull As Boolean = False) As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function

' --- Logical Grouping ---
'@Description: 
'@Param: 
'@Returns: 

Public Function OpenGroup() As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function CloseGroup() As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function AddAnd() As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function AddOr() As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function AddOrGroup() As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function AddAndGroup() As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function

' --- Other clauses ---
'@Description: 
'@Param: 
'@Returns: 

Public Function AddAlias(tableAlias As String) As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function ' NYI
'@Description: 
'@Param: 
'@Returns: 

Public Function UseSafeIdentifiers(Optional ByVal useSafe As Boolean = True) As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function Join(ByVal joinTable As String, ByVal onClause As String, Optional ByVal joinType As String = "INNER") As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function GroupBy(ByVal columns As String) As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function OrderBy(ByVal columns As String, Optional ByVal descending As Boolean = False) As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function
'@Description: 
'@Param: 
'@Returns: 

Public Function Top(ByVal n As Long) As IQueryBuilder: Err.Raise 513, "IQueryBuilder": End Function

' --- Build & Debug ---
'@Description: 
'@Param: 
'@Returns: 

Public Function Build() As Variant: Err.Raise 513, "IQueryBuilder": End Function ' Array(SQLString, ParametersArray)
'@Description: 
'@Param: 
'@Returns: 

Public Function ToSqlString() As String: Err.Raise 513, "IQueryBuilder": End Function
Public Function ToDebugString() As String: Err.Raise 513, "IQueryBuilder": End Function 