' Migrated to apex-metier/database/interfaces - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

' ==========================================================================
' Interface : IDbDriver
' Version : 2.0
' Purpose : Defines contract for database-specific behaviors and syntax.
' ==========================================================================

' --- Connection String ---'@Description: 
'@Param: 
'@Returns: 


Function GetConnectionString(server As String, dbName As String, Optional user As String = "", Optional pwd As String = "", Optional useWindowsAuth As Boolean = True) As String: Err.Raise vbObjectError + 1301, "IDbDriver": End Function

' --- Parameter Formatting ---
Property Get ParameterMarker() As String: Err.Raise vbObjectError + 1301, "IDbDriver": End Property'@Description: 
'@Param: 
'@Returns: 


Function FormatParameterName(baseName As String) As String: Err.Raise vbObjectError + 1301, "IDbDriver": End Function
Function FormatSqlDate(dtValue As Date) As String: Err.Raise vbObjectError + 1301, "IDbDriver": End'@Description: 
'@Param: 
'@Returns: 

 Function
Function FormatSqlString(value As String) As String: Err.Raise vbObjectError + 1301, "IDbDriver": End Function ' Nouveau
Function FormatSqlNumber(value As Variant) As String: Err.Raise vbObjectError + 1301, "IDbDriver": End Function ' Nouveau'@Description: 
'@Param: 
'@Returns: 


Function EscapeIdentifier(identifier As String) As String: Err.Raise vbObjectError + 1301, "IDbDriver": End Function ' Nouveau

' --- SQL Specific Clauses ---'@Description: 
'@Param: 
'@Returns: 


Function GetTopClause(n As Long) As String: Err.Raise vbObjectError + 1301, "IDbDriver": End Function
Function GetLimitClause(n As Long) As String: Err.Raise vbObjectError + 1301, "IDbDriver": End'@Description: 
'@Param: 
'@Returns: 

 Function
Function GetLastIdentitySql() As String: Err.Raise vbObjectError + 1301, "IDbDriver": End Function
Function GetPaginationSql(sql As String, pageSize As Long, pageNumber As Long) As String: Err.Raise vbObjectError + 1301, "IDbDriver": End Function ' Nouveau

' --- Batch Processing ---'@Description: 
'@Param: 
'@Returns: 


Function SupportsBatchOperations() As Boolean: Err.Raise vbObjectError + 1301, "IDbDriver": End Function ' Nouveau
Function GetBatchInsertSql(tableName As String, columnNames() As String, rowCount As Long) As String: Err.Raise vbObjectError + 1301, "IDbDriver": End Function ' Nouveau

' --- Schema Information ---'@Description: 
'@Param: 
'@Returns: 


Function GetTableColumnsSql(tableName As String) As String: Err.Raise vbObjectError + 1301, "IDbDriver": End Function ' Nouveau
Function GetTableListSql() As String: Err.Raise vbObjectError + 1301, "IDbDriver": End Function ' Nouveau'@Description: 
'@Param: 
'@Returns: 


Function GetPrimaryKeysSql(tableName As String) As String: Err.Raise vbObjectError + 1301, "IDbDriver": End Function ' Nouveau
Function GetForeignKeysSql(tableName As String) As String: Err.Raise vbObjectError + 1301, "IDbDriver": End Function ' Nouveau

' --- Advanced Features ---'@Description: 
'@Param: 
'@Returns: 


Function SupportsTransactionIsolationLevel() As Boolean: Err.Raise vbObjectError + 1301, "IDbDriver": End Function ' Nouveau
Function GetSetIsolationLevelSql(level As Long) As String: Err.Raise vbObjectError + 1301, "IDbDriver": End Function ' Nouveau'@Description: 
'@Param: 
'@Returns: 


Function SupportsStoredProcedures() As Boolean: Err.Raise vbObjectError + 1301, "IDbDriver": End Function ' Nouveau
Function SupportsUnions() As Boolean: Err.Raise vbObjectError + 1301, "IDbDriver": End Function ' Nouveau 