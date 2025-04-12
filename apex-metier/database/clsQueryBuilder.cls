' Migrated to apex-metier/database - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

' ==========================================================================
' Class : clsQueryBuilder
' Version : 5.0 - Phase 3 - Integrated with IDbDriver
' Implements: IQueryBuilder (v4.0)
' Purpose : Advanced SQL Query Builder: Grouping, Safe IDs, Debug String, Driver Integration.
' Requires : modFrameworkUtils, IDbDriver
' ==========================================================================
Implements IQueryBuilder

' --- Member Variables ---
Private m_Select As String
Private m_From As String
Private m_Where As Collection ' Stores Strings: "(", ")", "AND", "OR" or Arrays: {Field, Op, Param/Value/Null}
Private m_Params As Collection ' Stores parameter arrays {ParamName, Value, Type, Size, Direction}
Private m_ParamIndex As Long
Private m_Joins As Collection
Private m_GroupBy As String
Private m_OrderBy As Collection ' Stores order by clauses {Array(Column, Direction)}
Private m_TopCount As Long
Private m_UseSafeIdentifiers As Boolean
Private m_SafeIdPrefix As String
Private m_SafeIdSuffix As String
Private m_LastError As Long ' Store last error number
Private m_Driver As IDbDriver ' Référence au driver de base de données
Private m_Aliases As Object ' Scripting.Dictionary - Stores table aliases
Private m_PageSize As Long ' Pour la pagination
Private m_PageNumber As Long ' Pour la pagination

' --- Constants ---
Private Const MAX_SQL_LENGTH As Long = 8000 ' Maximum safe SQL length
Private Const MAX_PARAMS As Long = 2000 ' Maximum safe number of parameters
Private Const INVALID_SQL_CHARS As String = "';""\%" ' Characters to check for in direct SQL injection

' --- Initialize ---
Private Sub Class_Initialize()
    m_Select = "*"
    Set m_Where = New Collection
    Set m_Params = New Collection
    Set m_Joins = New Collection
    Set m_OrderBy = New Collection
    Set m_Aliases = CreateObject("Scripting.Dictionary")
    m_ParamIndex = 0
    m_TopCount = 0
    m_UseSafeIdentifiers = False
    m_SafeIdPrefix = ""
    m_SafeIdSuffix = ""
    m_LastError = 0
    m_PageSize = 0
    m_PageNumber = 0
End Sub

' --- Driver Methods ---
'@Description: 
'@Param: 
'@Returns: 

Public Sub SetDriver(driver As IDbDriver)
    Set m_Driver = driver
End Sub

' --- Helper Methods for Aliases ---
'@Description: 
'@Param: 
'@Returns: 

Private Function GetTableAlias(tableName As String) As String
    If m_Aliases.Exists(tableName) Then
        GetTableAlias = m_Aliases(tableName)
    Else
        GetTableAlias = tableName
    End If
End Function

' --- IQueryBuilder Implementation ---
'@Description: 
'@Param: 
'@Returns: 

Private Function IQueryBuilder_SelectColumns(ByVal columns As String) As IQueryBuilder
    ' Validate input
    If Len(Trim(columns)) = 0 Then
        m_Select = "*"
    Else
        ' Validate for basic SQL injection
        If ContainsSqlInjection(columns) Then
            Debug.Print "WARNING: Potentially unsafe SQL in columns: " & columns
            columns = SafeguardSqlInput(columns)
        End If
        
        m_Select = columns
    End If
    
    Set IQueryBuilder_SelectColumns = Me
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_FromTable(ByVal tableName As String) As IQueryBuilder
    ' Validate input
    If Len(Trim(tableName)) = 0 Then
        ' Cannot build valid SQL without a table, use a dummy table
        Debug.Print "ERROR: Empty table name provided to FromTable"
        m_From = "DUAL" ' Oracle-style dummy table, will fail on other DBs
    Else
        ' Validate for basic SQL injection
        If ContainsSqlInjection(tableName) Then
            Debug.Print "WARNING: Potentially unsafe SQL in table name: " & tableName
            tableName = SafeguardSqlInput(tableName)
        End If
        
        m_From = tableName
    End If
    
    Set IQueryBuilder_FromTable = Me
End Function

' --- New Pagination Methods ---
'@Description: 
'@Param: 
'@Returns: 

Public Function SetPagination(ByVal pageSize As Long, ByVal pageNumber As Long) As IQueryBuilder
    If pageSize > 0 And pageNumber > 0 Then
        m_PageSize = pageSize
        m_PageNumber = pageNumber
    Else
        m_PageSize = 0
        m_PageNumber = 0
    End If
    
    Set SetPagination = Me
End Function

' --- WHERE Clauses ---
' Helper function to add logical operator (AND, OR) when needed
'@Description: 
'@Param: 
'@Returns: 

Private Function AddLogicalOperator(op As String) As IQueryBuilder
    If m_Where.Count > 0 Then
        Dim lastItem As Variant
        Dim lastItemValid As Boolean
        
        ' Get the last item in the WHERE collection
        On Error GoTo GetLastItemError
        lastItem = m_Where(m_Where.Count)
        lastItemValid = True
        
GetLastItemResume:
        ' Add operator only after a condition or closing parenthesis
        If lastItemValid And (IsArray(lastItem) Or lastItem = ")") Then
            m_Where.Add UCase$(op)
        End If
    End If
    
    Set AddLogicalOperator = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
GetLastItemError:
    ' Handle the error - lastItem is not valid
    lastItemValid = False
    m_LastError = Err.Number
    Debug.Print "Error in AddLogicalOperator: " & Err.Description
    Resume GetLastItemResume
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_AddWhere(ByVal field As String, ByVal operator As String, ByVal value As Variant, Optional ByVal paramType As ADODB.DataTypeEnum = adVariant, Optional ByVal paramSize As Long = 0) As IQueryBuilder
    On Error GoTo AddWhereError
    
    ' Validate inputs
    If Len(Trim(field)) = 0 Then
        Debug.Print "WARNING: Empty field name in AddWhere"
        GoTo AddWhereExit
    End If
    
    ' Check for SQL injection in field and operator
    If ContainsSqlInjection(field) Then
        Debug.Print "WARNING: Potentially unsafe SQL in field: " & field
        field = SafeguardSqlInput(field)
    End If
    
    ' Sanitize operator
    operator = Trim(UCase(operator))
    If Not IsValidOperator(operator) Then
        Debug.Print "WARNING: Invalid operator in AddWhere: " & operator
        operator = "="  ' Default to equals
    End If
    
    ' Check for parameter limit
    If m_Params.Count >= MAX_PARAMS Then
        Debug.Print "ERROR: Maximum parameter count reached in QueryBuilder"
        GoTo AddWhereExit
    End If
    
    ' Add logical operator (AND) if needed
    Call AddLogicalOperator("AND")
    
    ' Generate parameter name and determine effective type
    Dim paramName As String
    Dim effectiveType As ADODB.DataTypeEnum
    
    paramName = GenerateParamName()
    
    ' Use provided type or guess from value
    If paramType = adVariant Then
        effectiveType = modFrameworkUtils.GuessAdoType(value)
    Else
        effectiveType = paramType
    End If
    
    ' Add parameter to parameters collection
    m_Params.Add Array(paramName, value, effectiveType, paramSize, adParamInput)
    
    ' Add the condition to WHERE clause
    m_Where.Add Array(field, operator, paramName)
    
AddWhereExit:
    Set IQueryBuilder_AddWhere = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
AddWhereError:
    m_LastError = Err.Number
    Debug.Print "Error in AddWhere: " & Err.Description
    Resume AddWhereExit
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_AddWhereClause(ByVal whereClause As String) As IQueryBuilder
    On Error GoTo AddWhereClauseError
    
    ' Validate input
    If Len(Trim(whereClause)) = 0 Then
        Debug.Print "WARNING: Empty WHERE clause provided"
        GoTo AddWhereClauseExit
    End If
    
    ' Check for SQL injection risks
    If ContainsSqlInjection(whereClause) Then
        Debug.Print "WARNING: Potentially unsafe SQL in WHERE clause: " & whereClause
        ' We don't automatically sanitize direct WHERE clauses as they may contain
        ' legitimate SQL syntax that would be broken by sanitization
        ' Instead, rely on the caller to ensure safety
    End If
    
    ' Add logical operator (AND) if needed
    Call AddLogicalOperator("AND")
    
    ' Add raw SQL clause
    m_Where.Add whereClause
    
AddWhereClauseExit:
    Set IQueryBuilder_AddWhereClause = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
AddWhereClauseError:
    m_LastError = Err.Number
    Debug.Print "Error in AddWhereClause: " & Err.Description
    Resume AddWhereClauseExit
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_AddWhereIn(ByVal field As String, ByVal values As Variant) As IQueryBuilder
    On Error GoTo AddWhereInError
    
    ' Validate field
    If Len(Trim(field)) = 0 Then
        Debug.Print "WARNING: Empty field name in AddWhereIn"
        GoTo AddWhereInExit
    End If
    
    ' Check for SQL injection in field
    If ContainsSqlInjection(field) Then
        Debug.Print "WARNING: Potentially unsafe SQL in field: " & field
        field = SafeguardSqlInput(field)
    End If
    
    ' Validate values is an array
    If Not IsArray(values) Then
        Debug.Print "ERROR: Values parameter must be an array in AddWhereIn"
        Err.Raise 5, "AddWhereIn", "Values parameter must be an array"
        GoTo AddWhereInExit
    End If
    
    ' Skip if empty array
    If UBound(values) < LBound(values) Then
        GoTo AddWhereInExit
    End If
    
    ' Check for parameter limit
    If m_Params.Count + (UBound(values) - LBound(values) + 1) > MAX_PARAMS Then
        Debug.Print "ERROR: Adding these values would exceed maximum parameter count"
        GoTo AddWhereInExit
    End If
    
    ' Add logical operator (AND) if needed
    Call AddLogicalOperator("AND")
    
    ' Build the list of parameters for IN clause
    Dim i As Long
    Dim paramName As String
    Dim paramList As String
    
    paramList = ""
    
    ' Create a parameter for each value
    For i = LBound(values) To UBound(values)
        paramName = GenerateParamName()
        
        ' Add parameter to parameters collection
        m_Params.Add Array(paramName, values(i), modFrameworkUtils.GuessAdoType(values(i)), 0, adParamInput)
        
        ' Build comma-separated list
        If paramList <> "" Then
            paramList = paramList & ","
        End If
        paramList = paramList & paramName
    Next i
    
    ' Add the IN condition to WHERE clause
    m_Where.Add Array(field, "IN", "(" & paramList & ")")
    
AddWhereInExit:
    Set IQueryBuilder_AddWhereIn = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
AddWhereInError:
    m_LastError = Err.Number
    Debug.Print "Error in AddWhereIn: " & Err.Description
    ' Don't re-raise if it was our own validation error
    If Err.Number <> 5 Then
        ' For unexpected errors
        Resume AddWhereInExit
    End If
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_AddWhereBetween(ByVal field As String, ByVal v1 As Variant, ByVal v2 As Variant, Optional ByVal pType As ADODB.DataTypeEnum = adVariant) As IQueryBuilder
    On Error GoTo AddWhereBetweenError
    
    ' Validate field
    If Len(Trim(field)) = 0 Then
        Debug.Print "WARNING: Empty field name in AddWhereBetween"
        GoTo AddWhereBetweenExit
    End If
    
    ' Check for SQL injection in field
    If ContainsSqlInjection(field) Then
        Debug.Print "WARNING: Potentially unsafe SQL in field: " & field
        field = SafeguardSqlInput(field)
    End If
    
    ' Check for parameter limit
    If m_Params.Count + 2 > MAX_PARAMS Then
        Debug.Print "ERROR: Maximum parameter count reached in QueryBuilder"
        GoTo AddWhereBetweenExit
    End If
    
    ' Add logical operator (AND) if needed
    Call AddLogicalOperator("AND")
    
    ' Generate parameter names for both values
    Dim paramName1 As String
    Dim paramName2 As String
    Dim type1 As ADODB.DataTypeEnum
    Dim type2 As ADODB.DataTypeEnum
    
    paramName1 = GenerateParamName()
    paramName2 = GenerateParamName()
    
    ' Determine parameter types
    If pType = adVariant Then
        type1 = modFrameworkUtils.GuessAdoType(v1)
    Else
        type1 = pType
    End If
    
    If pType = adVariant Then
        type2 = modFrameworkUtils.GuessAdoType(v2)
    Else
        type2 = pType
    End If
    
    ' Add parameters to parameters collection
    m_Params.Add Array(paramName1, v1, type1, 0, adParamInput)
    m_Params.Add Array(paramName2, v2, type2, 0, adParamInput)
    
    ' Add the BETWEEN condition to WHERE clause
    m_Where.Add Array(field, "BETWEEN", paramName1 & " AND " & paramName2)
    
AddWhereBetweenExit:
    Set IQueryBuilder_AddWhereBetween = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
AddWhereBetweenError:
    m_LastError = Err.Number
    Debug.Print "Error in AddWhereBetween: " & Err.Description
    Resume AddWhereBetweenExit
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_AddWhereLike(ByVal field As String, ByVal pattern As String, Optional ByVal escapeChar As String = "") As IQueryBuilder
    On Error GoTo AddWhereLikeError
    
    ' Validate inputs
    If Len(Trim(field)) = 0 Then
        Debug.Print "WARNING: Empty field name in AddWhereLike"
        GoTo AddWhereLikeExit
    End If
    
    If Len(pattern) = 0 Then
        Debug.Print "WARNING: Empty pattern in AddWhereLike"
        pattern = "%" ' Default to match anything
    End If
    
    If Len(escapeChar) > 1 Then
        Debug.Print "WARNING: Escape character must be a single character. Using first character."
        escapeChar = Left$(escapeChar, 1)
    End If
    
    ' Check for SQL injection in field
    If ContainsSqlInjection(field) Then
        Debug.Print "WARNING: Potentially unsafe SQL in field: " & field
        field = SafeguardSqlInput(field)
    End If
    
    ' Check for parameter limit
    If m_Params.Count + 1 > MAX_PARAMS Then
        Debug.Print "ERROR: Maximum parameter count reached in QueryBuilder"
        GoTo AddWhereLikeExit
    End If
    
    ' Add logical operator (AND) if needed
    Call AddLogicalOperator("AND")
    
    ' Generate parameter name
    Dim paramName As String
    Dim likeClause As String
    
    paramName = GenerateParamName()
    likeClause = paramName
    
    ' Add escape clause if provided
    If escapeChar <> "" Then
        likeClause = likeClause & " ESCAPE '" & escapeChar & "'"
    End If
    
    ' Add parameter to parameters collection
    m_Params.Add Array(paramName, pattern, adVarWChar, Len(pattern), adParamInput)
    
    ' Add the LIKE condition to WHERE clause
    m_Where.Add Array(field, "LIKE", likeClause)
    
AddWhereLikeExit:
    Set IQueryBuilder_AddWhereLike = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
AddWhereLikeError:
    m_LastError = Err.Number
    Debug.Print "Error in AddWhereLike: " & Err.Description
    Resume AddWhereLikeExit
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_AddWhereIsNull(ByVal field As String, Optional ByVal isNotNull As Boolean = False) As IQueryBuilder
    On Error GoTo AddWhereIsNullError
    
    ' Validate field
    If Len(Trim(field)) = 0 Then
        Debug.Print "WARNING: Empty field name in AddWhereIsNull"
        GoTo AddWhereIsNullExit
    End If
    
    ' Check for SQL injection in field
    If ContainsSqlInjection(field) Then
        Debug.Print "WARNING: Potentially unsafe SQL in field: " & field
        field = SafeguardSqlInput(field)
    End If
    
    ' Add logical operator (AND) if needed
    Call AddLogicalOperator("AND")
    
    ' Determine operator
    Dim operator As String
    
    If isNotNull Then
        operator = "IS NOT NULL"
    Else
        operator = "IS NULL"
    End If
    
    ' Add the IS NULL condition to WHERE clause
    m_Where.Add Array(field, operator, Null)
    
AddWhereIsNullExit:
    Set IQueryBuilder_AddWhereIsNull = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
AddWhereIsNullError:
    m_LastError = Err.Number
    Debug.Print "Error in AddWhereIsNull: " & Err.Description
    Resume AddWhereIsNullExit
End Function

' --- Grouping Operations ---
'@Description: 
'@Param: 
'@Returns: 

Private Function IQueryBuilder_OpenGroup() As IQueryBuilder
    On Error GoTo OpenGroupError
    
    ' Add logical operator (AND) if needed
    Call AddLogicalOperator("AND")
    
    ' Add opening parenthesis
    m_Where.Add "("
    
OpenGroupExit:
    Set IQueryBuilder_OpenGroup = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
OpenGroupError:
    m_LastError = Err.Number
    Debug.Print "Error in OpenGroup: " & Err.Description
    Resume OpenGroupExit
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_CloseGroup() As IQueryBuilder
    On Error GoTo CloseGroupError
    
    ' Only add closing parenthesis if there's content and it makes sense syntactically
    If m_Where.Count > 0 Then
        Dim lastItem As Variant
        Dim lastItemValid As Boolean
        
        ' Get the last item in the WHERE collection
        On Error Resume Next
        lastItem = m_Where(m_Where.Count)
        lastItemValid = (Err.Number = 0)
        On Error GoTo CloseGroupError
        
        ' Only add closing parenthesis after a condition or another closing parenthesis
        If lastItemValid And (IsArray(lastItem) Or lastItem = ")") Then
            m_Where.Add ")"
        End If
    End If
    
CloseGroupExit:
    Set IQueryBuilder_CloseGroup = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
CloseGroupError:
    m_LastError = Err.Number
    Debug.Print "Error in CloseGroup: " & Err.Description
    Resume CloseGroupExit
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_AddAnd() As IQueryBuilder
    On Error GoTo AddAndError
    
    AddOperator "AND"
    
AddAndExit:
    Set IQueryBuilder_AddAnd = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
AddAndError:
    m_LastError = Err.Number
    Debug.Print "Error in AddAnd: " & Err.Description
    Resume AddAndExit
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_AddOr() As IQueryBuilder
    On Error GoTo AddOrError
    
    AddOperator "OR"
    
AddOrExit:
    Set IQueryBuilder_AddOr = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
AddOrError:
    m_LastError = Err.Number
    Debug.Print "Error in AddOr: " & Err.Description
    Resume AddOrExit
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_AddOrGroup() As IQueryBuilder
    On Error GoTo AddOrGroupError
    
    AddOperator "OR"
    m_Where.Add "("
    
AddOrGroupExit:
    Set IQueryBuilder_AddOrGroup = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
AddOrGroupError:
    m_LastError = Err.Number
    Debug.Print "Error in AddOrGroup: " & Err.Description
    Resume AddOrGroupExit
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_AddAndGroup() As IQueryBuilder
    On Error GoTo AddAndGroupError
    
    AddOperator "AND"
    m_Where.Add "("
    
AddAndGroupExit:
    Set IQueryBuilder_AddAndGroup = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
AddAndGroupError:
    m_LastError = Err.Number
    Debug.Print "Error in AddAndGroup: " & Err.Description
    Resume AddAndGroupExit
End Function

' --- Other Clauses ---
'@Description: 
'@Param: 
'@Returns: 

Private Function IQueryBuilder_AddAlias(tableAlias As String) As IQueryBuilder
    On Error GoTo AddAliasError
    
    ' Validate input
    If Len(Trim(tableAlias)) = 0 Then
        Debug.Print "WARNING: Empty alias name"
        GoTo AddAliasExit
    End If
    
    ' Check for SQL injection
    If ContainsSqlInjection(tableAlias) Then
        Debug.Print "WARNING: Potentially unsafe SQL in alias: " & tableAlias
        tableAlias = SafeguardSqlInput(tableAlias)
    End If
    
    ' Parse "table AS alias" format
    Dim parts() As String
    Dim tableName As String
    Dim aliasName As String
    
    If InStr(1, tableAlias, " AS ", vbTextCompare) > 0 Then
        parts = Split(tableAlias, " AS ", , vbTextCompare)
        tableName = Trim(parts(0))
        aliasName = Trim(parts(1))
        
        ' Store in aliases dictionary
        If Not m_Aliases.Exists(tableName) Then
            m_Aliases.Add tableName, aliasName
        Else
            m_Aliases(tableName) = aliasName
        End If
    End If
    
AddAliasExit:
    Set IQueryBuilder_AddAlias = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
AddAliasError:
    m_LastError = Err.Number
    Debug.Print "Error in AddAlias: " & Err.Description
    Resume AddAliasExit
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_UseSafeIdentifiers(Optional ByVal useSafe As Boolean = True) As IQueryBuilder
    On Error GoTo UseSafeIdentifiersError
    
    m_UseSafeIdentifiers = useSafe
    
    If useSafe Then
        m_SafeIdPrefix = "["
        m_SafeIdSuffix = "]"
    Else
        m_SafeIdPrefix = ""
        m_SafeIdSuffix = ""
    End If
    
UseSafeIdentifiersExit:
    Set IQueryBuilder_UseSafeIdentifiers = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
UseSafeIdentifiersError:
    m_LastError = Err.Number
    Debug.Print "Error in UseSafeIdentifiers: " & Err.Description
    Resume UseSafeIdentifiersExit
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_Join(ByVal joinTable As String, ByVal onClause As String, Optional ByVal joinType As String = "INNER") As IQueryBuilder
    On Error GoTo JoinError
    
    ' Validate inputs
    If Len(Trim(joinTable)) = 0 Then
        Debug.Print "WARNING: Empty join table name"
        GoTo JoinExit
    End If
    
    If Len(Trim(onClause)) = 0 Then
        Debug.Print "WARNING: Empty ON clause in join"
        GoTo JoinExit
    End If
    
    ' Validate join type
    joinType = UCase$(Trim$(joinType))
    If Not IsValidJoinType(joinType) Then
        Debug.Print "WARNING: Invalid join type: " & joinType & ". Using INNER."
        joinType = "INNER"
    End If
    
    ' Check for SQL injection
    If ContainsSqlInjection(joinTable) Then
        Debug.Print "WARNING: Potentially unsafe SQL in join table: " & joinTable
        joinTable = SafeguardSqlInput(joinTable)
    End If
    
    ' We don't automatically sanitize the ON clause as it may contain
    ' legitimate SQL syntax that would be broken by sanitization
    ' Instead, warn about potential issues
    If ContainsSqlInjection(onClause) Then
        Debug.Print "WARNING: Potentially unsafe SQL in ON clause: " & onClause
    End If
    
    ' Add the join
    m_Joins.Add joinType & " JOIN " & joinTable & " ON " & onClause
    
JoinExit:
    Set IQueryBuilder_Join = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
JoinError:
    m_LastError = Err.Number
    Debug.Print "Error in Join: " & Err.Description
    Resume JoinExit
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_GroupBy(ByVal columns As String) As IQueryBuilder
    On Error GoTo GroupByError
    
    ' Validate input
    If Len(Trim(columns)) = 0 Then
        Debug.Print "WARNING: Empty columns in GroupBy"
        GoTo GroupByExit
    End If
    
    ' Check for SQL injection
    If ContainsSqlInjection(columns) Then
        Debug.Print "WARNING: Potentially unsafe SQL in GROUP BY columns: " & columns
        columns = SafeguardSqlInput(columns)
    End If
    
    m_GroupBy = columns
    
GroupByExit:
    Set IQueryBuilder_GroupBy = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
GroupByError:
    m_LastError = Err.Number
    Debug.Print "Error in GroupBy: " & Err.Description
    Resume GroupByExit
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_OrderBy(ByVal columns As String, Optional ByVal descending As Boolean = False) As IQueryBuilder
    On Error GoTo OrderByError
    
    ' Validate input
    If Len(Trim(columns)) = 0 Then
        Debug.Print "WARNING: Empty columns in OrderBy"
        GoTo OrderByExit
    End If
    
    ' Check for SQL injection
    If ContainsSqlInjection(columns) Then
        Debug.Print "WARNING: Potentially unsafe SQL in ORDER BY columns: " & columns
        columns = SafeguardSqlInput(columns)
    End If
    
    Dim direction As String
    
    If descending Then
        direction = "DESC"
    Else
        direction = "ASC"
    End If
    
    m_OrderBy.Add Array(columns, direction)
    
OrderByExit:
    Set IQueryBuilder_OrderBy = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
OrderByError:
    m_LastError = Err.Number
    Debug.Print "Error in OrderBy: " & Err.Description
    Resume OrderByExit
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_Top(ByVal n As Long) As IQueryBuilder
    On Error GoTo TopError
    
    If n > 0 Then
        m_TopCount = n
    Else
        m_TopCount = 0
    End If
    
TopExit:
    Set IQueryBuilder_Top = Me
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
TopError:
    m_LastError = Err.Number
    Debug.Print "Error in Top: " & Err.Description
    Resume TopExit
End Function

' --- Build & Debug ---
'@Description: 
'@Param: 
'@Returns: 

Private Function IQueryBuilder_Build() As Variant
    On Error GoTo BuildError
    
    Dim sqlParts As Object
    Dim finalParams() As Variant
    Dim sqlString As String
    Dim i As Long
    
    ' Create Dictionary to hold SQL parts
    Set sqlParts = CreateObject("Scripting.Dictionary")
    
    ' Build each clause
    BuildSelectClause sqlParts, Me
    BuildFromClause sqlParts, Me
    BuildJoinClause sqlParts, Me
    BuildWhereClause sqlParts, Me
    BuildGroupByClause sqlParts, Me
    BuildOrderByClause sqlParts, Me
    
    ' Concatenate all parts
    sqlString = sqlParts("SELECT") & _
                sqlParts("FROM") & _
                sqlParts("JOIN") & _
                sqlParts("WHERE") & _
                sqlParts("GROUPBY") & _
                sqlParts("ORDERBY")
    
    ' Apply pagination if driver is available and pagination is requested
    If Not m_Driver Is Nothing And m_PageSize > 0 And m_PageNumber > 0 Then
        sqlString = m_Driver.GetPaginationSql(sqlString, m_PageSize, m_PageNumber)
    End If
    
    ' Check for SQL length limit
    If Len(sqlString) > MAX_SQL_LENGTH Then
        Debug.Print "WARNING: Generated SQL exceeds recommended length limit: " & Len(sqlString) & " chars"
    End If
    
    ' Prepare parameters array if any
    If m_Params.Count > 0 Then
        ReDim finalParams(0 To m_Params.Count - 1)
        
        For i = 1 To m_Params.Count
            finalParams(i - 1) = m_Params(i)
        Next i
    Else
        finalParams = Empty
    End If
    
    ' Return array with SQL and parameters
    IQueryBuilder_Build = Array(Trim$(sqlString), finalParams)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
BuildError:
    m_LastError = Err.Number
    Debug.Print "Error in Build: " & Err.Description
    ' Return an empty, but valid result
    IQueryBuilder_Build = Array("SELECT 1 WHERE 1=0", Empty)
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_ToSqlString() As String
    On Error GoTo ToSqlStringError
    
    Dim result As Variant
    
    result = IQueryBuilder_Build()
    IQueryBuilder_ToSqlString = result(0)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ToSqlStringError:
    m_LastError = Err.Number
    Debug.Print "Error in ToSqlString: " & Err.Description
    IQueryBuilder_ToSqlString = ""
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IQueryBuilder_ToDebugString() As String
    On Error GoTo DebugStringError
    
    Dim buildResult As Variant
    Dim sql As String
    Dim params As Variant
    Dim i As Long
    Dim paramInfo As String
    Dim paramItem As Variant
    
    ' Get SQL and parameters
    buildResult = Me.Build()
    sql = buildResult(0)
    params = buildResult(1)
    
    ' Format parameters info
    paramInfo = ""
    
    If Not IsEmpty(params) Then
        If IsArray(params) Then
            paramInfo = paramInfo & vbCrLf & "-- Parameters (" & UBound(params) + 1 & ") --"
            
            For i = LBound(params) To UBound(params)
                ' Safely access parameter item
                On Error Resume Next
                paramItem = params(i)
                
                If Err.Number = 0 Then
                    Dim paramStr As String
                    paramStr = "Name=" & SafeGetArrayItem(paramItem, 0) & _
                              ", Type=" & SafeGetArrayItem(paramItem, 2) & _
                              ", Size=" & SafeGetArrayItem(paramItem, 3) & _
                              ", Direction=" & SafeGetArrayItem(paramItem, 4) & _
                              ", Value='" & Format_Value(SafeGetArrayItem(paramItem, 1)) & "'"
                              
                    paramInfo = paramInfo & vbCrLf & paramStr
                Else
                    paramInfo = paramInfo & vbCrLf & "Invalid parameter at index " & i
                    Err.Clear
                End If
                On Error GoTo DebugStringError
            Next i
        Else
            paramInfo = vbCrLf & "-- Invalid Parameters Object --"
        End If
    Else
        paramInfo = vbCrLf & "-- No Parameters --"
    End If
    
    IQueryBuilder_ToDebugString = "-- Generated SQL --" & vbCrLf & sql & paramInfo & vbCrLf & "-------------------"
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
DebugStringError:
    m_LastError = Err.Number
    IQueryBuilder_ToDebugString = "Error generating Debug String: " & Err.Description
End Function

' --- Private Build Helpers ---
'@Description: 
'@Param: 
'@Returns: 

Private Function SafeId(identifier As String) As String
    ' Use driver if available
    If Not m_Driver Is Nothing Then
        If m_UseSafeIdentifiers And identifier <> "*" Then
            SafeId = m_Driver.EscapeIdentifier(identifier)
        Else
            SafeId = identifier
        End If
    Else
        ' Fallback to original behavior
        If m_UseSafeIdentifiers And identifier <> "*" Then
            SafeId = m_SafeIdPrefix & identifier & m_SafeIdSuffix
        Else
            SafeId = identifier
        End If
    End If
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Sub BuildSelectClause(sqlParts As Object, instance As clsQueryBuilder)
    Dim sql As String
    
    sql = "SELECT "
    
    ' Add TOP or LIMIT using driver if available
    If Not instance.m_Driver Is Nothing And instance.m_TopCount > 0 Then
        sql = sql & instance.m_Driver.GetTopClause(instance.m_TopCount) & " "
    ElseIf instance.m_TopCount > 0 Then
        ' Fallback if no driver
        sql = sql & "TOP " & instance.m_TopCount & " "
    End If
    
    sql = sql & instance.m_Select
    sqlParts("SELECT") = sql
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub BuildFromClause(sqlParts As Object, instance As clsQueryBuilder)
    On Error GoTo BuildFromClauseError
    
    sqlParts("FROM") = " FROM " & SafeId(instance.m_From)
    
    ' TODO: Handle table alias
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    
BuildFromClauseError:
    Debug.Print "Error in BuildFromClause: " & Err.Description
    sqlParts("FROM") = " FROM DUAL" ' Generic fallback
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub BuildJoinClause(sqlParts As Object, instance As clsQueryBuilder)
    On Error GoTo BuildJoinClauseError
    
    Dim sql As String
    Dim item As Variant
    
    sql = ""
    
    If instance.m_Joins.Count > 0 Then
        For Each item In instance.m_Joins
            sql = sql & " " & item
        Next item
    End If
    
    sqlParts("JOIN") = sql
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    
BuildJoinClauseError:
    Debug.Print "Error in BuildJoinClause: " & Err.Description
    sqlParts("JOIN") = ""
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub BuildWhereClause(sqlParts As Object, instance As clsQueryBuilder)
    On Error GoTo BuildWhereClauseError
    
    Dim sql As String
    Dim item As Variant
    Dim clausePart As String
    Dim needsOp As Boolean
    
    sql = ""
    
    If instance.m_Where.Count > 0 Then
        For Each item In instance.m_Where
            If IsArray(item) Then
                ' Build condition from field, operator, and value
                clausePart = BuildWherePart(item, instance)
                
                If needsOp And sql <> "" And Right$(Trim$(sql), 1) <> "(" Then
                    sql = sql & " "
                End If
                
                sql = sql & clausePart
                needsOp = True
            Else
                ' Handle operators and parentheses
                Select Case item
                    Case "(", ")"
                        If item = "(" Then
                            If needsOp And sql <> "" And Right$(Trim$(sql), 1) <> "(" Then
                                sql = sql & " "
                            End If
                            sql = sql & item & " "
                        Else ' ")"
                            sql = sql & " " & item & " "
                        End If
                        needsOp = False
                    Case Else ' "AND", "OR"
                        If needsOp Then
                            sql = Trim$(sql) & " " & item & " "
                            needsOp = False
                        End If
                End Select
            End If
        Next item
        
        sqlParts("WHERE") = " WHERE " & Trim$(sql)
    Else
        sqlParts("WHERE") = ""
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    
BuildWhereClauseError:
    Debug.Print "Error in BuildWhereClause: " & Err.Description
    sqlParts("WHERE") = ""
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Function BuildWherePart(clauseParts As Variant, instance As clsQueryBuilder) As String
    On Error GoTo BuildWherePartError
    
    Dim field As String
    Dim op As String
    Dim val As Variant
    
    ' Safely get clause parts
    If IsArray(clauseParts) Then
        If UBound(clauseParts) >= LBound(clauseParts) + 2 Then
            field = SafeId(clauseParts(0))
            op = UCase$(clauseParts(1))
            val = clauseParts(2)
            
            ' Build clause based on operator type
            Select Case op
                Case "IS NULL", "IS NOT NULL"
                    BuildWherePart = field & " " & op
                    
                Case "IN", "BETWEEN"
                    BuildWherePart = field & " " & op & " " & val
                    
                Case "LIKE"
                    BuildWherePart = field & " " & op & " " & val
                    
                Case Else
                    BuildWherePart = field & " " & op & " " & val
            End Select
        Else
            BuildWherePart = "1=1" ' Default to always true if invalid array
        End If
    Else
        BuildWherePart = "1=1" ' Default to always true if not an array
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
BuildWherePartError:
    Debug.Print "Error in BuildWherePart: " & Err.Description
    BuildWherePart = "1=1" ' Default to always true on error
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Sub BuildGroupByClause(sqlParts As Object, instance As clsQueryBuilder)
    On Error GoTo BuildGroupByClauseError
    
    If instance.m_GroupBy <> "" Then
        sqlParts("GROUPBY") = " GROUP BY " & instance.m_GroupBy
    Else
        sqlParts("GROUPBY") = ""
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    
BuildGroupByClauseError:
    Debug.Print "Error in BuildGroupByClause: " & Err.Description
    sqlParts("GROUPBY") = ""
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub BuildOrderByClause(sqlParts As Object, instance As clsQueryBuilder)
    On Error GoTo BuildOrderByClauseError
    
    Dim sql As String
    Dim item As Variant
    
    sql = ""
    
    If instance.m_OrderBy.Count > 0 Then
        For Each item In instance.m_OrderBy
            If IsArray(item) And UBound(item) >= 1 Then
                If sql <> "" Then
                    sql = sql & ", "
                End If
                
                sql = sql & item(0) & " " & item(1)
            End If
        Next item
        
        If sql <> "" Then
            sqlParts("ORDERBY") = " ORDER BY " & sql
        Else
            sqlParts("ORDERBY") = ""
        End If
    Else
        sqlParts("ORDERBY") = ""
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    
BuildOrderByClauseError:
    Debug.Print "Error in BuildOrderByClause: " & Err.Description
    sqlParts("ORDERBY") = ""
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Function GenerateParamName() As String
    ' Use driver if available
    If Not m_Driver Is Nothing Then
        Dim baseName As String
        baseName = "p" & m_ParamIndex
        m_ParamIndex = m_ParamIndex + 1
        GenerateParamName = m_Driver.FormatParameterName(baseName)
    Else
        ' Fallback to original behavior
        GenerateParamName = "@p" & m_ParamIndex
        m_ParamIndex = m_ParamIndex + 1
    End If
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Sub AddOperator(op As String)
    On Error GoTo AddOperatorError
    
    If m_Where.Count > 0 Then
        Dim lastItem As Variant
        Dim lastItemValid As Boolean
        
        ' Get the last item in the WHERE collection
        On Error Resume Next
        lastItem = m_Where(m_Where.Count)
        lastItemValid = (Err.Number = 0)
        On Error GoTo AddOperatorError
        
        ' Add operator only after a condition or closing parenthesis
        If lastItemValid And (IsArray(lastItem) Or lastItem = ")") Then
            m_Where.Add UCase$(op)
        End If
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    
AddOperatorError:
    m_LastError = Err.Number
    Debug.Print "Error in AddOperator: " & Err.Description
End Sub

' --- Security and Validation Helpers ---

' Check if a string might contain SQL injection attempts
'@Description: 
'@Param: 
'@Returns: 

Private Function ContainsSqlInjection(input As String) As Boolean
    Dim i As Integer
    
    ' Default to safe
    ContainsSqlInjection = False
    
    ' Check for basic SQL injection patterns
    If InStr(1, input, "--;", vbTextCompare) > 0 Then ContainsSqlInjection = True
    If InStr(1, input, "/*", vbTextCompare) > 0 Then ContainsSqlInjection = True
    If InStr(1, input, "*/", vbTextCompare) > 0 Then ContainsSqlInjection = True
    If InStr(1, input, "@@", vbTextCompare) > 0 Then ContainsSqlInjection = True
    If InStr(1, input, "EXEC(", vbTextCompare) > 0 Then ContainsSqlInjection = True
    If InStr(1, input, "EXECUTE(", vbTextCompare) > 0 Then ContainsSqlInjection = True
    If InStr(1, input, "UNION ", vbTextCompare) > 0 Then ContainsSqlInjection = True
    If InStr(1, input, "DROP ", vbTextCompare) > 0 Then ContainsSqlInjection = True
    If InStr(1, input, "DELETE ", vbTextCompare) > 0 Then ContainsSqlInjection = True
    If InStr(1, input, "INSERT ", vbTextCompare) > 0 Then ContainsSqlInjection = True
    If InStr(1, input, "UPDATE ", vbTextCompare) > 0 Then ContainsSqlInjection = True
    
    ' Check for suspicious character combinations
    For i = 1 To Len(INVALID_SQL_CHARS)
        If InStr(1, input, Mid$(INVALID_SQL_CHARS, i, 1)) > 0 Then
            ContainsSqlInjection = True
            Exit'@Description: 
'@Param: 
'@Returns: 

 Function
        End If
    Next i
End Function

' Basic sanitization of input for SQL
'@Description: 
'@Param: 
'@Returns: 

Private Function SafeguardSqlInput(input As String) As String
    Dim result As String
    
    ' Remove potentially dangerous characters
    result = input
    result = Replace(result, "'", "")
    result = Replace(result, """", "")
    result = Replace(result, ";", "")
    result = Replace(result, "--", "")
    result = Replace(result, "/*", "")
    result = Replace(result, "*/", "")
    result = Replace(result, "@@", "@")
    result = Replace(result, "%", "[%]") ' Escape % for LIKE patterns
    result = Replace(result, "_", "[_]") ' Escape _ for LIKE patterns
    
    SafeguardSqlInput = result
End Function

' Check if operator is valid
'@Description: 
'@Param: 
'@Returns: 

Private Function IsValidOperator(op As String) As Boolean
    ' List of allowed operators
    Select Case op
        Case "=", "<>", ">", "<", ">=", "<=", "LIKE", "NOT LIKE", "IN", "NOT IN", "IS NULL", "IS NOT NULL", "BETWEEN"
            IsValidOperator = True
        Case Else
            IsValidOperator = False
    End Select
End Function

' Check if join type is valid
'@Description: 
'@Param: 
'@Returns: 

Private Function IsValidJoinType(joinType As String) As Boolean
    ' List of allowed join types
    Select Case joinType
        Case "INNER", "LEFT", "RIGHT", "FULL", "CROSS", "LEFT OUTER", "RIGHT OUTER", "FULL OUTER"
            IsValidJoinType = True
        Case Else
            IsValidJoinType = False
    End Select
End Function

' Safely get item from an array, with error handling
'@Description: 
'@Param: 
'@Returns: 

Private Function SafeGetArrayItem(arr As Variant, index As Long) As Variant
    On Error Resume Next
    
    If IsArray(arr) Then
        If index >= LBound(arr) And index <= UBound(arr) Then
            SafeGetArrayItem = arr(index)
            
            If Err.Number <> 0 Then
                SafeGetArrayItem = "[Access Error]"
                Err.Clear
            End If
        Else
            SafeGetArrayItem = "[Out of Bounds]"
        End If
    Else
        SafeGetArrayItem = "[Not an Array]"
    End If
    
    On Error GoTo 0
End Function

' Format a value safely for display in debug output
'@Description: 
'@Param: 
'@Returns: 

Private Function Format_Value(value As Variant) As String
    On Error Resume Next
    
    If IsNull(value) Then
        Format_Value = "NULL"
    ElseIf IsEmpty(value) Then
        Format_Value = "EMPTY"
    ElseIf VarType(value) = vbError Then
        Format_Value = "ERROR #" & CStr(value)
    ElseIf IsObject(value) Then
        Format_Value = "OBJECT:" & TypeName(value)
    ElseIf IsArray(value) Then
        Format_Value = "ARRAY[" & (UBound(value) - LBound(value) + 1) & "]"
    Else
        ' Try to convert to string
        Format_Value = CStr(value)
        
        ' Truncate if too long
        If Len(Format_Value) > 50 Then
            Format_Value = Left$(Format_Value, 47) & "..."
        End If
    End If
    
    If Err.Number <> 0 Then
        Format_Value = "[Value conversion error]"
        Err.Clear
    End If
    
    On Error GoTo 0
End Function

' --- Class Cleanup ---
Private Sub Class_Terminate()
    Set m_Where = Nothing
    Set m_Params = Nothing
    Set m_Joins = Nothing
    Set m_OrderBy = Nothing
End Sub

' --- New Batch Processing Support ---
'@Description: 
'@Param: 
'@Returns: 

Public Function SupportsBatchOperations() As Boolean
    If Not m_Driver Is Nothing Then
        SupportsBatchOperations = m_Driver.SupportsBatchOperations()
    Else
        SupportsBatchOperations = False
    End If
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetBatchInsertSql(tableName As String, columnNames() As String, rowCount As Long) As String
    If Not m_Driver Is Nothing Then
        GetBatchInsertSql = m_Driver.GetBatchInsertSql(tableName, columnNames, rowCount)
    Else
        ' Fallback implementation
        Dim sql As String
        Dim i As Long, j As Long
        Dim placeholders As String
        
        sql = "INSERT INTO " & SafeId(tableName) & " ("
        
        ' Add column names
        For i = LBound(columnNames) To UBound(columnNames)
            sql = sql & SafeId(columnNames(i))
            If i < UBound(columnNames) Then sql = sql & ", "
        Next i
        
        sql = sql & ") VALUES "
        
        ' Add placeholders for values
        For i = 1 To rowCount
            placeholders = "("
            For j = LBound(columnNames) To UBound(columnNames)
                placeholders = placeholders & "?"
                If j < UBound(columnNames) Then placeholders = placeholders & ", "
            Next j
            placeholders = placeholders & ")"
            
            sql = sql & placeholders
            If i < rowCount Then sql = sql & ", "
        Next i
        
        GetBatchInsertSql = sql
    End If
End Function 