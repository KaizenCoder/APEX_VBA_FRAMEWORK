' Migrated to apex-metier/database - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit
' ==========================================================================
' Class : clsAccessDriver
' Version : 2.0
' Implements: IDbDriver
' Purpose : Implements DB-specific logic for MS Access (ACE/Jet).
' ==========================================================================
Implements IDbDriver

' --- Connection String ---
Private Function IDbDriver_GetConnectionString(server As String, dbName As String, Optional user As String = "", Optional pwd As String = "", Optional useWindowsAuth As Boolean = True) As String
    IDbDriver_GetConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbName & ";"
    If pwd <> "" Then IDbDriver_GetConnectionString = IDbDriver_GetConnectionString & "Jet OLEDB:Database Password=" & pwd & ";"
End Function

' --- Parameter Formatting ---
Private Property Get IDbDriver_ParameterMarker() As String
    IDbDriver_ParameterMarker = "?"
End Property

Private Function IDbDriver_FormatParameterName(baseName As String) As String
    IDbDriver_FormatParameterName = "?" ' Access uses positional markers
End Function

Private Function IDbDriver_FormatSqlDate(dtValue As Date) As String
    IDbDriver_FormatSqlDate = "#" & Format$(dtValue, "yyyy-mm-dd hh:nn:ss") & "#"
End Function

Private Function IDbDriver_FormatSqlString(value As String) As String
    ' Doubler les apostrophes pour échapper
    IDbDriver_FormatSqlString = "'" & Replace(value, "'", "''") & "'"
End Function

Private Function IDbDriver_FormatSqlNumber(value As Variant) As String
    If IsNumeric(value) Then
        IDbDriver_FormatSqlNumber = CStr(value)
    Else
        IDbDriver_FormatSqlNumber = "NULL"
    End If
End Function

Private Function IDbDriver_EscapeIdentifier(identifier As String) As String
    ' Access utilise les crochets [] pour échapper les identifiants
    IDbDriver_EscapeIdentifier = "[" & Replace(identifier, "]", "]]") & "]"
End Function

' --- SQL Specific Clauses ---
Private Function IDbDriver_GetTopClause(n As Long) As String
    If n > 0 Then IDbDriver_GetTopClause = "TOP " & n Else IDbDriver_GetTopClause = ""
End Function

Private Function IDbDriver_GetLimitClause(n As Long) As String
    IDbDriver_GetLimitClause = "" ' Access uses TOP
End Function

Private Function IDbDriver_GetLastIdentitySql() As String
    IDbDriver_GetLastIdentitySql = "SELECT @@IDENTITY"
End Function

Private Function IDbDriver_GetPaginationSql(sql As String, pageSize As Long, pageNumber As Long) As String
    ' Access ne supporte pas directement la pagination
    ' On utilise un TOP avec sous-requête pour simuler
    If pageSize <= 0 Or pageNumber <= 0 Then
        IDbDriver_GetPaginationSql = sql
        Exit Function
    End If
    
    Dim offset As Long
    offset = (pageNumber - 1) * pageSize
    
    ' Si c'est la première page, on peut simplement utiliser TOP
    If offset = 0 Then
        ' Insérer TOP après SELECT
        Dim selectPos As Long
        selectPos = InStr(1, UCase$(sql), "SELECT")
        If selectPos > 0 Then
            IDbDriver_GetPaginationSql = "SELECT TOP " & pageSize & " " & Mid$(sql, selectPos + 7)
        Else
            IDbDriver_GetPaginationSql = sql
        End If
    Else
        ' Pour Access, la pagination avancée nécessite une sous-requête avec un ID unique
        ' C'est une approximation, car Access ne supporte pas OFFSET/FETCH
        IDbDriver_GetPaginationSql = sql
    End If
End Function

' --- Batch Processing ---
Private Function IDbDriver_SupportsBatchOperations() As Boolean
    IDbDriver_SupportsBatchOperations = False ' Access ne supporte pas bien les opérations par lots
End Function

Private Function IDbDriver_GetBatchInsertSql(tableName As String, columnNames() As String, rowCount As Long) As String
    ' Access ne supporte pas vraiment les insertions par lots efficaces
    ' On construit une requête INSERT standard
    Dim sql As String
    Dim i As Long, j As Long
    Dim placeholders As String
    
    sql = "INSERT INTO " & IDbDriver_EscapeIdentifier(tableName) & " ("
    
    ' Ajouter les noms de colonnes
    For i = LBound(columnNames) To UBound(columnNames)
        sql = sql & IDbDriver_EscapeIdentifier(columnNames(i))
        If i < UBound(columnNames) Then sql = sql & ", "
    Next i
    
    sql = sql & ") VALUES "
    
    ' Ajouter les valeurs pour chaque ligne
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
    
    IDbDriver_GetBatchInsertSql = sql
End Function

' --- Schema Information ---
Private Function IDbDriver_GetTableColumnsSql(tableName As String) As String
    IDbDriver_GetTableColumnsSql = "SELECT * FROM " & IDbDriver_EscapeIdentifier(tableName) & " WHERE 1=0"
End Function

Private Function IDbDriver_GetTableListSql() As String
    IDbDriver_GetTableListSql = "SELECT Name FROM MSysObjects WHERE Type=1 AND Flags=0"
End Function

Private Function IDbDriver_GetPrimaryKeysSql(tableName As String) As String
    ' Access ne fournit pas facilement cette information via SQL
    ' On peut utiliser des méthodes ADO alternatives
    IDbDriver_GetPrimaryKeysSql = ""
End Function

Private Function IDbDriver_GetForeignKeysSql(tableName As String) As String
    ' Access ne fournit pas facilement cette information via SQL
    ' On peut utiliser des méthodes ADO alternatives
    IDbDriver_GetForeignKeysSql = ""
End Function

' --- Advanced Features ---
Private Function IDbDriver_SupportsTransactionIsolationLevel() As Boolean
    IDbDriver_SupportsTransactionIsolationLevel = False
End Function

Private Function IDbDriver_GetSetIsolationLevelSql(level As Long) As String
    IDbDriver_GetSetIsolationLevelSql = "" ' Non supporté par Access
End Function

Private Function IDbDriver_SupportsStoredProcedures() As Boolean
    IDbDriver_SupportsStoredProcedures = False ' Access utilise des requêtes nommées, pas des procédures stockées
End Function

Private Function IDbDriver_SupportsUnions() As Boolean
    IDbDriver_SupportsUnions = True
End Function 