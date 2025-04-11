' Migrated to apex-metier/database - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit
' ==========================================================================
' Class : clsDbAccessor
' Version : 5.0 - Phase 3 - Plugin System and Performance Optimizations
' Implements: IDbAccessorBase (v3.1)
' Purpose : Industrialized ADO layer: Driver injection support, refined logging,
' helpers, robust parameters, large volume handling via GetRecordset.
' Requires : ADO, All Framework Interfaces/Classes, modFrameworkUtils
' ==========================================================================

Implements IDbAccessorBase

' --- Member Variables ---
Private m_cn As ADODB.Connection
Private m_connString As String
Private m_isConnected As Boolean
Private m_isInTransaction As Boolean
Private m_Logger As ILoggerBase
Private m_SlowQueryThresholdMillisec As Long
Private m_Driver As IDbDriver ' Store the injected driver
Private m_LastError As Long ' Store last error number for checking
Private m_TransactionNestLevel As Long ' Track nested transactions
Private m_PluginManager As clsPluginManager ' Gestionnaire de plugins
Private m_UsePlugins As Boolean ' Activation des plugins
Private m_StatementCache As Object ' Scripting.Dictionary - Cache des commandes préparées
Private m_StatementCacheEnabled As Boolean ' Activation du cache de commandes
Private m_BatchSize As Long ' Taille de lot pour les opérations par lots

' --- Constants ---
Private Const DEFAULT_SLOW_QUERY_MS As Long = 500
Private Const MAX_PARAM_SIZE As Long = 8000 ' Maximum safe parameter size
Private Const DB_OPERATION_TIMEOUT As Long = 30 ' 30 seconds default command timeout
Private Const DEFAULT_BATCH_SIZE As Long = 1000 ' Taille de lot par défaut
Private Const MAX_PREPARED_STATEMENTS As Long = 50 ' Nombre maximal de déclarations préparées en cache

' --- Interface Implementation ---

' --- Initialization & Configuration ---
Private Sub IDbAccessorBase_Initialize(Optional ByVal configSource As Variant)
    Set m_Logger = Nothing
    Set m_Driver = Nothing
    m_SlowQueryThresholdMillisec = DEFAULT_SLOW_QUERY_MS
    m_LastError = 0
    m_TransactionNestLevel = 0
    m_UsePlugins = False
    Set m_PluginManager = Nothing
    Set m_StatementCache = CreateObject("Scripting.Dictionary")
    m_StatementCacheEnabled = True
    m_BatchSize = DEFAULT_BATCH_SIZE
End Sub

Private Sub IDbAccessorBase_SetLogger(ByVal loggerInstance As ILoggerBase)
    Set m_Logger = loggerInstance
    LogIfAble Me, "SetLogger", LogLevelDebug, "Logger set: " & TypeName(loggerInstance), "CONFIG"
    
    ' Si le gestionnaire de plugins existe, on lui définit aussi le logger
    If Not m_PluginManager Is Nothing Then
        m_PluginManager.SetLogger loggerInstance
    End If
End Sub

Public Sub IDbAccessorBase_SetDriver(ByVal dbDriver As IDbDriver)
    Set m_Driver = dbDriver
    LogIfAble Me, "SetDriver", LogLevelDebug, "Driver set: " & TypeName(dbDriver), "CONFIG"
End Sub

' --- Plugin Management (New) ---
Public Sub EnablePlugins(Optional ByVal enable As Boolean = True)
    m_UsePlugins = enable
    
    If m_UsePlugins And m_PluginManager Is Nothing Then
        Set m_PluginManager = New clsPluginManager
        m_PluginManager.Initialize m_Logger
        LogIfAble Me, "EnablePlugins", LogLevelInfo, "Plugin system enabled", "PLUGIN"
    ElseIf Not m_UsePlugins And Not m_PluginManager Is Nothing Then
        LogIfAble Me, "EnablePlugins", LogLevelInfo, "Plugin system disabled", "PLUGIN"
    End If
End Sub

Public Function RegisterPlugin(plugin As IPlugin, Optional ByVal pluginId As String = "") As String
    If Not m_UsePlugins Then
        EnablePlugins True
    End If
    
    If m_PluginManager Is Nothing Then
        LogIfAble Me, "RegisterPlugin", LogLevelError, "Plugin manager not initialized", "PLUGIN"
        RegisterPlugin = ""
        Exit Function
    End If
    
    RegisterPlugin = m_PluginManager.RegisterPlugin(plugin, pluginId)
End Function

Public Function GetPlugin(ByVal pluginId As String) As IPlugin
    If Not m_UsePlugins Or m_PluginManager Is Nothing Then
        Set GetPlugin = Nothing
        Exit Function
    End If
    
    Set GetPlugin = m_PluginManager.GetPlugin(pluginId)
End Function

Public Function UnregisterPlugin(ByVal pluginId As String) As Boolean
    If Not m_UsePlugins Or m_PluginManager Is Nothing Then
        UnregisterPlugin = False
        Exit Function
    End If
    
    UnregisterPlugin = m_PluginManager.UnregisterPlugin(pluginId)
End Function

' --- Performance Settings (New) ---
Public Sub EnableStatementCache(Optional ByVal enable As Boolean = True)
    m_StatementCacheEnabled = enable
    
    If Not m_StatementCacheEnabled Then
        ' Clear the cache when disabling
        ClearStatementCache
    End If
    
    LogIfAble Me, "EnableStatementCache", LogLevelInfo, "Statement cache " & IIf(enable, "enabled", "disabled"), "PERFORMANCE"
End Sub

Public Sub ClearStatementCache()
    ' Nettoyer les ressources COM
    Dim key As Variant
    Dim cmd As ADODB.Command
    
    If Not m_StatementCache Is Nothing Then
        For Each key In m_StatementCache.Keys
            On Error Resume Next
            Set cmd = m_StatementCache(key)
            Set cmd = Nothing
            On Error GoTo 0
        Next key
        
        m_StatementCache.RemoveAll
    End If
    
    LogIfAble Me, "ClearStatementCache", LogLevelInfo, "Statement cache cleared", "PERFORMANCE"
End Sub

Public Property Let BatchSize(ByVal value As Long)
    If value > 0 Then
        m_BatchSize = value
    Else
        m_BatchSize = DEFAULT_BATCH_SIZE
    End If
    
    LogIfAble Me, "BatchSize", LogLevelInfo, "Batch size set to " & m_BatchSize, "PERFORMANCE"
End Property

Public Property Get BatchSize() As Long
    BatchSize = m_BatchSize
End Property

' --- New Batch Methods ---
Public Function ExecuteBulkInsert(ByVal tableName As String, ByVal data As Object, ByVal columnNames() As String) As Long
    ' Le paramètre data peut être soit un ADO Recordset, soit une Collection de tableaux
    Dim recordCount As Long
    Dim columnCount As Long
    Dim batchCount As Long
    Dim totalInserted As Long
    Dim sqlTemplate As String
    Dim currentBatchSize As Long
    Dim i As Long, j As Long, k As Long
    Dim currentRow As Long
    Dim batchParams As Collection
    Dim batchParamArray() As Variant
    Dim rs As ADODB.Recordset
    Dim isRecordset As Boolean
    
    ' Vérifications préliminaires
    If Not PreExecutionChecks(Me, "ExecuteBulkInsert") Then
        ExecuteBulkInsert = -1
        Exit Function
    End If
    
    ' Déterminer le type de données
    isRecordset = (TypeName(data) = "Recordset")
    
    ' Obtenir un compte approximatif des enregistrements
    If isRecordset Then
        Set rs = data
        On Error Resume Next
        recordCount = rs.RecordCount
        If Err.Number <> 0 Or recordCount < 0 Then recordCount = 100 ' Valeur par défaut
        On Error GoTo 0
    Else
        ' Supposons que c'est une Collection
        recordCount = data.Count
    End If
    
    columnCount = UBound(columnNames) - LBound(columnNames) + 1
    
    ' Créer le modèle SQL pour les lots, en utilisant le driver si disponible
    If Not m_Driver Is Nothing And m_Driver.SupportsBatchOperations() Then
        ' Utiliser la capacité du pilote pour l'insertion par lots
        currentBatchSize = m_BatchSize
        sqlTemplate = m_Driver.GetBatchInsertSql(tableName, columnNames, currentBatchSize)
    Else
        ' Créer une requête INSERT standard
        sqlTemplate = "INSERT INTO " & tableName & " ("
        
        ' Ajouter les noms de colonnes
        For i = LBound(columnNames) To UBound(columnNames)
            If i > LBound(columnNames) Then sqlTemplate = sqlTemplate & ", "
            sqlTemplate = sqlTemplate & columnNames(i)
        Next i
        
        sqlTemplate = sqlTemplate & ") VALUES ("
        
        ' Ajouter les paramètres
        For i = LBound(columnNames) To UBound(columnNames)
            If i > LBound(columnNames) Then sqlTemplate = sqlTemplate & ", "
            sqlTemplate = sqlTemplate & "?"
        Next i
        
        sqlTemplate = sqlTemplate & ")"
        
        ' Insertion ligne par ligne pour la compatibilité
        currentBatchSize = 1
    End If
    
    ' Calculer le nombre de lots
    batchCount = (recordCount + currentBatchSize - 1) \ currentBatchSize
    
    ' Traiter chaque lot
    totalInserted = 0
    currentRow = 0
    
    For i = 1 To batchCount
        ' Préparer les paramètres du lot
        Set batchParams = New Collection
        
        ' Définir la taille de ce lot (dernier lot peut être plus petit)
        Dim remainingRows As Long
        remainingRows = recordCount - currentRow
        If remainingRows < currentBatchSize Then
            currentBatchSize = remainingRows
        End If
        
        If currentBatchSize <= 0 Then
            Exit For ' Plus d'enregistrements à traiter
        End If
        
        ' Pour les drivers qui ne supportent pas les insertions par lots
        If m_Driver Is Nothing Or Not m_Driver.SupportsBatchOperations() Then
            ' Insérer ligne par ligne
            For j = 1 To currentBatchSize
                If isRecordset Then
                    ' Paramètres depuis Recordset
                    ReDim batchParamArray(LBound(columnNames) To UBound(columnNames))
                    For k = LBound(columnNames) To UBound(columnNames)
                        ' Créer un tableau de paramètres pour chaque colonne
                        batchParamArray(k) = Array(columnNames(k), rs.Fields(columnNames(k)).Value, GetFieldDataType(rs.Fields(columnNames(k))))
                    Next k
                    
                    ' Insérer cette ligne
                    Dim rowsInserted As Long
                    rowsInserted = IDbAccessorBase_ExecuteNonQuery(sqlTemplate, batchParamArray)
                    
                    If rowsInserted > 0 Then
                        totalInserted = totalInserted + rowsInserted
                    End If
                    
                    ' Passer à l'enregistrement suivant
                    If Not rs.EOF Then rs.MoveNext
                Else
                    ' Paramètres depuis Collection d'Arrays
                    Dim rowData As Variant
                    rowData = data(currentRow + j)
                    
                    ReDim batchParamArray(LBound(columnNames) To UBound(columnNames))
                    For k = LBound(columnNames) To UBound(columnNames)
                        ' Créer un tableau de paramètres pour chaque colonne
                        batchParamArray(k) = Array(columnNames(k), rowData(k - LBound(columnNames)), modFrameworkUtils.GuessAdoType(rowData(k - LBound(columnNames))))
                    Next k
                    
                    ' Insérer cette ligne
                    rowsInserted = IDbAccessorBase_ExecuteNonQuery(sqlTemplate, batchParamArray)
                    
                    If rowsInserted > 0 Then
                        totalInserted = totalInserted + rowsInserted
                    End If
                End If
            Next j
        Else
            ' Pour les drivers supportant les insertions par lots
            ' TODO: Implémentation de l'insertion par lots
        End If
        
        ' Mettre à jour la position actuelle
        currentRow = currentRow + currentBatchSize
    Next i
    
    ' Vérifier si des plugins sont intéressés par cet événement
    If m_UsePlugins And Not m_PluginManager Is Nothing Then
        Dim eventData As Object
        Set eventData = CreateObject("Scripting.Dictionary")
        eventData.Add "TableName", tableName
        eventData.Add "RowsInserted", totalInserted
        m_PluginManager.TriggerEvent "DB_BULK_INSERT_COMPLETE", eventData
    End If
    
    ExecuteBulkInsert = totalInserted
End Function

' --- Connection Management ---
Private Function IDbAccessorBase_Connect(ByVal connectionString As String, Optional ByVal maxRetries As Long = 1, Optional ByVal retryDelaySeconds As Long = 2) As Boolean
    Dim retryCount As Long
    Dim connectedSuccessfully As Boolean
    Dim startTime As Single
    Dim duration As Single
    
    ' Validate parameters
    If maxRetries < 1 Then maxRetries = 1
    If retryDelaySeconds < 1 Then retryDelaySeconds = 1
    
    ' Check for null or empty connection string
    If Len(Trim(connectionString)) = 0 Then
        LogIfAble Me, "Connect", LogLevelError, "Connection string is empty", "CONNECTION"
        m_LastError = -1
        IDbAccessorBase_Connect = False
        Exit Function
    End If
    
    LogIfAble Me, "Connect", LogLevelInfo, "Connecting (Retries=" & maxRetries & ")...", "CONNECTION"
    startTime = Timer
    
    On Error GoTo ConnectError
    
    ' Disconnect if already connected
    If m_isConnected Then
        IDbAccessorBase_Disconnect
    End If
    
    ' Prepare connection
    m_connString = connectionString
    Set m_cn = New ADODB.Connection
    m_cn.CursorLocation = adUseClient
    
    ' Set a reasonable command timeout
    m_cn.CommandTimeout = DB_OPERATION_TIMEOUT
    
    ' Try connection with retries
    For retryCount = 1 To maxRetries
        On Error Resume Next
        m_cn.ConnectionString = m_connString
        m_cn.Open
        m_LastError = Err.Number
        connectedSuccessfully = (Err.Number = 0 And m_cn.State = adStateOpen)
        On Error GoTo ConnectError
        
        If connectedSuccessfully Then
            If retryCount > 1 Then
                LogIfAble Me, "Connect", LogLevelWarning, "Connected after " & retryCount & " attempts.", "CONNECTION"
            End If
            Exit For
        End If
        
        ' Log specific connection error
        Dim errorDetail As String
        errorDetail = "Err " & Err.Number & ": " & Err.Description
        LogIfAble Me, "Connect", LogLevelWarning, "Attempt " & retryCount & "/" & maxRetries & " failed. " & errorDetail, "CONNECTION"
        
        ' Clear error
        Err.Clear
        
        ' Wait before next retry if more attempts remain
        If retryCount < maxRetries Then
            Application.Wait Now + TimeValue("00:00:" & Application.Max(1, retryDelaySeconds))
        End If
    Next retryCount
    
    ' Update state and return result
    m_isConnected = connectedSuccessfully
    
    ' Reset transaction state
    If m_isConnected Then
        m_isInTransaction = False
        m_TransactionNestLevel = 0
        
        ' Notify plugins of connection
        If m_UsePlugins And Not m_PluginManager Is Nothing Then
            Dim eventData As Object
            Set eventData = CreateObject("Scripting.Dictionary")
            eventData.Add "ConnectionString", ConnectionStringWithoutPassword(m_connString)
            m_PluginManager.TriggerEvent "DB_CONNECTED", eventData
        End If
    End If
    
    IDbAccessorBase_Connect = m_isConnected
    
    ' Calculate duration and log result
    duration = (Timer - startTime) * 1000
    
    If m_isConnected Then
        LogIfAble Me, "Connect", LogLevelInfo, "Connected (" & Round(duration, 0) & " ms).", "CONNECTION"
    Else
        HandleDbError Me, "Connect", Err, "Failed after " & maxRetries & " attempts."
        GoTo ConnectCleanFail
    End If
    
    Exit Function
    
ConnectError:
    m_LastError = Err.Number
    HandleDbError Me, "Connect", Err
    
ConnectCleanFail:
    m_isConnected = False
    m_isInTransaction = False
    m_TransactionNestLevel = 0
    IDbAccessorBase_Connect = False
    
    If Not m_cn Is Nothing Then
        On Error Resume Next
        If m_cn.State = adStateOpen Then m_cn.Close
        On Error GoTo 0
        Set m_cn = Nothing
    End If
End Function

Private Sub IDbAccessorBase_Disconnect()
    LogIfAble Me, "Disconnect", LogLevelInfo, "Disconnect requested.", "CONNECTION"
    
    On Error GoTo DisconnectError
    
    ' Notify plugins of disconnection
    If m_UsePlugins And Not m_PluginManager Is Nothing Then
        m_PluginManager.TriggerEvent "DB_DISCONNECTING", Nothing
    End If
    
    ' Clear statement cache
    ClearStatementCache
    
    If Not m_cn Is Nothing Then
        If m_cn.State = adStateOpen Then
            ' If transaction is active, roll it back first
            If m_isInTransaction Then
                LogIfAble Me, "Disconnect", LogLevelWarning, "Rolling back active transaction.", "TRANSACTION"
                SafeRollbackTransaction
            End If
            
            ' Close the connection
            m_cn.Close
        End If
        
        Set m_cn = Nothing
    End If
    
    ' Release the driver reference
    Set m_Driver = Nothing
    
    ' Update state
    m_isConnected = False
    m_isInTransaction = False
    m_TransactionNestLevel = 0
    
    LogIfAble Me, "Disconnect", LogLevelInfo, "Connection closed.", "CONNECTION"
    Exit Sub
    
DisconnectError:
    m_LastError = Err.Number
    HandleDbError Me, "Disconnect", Err, "Error during disconnect"
    
    ' Make sure we reset state
    m_isConnected = False
    m_isInTransaction = False
    m_TransactionNestLevel = 0
    
    ' Try really hard to release connection
    On Error Resume Next
    If Not m_cn Is Nothing Then
        If m_cn.State = adStateOpen Then m_cn.Close
        Set m_cn = Nothing
    End If
    
    Set m_Driver = Nothing
    On Error GoTo 0
End Sub

' --- Execution Methods ---
Private Function IDbAccessorBase_ExecuteNonQuery(ByVal sql As String, Optional ByVal params As Variant) As Long
    Dim cmd As ADODB.Command
    Dim recordsAffected As Long
    Dim startTime As Single
    Dim duration As Single
    Dim context As String
    Dim finalSql As String
    Dim cacheKey As String
    Dim fromCache As Boolean
    Dim eventDataBefore As Object
    Dim eventDataAfter As Object
    
    context = "ExecuteNonQuery"
    finalSql = sql
    IDbAccessorBase_ExecuteNonQuery = -1
    fromCache = False
    
    ' Validate SQL
    If Len(Trim(sql)) = 0 Then
        LogIfAble Me, context, LogLevelError, "SQL is empty", "SQL_ERROR"
        m_LastError = -1
        Exit Function
    End If
    
    ' Run pre-execution checks
    If Not PreExecutionChecks(Me, context) Then
        Exit Function
    End If
    
    ' Apply driver-specific SQL transformations if driver available
    If Not m_Driver Is Nothing Then
        ' TODO: Implement driver-specific SQL modifications
    End If
    
    ' Notify plugins before query execution
    If m_UsePlugins And Not m_PluginManager Is Nothing Then
        Set eventDataBefore = CreateObject("Scripting.Dictionary")
        eventDataBefore.Add "SQL", finalSql
        If Not IsMissing(params) Then eventDataBefore.Add "Params", params
        Dim eventResult As Collection
        Set eventResult = m_PluginManager.TriggerEvent("DB_BEFORE_NONQUERY", eventDataBefore)
        
        ' Si un plugin a renvoyé un résultat direct, on l'utilise
        If eventResult.Count > 0 Then
            Dim resultKey As Variant
            For Each resultKey In eventResult
                ' Vérifier si un plugin a retourné un résultat valide
                If Not IsNull(eventResult(resultKey)) Then
                    If IsNumeric(eventResult(resultKey)) Then
                        IDbAccessorBase_ExecuteNonQuery = CLng(eventResult(resultKey))
                        LogIfAble Me, context, LogLevelInfo, "Result provided by plugin '" & resultKey & "'", "PLUGIN"
                        Exit Function
                    End If
                End If
            Next resultKey
        End If
    End If
    
    LogIfAble Me, context, LogLevelDebug, "SQL: " & finalSql, "SQL"
    
    On Error GoTo ExecNonQueryError
    
    startTime = Timer
    
    ' Try to get prepared statement from cache if enabled
    If m_StatementCacheEnabled Then
        cacheKey = GetStatementCacheKey(finalSql, params)
        If m_StatementCache.Exists(cacheKey) Then
            Set cmd = m_StatementCache(cacheKey)
            fromCache = True
            LogIfAble Me, context, LogLevelDebug, "Using cached statement", "CACHE"
        End If
    End If
    
    ' Create new command if not from cache
    If cmd Is Nothing Then
        ' Create and configure command
        Set cmd = New ADODB.Command
        With cmd
            .ActiveConnection = m_cn
            .CommandText = finalSql
            .CommandType = adCmdText
            .CommandTimeout = DB_OPERATION_TIMEOUT
            
            ' Add parameters if provided
            If Not IsMissing(params) Then
                If Not AddParameters(Me, cmd, params) Then
                    LogIfAble Me, context, LogLevelError, "Failed to add parameters", "SQL_ERROR"
                    GoTo ExecNonQueryCleanExit
                End If
            End If
            
            ' Store in cache if enabled
            If m_StatementCacheEnabled And Not fromCache Then
                ' Don't let cache grow too large
                If m_StatementCache.Count >= MAX_PREPARED_STATEMENTS Then
                    ' Remove oldest statement (first key)
                    If m_StatementCache.Count > 0 Then
                        m_StatementCache.Remove m_StatementCache.Keys()(0)
                    End If
                End If
                
                ' Add to cache
                If Not m_StatementCache.Exists(cacheKey) Then
                    m_StatementCache.Add cacheKey, cmd
                End If
            End If
        End With
    End If
    
    ' Execute the command
    cmd.Execute recordsAffected, , adExecuteNoRecords
    
    ' Calculate duration
    duration = (Timer - startTime) * 1000
    
    ' Set result
    IDbAccessorBase_ExecuteNonQuery = recordsAffected
    
    ' Log and check for slow query
    LogIfAble Me, context, LogLevelInfo, "Executed (" & Round(duration, 0) & " ms). Affected: " & recordsAffected, "SQL_RESULT"
    CheckSlowQuery Me, context, duration, finalSql
    
    ' Notify plugins after query execution
    If m_UsePlugins And Not m_PluginManager Is Nothing Then
        Set eventDataAfter = CreateObject("Scripting.Dictionary")
        eventDataAfter.Add "SQL", finalSql
        If Not IsMissing(params) Then eventDataAfter.Add "Params", params
        eventDataAfter.Add "Duration", duration
        eventDataAfter.Add "Result", recordsAffected
        m_PluginManager.TriggerEvent "DB_AFTER_NONQUERY", eventDataAfter
    End If
    
ExecNonQueryCleanExit:
    ' Do not destroy cached commands
    If Not fromCache And Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    Exit Function
    
ExecNonQueryError:
    m_LastError = Err.Number
    duration = (Timer - startTime) * 1000
    HandleDbError Me, context, Err, "(Duration: " & Round(duration, 0) & " ms) SQL: " & finalSql
    Resume ExecNonQueryCleanExit
End Function

Private Function IDbAccessorBase_GetRecordset(ByVal sql As String, Optional ByVal params As Variant) As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim startTime As Single
    Dim duration As Single
    Dim context As String
    Dim finalSql As String
    Dim cacheKey As String
    Dim fromCache As Boolean
    Dim eventDataBefore As Object
    Dim eventDataAfter As Object
    
    context = "GetRecordset"
    finalSql = sql
    Set IDbAccessorBase_GetRecordset = Nothing
    fromCache = False
    
    ' Validate SQL
    If Len(Trim(sql)) = 0 Then
        LogIfAble Me, context, LogLevelError, "SQL is empty", "SQL_ERROR"
        m_LastError = -1
        Exit Function
    End If
    
    ' Run pre-execution checks
    If Not PreExecutionChecks(Me, context) Then
        Exit Function
    End If
    
    ' Apply driver-specific SQL transformations if driver available
    If Not m_Driver Is Nothing Then
        ' TODO: Implement driver-specific SQL modifications
    End If
    
    ' Notify plugins before query execution
    If m_UsePlugins And Not m_PluginManager Is Nothing Then
        Set eventDataBefore = CreateObject("Scripting.Dictionary")
        eventDataBefore.Add "SQL", finalSql
        If Not IsMissing(params) Then eventDataBefore.Add "Params", params
        Dim eventResult As Collection
        Set eventResult = m_PluginManager.TriggerEvent("DB_BEFORE_QUERY", eventDataBefore)
        
        ' Si un plugin a renvoyé un résultat direct, on l'utilise
        If eventResult.Count > 0 Then
            Dim resultKey As Variant
            For Each resultKey In eventResult
                ' Vérifier si un plugin a retourné un recordset valide
                If Not IsNull(eventResult(resultKey)) Then
                    If TypeName(eventResult(resultKey)) = "Recordset" Then
                        Set IDbAccessorBase_GetRecordset = eventResult(resultKey)
                        LogIfAble Me, context, LogLevelInfo, "Recordset provided by plugin '" & resultKey & "'", "PLUGIN"
                        Exit Function
                    End If
                End If
            Next resultKey
        End If
    End If
    
    LogIfAble Me, context, LogLevelDebug, "SQL: " & finalSql, "SQL"
    
    On Error GoTo GetRecordsetError
    
    startTime = Timer
    
    ' Try to get prepared statement from cache if enabled
    If m_StatementCacheEnabled Then
        cacheKey = GetStatementCacheKey(finalSql, params)
        If m_StatementCache.Exists(cacheKey) Then
            Set cmd = m_StatementCache(cacheKey)
            fromCache = True
            LogIfAble Me, context, LogLevelDebug, "Using cached statement", "CACHE"
        End If
    End If
    
    ' Create new command if not from cache
    If cmd Is Nothing Then
        ' Create and configure command
        Set cmd = New ADODB.Command
        With cmd
            .ActiveConnection = m_cn
            .CommandText = finalSql
            .CommandType = adCmdText
            .CommandTimeout = DB_OPERATION_TIMEOUT
            
            ' Add parameters if provided
            If Not IsMissing(params) Then
                If Not AddParameters(Me, cmd, params) Then
                    LogIfAble Me, context, LogLevelError, "Failed to add parameters", "SQL_ERROR"
                    GoTo GetRecordsetCleanExit
                End If
            End If
            
            ' Store in cache if enabled
            If m_StatementCacheEnabled And Not fromCache Then
                ' Don't let cache grow too large
                If m_StatementCache.Count >= MAX_PREPARED_STATEMENTS Then
                    ' Remove oldest statement (first key)
                    If m_StatementCache.Count > 0 Then
                        m_StatementCache.Remove m_StatementCache.Keys()(0)
                    End If
                End If
                
                ' Add to cache
                If Not m_StatementCache.Exists(cacheKey) Then
                    m_StatementCache.Add cacheKey, cmd
                End If
            End If
        End With
    End If
    
    ' Create and configure recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    ' Open recordset
    rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    ' Calculate duration
    duration = (Timer - startTime) * 1000
    
    ' Return the OPEN recordset (CALLER MUST CLOSE IT)
    Set IDbAccessorBase_GetRecordset = rs
    
    ' Log and check for slow query
    LogIfAble Me, context, LogLevelInfo, "Recordset opened (" & Round(duration, 0) & " ms). CALLER MUST CLOSE.", "SQL_RESULT"
    CheckSlowQuery Me, context, duration, finalSql
    
    ' Notify plugins after query execution
    If m_UsePlugins And Not m_PluginManager Is Nothing Then
        Set eventDataAfter = CreateObject("Scripting.Dictionary")
        eventDataAfter.Add "SQL", finalSql
        If Not IsMissing(params) Then eventDataAfter.Add "Params", params
        eventDataAfter.Add "Duration", duration
        eventDataAfter.Add "Result", rs
        m_PluginManager.TriggerEvent "DB_AFTER_QUERY", eventDataAfter
    End If
    
GetRecordsetCleanExit:
    ' Do not destroy cached commands
    If Not fromCache And Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    Exit Function ' DON'T CLOSE RS - It's returned to caller
    
GetRecordsetError:
    m_LastError = Err.Number
    duration = (Timer - startTime) * 1000
    HandleDbError Me, context, Err, "(Duration: " & Round(duration, 0) & " ms) SQL: " & finalSql
    
    ' Clean up if recordset was created
    If Not rs Is Nothing Then
        On Error Resume Next
        If rs.State = adStateOpen Then rs.Close
        On Error GoTo 0
        Set rs = Nothing
    End If
    
    Set IDbAccessorBase_GetRecordset = Nothing
    Resume GetRecordsetCleanExit
End Function

' --- Private Utility Methods ---
Private Function GetStatementCacheKey(ByVal sql As String, Optional ByVal params As Variant) As String
    ' Create a unique key for caching prepared statements
    Dim key As String
    key = UCase$(Trim$(sql))
    
    If Not IsMissing(params) And IsArray(params) Then
        key = key & "_" & UBound(params) - LBound(params) + 1 & "params"
    End If
    
    GetStatementCacheKey = key
End Function

Private Function ConnectionStringWithoutPassword(ByVal connString As String) As String
    ' Masque les mots de passe dans les chaînes de connexion pour la journalisation
    Dim result As String
    result = connString
    
    Dim pwdPatterns() As String
    pwdPatterns = Array("Password=", "PWD=", "Jet OLEDB:Database Password=")
    
    Dim i As Long
    For i = LBound(pwdPatterns) To UBound(pwdPatterns)
        Dim pattern As String
        pattern = pwdPatterns(i)
        
        Dim startPos As Long
        startPos = InStr(1, result, pattern, vbTextCompare)
        
        If startPos > 0 Then
            Dim endPos As Long
            endPos = InStr(startPos + Len(pattern), result, ";")
            
            If endPos = 0 Then
                endPos = Len(result) + 1
            End If
            
            result = Left$(result, startPos + Len(pattern) - 1) & "********" & Mid$(result, endPos)
        End If
    Next i
    
    ConnectionStringWithoutPassword = result
End Function

Private Function GetFieldDataType(field As ADODB.Field) As ADODB.DataTypeEnum
    ' Obtenir le type ADO d'un champ de recordset
    On Error Resume Next
    GetFieldDataType = field.Type
    If Err.Number <> 0 Then GetFieldDataType = adVariant
    On Error GoTo 0
End Function

' --- Transaction Methods - Modified for Plugin Support ---
Private Sub IDbAccessorBase_BeginTrans()
    If Not PreExecutionChecks(Me, "BeginTrans") Then Exit Sub
    
    ' Support for nested transactions
    If m_isInTransaction Then
        m_TransactionNestLevel = m_TransactionNestLevel + 1
        LogIfAble Me, "BeginTrans", LogLevelInfo, "Nested transaction started (level " & m_TransactionNestLevel & ")", "TRANSACTION"
        Exit Sub
    End If
    
    On Error GoTo BeginTransError
    
    ' Start actual transaction
    m_cn.BeginTrans
    m_isInTransaction = True
    m_TransactionNestLevel = 1
    
    ' Notify plugins
    If m_UsePlugins And Not m_PluginManager Is Nothing Then
        m_PluginManager.TriggerEvent "DB_TRANSACTION_BEGIN", Nothing
    End If
    
    LogIfAble Me, "BeginTrans", LogLevelInfo, "Transaction started.", "TRANSACTION"
    Exit Sub
    
BeginTransError:
    m_LastError = Err.Number
    m_isInTransaction = False
    m_TransactionNestLevel = 0
    HandleDbError Me, "BeginTrans", Err
End Sub

Private Sub IDbAccessorBase_CommitTrans()
    If Not PreExecutionChecks(Me, "CommitTrans") Then Exit Sub
    
    If Not m_isInTransaction Then
        HandleDbError Me, "CommitTrans", 0, "No active transaction"
        Exit Sub
    End If
    
    ' Handle nested transactions
    If m_TransactionNestLevel > 1 Then
        m_TransactionNestLevel = m_TransactionNestLevel - 1
        LogIfAble Me, "CommitTrans", LogLevelInfo, "Nested transaction committed (level " & m_TransactionNestLevel & " remaining)", "TRANSACTION"
        Exit Sub
    End If
    
    On Error GoTo CommitTransError
    
    ' Commit actual transaction
    m_cn.CommitTrans
    m_isInTransaction = False
    m_TransactionNestLevel = 0
    
    ' Notify plugins
    If m_UsePlugins And Not m_PluginManager Is Nothing Then
        m_PluginManager.TriggerEvent "DB_TRANSACTION_COMMIT", Nothing
    End If
    
    LogIfAble Me, "CommitTrans", LogLevelInfo, "Transaction committed.", "TRANSACTION"
    Exit Sub
    
CommitTransError:
    m_LastError = Err.Number
    HandleDbError Me, "CommitTrans", Err
    LogIfAble Me, "CommitTrans", LogLevelFatal, "Commit FAILED.", "TRANSACTION"
    
    ' Try to rollback on commit failure
    SafeRollbackTransaction
End Sub

Private Sub IDbAccessorBase_RollbackTrans()
    If Not PreExecutionChecks(Me, "RollbackTrans") Then Exit Sub
    
    If Not m_isInTransaction Then
        HandleDbError Me, "RollbackTrans", 0, "No active transaction"
        Exit Sub
    End If
    
    ' Handle nested transactions - a rollback at any level cancels the entire transaction
    On Error GoTo RollbackTransError
    
    ' Rollback actual transaction
    m_cn.RollbackTrans
    m_isInTransaction = False
    m_TransactionNestLevel = 0
    
    ' Notify plugins
    If m_UsePlugins And Not m_PluginManager Is Nothing Then
        m_PluginManager.TriggerEvent "DB_TRANSACTION_ROLLBACK", Nothing
    End If
    
    LogIfAble Me, "RollbackTrans", LogLevelWarning, "Transaction rolled back.", "TRANSACTION"
    Exit Sub
    
RollbackTransError:
    m_LastError = Err.Number
    HandleDbError Me, "RollbackTrans", Err
    m_isInTransaction = False ' Reset state even on error
    m_TransactionNestLevel = 0
End Sub

Private Sub SafeRollbackTransaction()
    ' Helper method to safely roll back a transaction without raising errors
    On Error Resume Next
    m_cn.RollbackTrans
    m_isInTransaction = False
    m_TransactionNestLevel = 0
    
    ' Notify plugins
    If m_UsePlugins And Not m_PluginManager Is Nothing Then
        m_PluginManager.TriggerEvent "DB_TRANSACTION_ROLLBACK", Nothing
    End If
    
    If Err.Number <> 0 Then
        LogIfAble Me, "SafeRollbackTransaction", LogLevelError, "Error during rollback: " & Err.Description, "TRANSACTION"
    Else
        LogIfAble Me, "SafeRollbackTransaction", LogLevelWarning, "Transaction rolled back (safe).", "TRANSACTION"
    End If
    On Error GoTo 0
End Sub

' --- Class Cleanup ---
Private Sub Class_Terminate()
    LogIfAble Me, "Class_Terminate", LogLevelDebug, "Terminating instance...", "SYSTEM"
    
    ' Notify plugins of termination
    If m_UsePlugins And Not m_PluginManager Is Nothing Then
        m_PluginManager.TriggerEvent "DB_ACCESSOR_TERMINATE", Nothing
    End If
    
    IDbAccessorBase_Disconnect
    ClearStatementCache
    
    Set m_Logger = Nothing
    Set m_PluginManager = Nothing
    Set m_StatementCache = Nothing
End Sub 