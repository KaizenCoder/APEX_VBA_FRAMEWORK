' Migrated to apex-metier/database/interfaces - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit
' ==========================================================================
' Interface : IDbAccessorBase
' Version : 3.1 - Added SetDriver
' Purpose : Defines the standard contract for ADO database accessor implementations.
' ==========================================================================

' --- Initialization & Configuration ---
Public Sub Initialize(Optional ByVal configSource As Variant): Err.Raise 513, "IDbAccessorBase": End Sub
Public Sub SetLogger(ByVal loggerInstance As ILoggerBase): Err.Raise 513, "IDbAccessorBase": End Sub
Public Sub SetDriver(ByVal dbDriver As IDbDriver) ' NEW
Err.Raise vbObjectError + 1101, "IDbAccessorBase", "SetDriver method not implemented."
End Sub

' --- Connection ---
Public Function Connect(ByVal connectionString As String, Optional ByVal maxRetries As Long = 1, Optional ByVal retryDelaySeconds As Long = 2) As Boolean: Err.Raise 513, "IDbAccessorBase": End Function
Public Sub Disconnect(): Err.Raise 513, "IDbAccessorBase": End Sub
Public Property Get IsConnected() As Boolean: Err.Raise 513, "IDbAccessorBase": End Property

' --- Execution ---
Public Function ExecuteNonQuery(ByVal sql As String, Optional ByVal params As Variant) As Long: Err.Raise 513, "IDbAccessorBase": End Function
Public Function GetRecordset(ByVal sql As String, Optional ByVal params As Variant) As ADODB.Recordset: Err.Raise 513, "IDbAccessorBase": End Function

' --- Helper Methods ---
Public Function ExecuteScalar(ByVal sql As String, Optional ByVal params As Variant) As Variant: Err.Raise 513, "IDbAccessorBase": End Function
Public Function ExecuteStoredProcedure(ByVal procName As String, Optional ByVal params As Variant) As ADODB.Recordset: Err.Raise 513, "IDbAccessorBase": End Function
Public Function ExecuteStoredProcedureNonQuery(ByVal procName As String, Optional ByVal params As Variant) As Long: Err.Raise 513, "IDbAccessorBase": End Function

' --- Query Builder Factory ---
Public Function QueryBuilder() As IQueryBuilder: Err.Raise 513, "IDbAccessorBase": End Function

' --- Transactions ---
Public Sub BeginTrans(): Err.Raise 513, "IDbAccessorBase": End Sub
Public Sub CommitTrans(): Err.Raise 513, "IDbAccessorBase": End Sub
Public Sub RollbackTrans(): Err.Raise 513, "IDbAccessorBase": End Sub
Public Property Get IsInTransaction() As Boolean: Err.Raise 513, "IDbAccessorBase": End Property 