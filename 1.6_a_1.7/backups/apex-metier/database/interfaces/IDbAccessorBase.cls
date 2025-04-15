'@Folder("APEX.Core.Interfaces.Database")
Option Explicit

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team


'==========================================================================
' Interface : IDbAccessorBase
' Purpose   : Interface de base pour l'accès aux données
' Author    : APEX Framework Team
' Date      : 2024-04-11
' Reference : IMPL-002
'==========================================================================

'@Description("Initialise l'accesseur avec la configuration fournie")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Initialize(Optional ByVal configSource As Variant)
End Sub

'@Description("Définit le logger à utiliser")
'@Description: 
'@Param: 
'@Returns: 

Public Sub SetLogger(ByVal loggerInstance As ILoggerBase)
End Sub

'@Description("Définit le driver de base de données")
'@Description: 
'@Param: 
'@Returns: 

Public Sub SetDriver(ByVal dbDriver As IDbDriver)
End Sub

'@Description("Établit la connexion avec la base de données")
'@Description: 
'@Param: 
'@Returns: 

Public Function Connect(ByVal connectionString As String, _
                       Optional ByVal maxRetries As Long = 1, _
                       Optional ByVal retryDelaySeconds As Long = 2) As Boolean
End Function

'@Description("Ferme la connexion avec la base de données")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Disconnect()
End Sub

'@Description("Indique si la connexion est active")
Public Property Get IsConnected() As Boolean
End Property

'@Description("Exécute une requête sans retour de données")
'@Description: 
'@Param: 
'@Returns: 

Public Function ExecuteNonQuery(ByVal sql As String, Optional ByVal params As Variant) As Long
End Function

'@Description("Exécute une requête et retourne un recordset")
'@Description: 
'@Param: 
'@Returns: 

Public Function GetRecordset(ByVal sql As String, Optional ByVal params As Variant) As ADODB.Recordset
End Function

'@Description("Exécute une requête et retourne une valeur unique")
'@Description: 
'@Param: 
'@Returns: 

Public Function ExecuteScalar(ByVal sql As String, Optional ByVal params As Variant) As Variant
End Function

'@Description("Exécute une procédure stockée et retourne un recordset")
'@Description: 
'@Param: 
'@Returns: 

Public Function ExecuteStoredProcedure(ByVal procName As String, Optional ByVal params As Variant) As ADODB.Recordset
End Function

'@Description("Exécute une procédure stockée sans retour de données")
'@Description: 
'@Param: 
'@Returns: 

Public Function ExecuteStoredProcedureNonQuery(ByVal procName As String, Optional ByVal params As Variant) As Long
End Function

'@Description("Retourne une instance du constructeur de requêtes")
'@Description: 
'@Param: 
'@Returns: 

Public Function QueryBuilder() As IQueryBuilder
End Function

'@Description("Démarre une transaction")
'@Description: 
'@Param: 
'@Returns: 

Public Sub BeginTrans()
End Sub

'@Description("Valide la transaction en cours")
'@Description: 
'@Param: 
'@Returns: 

Public Sub CommitTrans()
End Sub

'@Description("Annule la transaction en cours")
'@Description: 
'@Param: 
'@Returns: 

Public Sub RollbackTrans()
End Sub

'@Description("Indique si une transaction est en cours")
Public Property Get IsInTransaction() As Boolean
End Property

'@Description("Configure les options de connexion")
'@Description: 
'@Param: 
'@Returns: 

Public Sub ConfigureConnection(ByVal timeout As Long, _
                             Optional ByVal enablePooling As Boolean = True, _
                             Optional ByVal maxPoolSize As Long = 100)
End Sub

'@Description("Valide une requête sans l'exécuter")
'@Description: 
'@Param: 
'@Returns: 

Public Function ValidateQuery(ByVal sql As String) As Boolean
End Function

'@Description("Retourne les métriques de performance")
'@Description: 
'@Param: 
'@Returns: 

Public Function GetPerformanceMetrics() As Collection
End Function 