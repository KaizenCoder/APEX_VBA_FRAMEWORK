VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConnectionPool"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("APEX.Metier.Database.Connection")
Option Explicit

'==========================================================================
' Classe    : clsConnectionPool
' Purpose   : Gestion du pool de connexions
' Author    : APEX Framework Team
' Date      : 2024-04-11
' Reference : PERF-001
'==========================================================================

Private Type TConnectionPoolSettings
    MinSize As Long
    MaxSize As Long
    IdleTimeout As Long
    MaxLifetime As Long
    CurrentSize As Long
    BusyConnections As Long
    LastCleanup As Date
End Type

Private Type TPooledConnection
    Connection As ADODB.Connection
    Created As Date
    LastUsed As Date
    InUse As Boolean
End Type

Private this As TConnectionPoolSettings
Private m_Connections As Collection
Private m_Logger As ILoggerBase
Private m_SecurityManager As clsSecurityManager

'@Description("Initialise le pool de connexions")
'@Description: 
'@Param: 
'@Returns: 

Public Function Initialize(ByVal securityManager As clsSecurityManager, _
                         Optional ByVal logger As ILoggerBase = Nothing) As Boolean
    ' Configuration par défaut
    With this
        .MinSize = 5
        .MaxSize = 20
        .IdleTimeout = 300 ' 5 minutes
        .MaxLifetime = 3600 ' 1 heure
        .CurrentSize = 0
        .BusyConnections = 0
        .LastCleanup = Now
    End With
    
    ' Initialisation des dépendances
    Set m_SecurityManager = securityManager
    If Not logger Is Nothing Then Set m_Logger = logger
    
    ' Création du pool initial
    Set m_Connections = New Collection
    Initialize = True
End Function

'@Description("Obtient une connexion du pool")
'@Description: 
'@Param: 
'@Returns: 

Public Function GetConnection(ByVal connectionString As String) As ADODB.Connection
    ' Nettoyage si nécessaire
    If NeedsCleanup Then CleanupPool
    
    ' Recherche d'une connexion disponible
    Dim conn As TPooledConnection
    conn = FindAvailableConnection
    
    ' Si aucune connexion disponible, en créer une nouvelle
    If conn.Connection Is Nothing Then
        If CanCreateNewConnection Then
            conn = CreateNewConnection(connectionString)
        Else
            LogPoolEvent "Pool de connexions saturé", "ERROR"
            Err.Raise 5, "clsConnectionPool", "Pool de connexions saturé"
        End If
    End If
    
    ' Marquer comme utilisée
    conn.InUse = True
    conn.LastUsed = Now
    this.BusyConnections = this.BusyConnections + 1
    
    Set GetConnection = conn.Connection
End Function

'@Description("Libère une connexion")
'@Description: 
'@Param: 
'@Returns: 

Public Sub ReleaseConnection(ByVal connection As ADODB.Connection)
    Dim i As Long
    For i = 1 To m_Connections.Count
        Dim conn As TPooledConnection
        conn = m_Connections(i)
        
        If conn.Connection Is connection Then
            conn.InUse = False
            conn.LastUsed = Now
            this.BusyConnections = this.BusyConnections - 1
            Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
        End If
    Next i
End Sub

'@Description("Nettoie le pool de connexions")
'@Description: 
'@Param: 
'@Returns: 

Private Sub CleanupPool()
    Dim i As Long
    For i = m_Connections.Count To 1 Step -1
        Dim conn As TPooledConnection
        conn = m_Connections(i)
        
        ' Fermer les connexions inactives ou trop anciennes
        If ShouldCloseConnection(conn) Then
            CloseConnection i
        End If
    Next i
    
    this.LastCleanup = Now
    LogPoolEvent "Nettoyage du pool effectué", "INFO"
End Sub

'@Description("Vérifie si une connexion doit être fermée")
'@Description: 
'@Param: 
'@Returns: 

Private Function ShouldCloseConnection(ByVal conn As TPooledConnection) As Boolean
    If conn.InUse Then Exit Function
    
    ' Vérifier le timeout d'inactivité
    If DateDiff("s", conn.LastUsed, Now) > this.IdleTimeout Then
        ShouldCloseConnection = True
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Vérifier la durée de vie maximale
    If DateDiff("s", conn.Created, Now) > this.MaxLifetime Then
        ShouldCloseConnection = True
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
End Function

'@Description("Ferme une connexion spécifique")
'@Description: 
'@Param: 
'@Returns: 

Private Sub CloseConnection(ByVal index As Long)
    Dim conn As TPooledConnection
    conn = m_Connections(index)
    
    If Not conn.Connection Is Nothing Then
        If conn.Connection.State = adStateOpen Then
            conn.Connection.Close
        End If
        Set conn.Connection = Nothing
    End If
    
    m_Connections.Remove index
    this.CurrentSize = this.CurrentSize - 1
End Sub

'@Description("Trouve une connexion disponible")
'@Description: 
'@Param: 
'@Returns: 

Private Function FindAvailableConnection() As TPooledConnection
    Dim i As Long
    For i = 1 To m_Connections.Count
        Dim conn As TPooledConnection
        conn = m_Connections(i)
        
        If Not conn.InUse Then
            FindAvailableConnection = conn
            Exit'@Description: 
'@Param: 
'@Returns: 

 Function
        End If
    Next i
End Function

'@Description("Crée une nouvelle connexion")
'@Description: 
'@Param: 
'@Returns: 

Private Function CreateNewConnection(ByVal connectionString As String) As TPooledConnection
    Dim conn As TPooledConnection
    
    ' Créer la connexion
    Set conn.Connection = New ADODB.Connection
    conn.Connection.ConnectionString = m_SecurityManager.DecryptConnectionString(connectionString)
    conn.Connection.Open
    
    ' Initialiser les métadonnées
    conn.Created = Now
    conn.LastUsed = Now
    conn.InUse = False
    
    ' Ajouter au pool
    m_Connections.Add conn
    this.CurrentSize = this.CurrentSize + 1
    
    CreateNewConnection = conn
    LogPoolEvent "Nouvelle connexion créée", "INFO"
End Function

'@Description("Vérifie si le pool peut créer une nouvelle connexion")
'@Description: 
'@Param: 
'@Returns: 

Private Function CanCreateNewConnection() As Boolean
    CanCreateNewConnection = (this.CurrentSize < this.MaxSize)
End Function

'@Description("Vérifie si le pool nécessite un nettoyage")
'@Description: 
'@Param: 
'@Returns: 

Private Function NeedsCleanup() As Boolean
    ' Nettoyer toutes les 5 minutes
    NeedsCleanup = (DateDiff("s", this.LastCleanup, Now) > 300)
End Function

'@Description("Journalise un événement du pool")
'@Description: 
'@Param: 
'@Returns: 

Private Sub LogPoolEvent(ByVal message As String, ByVal level As String)
    If m_Logger Is Nothing Then Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    
    m_Logger.LogEvent _
        component:="ConnectionPool", _
        eventType:="POOL", _
        message:=message & " [Size:" & this.CurrentSize & "/" & this.MaxSize & "]", _
        level:=level
End Sub

' Propriétés de configuration
'@Description("Définit la taille minimale du pool")
Public Property Let MinSize(ByVal value As Long)
    this.MinSize = value
End Property

'@Description("Définit la taille maximale du pool")
Public Property Let MaxSize(ByVal value As Long)
    this.MaxSize = value
End Property

'@Description("Définit le timeout d'inactivité")
Public Property Let IdleTimeout(ByVal value As Long)
    this.IdleTimeout = value
End Property

'@Description("Définit la durée de vie maximale")
Public Property Let MaxLifetime(ByVal value As Long)
    this.MaxLifetime = value
End Property

'@Description("Retourne les statistiques du pool")
'@Description: 
'@Param: 
'@Returns: 

Public Function GetStatistics() As Collection
    Dim stats As New Collection
    With stats
        .Add this.CurrentSize, "CurrentSize"
        .Add this.BusyConnections, "BusyConnections"
        .Add this.MaxSize, "MaxSize"
        .Add this.MinSize, "MinSize"
        .Add this.IdleTimeout, "IdleTimeout"
        .Add this.MaxLifetime, "MaxLifetime"
    End With
    Set GetStatistics = stats
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Sub Class_Terminate()
    ' Fermer toutes les connexions
    Dim i As Long
    For i = m_Connections.Count To 1 Step -1
        CloseConnection i
    Next i
    
    Set m_Connections = Nothing
    Set m_Logger = Nothing
    Set m_SecurityManager = Nothing
End Sub 