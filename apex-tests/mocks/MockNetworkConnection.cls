VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockNetworkConnection"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("APEX.Tests.Mocks")
Option Explicit

Private Type TConnectionState
    IsConnected As Boolean
    LatencyMs As Long
    PacketLossRate As Double
    ErrorRate As Double
    LastError As String
    TimeoutMs As Long
    RetryCount As Long
    MaxRetries As Long
End Type

Private Type TConnectionStats
    TotalRequests As Long
    SuccessfulRequests As Long
    FailedRequests As Long
    TimeoutRequests As Long
    TotalLatency As Double
    MaxLatency As Long
    MinLatency As Long
    LastRequestTime As Date
End Type

Private this As TConnectionState
Private m_Stats As TConnectionStats
Private m_Logger As ILoggerBase
Private m_MetricsCollector As clsMetricsCollector

'@Description("Initialise la connexion mock")
'@Description: 
'@Param: 
'@Returns: 

Public Function Initialize(Optional ByVal logger As ILoggerBase = Nothing, _
                         Optional ByVal metricsCollector As clsMetricsCollector = Nothing) As Boolean
    Set m_Logger = logger
    Set m_MetricsCollector = metricsCollector
    
    ' Configuration par défaut
    With this
        .IsConnected = True
        .LatencyMs = 50
        .PacketLossRate = 0
        .ErrorRate = 0
        .TimeoutMs = 5000
        .RetryCount = 0
        .MaxRetries = 3
    End With
    
    Initialize = True
End Function

'@Description("Configure les paramètres de simulation")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Configure(ByVal latencyMs As Long, _
                    ByVal packetLossRate As Double, _
                    ByVal errorRate As Double, _
                    ByVal timeoutMs As Long)
    this.LatencyMs = latencyMs
    this.PacketLossRate = packetLossRate
    this.ErrorRate = errorRate
    this.TimeoutMs = timeoutMs
End Sub

'@Description("Simule une requête réseau")
'@Description: 
'@Param: 
'@Returns: 

Public Function SendRequest(ByVal request As String) As String
    On Error GoTo ErrorHandler
    
    ' Incrémenter les statistiques
    m_Stats.TotalRequests = m_Stats.TotalRequests + 1
    m_Stats.LastRequestTime = Now
    
    ' Simuler la latence
    Sleep this.LatencyMs
    
    ' Mettre à jour les métriques de latence
    UpdateLatencyStats this.LatencyMs
    
    ' Simuler la perte de paquets
    If SimulatePacketLoss Then
        this.LastError = "Packet loss"
        RaiseError 35001, "Simulated packet loss"
    End If
    
    ' Simuler les erreurs
    If SimulateError Then
        this.LastError = "Network error"
        RaiseError 35002, "Simulated network error"
    End If
    
    ' Simuler le timeout
    If this.LatencyMs > this.TimeoutMs Then
        this.LastError = "Timeout"
        RaiseError 35003, "Request timeout"
    End If
    
    ' Succès
    m_Stats.SuccessfulRequests = m_Stats.SuccessfulRequests + 1
    SendRequest = "SUCCESS: " & request
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_Stats.FailedRequests = m_Stats.FailedRequests + 1
    If Err.Number = 35003 Then
        m_Stats.TimeoutRequests = m_Stats.TimeoutRequests + 1
    End If
    
    LogError "SendRequest", Err.Description
    
    ' Tentative de retry
    If this.RetryCount < this.MaxRetries Then
        this.RetryCount = this.RetryCount + 1
        Resume
    End If
    
    SendRequest = "ERROR: " & Err.Description
End Function

'@Description("Simule la perte de paquets")
'@Description: 
'@Param: 
'@Returns: 

Private Function SimulatePacketLoss() As Boolean
    SimulatePacketLoss = (Rnd < this.PacketLossRate)
End Function

'@Description("Simule les erreurs réseau")
'@Description: 
'@Param: 
'@Returns: 

Private Function SimulateError() As Boolean
    SimulateError = (Rnd < this.ErrorRate)
End Function

'@Description("Met à jour les statistiques de latence")
'@Description: 
'@Param: 
'@Returns: 

Private Sub UpdateLatencyStats(ByVal latency As Long)
    With m_Stats
        .TotalLatency = .TotalLatency + latency
        
        If latency > .MaxLatency Then
            .MaxLatency = latency
        End If
        
        If .MinLatency = 0 Or latency < .MinLatency Then
            .MinLatency = latency
        End If
    End With
    
    ' Enregistrer les métriques
    If Not m_MetricsCollector Is Nothing Then
        m_MetricsCollector.RecordMetric "NetworkLatency", latency
    End If
End Sub

'@Description("Obtient les statistiques de connexion")
'@Description: 
'@Param: 
'@Returns: 

Public Function GetStats() As String
    Dim avgLatency As Double
    If m_Stats.TotalRequests > 0 Then
        avgLatency = m_Stats.TotalLatency / m_Stats.TotalRequests
    End If
    
    GetStats = "=== Statistiques Réseau ===" & vbNewLine & _
               "Total requêtes: " & m_Stats.TotalRequests & vbNewLine & _
               "Réussies: " & m_Stats.SuccessfulRequests & vbNewLine & _
               "Échouées: " & m_Stats.FailedRequests & vbNewLine & _
               "Timeouts: " & m_Stats.TimeoutRequests & vbNewLine & _
               "Latence moyenne: " & Format(avgLatency, "0.000") & " ms" & vbNewLine & _
               "Latence max: " & m_Stats.MaxLatency & " ms" & vbNewLine & _
               "Latence min: " & m_Stats.MinLatency & " ms" & vbNewLine & _
               "Dernière requête: " & Format(m_Stats.LastRequestTime, "yyyy-mm-dd hh:nn:ss")
End Function

'@Description("Journalise une erreur")
'@Description: 
'@Param: 
'@Returns: 

Private Sub LogError(ByVal method As String, ByVal description As String)
    If Not m_Logger Is Nothing Then
        m_Logger.Log "ERROR", "MockNetwork." & method, description
    End If
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub Sleep(ByVal milliseconds As Long)
    Dim endTime As Date
    endTime = DateAdd("s", milliseconds / 1000#, Now())
    Do While Now() < endTime
        DoEvents
    Loop
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub RaiseError(ByVal number As Long, ByVal description As String)
    Err.Raise number, "MockNetworkConnection", description
End Sub 