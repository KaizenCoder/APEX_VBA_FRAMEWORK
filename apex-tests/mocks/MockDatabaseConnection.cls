VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockDatabaseConnection"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("APEX.Tests.Mocks")
Option Explicit

Private Type TMockSettings
    ErrorRate As Double
    Latency As Long
    EnableDeadlocks As Boolean
    LastError As String
    QueryCount As Long
    SuccessCount As Long
    ErrorCount As Long
    DeadlockCount As Long
    TimeoutCount As Long
End Type

Private this As TMockSettings
Private Const TIMEOUT_THRESHOLD As Long = 5000 'ms

'@Description("Initialise la connexion mock")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Initialize()
    With this
        .ErrorRate = 0
        .Latency = 0
        .EnableDeadlocks = False
        .LastError = ""
        .QueryCount = 0
        .SuccessCount = 0
        .ErrorCount = 0
        .DeadlockCount = 0
        .TimeoutCount = 0
    End With
End Sub

'@Description("Configure le taux d'erreur (0-1)")
'@Description: 
'@Param: 
'@Returns: 

Public Sub SetErrorRate(ByVal rate As Double)
    If rate < 0 Or rate > 1 Then
        Err.Raise 5, "MockDatabaseConnection", "Taux d'erreur invalide (doit être entre 0 et 1)"
    End If
    this.ErrorRate = rate
End Sub

'@Description("Configure la latence en millisecondes")
'@Description: 
'@Param: 
'@Returns: 

Public Sub SetLatency(ByVal latencyMs As Long)
    If latencyMs < 0 Then
        Err.Raise 5, "MockDatabaseConnection", "Latence invalide"
    End If
    this.Latency = latencyMs
End Sub

'@Description("Active/désactive la simulation de deadlocks")
'@Description: 
'@Param: 
'@Returns: 

Public Sub EnableDeadlockSimulation(ByVal enable As Boolean)
    this.EnableDeadlocks = enable
End Sub

'@Description("Exécute une requête simulée")
'@Description: 
'@Param: 
'@Returns: 

Public Function ExecuteQuery(ByVal query As String) As Variant
    this.QueryCount = this.QueryCount + 1
    
    'Simuler la latence
    If this.Latency > 0 Then
        Sleep this.Latency
    End If
    
    'Vérifier le timeout
    If this.Latency > TIMEOUT_THRESHOLD Then
        this.TimeoutCount = this.TimeoutCount + 1
        this.LastError = "Timeout après " & this.Latency & "ms"
        Err.Raise vbObjectError + 1, "MockDatabaseConnection", this.LastError
    End If
    
    'Simuler les erreurs aléatoires
    If Rnd() < this.ErrorRate Then
        this.ErrorCount = this.ErrorCount + 1
        
        'Simuler différents types d'erreurs
        Select Case Int(Rnd() * 3)
            Case 0
                this.LastError = "Erreur de connexion"
                Err.Raise vbObjectError + 2, "MockDatabaseConnection", this.LastError
            Case 1
                this.LastError = "Erreur de syntaxe SQL"
                Err.Raise vbObjectError + 3, "MockDatabaseConnection", this.LastError
            Case 2
                If this.EnableDeadlocks Then
                    this.DeadlockCount = this.DeadlockCount + 1
                    this.LastError = "Deadlock détecté"
                    Err.Raise vbObjectError + 4, "MockDatabaseConnection", this.LastError
                Else
                    this.LastError = "Erreur d'exécution"
                    Err.Raise vbObjectError + 5, "MockDatabaseConnection", this.LastError
                End If
        End Select
    End If
    
    'Simuler un résultat réussi
    this.SuccessCount = this.SuccessCount + 1
    ExecuteQuery = Array("Résultat simulé pour: " & query)
End Function

'@Description("Retourne les statistiques de la connexion mock")
'@Description: 
'@Param: 
'@Returns: 

Public Function GetStats() As String
    Dim stats As String
    stats = "Statistiques de la connexion mock:" & vbNewLine & _
            "- Requêtes totales: " & this.QueryCount & vbNewLine & _
            "- Succès: " & this.SuccessCount & vbNewLine & _
            "- Erreurs: " & this.ErrorCount & vbNewLine & _
            "- Deadlocks: " & this.DeadlockCount & vbNewLine & _
            "- Timeouts: " & this.TimeoutCount & vbNewLine & _
            "- Dernière erreur: " & this.LastError
    GetStats = stats
End Function

'@Description("Simule un délai en millisecondes")
'@Description: 
'@Param: 
'@Returns: 

Private Sub Sleep(ByVal milliseconds As Long)
    Dim endTime As Date
    endTime = DateAdd("s", milliseconds / 1000#, Now())
    Do While Now() < endTime
        DoEvents
    Loop
End Sub 