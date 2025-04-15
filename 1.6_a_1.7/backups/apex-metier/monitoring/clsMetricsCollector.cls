VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMetricsCollector"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("APEX.Metier.Monitoring")
Option Explicit

'==========================================================================
' Classe    : clsMetricsCollector
' Purpose   : Collecte et analyse des métriques de performance
' Author    : APEX Framework Team
' Date      : 2024-04-11
' Reference : PERF-002
'==========================================================================

Private Type TMetrics
    QueryCount As Long
    AvgExecutionTime As Double
    ErrorCount As Long
    ConnectionUsage As Double
    LastUpdate As Date
    PoolSize As Long
    ActiveConnections As Long
    QueuedRequests As Long
    CacheHitRate As Double
    MemoryUsage As Double
End Type

Private Type TPerformanceThresholds
    MaxQueryTime As Double
    MaxPoolSize As Long
    MaxQueueSize As Long
    MinCacheHitRate As Double
    MaxMemoryUsage As Double
End Type

Private this As TMetrics
Private m_Thresholds As TPerformanceThresholds
Private m_Logger As ILoggerBase
Private m_MetricsHistory As Collection

'@Description("Initialise le collecteur de métriques")
'@Description: 
'@Param: 
'@Returns: 

Public Function Initialize(Optional ByVal logger As ILoggerBase = Nothing) As Boolean
    ' Configuration par défaut
    With m_Thresholds
        .MaxQueryTime = 5000 ' 5 secondes
        .MaxPoolSize = 50
        .MaxQueueSize = 100
        .MinCacheHitRate = 0.7 ' 70%
        .MaxMemoryUsage = 1024 * 1024 * 512 ' 512 MB
    End With
    
    ' Initialisation des collections
    Set m_MetricsHistory = New Collection
    
    ' Logger
    If Not logger Is Nothing Then Set m_Logger = logger
    
    Initialize = True
End Function

'@Description("Enregistre une nouvelle mesure de performance")
'@Description: 
'@Param: 
'@Returns: 

Public Sub RecordMetric(ByVal metricType As String, ByVal value As Variant)
    Select Case metricType
        Case "QueryExecution"
            UpdateQueryMetrics CDbl(value)
        Case "ConnectionUsage"
            UpdateConnectionMetrics CDbl(value)
        Case "Error"
            UpdateErrorMetrics
        Case "CacheHit"
            UpdateCacheMetrics CBool(value)
        Case "MemoryUsage"
            UpdateMemoryMetrics CDbl(value)
    End Select
    
    ' Vérification des seuils
    CheckThresholds
    
    ' Sauvegarde historique
    SaveMetricsSnapshot
End Sub

'@Description("Obtient un rapport de performance")
'@Description: 
'@Param: 
'@Returns: 

Public Function GetPerformanceReport() As String
    Dim report As String
    
    report = "=== Rapport de Performance ===" & vbNewLine & _
             "Date: " & Format(Now, "yyyy-mm-dd hh:mm:ss") & vbNewLine & _
             vbNewLine & _
             "1. Métriques de Requêtes" & vbNewLine & _
             "   - Nombre total: " & this.QueryCount & vbNewLine & _
             "   - Temps moyen: " & Format(this.AvgExecutionTime, "0.00") & " ms" & vbNewLine & _
             vbNewLine & _
             "2. Connexions" & vbNewLine & _
             "   - Utilisation: " & Format(this.ConnectionUsage * 100, "0.0") & "%" & vbNewLine & _
             "   - Pool actif: " & this.PoolSize & vbNewLine & _
             "   - Connexions actives: " & this.ActiveConnections & vbNewLine & _
             vbNewLine & _
             "3. Cache" & vbNewLine & _
             "   - Taux de succès: " & Format(this.CacheHitRate * 100, "0.0") & "%" & vbNewLine & _
             vbNewLine & _
             "4. Ressources" & vbNewLine & _
             "   - Mémoire utilisée: " & Format(this.MemoryUsage / 1024 / 1024, "0.0") & " MB" & vbNewLine & _
             "   - Erreurs: " & this.ErrorCount
    
    GetPerformanceReport = report
End Function

'@Description("Exporte les métriques vers Excel")
'@Description: 
'@Param: 
'@Returns: 

Public Sub ExportMetricsToExcel(ByVal targetWorkbook As Workbook)
    Dim ws As Worksheet
    
    ' Créer ou obtenir la feuille de métriques
    On Error Resume Next
    Set ws = targetWorkbook.Sheets("Performance_Metrics")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = targetWorkbook.Sheets.Add
        ws.Name = "Performance_Metrics"
        
        ' En-têtes
        ws.Range("A1").value = "Date"
        ws.Range("B1").value = "Requêtes"
        ws.Range("C1").value = "Temps Moyen (ms)"
        ws.Range("D1").value = "Utilisation Connexions"
        ws.Range("E1").value = "Taux Cache"
        ws.Range("F1").value = "Mémoire (MB)"
        ws.Range("G1").value = "Erreurs"
    End If
    
    ' Ajouter les données
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ws.Cells(nextRow, 1).value = Now
    ws.Cells(nextRow, 2).value = this.QueryCount
    ws.Cells(nextRow, 3).value = this.AvgExecutionTime
    ws.Cells(nextRow, 4).value = this.ConnectionUsage
    ws.Cells(nextRow, 5).value = this.CacheHitRate
    ws.Cells(nextRow, 6).value = this.MemoryUsage / 1024 / 1024
    ws.Cells(nextRow, 7).value = this.ErrorCount
    
    ' Formatage
    ws.Range("A" & nextRow).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Range("C" & nextRow).NumberFormat = "0.00"
    ws.Range("D" & nextRow & ":E" & nextRow).NumberFormat = "0.0%"
    ws.Range("F" & nextRow).NumberFormat = "0.0"
    
    ' Sauvegarder
    targetWorkbook.Save
End Sub

'@Description("Met à jour les métriques de requêtes")
'@Description: 
'@Param: 
'@Returns: 

Private Sub UpdateQueryMetrics(ByVal executionTime As Double)
    ' Moyenne mobile pondérée
    this.AvgExecutionTime = (this.AvgExecutionTime * this.QueryCount + executionTime) / (this.QueryCount + 1)
    this.QueryCount = this.QueryCount + 1
End Sub

'@Description("Met à jour les métriques de connexions")
'@Description: 
'@Param: 
'@Returns: 

Private Sub UpdateConnectionMetrics(ByVal usage As Double)
    this.ConnectionUsage = usage
    this.LastUpdate = Now
End Sub

'@Description("Met à jour le compteur d'erreurs")
'@Description: 
'@Param: 
'@Returns: 

Private Sub UpdateErrorMetrics()
    this.ErrorCount = this.ErrorCount + 1
End Sub

'@Description("Met à jour les métriques de cache")
'@Description: 
'@Param: 
'@Returns: 

Private Sub UpdateCacheMetrics(ByVal isHit As Boolean)
    Static totalRequests As Long
    totalRequests = totalRequests + 1
    
    If isHit Then
        this.CacheHitRate = (this.CacheHitRate * (totalRequests - 1) + 1) / totalRequests
    Else
        this.CacheHitRate = this.CacheHitRate * (totalRequests - 1) / totalRequests
    End If
End Sub

'@Description("Met à jour les métriques de mémoire")
'@Description: 
'@Param: 
'@Returns: 

Private Sub UpdateMemoryMetrics(ByVal usage As Double)
    this.MemoryUsage = usage
End Sub

'@Description("Vérifie les seuils de performance")
'@Description: 
'@Param: 
'@Returns: 

Private Sub CheckThresholds()
    ' Vérification du temps de requête
    If this.AvgExecutionTime > m_Thresholds.MaxQueryTime Then
        LogAlert "Temps de requête moyen élevé", "WARN"
    End If
    
    ' Vérification de l'utilisation des connexions
    If this.PoolSize > m_Thresholds.MaxPoolSize Then
        LogAlert "Taille du pool de connexions critique", "ERROR"
    End If
    
    ' Vérification du cache
    If this.CacheHitRate < m_Thresholds.MinCacheHitRate Then
        LogAlert "Taux de succès du cache faible", "WARN"
    End If
    
    ' Vérification de la mémoire
    If this.MemoryUsage > m_Thresholds.MaxMemoryUsage Then
        LogAlert "Utilisation mémoire critique", "ERROR"
    End If
End Sub

'@Description("Enregistre une alerte dans les logs")
'@Description: 
'@Param: 
'@Returns: 

Private Sub LogAlert(ByVal message As String, ByVal level As String)
    If Not m_Logger Is Nothing Then
        m_Logger.Log level, "MetricsCollector", message
    End If
End Sub

'@Description("Sauvegarde un instantané des métriques")
'@Description: 
'@Param: 
'@Returns: 

Private Sub SaveMetricsSnapshot()
    Dim snapshot As TMetrics
    snapshot = this
    
    ' Limiter l'historique à 1000 entrées
    If m_MetricsHistory.Count >= 1000 Then
        RemoveOldestMetric
    End If
    
    m_MetricsHistory.Add snapshot
End Sub

'@Description("Supprime la plus ancienne métrique")
'@Description: 
'@Param: 
'@Returns: 

Private Sub RemoveOldestMetric()
    If m_MetricsHistory.Count > 0 Then
        m_MetricsHistory.Remove 1
    End If
End Sub 