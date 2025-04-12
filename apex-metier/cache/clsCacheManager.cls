VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCacheManager"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("APEX.Metier.Cache")
Option Explicit

'==========================================================================
' Classe    : clsCacheManager
' Purpose   : Gestion du cache pour optimiser les performances
' Author    : APEX Framework Team
' Date      : 2024-04-11
' Reference : PERF-003
'==========================================================================

Private Type TCacheEntry
    Key As String
    Value As Variant
    ExpiresAt As Date
    LastAccessed As Date
    AccessCount As Long
End Type

Private Type TCacheSettings
    MaxSize As Long
    DefaultTTL As Long ' Secondes
    CleanupInterval As Long ' Secondes
    LastCleanup As Date
End Type

Private Type TCacheStats
    HitCount As Long
    MissCount As Long
    EvictionCount As Long
    TotalSize As Long
End Type

Private this As TCacheSettings
Private m_Cache As Collection
Private m_Stats As TCacheStats
Private m_Logger As ILoggerBase
Private m_MetricsCollector As clsMetricsCollector

'@Description("Initialise le gestionnaire de cache")
'@Description: 
'@Param: 
'@Returns: 

Public Function Initialize(Optional ByVal logger As ILoggerBase = Nothing, _
                         Optional ByVal metricsCollector As clsMetricsCollector = Nothing) As Boolean
    ' Configuration par défaut
    With this
        .MaxSize = 1000 ' Entrées maximum
        .DefaultTTL = 3600 ' 1 heure
        .CleanupInterval = 300 ' 5 minutes
        .LastCleanup = Now
    End With
    
    ' Initialisation des dépendances
    Set m_Logger = logger
    Set m_MetricsCollector = metricsCollector
    
    ' Création du cache
    Set m_Cache = New Collection
    
    Initialize = True
End Function

'@Description("Configure les paramètres du cache")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Configure(ByVal maxSize As Long, _
                    ByVal defaultTTL As Long, _
                    ByVal cleanupInterval As Long)
    this.MaxSize = maxSize
    this.DefaultTTL = defaultTTL
    this.CleanupInterval = cleanupInterval
End Sub

'@Description("Ajoute ou met à jour une entrée dans le cache")
'@Description: 
'@Param: 
'@Returns: 

Public Function SetValue(ByVal key As String, _
                        ByVal value As Variant, _
                        Optional ByVal ttl As Long = -1) As Boolean
    On Error GoTo ErrorHandler
    
    ' Nettoyage si nécessaire
    If NeedsCleanup Then CleanupCache
    
    ' Vérifier la taille du cache
    If m_Cache.Count >= this.MaxSize Then
        EvictOldestEntry
    End If
    
    ' Créer l'entrée
    Dim entry As TCacheEntry
    entry.Key = key
    entry.Value = value
    entry.LastAccessed = Now
    entry.AccessCount = 0
    
    ' Définir l'expiration
    If ttl = -1 Then ttl = this.DefaultTTL
    entry.ExpiresAt = DateAdd("s", ttl, Now)
    
    ' Ajouter au cache
    If EntryExists(key) Then
        RemoveEntry key
    End If
    AddEntry entry
    
    ' Métriques
    m_Stats.TotalSize = m_Stats.TotalSize + 1
    If Not m_MetricsCollector Is Nothing Then
        m_MetricsCollector.RecordMetric "CacheSize", m_Stats.TotalSize
    End If
    
    SetValue = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    LogError "SetValue", Err.Description
    SetValue = False
End Function

'@Description("Récupère une valeur du cache")
'@Description: 
'@Param: 
'@Returns: 

Public Function GetValue(ByVal key As String, Optional ByRef found As Boolean = False) As Variant
    On Error GoTo ErrorHandler
    
    found = False
    
    ' Vérifier si l'entrée existe
    If Not EntryExists(key) Then
        m_Stats.MissCount = m_Stats.MissCount + 1
        If Not m_MetricsCollector Is Nothing Then
            m_MetricsCollector.RecordMetric "CacheMiss", 1
        End If
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Récupérer l'entrée
    Dim entry As TCacheEntry
    entry = GetEntry(key)
    
    ' Vérifier l'expiration
    If entry.ExpiresAt < Now Then
        RemoveEntry key
        m_Stats.MissCount = m_Stats.MissCount + 1
        If Not m_MetricsCollector Is Nothing Then
            m_MetricsCollector.RecordMetric "CacheMiss", 1
        End If
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Mettre à jour les statistiques
    entry.LastAccessed = Now
    entry.AccessCount = entry.AccessCount + 1
    UpdateEntry key, entry
    
    m_Stats.HitCount = m_Stats.HitCount + 1
    If Not m_MetricsCollector Is Nothing Then
        m_MetricsCollector.RecordMetric "CacheHit", 1
    End If
    
    found = True
    GetValue = entry.Value
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    LogError "GetValue", Err.Description
End Function

'@Description("Supprime une entrée du cache")
'@Description: 
'@Param: 
'@Returns: 

Public Function RemoveValue(ByVal key As String) As Boolean
    On Error GoTo ErrorHandler
    
    If EntryExists(key) Then
        RemoveEntry key
        m_Stats.TotalSize = m_Stats.TotalSize - 1
        If Not m_MetricsCollector Is Nothing Then
            m_MetricsCollector.RecordMetric "CacheSize", m_Stats.TotalSize
        End If
        RemoveValue = True
    End If
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    LogError "RemoveValue", Err.Description
    RemoveValue = False
End Function

'@Description("Vide le cache")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Clear()
    Set m_Cache = New Collection
    m_Stats.TotalSize = 0
    If Not m_MetricsCollector Is Nothing Then
        m_MetricsCollector.RecordMetric "CacheSize", 0
    End If
End Sub

'@Description("Obtient les statistiques du cache")
'@Description: 
'@Param: 
'@Returns: 

Public Function GetStats() As String
    Dim hitRate As Double
    If (m_Stats.HitCount + m_Stats.MissCount) > 0 Then
        hitRate = m_Stats.HitCount / (m_Stats.HitCount + m_Stats.MissCount)
    End If
    
    GetStats = "=== Statistiques Cache ===" & vbNewLine & _
               "Taille actuelle: " & m_Stats.TotalSize & vbNewLine & _
               "Succès: " & m_Stats.HitCount & vbNewLine & _
               "Échecs: " & m_Stats.MissCount & vbNewLine & _
               "Taux de succès: " & Format(hitRate, "0.0%") & vbNewLine & _
               "Évictions: " & m_Stats.EvictionCount
End Function

'@Description("Vérifie si le cache nécessite un nettoyage")
'@Description: 
'@Param: 
'@Returns: 

Private Function NeedsCleanup() As Boolean
    NeedsCleanup = DateDiff("s", this.LastCleanup, Now) >= this.CleanupInterval
End Function

'@Description("Nettoie les entrées expirées")
'@Description: 
'@Param: 
'@Returns: 

Private Sub CleanupCache()
    Dim key As Variant
    Dim entry As TCacheEntry
    Dim keysToRemove As Collection
    Set keysToRemove = New Collection
    
    ' Identifier les entrées expirées
    For Each key In m_Cache.Keys
        entry = m_Cache.Item(key)
        If entry.ExpiresAt < Now Then
            keysToRemove.Add key
        End If
    Next key
    
    ' Supprimer les entrées expirées
    For Each key In keysToRemove
        RemoveEntry key
        m_Stats.EvictionCount = m_Stats.EvictionCount + 1
        m_Stats.TotalSize = m_Stats.TotalSize - 1
    Next key
    
    this.LastCleanup = Now
    
    If Not m_MetricsCollector Is Nothing Then
        m_MetricsCollector.RecordMetric "CacheEvictions", keysToRemove.Count
        m_MetricsCollector.RecordMetric "CacheSize", m_Stats.TotalSize
    End If
End Sub

'@Description("Supprime l'entrée la plus ancienne")
'@Description: 
'@Param: 
'@Returns: 

Private Sub EvictOldestEntry()
    Dim oldestKey As String
    Dim oldestAccess As Date
    Dim entry As TCacheEntry
    Dim key As Variant
    
    oldestAccess = Now
    
    ' Trouver l'entrée la plus ancienne
    For Each key In m_Cache.Keys
        entry = m_Cache.Item(key)
        If entry.LastAccessed < oldestAccess Then
            oldestAccess = entry.LastAccessed
            oldestKey = key
        End If
    Next key
    
    ' Supprimer l'entrée
    If Len(oldestKey) > 0 Then
        RemoveEntry oldestKey
        m_Stats.EvictionCount = m_Stats.EvictionCount + 1
        m_Stats.TotalSize = m_Stats.TotalSize - 1
        
        If Not m_MetricsCollector Is Nothing Then
            m_MetricsCollector.RecordMetric "CacheEvictions", 1
            m_MetricsCollector.RecordMetric "CacheSize", m_Stats.TotalSize
        End If
    End If
End Sub

'@Description("Vérifie si une entrée existe")
'@Description: 
'@Param: 
'@Returns: 

Private Function EntryExists(ByVal key As String) As Boolean
    On Error Resume Next
    EntryExists = Not m_Cache.Item(key) Is Nothing
    On Error GoTo 0
End Function

'@Description("Récupère une entrée")
'@Description: 
'@Param: 
'@Returns: 

Private Function GetEntry(ByVal key As String) As TCacheEntry
    GetEntry = m_Cache.Item(key)
End Function

'@Description("Ajoute une entrée")
'@Description: 
'@Param: 
'@Returns: 

Private Sub AddEntry(ByRef entry As TCacheEntry)
    m_Cache.Add entry, entry.Key
End Sub

'@Description("Met à jour une entrée")
'@Description: 
'@Param: 
'@Returns: 

Private Sub UpdateEntry(ByVal key As String, ByRef entry As TCacheEntry)
    m_Cache.Remove key
    m_Cache.Add entry, key
End Sub

'@Description("Supprime une entrée")
'@Description: 
'@Param: 
'@Returns: 

Private Sub RemoveEntry(ByVal key As String)
    m_Cache.Remove key
End Sub

'@Description("Journalise une erreur")
'@Description: 
'@Param: 
'@Returns: 

Private Sub LogError(ByVal method As String, ByVal description As String)
    If Not m_Logger Is Nothing Then
        m_Logger.Log "ERROR", "CacheManager." & method, description
    End If
End Sub 