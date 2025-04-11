' Migrated to apex-core - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit
' ==========================================================================
' Class : clsPluginManager
' Version : 1.0
' Purpose : Gestionnaire central pour les plugins du framework
' Requires : IPlugin, Scripting.Dictionary
' ==========================================================================

Private m_Plugins As Object ' Scripting.Dictionary
Private m_Logger As ILoggerBase
Private m_IsInitialized As Boolean
Private m_EventSubscriptions As Object ' Scripting.Dictionary - stocke les abonnements d'événements

' --- Initialisation ---
Private Sub Class_Initialize()
    Set m_Plugins = CreateObject("Scripting.Dictionary")
    Set m_EventSubscriptions = CreateObject("Scripting.Dictionary")
    m_IsInitialized = False
End Sub

Public Sub Initialize(Optional ByVal logger As ILoggerBase = Nothing)
    If m_IsInitialized Then Exit Sub
    Set m_Logger = logger
    LogIfAble "Initialize", LogLevelInfo, "Plugin Manager initialized", "PLUGIN"
    m_IsInitialized = True
End Sub

' --- Gestion de plugins ---
Public Function RegisterPlugin(plugin As IPlugin, Optional ByVal pluginId As String = "") As String
    Dim pluginIdToUse As String
    
    If Not m_IsInitialized Then Initialize
    
    ' Générer un ID si non fourni
    If pluginId = "" Then pluginIdToUse = plugin.PluginName & "_" & Format(Now, "yyyymmddhhnnss") Else pluginIdToUse = pluginId
    
    ' Vérifier si l'ID existe déjà
    If m_Plugins.Exists(pluginIdToUse) Then
        LogIfAble "RegisterPlugin", LogLevelWarning, "Plugin ID '" & pluginIdToUse & "' already exists. Overwriting.", "PLUGIN"
        UnregisterPlugin pluginIdToUse
    End If
    
    ' Initialiser le plugin
    On Error Resume Next
    Dim initSuccess As Boolean
    initSuccess = plugin.Initialize
    If Err.Number <> 0 Then
        LogIfAble "RegisterPlugin", LogLevelError, "Failed to initialize plugin '" & plugin.PluginName & "': " & Err.Description, "PLUGIN"
        Set plugin = Nothing
        RegisterPlugin = ""
        Exit Function
    End If
    On Error GoTo 0
    
    ' Enregistrer le plugin
    m_Plugins.Add pluginIdToUse, plugin
    LogIfAble "RegisterPlugin", LogLevelInfo, "Plugin '" & plugin.PluginName & "' v" & plugin.PluginVersion & " registered with ID: " & pluginIdToUse, "PLUGIN"
    
    RegisterPlugin = pluginIdToUse
End Function

Public Function GetPlugin(ByVal pluginId As String) As IPlugin
    If Not m_IsInitialized Or Not m_Plugins.Exists(pluginId) Then
        Set GetPlugin = Nothing
        Exit Function
    End If
    
    Set GetPlugin = m_Plugins(pluginId)
End Function

Public Function UnregisterPlugin(ByVal pluginId As String) As Boolean
    If Not m_IsInitialized Or Not m_Plugins.Exists(pluginId) Then
        UnregisterPlugin = False
        Exit Function
    End If
    
    Dim plugin As IPlugin
    Set plugin = m_Plugins(pluginId)
    
    ' Nettoyer les abonnements d'événements pour ce plugin
    RemovePluginSubscriptions pluginId
    
    ' Terminer le plugin
    On Error Resume Next
    plugin.Terminate
    If Err.Number <> 0 Then
        LogIfAble "UnregisterPlugin", LogLevelWarning, "Error during plugin '" & pluginId & "' termination: " & Err.Description, "PLUGIN"
    End If
    On Error GoTo 0
    
    ' Supprimer le plugin du dictionnaire
    m_Plugins.Remove pluginId
    LogIfAble "UnregisterPlugin", LogLevelInfo, "Plugin '" & pluginId & "' unregistered", "PLUGIN"
    
    UnregisterPlugin = True
End Function

Public Function GetAllPlugins() As Collection
    Dim result As New Collection
    Dim key As Variant
    
    If Not m_IsInitialized Then
        Set GetAllPlugins = result
        Exit Function
    End If
    
    For Each key In m_Plugins.Keys
        result.Add m_Plugins(key), CStr(key)
    Next key
    
    Set GetAllPlugins = result
End Function

' --- Gestion des événements ---
Public Sub SubscribeToEvent(ByVal pluginId As String, ByVal eventName As String)
    If Not m_IsInitialized Or Not m_Plugins.Exists(pluginId) Then Exit Sub
    
    Dim eventKey As String
    eventKey = UCase$(eventName)
    
    If Not m_EventSubscriptions.Exists(eventKey) Then
        m_EventSubscriptions.Add eventKey, CreateObject("Scripting.Dictionary")
    End If
    
    Dim eventSubscribers As Object ' Dictionary
    Set eventSubscribers = m_EventSubscriptions(eventKey)
    
    If Not eventSubscribers.Exists(pluginId) Then
        eventSubscribers.Add pluginId, True
        LogIfAble "SubscribeToEvent", LogLevelDebug, "Plugin '" & pluginId & "' subscribed to event '" & eventName & "'", "PLUGIN"
    End If
End Sub

Public Sub UnsubscribeFromEvent(ByVal pluginId As String, ByVal eventName As String)
    If Not m_IsInitialized Then Exit Sub
    
    Dim eventKey As String
    eventKey = UCase$(eventName)
    
    If Not m_EventSubscriptions.Exists(eventKey) Then Exit Sub
    
    Dim eventSubscribers As Object ' Dictionary
    Set eventSubscribers = m_EventSubscriptions(eventKey)
    
    If eventSubscribers.Exists(pluginId) Then
        eventSubscribers.Remove pluginId
        LogIfAble "UnsubscribeFromEvent", LogLevelDebug, "Plugin '" & pluginId & "' unsubscribed from event '" & eventName & "'", "PLUGIN"
    End If
    
    ' Nettoyer les dictionnaires vides
    If eventSubscribers.Count = 0 Then
        m_EventSubscriptions.Remove eventKey
    End If
End Sub

Private Sub RemovePluginSubscriptions(ByVal pluginId As String)
    If Not m_IsInitialized Then Exit Sub
    
    Dim eventKey As Variant
    Dim eventSubscribers As Object ' Dictionary
    Dim keysToRemove As New Collection
    
    For Each eventKey In m_EventSubscriptions.Keys
        Set eventSubscribers = m_EventSubscriptions(eventKey)
        If eventSubscribers.Exists(pluginId) Then
            eventSubscribers.Remove pluginId
            LogIfAble "RemovePluginSubscriptions", LogLevelDebug, "Plugin '" & pluginId & "' unsubscribed from event '" & eventKey & "'", "PLUGIN"
            
            ' Marquer les événements vides pour suppression
            If eventSubscribers.Count = 0 Then
                keysToRemove.Add eventKey
            End If
        End If
    Next eventKey
    
    ' Nettoyer les dictionnaires vides
    Dim key As Variant
    For Each key In keysToRemove
        m_EventSubscriptions.Remove key
    Next key
End Sub

Public Function TriggerEvent(ByVal eventName As String, Optional ByVal eventData As Variant) As Collection
    Dim results As New Collection
    
    If Not m_IsInitialized Then
        Set TriggerEvent = results
        Exit Function
    End If
    
    Dim eventKey As String
    eventKey = UCase$(eventName)
    
    If Not m_EventSubscriptions.Exists(eventKey) Then
        LogIfAble "TriggerEvent", LogLevelDebug, "Event '" & eventName & "' triggered but no subscribers", "PLUGIN"
        Set TriggerEvent = results
        Exit Function
    End If
    
    Dim eventSubscribers As Object ' Dictionary
    Set eventSubscribers = m_EventSubscriptions(eventKey)
    
    Dim pluginId As Variant
    Dim plugin As IPlugin
    Dim result As Variant
    
    LogIfAble "TriggerEvent", LogLevelDebug, "Event '" & eventName & "' triggered with " & eventSubscribers.Count & " subscribers", "PLUGIN"
    
    For Each pluginId In eventSubscribers.Keys
        If m_Plugins.Exists(pluginId) Then
            Set plugin = m_Plugins(pluginId)
            
            On Error Resume Next
            result = plugin.HandleEvent(eventName, eventData)
            If Err.Number <> 0 Then
                LogIfAble "TriggerEvent", LogLevelWarning, "Plugin '" & pluginId & "' event handler error: " & Err.Description, "PLUGIN"
                result = Null
            End If
            On Error GoTo 0
            
            results.Add result, CStr(pluginId)
        End If
    Next pluginId
    
    Set TriggerEvent = results
End Function

' --- Utilitaires ---
Public Sub SetLogger(ByVal logger As ILoggerBase)
    Set m_Logger = logger
End Sub

Private Sub LogIfAble(ByVal sourceMethod As String, ByVal level As LogLevelEnum, ByVal msg As String, ByVal category As String)
    If m_Logger Is Nothing Then Exit Sub
    
    On Error Resume Next
    m_Logger.LogMessage msg, level, category, "clsPluginManager." & sourceMethod
    On Error GoTo 0
End Sub

' --- Nettoyage ---
Private Sub Class_Terminate()
    Dim pluginId As Variant
    
    ' Terminer tous les plugins
    For Each pluginId In m_Plugins.Keys
        UnregisterPlugin CStr(pluginId)
    Next pluginId
    
    Set m_Plugins = Nothing
    Set m_EventSubscriptions = Nothing
    Set m_Logger = Nothing
    
    m_IsInitialized = False
End Sub 