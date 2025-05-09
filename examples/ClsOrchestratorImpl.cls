'@Class
'@ModuleDescription("Impl�mentation de r�f�rence de l'orchestrateur APEX")
'@Version(1.0)
'@Author("Assistant")
'@LastSync("2024-04-14 17:30")
'@CoordinationPoint("Exemple d'impl�mentation")

Option Explicit

Implements IApexOrchestrator

'Private Types pour l'�tat interne
Private Type TComponentEntry
    Component As IApexComponent
    Dependencies As Collection
    IsActive As Boolean
    LastError As String
    StartTime As Double
End Type

Private Type TEventHandlerEntry
    Handler As IEventHandler
    Priority As Long
End Type

Private Type TEventHandlerMap
    Handlers As Collection 'Collection de TEventHandlerEntry
End Type

Private Type TOrchestrator
    Status As ApexOrchestratorStatus
    Components As Dictionary 'Cl�: Nom, Valeur: TComponentEntry
    EventHandlers As Dictionary 'Cl�: eventName, Valeur: TEventHandlerMap
    Config As Dictionary
    StartTime As Double
    ExcelRouter As IExcelEventRouter
    Logger As Object
    Metrics As Dictionary
    LastError As String
    ErrorCount As Long
    HasShutdownStarted As Boolean
End Type

'�tat interne
Private this As TOrchestrator

'Initialisation
Private Sub Class_Initialize()
    this.Status = NotInitialized
    Set this.Components = New Dictionary
    Set this.EventHandlers = New Dictionary
    Set this.Metrics = New Dictionary
    this.ErrorCount = 0
    this.HasShutdownStarted = False
End Sub

'Nettoyage
Private Sub Class_Terminate()
    Set this.Components = Nothing
    Set this.EventHandlers = Nothing
    Set this.Config = Nothing
    Set this.ExcelRouter = Nothing
    Set this.Logger = Nothing
    Set this.Metrics = Nothing
End Sub

'Impl�mentation IApexOrchestrator

Public Function Initialize(ByVal config As Dictionary) As Boolean Implements IApexOrchestrator.Initialize
    On Error GoTo ErrorHandler
    
    'Validation des �tats
    If this.Status <> NotInitialized Then
        this.LastError = "D�j� initialis�"
        Initialize = False
        Exit Function
    End If
    
    'Mise � jour de l'�tat
    this.Status = Initializing
    LogStatus "Initialisation de l'orchestrateur"
    
    'Validation de la configuration
    If config Is Nothing Then
        this.LastError = "Configuration manquante"
        this.Status = Error
        Initialize = False
        Exit Function
    End If
    
    'Copie de la configuration
    Set this.Config = New Dictionary
    Dim key As Variant
    For Each key In config.Keys
        this.Config.Add key, config(key)
    Next key
    
    'Initialisation du logger
    If config.Exists("loggerFactory") Then
        Set this.Logger = config("loggerFactory").CreateLogger("Orchestrator")
    End If
    
    'Initialisation des m�triques
    this.Metrics.Add "startTime", Now
    this.Metrics.Add "componentCount", 0
    this.Metrics.Add "eventHandlerCount", 0
    this.Metrics.Add "eventCount", 0
    this.Metrics.Add "errorCount", 0
    
    'Initialisation du routeur Excel si configur�
    If config.Exists("excelRouterFactory") Then
        Set this.ExcelRouter = config("excelRouterFactory").CreateExcelRouter
        this.ExcelRouter.Initialize config
        LogInfo "Routeur Excel initialis�"
    End If
    
    'Mise � jour de l'�tat
    this.Status = Ready
    LogStatus "Orchestrateur initialis� avec succ�s"
    
    Initialize = True
    Exit Function
    
ErrorHandler:
    this.LastError = "Erreur d'initialisation: " & Err.Description
    this.Status = Error
    this.ErrorCount = this.ErrorCount + 1
    LogError this.LastError
    Initialize = False
End Function

Public Function Start() As Boolean Implements IApexOrchestrator.Start
    On Error GoTo ErrorHandler
    
    'Validation des �tats
    If this.Status <> Ready Then
        this.LastError = "L'orchestrateur n'est pas pr�t � d�marrer"
        Start = False
        Exit Function
    End If
    
    'Mise � jour de l'�tat
    this.Status = Starting
    LogStatus "D�marrage de l'orchestrateur"
    
    'D�marrage des composants par ordre de d�pendance
    Dim componentStartOrder As Collection
    Set componentStartOrder = ResolveComponentStartOrder
    
    Dim i As Long
    Dim componentName As String
    Dim startSuccess As Boolean
    
    startSuccess = True
    
    For i = 1 To componentStartOrder.Count
        componentName = componentStartOrder(i)
        
        If this.Components.Exists(componentName) Then
            Dim entry As TComponentEntry
            entry = this.Components(componentName)
            
            LogInfo "D�marrage du composant: " & componentName
            
            'D�marrage du composant
            entry.StartTime = MicroTimer
            If Not entry.Component.Start Then
                this.LastError = "�chec du d�marrage du composant: " & componentName
                entry.IsActive = False
                entry.LastError = this.LastError
                startSuccess = False
                LogError this.LastError
            Else
                entry.IsActive = True
                LogInfo "Composant d�marr�: " & componentName
            End If
            
            'Mise � jour de l'entr�e
            this.Components(componentName) = entry
            
            'Si un �chec critique, on arr�te le d�marrage
            If Not startSuccess And IsCriticalComponent(componentName) Then
                this.LastError = "�chec du d�marrage d'un composant critique: " & componentName
                this.Status = Error
                Start = False
                Exit Function
            End If
        End If
    Next i
    
    'Mise � jour de l'�tat si tout s'est bien pass�
    If startSuccess Then
        this.Status = Running
        this.StartTime = MicroTimer
        LogStatus "Orchestrateur d�marr� avec succ�s"
    Else
        this.Status = Error
        LogStatus "D�marrage partiel avec erreurs"
    End If
    
    Start = startSuccess
    Exit Function
    
ErrorHandler:
    this.LastError = "Erreur de d�marrage: " & Err.Description
    this.Status = Error
    this.ErrorCount = this.ErrorCount + 1
    LogError this.LastError
    Start = False
End Function

Public Function Stop() As Boolean Implements IApexOrchestrator.Stop
    On Error GoTo ErrorHandler
    
    '�viter les arr�ts multiples
    If this.HasShutdownStarted Then
        LogWarning "Arr�t d�j� en cours"
        Stop = False
        Exit Function
    End If
    
    this.HasShutdownStarted = True
    
    'Mise � jour de l'�tat
    this.Status = Stopping
    LogStatus "Arr�t de l'orchestrateur"
    
    'Arr�t des composants dans l'ordre inverse de d�marrage
    Dim componentStartOrder As Collection
    Set componentStartOrder = ResolveComponentStartOrder
    
    Dim i As Long
    Dim componentName As String
    Dim stopSuccess As Boolean
    
    stopSuccess = True
    
    For i = componentStartOrder.Count To 1 Step -1
        componentName = componentStartOrder(i)
        
        If this.Components.Exists(componentName) Then
            Dim entry As TComponentEntry
            entry = this.Components(componentName)
            
            If entry.IsActive Then
                LogInfo "Arr�t du composant: " & componentName
                
                'Arr�t du composant
                If Not entry.Component.Stop Then
                    this.LastError = "�chec de l'arr�t du composant: " & componentName
                    entry.LastError = this.LastError
                    stopSuccess = False
                    LogError this.LastError
                Else
                    entry.IsActive = False
                    LogInfo "Composant arr�t�: " & componentName
                End If
                
                'Mise � jour de l'entr�e
                this.Components(componentName) = entry
            End If
        End If
    Next i
    
    'Arr�t du routeur Excel si pr�sent
    If Not this.ExcelRouter Is Nothing Then
        On Error Resume Next
        this.ExcelRouter.Finalize
        On Error GoTo ErrorHandler
        Set this.ExcelRouter = Nothing
    End If
    
    'Mise � jour de l'�tat
    this.Status = Stopped
    LogStatus "Orchestrateur arr�t�"
    
    'Suppression des r�f�rences
    Set this.Components = New Dictionary
    Set this.EventHandlers = New Dictionary
    
    Stop = True
    Exit Function
    
ErrorHandler:
    this.LastError = "Erreur d'arr�t: " & Err.Description
    this.Status = Error
    this.ErrorCount = this.ErrorCount + 1
    LogError this.LastError
    Stop = False
End Function

Public Function RouteEvent(ByVal eventName As String, ByVal eventData As Variant) As Boolean Implements IApexOrchestrator.RouteEvent
    On Error GoTo ErrorHandler
    
    'Validation des �tats
    If this.Status <> Running And this.Status <> Paused Then
        this.LastError = "L'orchestrateur n'est pas en cours d'ex�cution"
        RouteEvent = False
        Exit Function
    End If
    
    'V�rification des gestionnaires pour cet �v�nement
    If Not this.EventHandlers.Exists(eventName) Then
        LogInfo "Aucun gestionnaire pour l'�v�nement: " & eventName
        RouteEvent = True 'Pas d'erreur, juste aucun gestionnaire
        Exit Function
    End If
    
    'R�cup�ration des gestionnaires et tri par priorit�
    Dim handlers As Collection
    Set handlers = GetSortedHandlersForEvent(eventName)
    
    'Routage de l'�v�nement � chaque gestionnaire
    Dim i As Long
    Dim handlerEntry As TEventHandlerEntry
    Dim handledSuccessfully As Boolean
    
    handledSuccessfully = False
    
    For i = 1 To handlers.Count
        Set handlerEntry = handlers(i)
        
        'V�rification si le gestionnaire peut traiter cet �v�nement
        If handlerEntry.Handler.CanHandleEvent(eventName) Then
            'Tentative de traitement
            If handlerEntry.Handler.HandleEvent(eventName, eventData) Then
                handledSuccessfully = True
                
                'Mise � jour des m�triques
                this.Metrics("eventCount") = this.Metrics("eventCount") + 1
                
                'Si configur� pour s'arr�ter au premier succ�s
                If GetStopOnFirstSuccess Then
                    Exit For
                End If
            End If
        End If
    Next i
    
    RouteEvent = handledSuccessfully
    Exit Function
    
ErrorHandler:
    this.LastError = "Erreur de routage d'�v�nement: " & Err.Description
    this.ErrorCount = this.ErrorCount + 1
    LogError this.LastError
    RouteEvent = False
End Function

Public Function RegisterEventHandler(ByVal eventName As String, ByVal handler As IEventHandler) As Boolean Implements IApexOrchestrator.RegisterEventHandler
    On Error GoTo ErrorHandler
    
    'Validation
    If handler Is Nothing Then
        this.LastError = "Gestionnaire d'�v�nements invalide"
        RegisterEventHandler = False
        Exit Function
    End If
    
    'Cr�ation de la map si elle n'existe pas
    If Not this.EventHandlers.Exists(eventName) Then
        Dim newMap As TEventHandlerMap
        Set newMap.Handlers = New Collection
        this.EventHandlers.Add eventName, newMap
    End If
    
    'R�cup�ration de la map
    Dim handlerMap As TEventHandlerMap
    handlerMap = this.EventHandlers(eventName)
    
    'Cr�ation de l'entr�e
    Dim entry As TEventHandlerEntry
    Set entry.Handler = handler
    entry.Priority = handler.GetPriority(eventName)
    
    'Ajout � la collection
    handlerMap.Handlers.Add entry
    
    'Mise � jour de la map
    this.EventHandlers(eventName) = handlerMap
    
    'Mise � jour des m�triques
    this.Metrics("eventHandlerCount") = this.Metrics("eventHandlerCount") + 1
    
    LogInfo "Gestionnaire enregistr� pour l'�v�nement: " & eventName & ", ID: " & handler.GetHandlerId
    
    RegisterEventHandler = True
    Exit Function
    
ErrorHandler:
    this.LastError = "Erreur d'enregistrement de gestionnaire: " & Err.Description
    this.ErrorCount = this.ErrorCount + 1
    LogError this.LastError
    RegisterEventHandler = False
End Function

Public Function UnregisterEventHandler(ByVal eventName As String, ByVal handler As IEventHandler) As Boolean Implements IApexOrchestrator.UnregisterEventHandler
    On Error GoTo ErrorHandler
    
    'Validation
    If handler Is Nothing Then
        this.LastError = "Gestionnaire d'�v�nements invalide"
        UnregisterEventHandler = False
        Exit Function
    End If
    
    If Not this.EventHandlers.Exists(eventName) Then
        LogInfo "Aucun gestionnaire pour l'�v�nement: " & eventName
        UnregisterEventHandler = True 'Pas d'erreur, juste aucun gestionnaire
        Exit Function
    End If
    
    'R�cup�ration de la map
    Dim handlerMap As TEventHandlerMap
    handlerMap = this.EventHandlers(eventName)
    
    'Recherche et suppression du gestionnaire
    Dim i As Long
    Dim entry As TEventHandlerEntry
    Dim handlerId As String
    
    handlerId = handler.GetHandlerId
    
    For i = handlerMap.Handlers.Count To 1 Step -1
        Set entry = handlerMap.Handlers(i)
        
        If entry.Handler.GetHandlerId = handlerId Then
            handlerMap.Handlers.Remove i
            
            'Mise � jour des m�triques
            this.Metrics("eventHandlerCount") = this.Metrics("eventHandlerCount") - 1
            
            LogInfo "Gestionnaire d�senregistr� pour l'�v�nement: " & eventName & ", ID: " & handlerId
            Exit For
        End If
    Next i
    
    'Mise � jour de la map
    this.EventHandlers(eventName) = handlerMap
    
    UnregisterEventHandler = True
    Exit Function
    
ErrorHandler:
    this.LastError = "Erreur de d�senregistrement de gestionnaire: " & Err.Description
    this.ErrorCount = this.ErrorCount + 1
    LogError this.LastError
    UnregisterEventHandler = False
End Function

Public Function RegisterComponent(ByVal component As IApexComponent) As Boolean Implements IApexOrchestrator.RegisterComponent
    On Error GoTo ErrorHandler
    
    'Validation
    If component Is Nothing Then
        this.LastError = "Composant invalide"
        RegisterComponent = False
        Exit Function
    End If
    
    Dim componentName As String
    componentName = component.GetName
    
    'V�rification si le composant existe d�j�
    If this.Components.Exists(componentName) Then
        this.LastError = "Composant d�j� enregistr�: " & componentName
        RegisterComponent = False
        Exit Function
    End If
    
    'Cr�ation de l'entr�e
    Dim entry As TComponentEntry
    Set entry.Component = component
    Set entry.Dependencies = component.GetDependencies
    entry.IsActive = False
    entry.LastError = ""
    
    'Ajout au dictionnaire
    this.Components.Add componentName, entry
    
    'Mise � jour des m�triques
    this.Metrics("componentCount") = this.Metrics("componentCount") + 1
    
    LogInfo "Composant enregistr�: " & componentName
    
    RegisterComponent = True
    Exit Function
    
ErrorHandler:
    this.LastError = "Erreur d'enregistrement de composant: " & Err.Description
    this.ErrorCount = this.ErrorCount + 1
    LogError this.LastError
    RegisterComponent = False
End Function

Public Function UnregisterComponent(ByVal componentName As String) As Boolean Implements IApexOrchestrator.UnregisterComponent
    On Error GoTo ErrorHandler
    
    'Validation
    If Not this.Components.Exists(componentName) Then
        LogInfo "Composant non trouv�: " & componentName
        UnregisterComponent = True 'Pas d'erreur, juste aucun composant
        Exit Function
    End If
    
    'V�rification si le composant est actif
    Dim entry As TComponentEntry
    entry = this.Components(componentName)
    
    If entry.IsActive Then
        'Tentative d'arr�t
        If Not entry.Component.Stop Then
            this.LastError = "�chec de l'arr�t du composant: " & componentName
            LogError this.LastError
            'On continue malgr� l'erreur
        End If
    End If
    
    'Suppression du composant
    this.Components.Remove componentName
    
    'Mise � jour des m�triques
    this.Metrics("componentCount") = this.Metrics("componentCount") - 1
    
    LogInfo "Composant d�senregistr�: " & componentName
    
    UnregisterComponent = True
    Exit Function
    
ErrorHandler:
    this.LastError = "Erreur de d�senregistrement de composant: " & Err.Description
    this.ErrorCount = this.ErrorCount + 1
    LogError this.LastError
    UnregisterComponent = False
End Function

Public Function GetComponentState(ByVal componentName As String) As Variant Implements IApexOrchestrator.GetComponentState
    On Error GoTo ErrorHandler
    
    'Validation
    If Not this.Components.Exists(componentName) Then
        LogWarning "Composant non trouv�: " & componentName
        Set GetComponentState = Nothing
        Exit Function
    End If
    
    'R�cup�ration de l'�tat
    Dim entry As TComponentEntry
    entry = this.Components(componentName)
    
    Set GetComponentState = entry.Component.GetState
    Exit Function
    
ErrorHandler:
    this.LastError = "Erreur de r�cup�ration d'�tat: " & Err.Description
    this.ErrorCount = this.ErrorCount + 1
    LogError this.LastError
    Set GetComponentState = Nothing
End Function

Public Function SetComponentState(ByVal componentName As String, ByVal state As Variant) As Boolean Implements IApexOrchestrator.SetComponentState
    On Error GoTo ErrorHandler
    
    'Validation
    If Not this.Components.Exists(componentName) Then
        LogWarning "Composant non trouv�: " & componentName
        SetComponentState = False
        Exit Function
    End If
    
    'R�cup�ration de l'entr�e
    Dim entry As TComponentEntry
    entry = this.Components(componentName)
    
    'Mise � jour de l'�tat
    SetComponentState = entry.Component.SetState(state)
    
    Exit Function
    
ErrorHandler:
    this.LastError = "Erreur de mise � jour d'�tat: " & Err.Description
    this.ErrorCount = this.ErrorCount + 1
    LogError this.LastError
    SetComponentState = False
End Function

Public Function IsComponentActive(ByVal componentName As String) As Boolean Implements IApexOrchestrator.IsComponentActive
    On Error GoTo ErrorHandler
    
    'Validation
    If Not this.Components.Exists(componentName) Then
        IsComponentActive = False
        Exit Function
    End If
    
    'R�cup�ration de l'entr�e
    Dim entry As TComponentEntry
    entry = this.Components(componentName)
    
    IsComponentActive = entry.IsActive
    Exit Function
    
ErrorHandler:
    IsComponentActive = False
End Function

Public Function GetRegisteredComponents() As Collection Implements IApexOrchestrator.GetRegisteredComponents
    On Error GoTo ErrorHandler
    
    Dim result As Collection
    Set result = New Collection
    
    Dim key As Variant
    For Each key In this.Components.Keys
        result.Add key
    Next key
    
    Set GetRegisteredComponents = result
    Exit Function
    
ErrorHandler:
    Set GetRegisteredComponents = New Collection
End Function

Public Function GetPerformanceMetrics() As Dictionary Implements IApexOrchestrator.GetPerformanceMetrics
    On Error GoTo ErrorHandler
    
    Dim result As Dictionary
    Set result = New Dictionary
    
    'Copie des m�triques de base
    Dim key As Variant
    For Each key In this.Metrics.Keys
        result.Add key, this.Metrics(key)
    Next key
    
    'Ajout des m�triques dynamiques
    If this.Status = Running Then
        result.Add "uptime", MicroTimer - this.StartTime
    End If
    
    result.Add "status", GetStatus
    result.Add "lastError", this.LastError
    result.Add "errorCount", this.ErrorCount
    
    Set GetPerformanceMetrics = result
    Exit Function
    
ErrorHandler:
    Set GetPerformanceMetrics = New Dictionary
End Function

Public Function GetStatus() As ApexOrchestratorStatus Implements IApexOrchestrator.GetStatus
    GetStatus = this.Status
End Function

Public Function HandleError(ByVal componentName As String, ByVal errorInfo As Dictionary) As Boolean Implements IApexOrchestrator.HandleError
    On Error GoTo ErrorHandler
    
    'Validation
    If Not this.Components.Exists(componentName) Then
        LogWarning "Composant non trouv� pour l'erreur: " & componentName
        HandleError = False
        Exit Function
    End If
    
    'Log de l'erreur
    Dim errorMsg As String
    If errorInfo.Exists("message") Then
        errorMsg = errorInfo("message")
    Else
        errorMsg = "Erreur non sp�cifi�e"
    End If
    
    LogError "Erreur dans le composant " & componentName & ": " & errorMsg
    
    'Mise � jour des m�triques
    this.ErrorCount = this.ErrorCount + 1
    
    'Gestion sp�cifique selon la gravit�
    If errorInfo.Exists("critical") And errorInfo("critical") = True Then
        If IsCriticalComponent(componentName) Then
            LogError "Erreur critique dans un composant critique: " & componentName
            this.Status = Error
        End If
    End If
    
    HandleError = True
    Exit Function
    
ErrorHandler:
    this.LastError = "Erreur de gestion d'erreur: " & Err.Description
    this.ErrorCount = this.ErrorCount + 1
    LogError this.LastError
    HandleError = False
End Function

Public Function SynchronizeState() As Boolean Implements IApexOrchestrator.SynchronizeState
    On Error GoTo ErrorHandler
    
    'Validation des �tats
    If this.Status <> Running Then
        this.LastError = "L'orchestrateur n'est pas en cours d'ex�cution"
        SynchronizeState = False
        Exit Function
    End If
    
    LogInfo "Synchronisation de l'�tat des composants"
    
    'Parcours des composants pour synchronisation
    Dim componentName As Variant
    Dim entry As TComponentEntry
    Dim syncSuccess As Boolean
    
    syncSuccess = True
    
    For Each componentName In this.Components.Keys
        entry = this.Components(componentName)
        
        If entry.IsActive Then
            'V�rification de l'�tat pr�t
            If Not entry.Component.IsReady Then
                LogWarning "Composant non pr�t: " & componentName
                syncSuccess = False
            End If
        End If
    Next componentName
    
    SynchronizeState = syncSuccess
    Exit Function
    
ErrorHandler:
    this.LastError = "Erreur de synchronisation: " & Err.Description
    this.ErrorCount = this.ErrorCount + 1
    LogError this.LastError
    SynchronizeState = False
End Function

'Fonctions priv�es

Private Function ResolveComponentStartOrder() As Collection
    Dim result As Collection
    Set result = New Collection
    
    'Simple impl�mentation pour exemple
    'Une impl�mentation r�elle devrait r�soudre les d�pendances
    Dim componentName As Variant
    For Each componentName In this.Components.Keys
        result.Add componentName
    Next componentName
    
    Set ResolveComponentStartOrder = result
End Function

Private Function IsCriticalComponent(ByVal componentName As String) As Boolean
    'D�termine si un composant est critique pour le syst�me
    'Exemple simple - � remplacer par une vraie impl�mentation
    
    'Par d�faut, consid�rons que les composants "Core" sont critiques
    IsCriticalComponent = InStr(1, componentName, "Core") > 0
End Function

Private Function GetSortedHandlersForEvent(ByVal eventName As String) As Collection
    Dim result As Collection
    Set result = New Collection
    
    If Not this.EventHandlers.Exists(eventName) Then
        Set GetSortedHandlersForEvent = result
        Exit Function
    End If
    
    Dim handlerMap As TEventHandlerMap
    handlerMap = this.EventHandlers(eventName)
    
    'Copie des handlers
    Dim i As Long
    For i = 1 To handlerMap.Handlers.Count
        result.Add handlerMap.Handlers(i)
    Next i
    
    'Tri par priorit� (impl�mentation simplifi�e)
    'Une vraie impl�mentation utiliserait un algorithme de tri efficace
    
    Dim j As Long, k As Long
    Dim temp As TEventHandlerEntry
    
    For j = 1 To result.Count - 1
        For k = j + 1 To result.Count
            If result(j).Priority < result(k).Priority Then
                Set temp = result(j)
                result.Remove j
                result.Add temp, , k
            End If
        Next k
    Next j
    
    Set GetSortedHandlersForEvent = result
End Function

Private Function GetStopOnFirstSuccess() As Boolean
    'R�cup�re la configuration de propagation d'�v�nements
    If this.Config.Exists("stopOnFirstSuccess") Then
        GetStopOnFirstSuccess = this.Config("stopOnFirstSuccess")
    Else
        GetStopOnFirstSuccess = False 'Par d�faut, on continue
    End If
End Function

Private Function MicroTimer() As Double
    'Fonction utilitaire pour la mesure pr�cise du temps
    '� remplacer par une impl�mentation r�elle
    MicroTimer = Timer
End Function

'Fonctions de logging

Private Sub LogInfo(ByVal message As String)
    If Not this.Logger Is Nothing Then
        this.Logger.LogInfo message
    End If
End Sub

Private Sub LogWarning(ByVal message As String)
    If Not this.Logger Is Nothing Then
        this.Logger.LogWarning message
    End If
End Sub

Private Sub LogError(ByVal message As String)
    If Not this.Logger Is Nothing Then
        this.Logger.LogError message
    End If
End Sub

Private Sub LogStatus(ByVal message As String)
    If Not this.Logger Is Nothing Then
        this.Logger.LogInfo "[STATUS] " & message & " - �tat: " & StatusToString(this.Status)
    End If
End Sub

Private Function StatusToString(ByVal status As ApexOrchestratorStatus) As String
    Select Case status
        Case NotInitialized: StatusToString = "Non initialis�"
        Case Initializing: StatusToString = "En cours d'initialisation"
        Case Ready: StatusToString = "Pr�t"
        Case Starting: StatusToString = "En cours de d�marrage"
        Case Running: StatusToString = "En cours d'ex�cution"
        Case Pausing: StatusToString = "En cours de pause"
        Case Paused: StatusToString = "En pause"
        Case Stopping: StatusToString = "En cours d'arr�t"
        Case Stopped: StatusToString = "Arr�t�"
        Case Error: StatusToString = "En erreur"
        Case Else: StatusToString = "�tat inconnu"
    End Select
End Function 