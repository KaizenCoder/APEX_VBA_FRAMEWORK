VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDependencyContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Interoperabilite.Integration")
'@ModuleDescription("Conteneur d'injection de dépendances pour gérer les services")
Option Explicit
' =========================================================================
' Classe  : clsDependencyContainer
' Version : 1.0
' Purpose : Conteneur d'injection de dépendances pour résoudre les composants
' Author  : APEX Framework Team
' Date    : 2024-04-11
' =========================================================================

' Constantes
Private Const CLASS_NAME As String = "clsDependencyContainer"
Private Const LOG_CATEGORY As String = "DependencyInjection"
Private Const ERR_NOT_INITIALIZED As Long = vbObjectError + 1001
Private Const ERR_NOT_REGISTERED As Long = vbObjectError + 1002
Private Const ERR_CIRCULAR_DEPENDENCY As Long = vbObjectError + 3402
Private Const ERR_TYPE_NOT_REGISTERED As Long = vbObjectError + 3101
Private Const ERR_INSTANCE_CREATION As Long = vbObjectError + 3103
Private Const ERR_RESOLUTION_FAILED As Long = vbObjectError + 3401
Private Const ERR_FACTORY_FAILED As Long = vbObjectError + 3403
Private Const MODULE_NAME As String = "clsDependencyContainer"
Private Const MAX_REGISTRATIONS As Long = 100
Private Const MAX_SERVICES As Long = 50

' Types d'enregistrement
Private Enum RegistrationType
    SingletonType = 0         ' Instance unique
    TransientType = 1         ' Nouvelle instance à chaque résolution
    FactoryType = 2           ' Création via fonction factory
End Enum

' Structure pour stocker les informations d'enregistrement
Private Type Registration
    InterfaceName As String      ' Nom de l'interface
    ImplementationClass As String ' Nom de la classe d'implémentation
    FactoryMethod As String      ' Nom de la méthode factory si applicable
    RegistrationType As RegistrationType ' Type d'enregistrement
    Instance As Object           ' Instance pour les singletons
End Type

' Collections pour stocker les enregistrements
Private m_registrations() As Registration  ' Tableau des enregistrements
Private m_count As Long                    ' Nombre d'enregistrements
Private m_initialized As Boolean           ' Indicateur d'initialisation
Private m_logger As Object                 ' Logger pour la journalisation
Private m_isResolving As Object           ' Dictionary pour détecter les dépendances circulaires

' Types et énumérations
Public Enum DependencyLifetime
    Transient = 0   ' Nouvelle instance à chaque résolution
    Singleton = 1   ' Instance unique pour toute l'application
End Enum

' Type pour stocker les informations sur une dépendance
Private Type DependencyInfo
    Name As String
    FactoryName As String
    Instance As Object
    Lifetime As DependencyLifetime
End Type

' Constantes
Private Const MAX_DEPENDENCIES As Long = 100

' Variables privées
Private m_dependencies() As DependencyInfo
Private m_dependencyCount As Long

'----------------------------------------------------------------------------------------
' Types personnalisés
'----------------------------------------------------------------------------------------
Private Type ServiceRegistration
    Name As String
    Instance As Object
    FactoryMethod As String
    Scope As String
    ImplementationName As String
    IsRegistered As Boolean
End Type

'----------------------------------------------------------------------------------------
' Variables privées
'----------------------------------------------------------------------------------------
Private m_services() As ServiceRegistration
Private m_serviceCount As Long
Private m_isInitialized As Boolean
Private m_lastError As String

' =========================================================================
' INITIALISATION
' =========================================================================

'@Description("Initialise le conteneur d'injection de dépendances")
Public Function Initialize(Optional ByVal logger As Object = Nothing) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier si déjà initialisé
    If m_isInitialized Then
        m_lastError = "Le conteneur est déjà initialisé"
        Initialize = True
        Exit Function
    End If
    
    ' Initialiser la collection de services
    ReDim m_services(1 To MAX_SERVICES) As ServiceRegistration
    m_serviceCount = 0
    
    ' Stocker le logger s'il est fourni
    If Not logger Is Nothing Then
        Set m_logger = logger
    End If
    
    ' Marquer comme initialisé
    m_isInitialized = True
    
    ' Log de l'initialisation
    LogInfo "Conteneur d'injection de dépendances initialisé avec succès"
    
    Initialize = True
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de l'initialisation du conteneur: " & Err.Description
    Initialize = False
End Function

'@Description("Vérifie si le conteneur est initialisé")
Public Function IsInitialized() As Boolean
    IsInitialized = m_isInitialized
End Function

'@Description("Enregistre les composants fondamentaux du système")
Private Sub RegisterCoreComponents()
    ' Auto-enregistrement du container
    RegisterInstance "IDependencyContainer", Me
    
    ' Journalisation - par défaut utilise la journalisation Excel
    RegisterSingleton "ILoggerBase", "clsExcelLogger"
    
    ' Gestion d'erreurs
    RegisterSingleton "IErrorHandler", "clsExcelErrorHandler"
End Sub

' =========================================================================
' ENREGISTREMENT
' =========================================================================

'@Description("Enregistre une classe comme implémentation d'une interface")
Public Function Register(ByVal interfaceName As String, ByVal implementationClass As String, Optional ByVal transient As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier l'initialisation
    If Not m_initialized Then
        Initialize
    End If
    
    ' Vérifier si l'interface est déjà enregistrée
    Dim idx As Long
    idx = FindRegistration(interfaceName)
    
    ' Si l'enregistrement existe déjà, le mettre à jour
    If idx > 0 Then
        With m_registrations(idx)
            .InterfaceName = interfaceName
            .ImplementationClass = implementationClass
            .FactoryMethod = ""
            If transient Then
                .RegistrationType = TransientType
            Else
                .RegistrationType = SingletonType
            End If
            Set .Instance = Nothing ' Réinitialiser l'instance si c'était un singleton
        End With
    Else
        ' Sinon, ajouter un nouvel enregistrement
        If m_count >= MAX_REGISTRATIONS Then
            If Not m_logger Is Nothing Then
                m_logger.LogError MODULE_NAME, "Nombre maximum d'enregistrements atteint"
            End If
            Register = False
            Exit Function
        End If
        
        m_count = m_count + 1
        With m_registrations(m_count)
            .InterfaceName = interfaceName
            .ImplementationClass = implementationClass
            .FactoryMethod = ""
            If transient Then
                .RegistrationType = TransientType
            Else
                .RegistrationType = SingletonType
            End If
            Set .Instance = Nothing
        End With
    End If
    
    ' Log de l'enregistrement
    If Not m_logger Is Nothing Then
        m_logger.LogInfo MODULE_NAME, "Interface '" & interfaceName & "' enregistrée avec implémentation '" & implementationClass & "'"
    End If
    
    Register = True
    Exit Function
    
ErrorHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError MODULE_NAME, "Erreur lors de l'enregistrement: " & Err.Description
    End If
    Register = False
End Function

'@Description("Enregistre une méthode factory pour créer des instances d'une interface")
Public Function RegisterFactory(ByVal interfaceName As String, ByVal factoryMethod As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier l'initialisation
    If Not m_initialized Then
        Initialize
    End If
    
    ' Vérifier si l'interface est déjà enregistrée
    Dim idx As Long
    idx = FindRegistration(interfaceName)
    
    ' Si l'enregistrement existe déjà, le mettre à jour
    If idx > 0 Then
        With m_registrations(idx)
            .InterfaceName = interfaceName
            .ImplementationClass = ""
            .FactoryMethod = factoryMethod
            .RegistrationType = FactoryType
            Set .Instance = Nothing ' Réinitialiser l'instance
        End With
    Else
        ' Sinon, ajouter un nouvel enregistrement
        If m_count >= MAX_REGISTRATIONS Then
            If Not m_logger Is Nothing Then
                m_logger.LogError MODULE_NAME, "Nombre maximum d'enregistrements atteint"
            End If
            RegisterFactory = False
            Exit Function
        End If
        
        m_count = m_count + 1
        With m_registrations(m_count)
            .InterfaceName = interfaceName
            .ImplementationClass = ""
            .FactoryMethod = factoryMethod
            .RegistrationType = FactoryType
            Set .Instance = Nothing
        End With
    End If
    
    ' Log de l'enregistrement
    If Not m_logger Is Nothing Then
        m_logger.LogInfo MODULE_NAME, "Interface '" & interfaceName & "' enregistrée avec factory '" & factoryMethod & "'"
    End If
    
    RegisterFactory = True
    Exit Function
    
ErrorHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError MODULE_NAME, "Erreur lors de l'enregistrement de factory: " & Err.Description
    End If
    RegisterFactory = False
End Function

'@Description("Enregistre une instance directement pour une interface")
Public Function RegisterInstance(ByVal interfaceName As String, ByVal instance As Object) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier l'initialisation
    If Not m_initialized Then
        Initialize
    End If
    
    ' Vérifier si l'interface est déjà enregistrée
    Dim idx As Long
    idx = FindRegistration(interfaceName)
    
    ' Si l'enregistrement existe déjà, le mettre à jour
    If idx > 0 Then
        With m_registrations(idx)
            .InterfaceName = interfaceName
            .ImplementationClass = TypeName(instance)
            .FactoryMethod = ""
            .RegistrationType = SingletonType
            Set .Instance = instance
        End With
    Else
        ' Sinon, ajouter un nouvel enregistrement
        If m_count >= MAX_REGISTRATIONS Then
            If Not m_logger Is Nothing Then
                m_logger.LogError MODULE_NAME, "Nombre maximum d'enregistrements atteint"
            End If
            RegisterInstance = False
            Exit Function
        End If
        
        m_count = m_count + 1
        With m_registrations(m_count)
            .InterfaceName = interfaceName
            .ImplementationClass = TypeName(instance)
            .FactoryMethod = ""
            .RegistrationType = SingletonType
            Set .Instance = instance
        End With
    End If
    
    ' Log de l'enregistrement
    If Not m_logger Is Nothing Then
        m_logger.LogInfo MODULE_NAME, "Instance '" & TypeName(instance) & "' enregistrée pour interface '" & interfaceName & "'"
    End If
    
    RegisterInstance = True
    Exit Function
    
ErrorHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError MODULE_NAME, "Erreur lors de l'enregistrement d'instance: " & Err.Description
    End If
    RegisterInstance = False
End Function

'@Description("Ajoute une dépendance à un enregistrement")
Public Sub AddDependency(ByVal interfaceName As String, ByVal dependencyName As String)
    EnsureInitialized
    
    If Not m_registrations.Exists(interfaceName) Then
        Err.Raise ERR_TYPE_NOT_REGISTERED, CLASS_NAME, _
                  "Type non enregistré: " & interfaceName
    End If
    
    Dim reg As Registration
    reg = m_registrations(interfaceName)
    
    ' Vérifier si la dépendance existe déjà
    Dim i As Long
    Dim exists As Boolean
    exists = False
    
    For i = 1 To reg.Dependencies.Count
        If reg.Dependencies(i) = dependencyName Then
            exists = True
            Exit For
        End If
    Next i
    
    ' Ajouter la dépendance si elle n'existe pas déjà
    If Not exists Then
        reg.Dependencies.Add dependencyName
    End If
    
    ' Mettre à jour l'enregistrement
    m_registrations(interfaceName) = reg
End Sub

' =========================================================================
' RÉSOLUTION
' =========================================================================

'@Description("Résout une interface pour obtenir une implémentation")
Public Function Resolve(ByVal interfaceName As String) As Object
    On Error GoTo ErrorHandler
    
    ' Vérifier l'initialisation
    If Not m_initialized Then
        Initialize
    End If
    
    ' Vérifier si l'interface est enregistrée
    Dim idx As Long
    idx = FindRegistration(interfaceName)
    
    ' Si l'interface n'est pas enregistrée, retourner Nothing
    If idx <= 0 Then
        If Not m_logger Is Nothing Then
            m_logger.LogWarning MODULE_NAME, "Interface '" & interfaceName & "' non enregistrée"
        End If
        Set Resolve = Nothing
        Exit Function
    End If
    
    ' Résoudre en fonction du type d'enregistrement
    With m_registrations(idx)
        Select Case .RegistrationType
            Case SingletonType
                ' Si c'est un singleton et que l'instance existe déjà, la retourner
                If Not .Instance Is Nothing Then
                    Set Resolve = .Instance
                Else
                    ' Sinon, créer l'instance et la stocker
                    Set .Instance = CreateInstance(.ImplementationClass)
                    Set Resolve = .Instance
                End If
                
            Case TransientType
                ' Pour les transients, créer une nouvelle instance à chaque fois
                Set Resolve = CreateInstance(.ImplementationClass)
                
            Case FactoryType
                ' Pour les factories, appeler la méthode factory
                Set Resolve = CallFactoryMethod(.FactoryMethod)
                
        End Select
    End With
    
    ' Log de la résolution
    If Not m_logger Is Nothing Then
        m_logger.LogInfo MODULE_NAME, "Interface '" & interfaceName & "' résolue avec succès"
    End If
    
    Exit Function
    
ErrorHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError MODULE_NAME, "Erreur lors de la résolution de '" & interfaceName & "': " & Err.Description
    End If
    Set Resolve = Nothing
End Function

' =========================================================================
' GESTION DU CONTENEUR
' =========================================================================

'@Description("Vérifie si une interface est enregistrée")
Public Function IsRegistered(ByVal interfaceName As String) As Boolean
    IsRegistered = (FindRegistration(interfaceName) > 0)
End Function

'@Description("Supprime un enregistrement")
Public Sub Unregister(ByVal interfaceName As String)
    EnsureInitialized
    
    If m_registrations.Exists(interfaceName) Then
        m_registrations.Remove interfaceName
    End If
End Sub

'@Description("Vide le conteneur")
Public Sub Clear()
    EnsureInitialized
    
    Dim i As Long
    For i = 1 To m_count
        Set m_registrations(i).Instance = Nothing
    Next i
    
    m_count = 0
    ReDim m_registrations(1 To MAX_REGISTRATIONS)
    
    ' Nettoyer toutes les dépendances
    Dim j As Long
    For j = 1 To m_dependencyCount
        Set m_dependencies(j).Instance = Nothing
    Next j
    
    ' Réinitialiser le compteur
    m_dependencyCount = 0
    
    ' Réinitialiser le tableau
    ReDim m_dependencies(1 To MAX_DEPENDENCIES)
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo MODULE_NAME, "Conteneur d'injection de dépendances vidé"
    End If
End Sub

'@Description("Obtient le nombre d'enregistrements")
Public Function GetRegistrationCount() As Long
    EnsureInitialized
    GetRegistrationCount = m_registrations.Count
End Function

'@Description("Résout tous les singletons pour les initialiser")
Public Sub ResolveAllSingletons()
    EnsureInitialized
    
    Dim key As Variant
    
    For Each key In m_registrations.Keys
        Dim reg As Registration
        reg = m_registrations(key)
        
        If reg.RegistrationType = SingletonType Then
            Dim instance As Object
            Set instance = Resolve(CStr(key))
        End If
    Next key
End Sub

' =========================================================================
' MÉTHODES PRIVÉES
' =========================================================================

' ===========================================================================
' Sub: LogInfo
' Purpose: Journalise un message d'information
' Parameters:
'   - message: Message à journaliser
' ===========================================================================
Private Sub LogInfo(ByVal message As String)
    On Error Resume Next
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo message, LOG_CATEGORY
    End If
End Sub

' ===========================================================================
' Sub: LogWarning
' Purpose: Journalise un message d'avertissement
' Parameters:
'   - message: Message à journaliser
' ===========================================================================
Private Sub LogWarning(ByVal message As String)
    On Error Resume Next
    
    If Not m_logger Is Nothing Then
        m_logger.LogWarning message, LOG_CATEGORY
    End If
End Sub

' ===========================================================================
' Sub: LogError
' Purpose: Journalise un message d'erreur
' Parameters:
'   - message: Message à journaliser
' ===========================================================================
Private Sub LogError(ByVal message As String)
    On Error Resume Next
    
    If Not m_logger Is Nothing Then
        m_logger.LogError message, LOG_CATEGORY
    Else
        Debug.Print "[ERROR] [" & LOG_CATEGORY & "] " & message
    End If
End Sub

'===========================================================
' UTILITAIRES PRIVÉS
'===========================================================

'@Description("Trouve l'index d'enregistrement pour une interface")
Private Function FindRegistration(ByVal interfaceName As String) As Long
    Dim i As Long
    For i = 1 To m_count
        If m_registrations(i).InterfaceName = interfaceName Then
            FindRegistration = i
            Exit Function
        End If
    Next i
    
    FindRegistration = 0 ' Non trouvé
End Function

'@Description("Crée une instance d'une classe par son nom")
Private Function CreateInstance(ByVal className As String) As Object
    On Error GoTo ErrorHandler
    
    ' Tenter de créer l'instance via Factory APEX si disponible
    Dim instance As Object
    Dim factoryAvailable As Boolean
    factoryAvailable = False
    
    On Error Resume Next
    ' Vérifier si le Factory APEX est disponible
    If Application.Run("modFactory.IsClassAvailable", className) Then
        ' Utiliser le Factory APEX pour créer l'instance
        Set instance = Application.Run("modFactory.CreateInstance", className)
        factoryAvailable = True
    End If
    On Error GoTo ErrorHandler
    
    ' Si la Factory n'est pas disponible, essayer de créer directement
    If Not factoryAvailable Then
        ' Créer l'instance via CreateObject
        On Error Resume Next
        Set instance = CreateObject("APEX_VBA." & className)
        If Err.Number <> 0 Then
            ' Si échoue, essayer via New
            Err.Clear
            Select Case className
                Case "clsDependencyContainer"
                    Set instance = New clsDependencyContainer
                Case "clsUnitOfWork"
                    Set instance = New clsUnitOfWork
                ' Ajouter d'autres cas selon les classes disponibles
                Case Else
                    Err.Raise 5, , "Classe non reconnue: " & className
            End Select
        End If
        On Error GoTo ErrorHandler
    End If
    
    Set CreateInstance = instance
    Exit Function
    
ErrorHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError MODULE_NAME, "Erreur lors de la création d'instance de '" & className & "': " & Err.Description
    End If
    Set CreateInstance = Nothing
End Function

'@Description("Appelle une méthode factory pour créer une instance")
Private Function CallFactoryMethod(ByVal factoryMethod As String) As Object
    On Error GoTo ErrorHandler
    
    Dim instance As Object
    
    ' Déterminer le module et la méthode
    Dim moduleName As String
    Dim methodName As String
    
    ' Si le nom contient un point, c'est module.méthode
    If InStr(factoryMethod, ".") > 0 Then
        moduleName = Left$(factoryMethod, InStr(factoryMethod, ".") - 1)
        methodName = Mid$(factoryMethod, InStr(factoryMethod, ".") + 1)
        
        ' Appeler via Application.Run avec module et méthode
        Set instance = Application.Run(factoryMethod)
    Else
        ' Sinon, c'est juste une méthode dans ce module (modApexBridge)
        ' Appeler via Application.Run pour la méthode
        Set instance = Application.Run("modApexBridge." & factoryMethod)
    End If
    
    Set CallFactoryMethod = instance
    Exit Function
    
ErrorHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError MODULE_NAME, "Erreur lors de l'appel de factory '" & factoryMethod & "': " & Err.Description
    End If
    Set CallFactoryMethod = Nothing
End Function

'===========================================================
' GESTION DES ERREURS
'===========================================================

'@Description("Définit le logger pour ce module")
Public Sub SetLogger(ByVal logger As Object)
    Set m_logger = logger
End Sub

Private Sub Class_Terminate()
    ' Libérer les références
    Clear
    Set m_logger = Nothing
End Sub

'===========================================================
' GESTION DES DÉPENDANCES
'===========================================================

'@Description("Enregistre une fabrique pour un type de dépendance")
Public Function RegisterFactory(ByVal dependencyName As String, ByVal factoryName As String, _
                               Optional ByVal asSingleton As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier l'initialisation
    If Not m_initialized Then
        If Not m_logger Is Nothing Then
            m_logger.LogError MODULE_NAME, "Le conteneur n'est pas initialisé"
        End If
        RegisterFactory = False
        Exit Function
    End If
    
    ' Vérifier si le nombre maximum de dépendances est atteint
    If m_dependencyCount >= MAX_DEPENDENCIES Then
        If Not m_logger Is Nothing Then
            m_logger.LogError MODULE_NAME, "Nombre maximum de dépendances atteint"
        End If
        RegisterFactory = False
        Exit Function
    End If
    
    ' Vérifier si la dépendance existe déjà
    Dim index As Long
    index = FindDependencyIndex(dependencyName)
    
    If index > 0 Then
        ' Mettre à jour la dépendance existante
        m_dependencies(index).FactoryName = factoryName
        m_dependencies(index).Lifetime = IIf(asSingleton, DependencyLifetime.Singleton, DependencyLifetime.Transient)
        Set m_dependencies(index).Instance = Nothing ' Réinitialiser l'instance
    Else
        ' Ajouter une nouvelle dépendance
        m_dependencyCount = m_dependencyCount + 1
        m_dependencies(m_dependencyCount).Name = dependencyName
        m_dependencies(m_dependencyCount).FactoryName = factoryName
        m_dependencies(m_dependencyCount).Lifetime = IIf(asSingleton, DependencyLifetime.Singleton, DependencyLifetime.Transient)
        Set m_dependencies(m_dependencyCount).Instance = Nothing
    End If
    
    ' Log du succès de l'enregistrement
    If Not m_logger Is Nothing Then
        m_logger.LogInfo MODULE_NAME, "Fabrique '" & factoryName & "' enregistrée pour '" & dependencyName & "' comme " & _
                                     IIf(asSingleton, "singleton", "transient")
    End If
    
    RegisterFactory = True
    Exit Function
    
ErrorHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError MODULE_NAME, "Erreur lors de l'enregistrement de la fabrique '" & factoryName & "' pour '" & _
                                     dependencyName & "': " & Err.Description
    End If
    RegisterFactory = False
End Function

'@Description("Enregistre une instance comme une dépendance")
Public Function RegisterInstance(ByVal dependencyName As String, ByVal instance As Object) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier l'initialisation
    If Not m_initialized Then
        If Not m_logger Is Nothing Then
            m_logger.LogError MODULE_NAME, "Le conteneur n'est pas initialisé"
        End If
        RegisterInstance = False
        Exit Function
    End If
    
    ' Vérifier si l'instance est fournie
    If instance Is Nothing Then
        If Not m_logger Is Nothing Then
            m_logger.LogError MODULE_NAME, "L'instance ne peut pas être Nothing"
        End If
        RegisterInstance = False
        Exit Function
    End If
    
    ' Vérifier si le nombre maximum de dépendances est atteint
    If m_dependencyCount >= MAX_DEPENDENCIES Then
        If Not m_logger Is Nothing Then
            m_logger.LogError MODULE_NAME, "Nombre maximum de dépendances atteint"
        End If
        RegisterInstance = False
        Exit Function
    End If
    
    ' Vérifier si la dépendance existe déjà
    Dim index As Long
    index = FindDependencyIndex(dependencyName)
    
    If index > 0 Then
        ' Mettre à jour la dépendance existante
        m_dependencies(index).FactoryName = ""
        m_dependencies(index).Lifetime = DependencyLifetime.Singleton
        Set m_dependencies(index).Instance = instance
    Else
        ' Ajouter une nouvelle dépendance
        m_dependencyCount = m_dependencyCount + 1
        m_dependencies(m_dependencyCount).Name = dependencyName
        m_dependencies(m_dependencyCount).FactoryName = ""
        m_dependencies(m_dependencyCount).Lifetime = DependencyLifetime.Singleton
        Set m_dependencies(m_dependencyCount).Instance = instance
    End If
    
    ' Log du succès de l'enregistrement
    If Not m_logger Is Nothing Then
        m_logger.LogInfo MODULE_NAME, "Instance enregistrée pour '" & dependencyName & "'"
    End If
    
    RegisterInstance = True
    Exit Function
    
ErrorHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError MODULE_NAME, "Erreur lors de l'enregistrement de l'instance pour '" & dependencyName & "': " & Err.Description
    End If
    RegisterInstance = False
End Function

'@Description("Résout une dépendance par son nom")
Public Function Resolve(ByVal dependencyName As String) As Object
    On Error GoTo ErrorHandler
    
    ' Vérifier l'initialisation
    If Not m_initialized Then
        If Not m_logger Is Nothing Then
            m_logger.LogError MODULE_NAME, "Le conteneur n'est pas initialisé"
        End If
        Set Resolve = Nothing
        Exit Function
    End If
    
    ' Trouver l'index de la dépendance
    Dim index As Long
    index = FindDependencyIndex(dependencyName)
    
    If index = 0 Then
        ' La dépendance n'existe pas
        If Not m_logger Is Nothing Then
            m_logger.LogError MODULE_NAME, "Dépendance '" & dependencyName & "' non trouvée"
        End If
        Set Resolve = Nothing
        Exit Function
    End If
    
    ' Vérifier si c'est un singleton avec une instance existante
    If m_dependencies(index).Lifetime = DependencyLifetime.Singleton And Not m_dependencies(index).Instance Is Nothing Then
        Set Resolve = m_dependencies(index).Instance
        Exit Function
    End If
    
    ' Résoudre en appelant la fabrique
    If Len(m_dependencies(index).FactoryName) > 0 Then
        Dim factoryResult As Object
        Set factoryResult = CallFactory(m_dependencies(index).FactoryName)
        
        If factoryResult Is Nothing Then
            If Not m_logger Is Nothing Then
                m_logger.LogError MODULE_NAME, "La fabrique '" & m_dependencies(index).FactoryName & "' a renvoyé Nothing"
            End If
            Set Resolve = Nothing
            Exit Function
        End If
        
        ' Stocker l'instance si c'est un singleton
        If m_dependencies(index).Lifetime = DependencyLifetime.Singleton Then
            Set m_dependencies(index).Instance = factoryResult
        End If
        
        Set Resolve = factoryResult
    Else
        ' Pas de fabrique et pas d'instance
        If Not m_logger Is Nothing Then
            m_logger.LogError MODULE_NAME, "Aucune fabrique ni instance pour '" & dependencyName & "'"
        End If
        Set Resolve = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError MODULE_NAME, "Erreur lors de la résolution de '" & dependencyName & "': " & Err.Description
    End If
    Set Resolve = Nothing
End Function

'@Description("Vérifie si une dépendance est enregistrée")
Public Function HasDependency(ByVal dependencyName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier l'initialisation
    If Not m_initialized Then
        If Not m_logger Is Nothing Then
            m_logger.LogWarning MODULE_NAME, "Le conteneur n'est pas initialisé"
        End If
        HasDependency = False
        Exit Function
    End If
    
    HasDependency = (FindDependencyIndex(dependencyName) > 0)
    Exit Function
    
ErrorHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError MODULE_NAME, "Erreur lors de la vérification de la dépendance '" & dependencyName & "': " & Err.Description
    End If
    HasDependency = False
End Function

'===========================================================
' MÉTHODES PRIVÉES
'===========================================================

'@Description("Trouve l'index d'une dépendance par son nom")
Private Function FindDependencyIndex(ByVal dependencyName As String) As Long
    Dim i As Long
    
    For i = 1 To m_dependencyCount
        If m_dependencies(i).Name = dependencyName Then
            FindDependencyIndex = i
            Exit Function
        End If
    Next i
    
    FindDependencyIndex = 0 ' Non trouvé
End Function

'@Description("Appelle une fabrique par son nom")
Private Function CallFactory(ByVal factoryName As String) As Object
    On Error GoTo ErrorHandler
    
    ' Vérifier si la fabrique existe dans une procédure publique du projet
    Dim result As Object
    
    ' Utiliser Application.Run pour appeler une procédure publique par son nom
    ' Attention : Application.Run ne fonctionne qu'avec des procédures publiques
    Set result = Application.Run(factoryName)
    
    Set CallFactory = result
    Exit Function
    
ErrorHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError MODULE_NAME, "Erreur lors de l'appel de la fabrique '" & factoryName & "': " & Err.Description
    End If
    Set CallFactory = Nothing
End Function

'===========================================================
' GESTIONNAIRE DE CLASSE
'===========================================================

Private Sub Class_Initialize()
    m_initialized = False
    ReDim m_registrations(1 To MAX_REGISTRATIONS)
    m_count = 0
    ReDim m_dependencies(1 To MAX_DEPENDENCIES)
    m_dependencyCount = 0
End Sub

Private Sub Class_Terminate()
    ' Nettoyer les références
    Clear
    Set m_logger = Nothing
End Sub

'----------------------------------------------------------------------------------------
' Enregistrement de services
'----------------------------------------------------------------------------------------

'@Description("Enregistre une instance d'objet comme service")
Public Function RegisterInstance(ByVal serviceName As String, ByVal instance As Object) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier l'initialisation
    If Not m_isInitialized Then
        m_lastError = "Le conteneur n'est pas initialisé"
        RegisterInstance = False
        Exit Function
    End If
    
    ' Vérifier les arguments
    If Len(serviceName) = 0 Then
        m_lastError = "Le nom du service ne peut pas être vide"
        RegisterInstance = False
        Exit Function
    End If
    
    If instance Is Nothing Then
        m_lastError = "L'instance ne peut pas être Nothing"
        RegisterInstance = False
        Exit Function
    End If
    
    ' Vérifier si le service existe déjà
    Dim existingIndex As Long
    existingIndex = FindServiceByName(serviceName)
    
    ' Si le service existe, mettre à jour l'instance
    If existingIndex > 0 Then
        Set m_services(existingIndex).Instance = instance
        m_services(existingIndex).FactoryMethod = ""
        LogInfo "Service '" & serviceName & "' mis à jour avec une nouvelle instance"
        RegisterInstance = True
        Exit Function
    End If
    
    ' Sinon, ajouter un nouveau service
    If m_serviceCount >= MAX_SERVICES Then
        m_lastError = "Nombre maximum de services atteint"
        RegisterInstance = False
        Exit Function
    End If
    
    ' Incrémenter le compteur et ajouter le service
    m_serviceCount = m_serviceCount + 1
    With m_services(m_serviceCount)
        .Name = serviceName
        Set .Instance = instance
        .FactoryMethod = ""
        .Scope = "Singleton"
        .IsRegistered = True
    End With
    
    LogInfo "Service '" & serviceName & "' enregistré avec succès"
    RegisterInstance = True
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de l'enregistrement du service '" & serviceName & "': " & Err.Description
    RegisterInstance = False
End Function

'@Description("Enregistre une factory pour un service")
Public Function RegisterFactory(ByVal serviceName As String, ByVal factoryMethod As String, _
                               Optional ByVal scope As String = "Transient", _
                               Optional ByVal implementationName As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier l'initialisation
    If Not m_isInitialized Then
        m_lastError = "Le conteneur n'est pas initialisé"
        RegisterFactory = False
        Exit Function
    End If
    
    ' Vérifier les arguments
    If Len(serviceName) = 0 Then
        m_lastError = "Le nom du service ne peut pas être vide"
        RegisterFactory = False
        Exit Function
    End If
    
    If Len(factoryMethod) = 0 Then
        m_lastError = "La méthode factory ne peut pas être vide"
        RegisterFactory = False
        Exit Function
    End If
    
    ' Vérifier si la méthode factory existe
    Dim isMethodValid As Boolean
    isMethodValid = IsMethodValid(factoryMethod)
    
    If Not isMethodValid Then
        m_lastError = "La méthode factory '" & factoryMethod & "' n'est pas valide"
        RegisterFactory = False
        Exit Function
    End If
    
    ' Vérifier si le service existe déjà
    Dim existingIndex As Long
    existingIndex = FindServiceByName(serviceName)
    
    ' Si le service existe, mettre à jour la factory
    If existingIndex > 0 Then
        With m_services(existingIndex)
            Set .Instance = Nothing
            .FactoryMethod = factoryMethod
            .Scope = scope
            If Len(implementationName) > 0 Then
                .ImplementationName = implementationName
            End If
        End With
        
        LogInfo "Service '" & serviceName & "' mis à jour avec une nouvelle factory"
        RegisterFactory = True
        Exit Function
    End If
    
    ' Sinon, ajouter un nouveau service
    If m_serviceCount >= MAX_SERVICES Then
        m_lastError = "Nombre maximum de services atteint"
        RegisterFactory = False
        Exit Function
    End If
    
    ' Incrémenter le compteur et ajouter le service
    m_serviceCount = m_serviceCount + 1
    With m_services(m_serviceCount)
        .Name = serviceName
        Set .Instance = Nothing
        .FactoryMethod = factoryMethod
        .Scope = scope
        .ImplementationName = implementationName
        .IsRegistered = True
    End With
    
    LogInfo "Factory pour le service '" & serviceName & "' enregistrée avec succès"
    RegisterFactory = True
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de l'enregistrement de la factory pour '" & serviceName & "': " & Err.Description
    RegisterFactory = False
End Function

'@Description("Vérifie si une méthode factory est valide")
Private Function IsMethodValid(ByVal methodName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Pour cette implémentation, nous supposons que la méthode est valide
    ' Dans une implémentation réelle, il faudrait utiliser la réflexion pour vérifier
    
    IsMethodValid = True
    Exit Function
    
ErrorHandler:
    IsMethodValid = False
End Function

'----------------------------------------------------------------------------------------
' Résolution de services
'----------------------------------------------------------------------------------------

'@Description("Résout un service par son nom")
Public Function ResolveByName(ByVal serviceName As String) As Object
    On Error GoTo ErrorHandler
    
    ' Vérifier l'initialisation
    If Not m_isInitialized Then
        m_lastError = "Le conteneur n'est pas initialisé"
        Set ResolveByName = Nothing
        Exit Function
    End If
    
    ' Trouver le service par son nom
    Dim serviceIndex As Long
    serviceIndex = FindServiceByName(serviceName)
    
    If serviceIndex <= 0 Then
        m_lastError = "Service '" & serviceName & "' non trouvé"
        LogWarning "Tentative de résolution d'un service non enregistré: " & serviceName
        Set ResolveByName = Nothing
        Exit Function
    End If
    
    ' Résoudre le service selon son type d'enregistrement
    Dim service As ServiceRegistration
    service = m_services(serviceIndex)
    
    ' Si une instance est déjà disponible et c'est un singleton, la retourner
    If Not service.Instance Is Nothing And service.Scope = "Singleton" Then
        Set ResolveByName = service.Instance
        Exit Function
    End If
    
    ' Si une factory est définie, l'utiliser pour créer une instance
    If Len(service.FactoryMethod) > 0 Then
        Dim factoryResult As Object
        
        ' Appeler la méthode factory
        Set factoryResult = Application.Run(service.FactoryMethod)
        
        ' Si c'est un singleton, stocker l'instance
        If service.Scope = "Singleton" And Not factoryResult Is Nothing Then
            Set m_services(serviceIndex).Instance = factoryResult
        End If
        
        ' Retourner l'instance créée
        Set ResolveByName = factoryResult
        Exit Function
    End If
    
    ' Si aucune méthode de résolution n'est disponible
    m_lastError = "Aucune méthode de résolution disponible pour le service '" & serviceName & "'"
    LogWarning "Impossible de résoudre le service '" & serviceName & "': aucune méthode de résolution"
    Set ResolveByName = Nothing
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de la résolution du service '" & serviceName & "': " & Err.Description
    LogError "Erreur lors de la résolution du service '" & serviceName & "': " & Err.Description
    Set ResolveByName = Nothing
End Function

'----------------------------------------------------------------------------------------
' Gestion des services
'----------------------------------------------------------------------------------------

'@Description("Trouve l'index d'un service par son nom")
Private Function FindServiceByName(ByVal serviceName As String) As Long
    Dim i As Long
    
    For i = 1 To m_serviceCount
        If m_services(i).IsRegistered And m_services(i).Name = serviceName Then
            FindServiceByName = i
            Exit Function
        End If
    Next i
    
    FindServiceByName = 0
End Function

'@Description("Supprime un service par son nom")
Public Function RemoveService(ByVal serviceName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier l'initialisation
    If Not m_isInitialized Then
        m_lastError = "Le conteneur n'est pas initialisé"
        RemoveService = False
        Exit Function
    End If
    
    ' Trouver le service par son nom
    Dim serviceIndex As Long
    serviceIndex = FindServiceByName(serviceName)
    
    If serviceIndex <= 0 Then
        m_lastError = "Service '" & serviceName & "' non trouvé"
        RemoveService = False
        Exit Function
    End If
    
    ' Nettoyer les références
    With m_services(serviceIndex)
        Set .Instance = Nothing
        .FactoryMethod = ""
        .IsRegistered = False
    End With
    
    LogInfo "Service '" & serviceName & "' supprimé avec succès"
    RemoveService = True
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de la suppression du service '" & serviceName & "': " & Err.Description
    RemoveService = False
End Function

'@Description("Supprime tous les services enregistrés")
Public Sub Clear()
    On Error Resume Next
    
    Dim i As Long
    
    For i = 1 To m_serviceCount
        If m_services(i).IsRegistered Then
            Set m_services(i).Instance = Nothing
            m_services(i).IsRegistered = False
        End If
    Next i
    
    ReDim m_services(1 To MAX_SERVICES) As ServiceRegistration
    m_serviceCount = 0
    
    LogInfo "Tous les services ont été supprimés"
End Sub

'----------------------------------------------------------------------------------------
' Propriétés
'----------------------------------------------------------------------------------------

'@Description("Vérifie si un service est enregistré")
Public Function IsServiceRegistered(ByVal serviceName As String) As Boolean
    IsServiceRegistered = (FindServiceByName(serviceName) > 0)
End Function

'@Description("Retourne le nombre de services enregistrés")
Public Property Get ServiceCount() As Long
    ServiceCount = m_serviceCount
End Property

'@Description("Indique si le conteneur est initialisé")
Public Property Get IsInitialized() As Boolean
    IsInitialized = m_isInitialized
End Property

'@Description("Retourne la dernière erreur survenue")
Public Property Get LastError() As String
    LastError = m_lastError
End Property

'----------------------------------------------------------------------------------------
' Journalisation
'----------------------------------------------------------------------------------------

'@Description("Enregistre un message d'information")
Private Sub LogInfo(ByVal message As String)
    On Error Resume Next
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo MODULE_NAME, message
    Else
        Debug.Print "[INFO] [" & MODULE_NAME & "] " & message
    End If
End Sub

'@Description("Enregistre un avertissement")
Private Sub LogWarning(ByVal message As String)
    On Error Resume Next
    
    If Not m_logger Is Nothing Then
        m_logger.LogWarning MODULE_NAME, message
    Else
        Debug.Print "[WARNING] [" & MODULE_NAME & "] " & message
    End If
End Sub

'@Description("Enregistre une erreur")
Private Sub LogError(ByVal message As String)
    On Error Resume Next
    
    If Not m_logger Is Nothing Then
        m_logger.LogError MODULE_NAME, message
    Else
        Debug.Print "[ERROR] [" & MODULE_NAME & "] " & message
    End If
End Sub

'----------------------------------------------------------------------------------------
' Classe Events
'----------------------------------------------------------------------------------------

Private Sub Class_Initialize()
    m_isInitialized = False
    m_serviceCount = 0
    m_lastError = ""
End Sub

Private Sub Class_Terminate()
    Clear
End Sub 