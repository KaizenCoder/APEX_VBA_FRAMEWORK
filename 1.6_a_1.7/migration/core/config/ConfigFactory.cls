VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@ModuleDescription("Factory pour le gestionnaire de configuration")
'@Folder("APEX.Core.Config")
'@Version 1.7.0
'@Author APEX Framework Team
Option Explicit

' =========================================================
' ConfigFactory
' Factory pour le gestionnaire de configuration
' =========================================================

' Cache des gestionnaires de configuration
Private m_configManagers As Object  ' Dictionnaire de configurations
Private m_defaultConfig As Object   ' Configuration par défaut
Private m_defaultLogger As Object   ' Logger par défaut
Private m_initialized As Boolean

' =========================================================
' Initialisation
' =========================================================
Private Sub Class_Initialize()
    ' Initialiser les caches
    Set m_configManagers = CreateObject("Scripting.Dictionary")
    m_initialized = False
End Sub

' =========================================================
' Méthodes publiques
' =========================================================

' Initialise la factory avec des paramètres par défaut
' @param defaultConfigPath Chemin vers la configuration par défaut
' @param logger Logger à utiliser (optionnel)
' @return True si l'initialisation a réussi, False sinon
Public Function Initialize(Optional ByVal defaultConfigPath As String = "", Optional ByVal logger As Object = Nothing) As Boolean
    On Error GoTo ErrorHandler
    
    ' Stocker le logger par défaut
    Set m_defaultLogger = logger
    
    ' Charger la configuration par défaut si spécifiée
    If Len(defaultConfigPath) > 0 Then
        Dim options As Object
        Set options = CreateObject("Scripting.Dictionary")
        If Not logger Is Nothing Then
            Set options("Logger") = logger
        End If
        
        Set m_defaultConfig = CreateConfigManager(defaultConfigPath, options)
        If m_defaultConfig Is Nothing Then
            Err.Raise vbObjectError + 1100, "ConfigFactory.Initialize", "Échec du chargement de la configuration par défaut"
        End If
    End If
    
    m_initialized = True
    Initialize = True
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de l'initialisation de ConfigFactory: " & Err.Description, "Initialize"
    Initialize = False
End Function

' Crée un gestionnaire de configuration
' @param configSource Source de configuration (chemin fichier, objet, etc.)
' @param options Options supplémentaires (optionnel)
' @return Gestionnaire de configuration initialisé
Public Function CreateConfigManager(ByVal configSource As Variant, Optional ByVal options As Object = Nothing) As IConfigManagerBase
    On Error GoTo ErrorHandler
    
    ' Créer un nouveau gestionnaire de configuration
    Dim configManager As New ConfigManager
    
    ' Ajouter le logger par défaut aux options si aucun n'est spécifié
    If options Is Nothing And Not m_defaultLogger Is Nothing Then
        Set options = CreateObject("Scripting.Dictionary")
        Set options("Logger") = m_defaultLogger
    End If
    
    ' Initialiser le gestionnaire
    If Not configManager.Initialize(configSource, options) Then
        Err.Raise vbObjectError + 1101, "ConfigFactory.CreateConfigManager", "Échec de l'initialisation du gestionnaire de configuration"
    End If
    
    ' Renvoyer le gestionnaire initialisé
    Set CreateConfigManager = configManager
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la création du ConfigManager: " & Err.Description, "CreateConfigManager"
    Set CreateConfigManager = Nothing
End Function

' Récupère un gestionnaire de configuration du cache ou en crée un nouveau
' @param configName Nom de la configuration
' @param configSource Source de configuration si nouvelle création
' @param options Options supplémentaires (optionnel)
' @return Gestionnaire de configuration
Public Function GetConfigManager(ByVal configName As String, Optional ByVal configSource As Variant = Null, Optional ByVal options As Object = Nothing) As IConfigManagerBase
    On Error GoTo ErrorHandler
    
    ' Vérifier si le gestionnaire existe déjà dans le cache
    If m_configManagers.Exists(configName) Then
        ' Renvoyer le gestionnaire du cache
        Set GetConfigManager = m_configManagers(configName)
    Else
        ' Vérifier que la source est spécifiée pour une nouvelle création
        If IsNull(configSource) Then
            Err.Raise vbObjectError + 1102, "ConfigFactory.GetConfigManager", "Source de configuration non spécifiée pour une nouvelle création"
        End If
        
        ' Créer un nouveau gestionnaire
        Dim configManager As IConfigManagerBase
        Set configManager = CreateConfigManager(configSource, options)
        
        ' Ajouter au cache
        m_configManagers.Add configName, configManager
        
        ' Renvoyer le nouveau gestionnaire
        Set GetConfigManager = configManager
    End If
    
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la récupération du ConfigManager '" & configName & "': " & Err.Description, "GetConfigManager"
    Set GetConfigManager = Nothing
End Function

' Récupère la configuration par défaut
' @return Gestionnaire de configuration par défaut
Public Function GetDefaultConfig() As IConfigManagerBase
    On Error GoTo ErrorHandler
    
    ' Vérifier si la configuration par défaut existe
    If m_defaultConfig Is Nothing Then
        Err.Raise vbObjectError + 1103, "ConfigFactory.GetDefaultConfig", "Aucune configuration par défaut définie"
    End If
    
    ' Renvoyer la configuration par défaut
    Set GetDefaultConfig = m_defaultConfig
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la récupération de la configuration par défaut: " & Err.Description, "GetDefaultConfig"
    Set GetDefaultConfig = Nothing
End Function

' Récupère tous les noms de configurations disponibles
' @return Collection de noms de configurations
Public Function GetConfigNames() As Collection
    On Error GoTo ErrorHandler
    
    ' Créer une collection pour les noms
    Dim result As New Collection
    
    ' Ajouter les noms au résultat
    Dim key As Variant
    For Each key In m_configManagers.Keys
        result.Add key
    Next
    
    Set GetConfigNames = result
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la récupération des noms de configurations: " & Err.Description, "GetConfigNames"
    Set GetConfigNames = New Collection
End Function

' Supprime une configuration du cache
' @param configName Nom de la configuration à supprimer
' @return True si la suppression a réussi, False sinon
Public Function RemoveConfig(ByVal configName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier si la configuration existe
    If m_configManagers.Exists(configName) Then
        ' Supprimer la configuration du cache
        m_configManagers.Remove configName
        RemoveConfig = True
    Else
        ' La configuration n'existe pas
        RemoveConfig = False
    End If
    
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la suppression de la configuration '" & configName & "': " & Err.Description, "RemoveConfig"
    RemoveConfig = False
End Function

' Définit la configuration par défaut
' @param configManager Gestionnaire de configuration à définir comme défaut
Public Sub SetDefaultConfig(ByVal configManager As IConfigManagerBase)
    On Error GoTo ErrorHandler
    
    ' Stocker la configuration par défaut
    Set m_defaultConfig = configManager
    Exit Sub
    
ErrorHandler:
    LogError "Erreur lors de la définition de la configuration par défaut: " & Err.Description, "SetDefaultConfig"
End Sub

' Définit le logger par défaut
' @param logger Logger à utiliser pour les nouvelles configurations
Public Sub SetDefaultLogger(ByVal logger As Object)
    On Error GoTo ErrorHandler
    
    ' Stocker le logger par défaut
    Set m_defaultLogger = logger
    Exit Sub
    
ErrorHandler:
    LogError "Erreur lors de la définition du logger par défaut: " & Err.Description, "SetDefaultLogger"
End Sub

' =========================================================
' Méthodes privées
' =========================================================

' Gestion des erreurs
Private Sub LogError(ByVal message As String, ByVal source As String)
    ' Logger l'erreur si un logger est disponible
    If Not m_defaultLogger Is Nothing Then
        On Error Resume Next
        m_defaultLogger.Error message, "ConfigFactory." & source
        On Error GoTo 0
    End If
    
    ' Également afficher dans Debug.Print pour le débogage
    Debug.Print "[ERROR][ConfigFactory." & source & "] " & message
End Sub 