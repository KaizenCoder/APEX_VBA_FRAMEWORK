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
Private m_defaultConfig As Object   ' Configuration par d�faut
Private m_defaultLogger As Object   ' Logger par d�faut
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
' M�thodes publiques
' =========================================================

' Initialise la factory avec des param�tres par d�faut
' @param defaultConfigPath Chemin vers la configuration par d�faut
' @param logger Logger � utiliser (optionnel)
' @return True si l'initialisation a r�ussi, False sinon
Public Function Initialize(Optional ByVal defaultConfigPath As String = "", Optional ByVal logger As Object = Nothing) As Boolean
    On Error GoTo ErrorHandler
    
    ' Stocker le logger par d�faut
    Set m_defaultLogger = logger
    
    ' Charger la configuration par d�faut si sp�cifi�e
    If Len(defaultConfigPath) > 0 Then
        Dim options As Object
        Set options = CreateObject("Scripting.Dictionary")
        If Not logger Is Nothing Then
            Set options("Logger") = logger
        End If
        
        Set m_defaultConfig = CreateConfigManager(defaultConfigPath, options)
        If m_defaultConfig Is Nothing Then
            Err.Raise vbObjectError + 1100, "ConfigFactory.Initialize", "�chec du chargement de la configuration par d�faut"
        End If
    End If
    
    m_initialized = True
    Initialize = True
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de l'initialisation de ConfigFactory: " & Err.Description, "Initialize"
    Initialize = False
End Function

' Cr�e un gestionnaire de configuration
' @param configSource Source de configuration (chemin fichier, objet, etc.)
' @param options Options suppl�mentaires (optionnel)
' @return Gestionnaire de configuration initialis�
Public Function CreateConfigManager(ByVal configSource As Variant, Optional ByVal options As Object = Nothing) As IConfigManagerBase
    On Error GoTo ErrorHandler
    
    ' Cr�er un nouveau gestionnaire de configuration
    Dim configManager As New ConfigManager
    
    ' Ajouter le logger par d�faut aux options si aucun n'est sp�cifi�
    If options Is Nothing And Not m_defaultLogger Is Nothing Then
        Set options = CreateObject("Scripting.Dictionary")
        Set options("Logger") = m_defaultLogger
    End If
    
    ' Initialiser le gestionnaire
    If Not configManager.Initialize(configSource, options) Then
        Err.Raise vbObjectError + 1101, "ConfigFactory.CreateConfigManager", "�chec de l'initialisation du gestionnaire de configuration"
    End If
    
    ' Renvoyer le gestionnaire initialis�
    Set CreateConfigManager = configManager
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la cr�ation du ConfigManager: " & Err.Description, "CreateConfigManager"
    Set CreateConfigManager = Nothing
End Function

' R�cup�re un gestionnaire de configuration du cache ou en cr�e un nouveau
' @param configName Nom de la configuration
' @param configSource Source de configuration si nouvelle cr�ation
' @param options Options suppl�mentaires (optionnel)
' @return Gestionnaire de configuration
Public Function GetConfigManager(ByVal configName As String, Optional ByVal configSource As Variant = Null, Optional ByVal options As Object = Nothing) As IConfigManagerBase
    On Error GoTo ErrorHandler
    
    ' V�rifier si le gestionnaire existe d�j� dans le cache
    If m_configManagers.Exists(configName) Then
        ' Renvoyer le gestionnaire du cache
        Set GetConfigManager = m_configManagers(configName)
    Else
        ' V�rifier que la source est sp�cifi�e pour une nouvelle cr�ation
        If IsNull(configSource) Then
            Err.Raise vbObjectError + 1102, "ConfigFactory.GetConfigManager", "Source de configuration non sp�cifi�e pour une nouvelle cr�ation"
        End If
        
        ' Cr�er un nouveau gestionnaire
        Dim configManager As IConfigManagerBase
        Set configManager = CreateConfigManager(configSource, options)
        
        ' Ajouter au cache
        m_configManagers.Add configName, configManager
        
        ' Renvoyer le nouveau gestionnaire
        Set GetConfigManager = configManager
    End If
    
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la r�cup�ration du ConfigManager '" & configName & "': " & Err.Description, "GetConfigManager"
    Set GetConfigManager = Nothing
End Function

' R�cup�re la configuration par d�faut
' @return Gestionnaire de configuration par d�faut
Public Function GetDefaultConfig() As IConfigManagerBase
    On Error GoTo ErrorHandler
    
    ' V�rifier si la configuration par d�faut existe
    If m_defaultConfig Is Nothing Then
        Err.Raise vbObjectError + 1103, "ConfigFactory.GetDefaultConfig", "Aucune configuration par d�faut d�finie"
    End If
    
    ' Renvoyer la configuration par d�faut
    Set GetDefaultConfig = m_defaultConfig
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la r�cup�ration de la configuration par d�faut: " & Err.Description, "GetDefaultConfig"
    Set GetDefaultConfig = Nothing
End Function

' R�cup�re tous les noms de configurations disponibles
' @return Collection de noms de configurations
Public Function GetConfigNames() As Collection
    On Error GoTo ErrorHandler
    
    ' Cr�er une collection pour les noms
    Dim result As New Collection
    
    ' Ajouter les noms au r�sultat
    Dim key As Variant
    For Each key In m_configManagers.Keys
        result.Add key
    Next
    
    Set GetConfigNames = result
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la r�cup�ration des noms de configurations: " & Err.Description, "GetConfigNames"
    Set GetConfigNames = New Collection
End Function

' Supprime une configuration du cache
' @param configName Nom de la configuration � supprimer
' @return True si la suppression a r�ussi, False sinon
Public Function RemoveConfig(ByVal configName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' V�rifier si la configuration existe
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

' D�finit la configuration par d�faut
' @param configManager Gestionnaire de configuration � d�finir comme d�faut
Public Sub SetDefaultConfig(ByVal configManager As IConfigManagerBase)
    On Error GoTo ErrorHandler
    
    ' Stocker la configuration par d�faut
    Set m_defaultConfig = configManager
    Exit Sub
    
ErrorHandler:
    LogError "Erreur lors de la d�finition de la configuration par d�faut: " & Err.Description, "SetDefaultConfig"
End Sub

' D�finit le logger par d�faut
' @param logger Logger � utiliser pour les nouvelles configurations
Public Sub SetDefaultLogger(ByVal logger As Object)
    On Error GoTo ErrorHandler
    
    ' Stocker le logger par d�faut
    Set m_defaultLogger = logger
    Exit Sub
    
ErrorHandler:
    LogError "Erreur lors de la d�finition du logger par d�faut: " & Err.Description, "SetDefaultLogger"
End Sub

' =========================================================
' M�thodes priv�es
' =========================================================

' Gestion des erreurs
Private Sub LogError(ByVal message As String, ByVal source As String)
    ' Logger l'erreur si un logger est disponible
    If Not m_defaultLogger Is Nothing Then
        On Error Resume Next
        m_defaultLogger.Error message, "ConfigFactory." & source
        On Error GoTo 0
    End If
    
    ' �galement afficher dans Debug.Print pour le d�bogage
    Debug.Print "[ERROR][ConfigFactory." & source & "] " & message
End Sub 