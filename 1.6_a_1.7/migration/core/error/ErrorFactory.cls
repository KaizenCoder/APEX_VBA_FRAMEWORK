VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@ModuleDescription("Factory pour le gestionnaire d'erreurs")
'@Folder("APEX.Core.Error")
'@Version 1.7.0
'@Author APEX Framework Team
Option Explicit

' =========================================================
' ErrorFactory
' Factory pour le gestionnaire d'erreurs
' =========================================================

' Cache du gestionnaire d'erreurs
Private m_defaultErrorHandler As IErrorHandlerBase
Private m_customErrorHandlers As Object  ' Dictionnaire de gestionnaires d'erreurs personnalis�s
Private m_defaultLogger As Object        ' ILoggerBase (par d�faut)
Private m_defaultConfig As Object        ' IConfigManagerBase (par d�faut)
Private m_initialized As Boolean

' =========================================================
' Initialisation
' =========================================================
Private Sub Class_Initialize()
    ' Initialiser le cache
    Set m_customErrorHandlers = CreateObject("Scripting.Dictionary")
    m_initialized = False
End Sub

' =========================================================
' M�thodes publiques
' =========================================================

' Initialise la factory avec des param�tres par d�faut
' @param logger Logger � utiliser (optionnel)
' @param config Configuration � utiliser (optionnel)
' @return True si l'initialisation a r�ussi, False sinon
Public Function Initialize(Optional ByVal logger As Object = Nothing, Optional ByVal config As Object = Nothing) As Boolean
    On Error GoTo ErrorHandler
    
    ' Stocker les d�pendances par d�faut
    Set m_defaultLogger = logger
    Set m_defaultConfig = config
    
    ' Cr�er le gestionnaire d'erreurs par d�faut
    Set m_defaultErrorHandler = CreateErrorHandler(logger, config)
    
    m_initialized = True
    Initialize = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur lors de l'initialisation de ErrorFactory: " & Err.Description
    Initialize = False
End Function

' Cr�e un nouveau gestionnaire d'erreurs
' @param logger Logger � utiliser (optionnel)
' @param config Configuration � utiliser (optionnel)
' @return Gestionnaire d'erreurs initialis�
Public Function CreateErrorHandler(Optional ByVal logger As Object = Nothing, Optional ByVal config As Object = Nothing) As IErrorHandlerBase
    On Error GoTo ErrorHandler
    
    ' Cr�er un nouveau gestionnaire d'erreurs
    Dim errorHandler As New ErrorHandler
    
    ' Pr�parer les options
    Dim options As Object
    Set options = CreateObject("Scripting.Dictionary")
    
    ' Utiliser le logger fourni ou le logger par d�faut
    If Not logger Is Nothing Then
        Set options("Logger") = logger
    ElseIf Not m_defaultLogger Is Nothing Then
        Set options("Logger") = m_defaultLogger
    End If
    
    ' Utiliser la configuration fournie ou la configuration par d�faut
    If Not config Is Nothing Then
        Set options("Config") = config
    ElseIf Not m_defaultConfig Is Nothing Then
        Set options("Config") = m_defaultConfig
    End If
    
    ' Initialiser le gestionnaire d'erreurs
    If Not errorHandler.Initialize(options) Then
        Err.Raise vbObjectError + 5000, "ErrorFactory.CreateErrorHandler", "�chec de l'initialisation du gestionnaire d'erreurs"
    End If
    
    ' Renvoyer le gestionnaire initialis�
    Set CreateErrorHandler = errorHandler
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur lors de la cr�ation du ErrorHandler: " & Err.Description
    Set CreateErrorHandler = Nothing
End Function

' R�cup�re un gestionnaire d'erreurs personnalis� ou en cr�e un nouveau
' @param handlerName Nom du gestionnaire d'erreurs
' @param logger Logger � utiliser si cr�ation (optionnel)
' @param config Configuration � utiliser si cr�ation (optionnel)
' @return Gestionnaire d'erreurs
Public Function GetErrorHandler(ByVal handlerName As String, Optional ByVal logger As Object = Nothing, Optional ByVal config As Object = Nothing) As IErrorHandlerBase
    On Error GoTo ErrorHandler
    
    ' V�rifier si le gestionnaire existe d�j� dans le cache
    If m_customErrorHandlers.Exists(handlerName) Then
        ' Renvoyer le gestionnaire du cache
        Set GetErrorHandler = m_customErrorHandlers(handlerName)
    Else
        ' Cr�er un nouveau gestionnaire
        Dim errorHandler As IErrorHandlerBase
        Set errorHandler = CreateErrorHandler(logger, config)
        
        ' Ajouter au cache
        m_customErrorHandlers.Add handlerName, errorHandler
        
        ' Renvoyer le nouveau gestionnaire
        Set GetErrorHandler = errorHandler
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur lors de la r�cup�ration du ErrorHandler '" & handlerName & "': " & Err.Description
    Set GetErrorHandler = Nothing
End Function

' R�cup�re le gestionnaire d'erreurs par d�faut
' @return Gestionnaire d'erreurs par d�faut
Public Function GetDefaultErrorHandler() As IErrorHandlerBase
    On Error GoTo ErrorHandler
    
    ' V�rifier si le gestionnaire par d�faut existe
    If m_defaultErrorHandler Is Nothing Then
        ' Cr�er un gestionnaire par d�faut si n�cessaire
        Set m_defaultErrorHandler = CreateErrorHandler(m_defaultLogger, m_defaultConfig)
    End If
    
    ' Renvoyer le gestionnaire par d�faut
    Set GetDefaultErrorHandler = m_defaultErrorHandler
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur lors de la r�cup�ration du gestionnaire d'erreurs par d�faut: " & Err.Description
    Set GetDefaultErrorHandler = Nothing
End Function

' D�finit le gestionnaire d'erreurs par d�faut
' @param errorHandler Gestionnaire d'erreurs � d�finir comme d�faut
Public Sub SetDefaultErrorHandler(ByVal errorHandler As IErrorHandlerBase)
    On Error Resume Next
    
    ' Stocker le gestionnaire par d�faut
    Set m_defaultErrorHandler = errorHandler
End Sub

' D�finit le logger par d�faut
' @param logger Logger � utiliser pour les nouveaux gestionnaires d'erreurs
Public Sub SetDefaultLogger(ByVal logger As Object)
    On Error Resume Next
    
    ' Stocker le logger par d�faut
    Set m_defaultLogger = logger
End Sub

' D�finit la configuration par d�faut
' @param config Configuration � utiliser pour les nouveaux gestionnaires d'erreurs
Public Sub SetDefaultConfig(ByVal config As Object)
    On Error Resume Next
    
    ' Stocker la configuration par d�faut
    Set m_defaultConfig = config
End Sub

' Supprime un gestionnaire d'erreurs personnalis� du cache
' @param handlerName Nom du gestionnaire d'erreurs � supprimer
' @return True si la suppression a r�ussi, False sinon
Public Function RemoveErrorHandler(ByVal handlerName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' V�rifier si le gestionnaire existe
    If m_customErrorHandlers.Exists(handlerName) Then
        ' Supprimer le gestionnaire du cache
        m_customErrorHandlers.Remove handlerName
        RemoveErrorHandler = True
    Else
        ' Le gestionnaire n'existe pas
        RemoveErrorHandler = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur lors de la suppression du ErrorHandler '" & handlerName & "': " & Err.Description
    RemoveErrorHandler = False
End Function

' =========================================================
' M�thodes utilitaires
' =========================================================

' G�re une erreur avec le gestionnaire par d�faut
' @param errNumber Num�ro de l'erreur
' @param errDescription Description de l'erreur
' @param errSource Source de l'erreur
' @param errModuleName Module o� l'erreur s'est produite
' @param errProcedureName Proc�dure o� l'erreur s'est produite
' @param errAdditionalInfo Informations suppl�mentaires sur l'erreur
' @param errSeverity Niveau de gravit� de l'erreur
' @return Code de r�sultat du traitement de l'erreur
Public Function HandleError(ByVal errNumber As Long, _
                           ByVal errDescription As String, _
                           Optional ByVal errSource As String = "", _
                           Optional ByVal errModuleName As String = "", _
                           Optional ByVal errProcedureName As String = "", _
                           Optional ByVal errAdditionalInfo As String = "", _
                           Optional ByVal errSeverity As Long = 2) As Long
    On Error Resume Next
    
    ' Obtenir le gestionnaire par d�faut
    Dim handler As IErrorHandlerBase
    Set handler = GetDefaultErrorHandler()
    
    ' Utiliser le gestionnaire pour traiter l'erreur
    If Not handler Is Nothing Then
        HandleError = handler.HandleError(errNumber, errDescription, errSource, errModuleName, _
                                       errProcedureName, 0, errAdditionalInfo, errSeverity)
    Else
        ' Gestionnaire par d�faut non disponible, utiliser Debug.Print
        Debug.Print "[ERROR] " & errNumber & ": " & errDescription & " (" & errSource & ")"
        HandleError = -1
    End If
End Function

' G�re l'erreur en cours avec le gestionnaire par d�faut
' @param errSource Source de l'erreur
' @param errModuleName Module o� l'erreur s'est produite
' @param errProcedureName Proc�dure o� l'erreur s'est produite
' @param errAdditionalInfo Informations suppl�mentaires sur l'erreur
' @param errSeverity Niveau de gravit� de l'erreur
' @return Code de r�sultat du traitement de l'erreur
Public Function HandleCurrentError(Optional ByVal errSource As String = "", _
                                  Optional ByVal errModuleName As String = "", _
                                  Optional ByVal errProcedureName As String = "", _
                                  Optional ByVal errAdditionalInfo As String = "", _
                                  Optional ByVal errSeverity As Long = 2) As Long
    On Error Resume Next
    
    ' Obtenir le gestionnaire par d�faut
    Dim handler As IErrorHandlerBase
    Set handler = GetDefaultErrorHandler()
    
    ' Utiliser le gestionnaire pour traiter l'erreur en cours
    If Not handler Is Nothing Then
        HandleCurrentError = handler.HandleCurrentError(errSource, errModuleName, _
                                                     errProcedureName, errAdditionalInfo, errSeverity)
    Else
        ' Gestionnaire par d�faut non disponible, utiliser Debug.Print
        Debug.Print "[CURRENT ERROR] " & Err.Number & ": " & Err.Description & " (" & Err.Source & ")"
        HandleCurrentError = -1
    End If
End Function

' Cr�e et l�ve une erreur personnalis�e avec le gestionnaire par d�faut
' @param errNumber Num�ro de l'erreur
' @param errDescription Description de l'erreur
' @param errSource Source de l'erreur
' @return Toujours False
Public Function RaiseError(ByVal errNumber As Long, _
                          ByVal errDescription As String, _
                          Optional ByVal errSource As String = "") As Boolean
    On Error Resume Next
    
    ' Obtenir le gestionnaire par d�faut
    Dim handler As IErrorHandlerBase
    Set handler = GetDefaultErrorHandler()
    
    ' Utiliser le gestionnaire pour lever l'erreur
    If Not handler Is Nothing Then
        RaiseError = handler.RaiseError(errNumber, errDescription, errSource)
    Else
        ' Gestionnaire par d�faut non disponible, lever directement
        Err.Raise errNumber, errSource, errDescription
        RaiseError = False
    End If
End Function

' Cr�e et l�ve une erreur typ�e avec le gestionnaire par d�faut
' @param errDescription Description de l'erreur
' @param errSource Source de l'erreur
' @param errType Type d'erreur
' @return Toujours False
Public Function RaiseTypedError(ByVal errDescription As String, _
                              Optional ByVal errSource As String = "", _
                              Optional ByVal errType As Long = 0) As Boolean
    On Error Resume Next
    
    ' Obtenir le gestionnaire par d�faut
    Dim handler As IErrorHandlerBase
    Set handler = GetDefaultErrorHandler()
    
    ' Utiliser le gestionnaire pour lever l'erreur typ�e
    If Not handler Is Nothing Then
        RaiseTypedError = handler.RaiseTypedError(errDescription, errSource, errType)
    Else
        ' Gestionnaire par d�faut non disponible, lever directement
        Dim errNumber As Long
        If errType = 0 Then
            errNumber = vbObjectError + 5500 ' ERROR_TYPE_SYSTEM
        Else
            errNumber = errType
        End If
        Err.Raise errNumber, errSource, errDescription
        RaiseTypedError = False
    End If
End Function 