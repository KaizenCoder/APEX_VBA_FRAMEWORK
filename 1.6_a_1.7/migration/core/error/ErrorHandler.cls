VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@ModuleDescription("Implémentation standard du gestionnaire d'erreurs")
'@Folder("APEX.Core.Error")
'@Version 1.7.0
'@Author APEX Framework Team
Option Explicit

Implements IErrorHandlerBase

' =========================================================
' ErrorHandler
' Implémentation standard du gestionnaire d'erreurs
' =========================================================

' Constantes pour les types d'erreurs (base pour RaiseTypedError)
Public Const ERROR_TYPE_VALIDATION As Long = vbObjectError + 5000
Public Const ERROR_TYPE_BUSINESS As Long = vbObjectError + 5100
Public Const ERROR_TYPE_DATA As Long = vbObjectError + 5200
Public Const ERROR_TYPE_SECURITY As Long = vbObjectError + 5300
Public Const ERROR_TYPE_CONFIGURATION As Long = vbObjectError + 5400
Public Const ERROR_TYPE_SYSTEM As Long = vbObjectError + 5500

' Constantes pour les niveaux de gravité des erreurs
Public Const ERROR_SEVERITY_CRITICAL As Long = 1
Public Const ERROR_SEVERITY_ERROR As Long = 2
Public Const ERROR_SEVERITY_WARNING As Long = 3
Public Const ERROR_SEVERITY_INFO As Long = 4

' Constantes pour les codes de retour du traitement des erreurs
Public Const ERROR_HANDLED_SUCCESS As Long = 0
Public Const ERROR_HANDLED_LOGGED As Long = 1
Public Const ERROR_HANDLED_NOTIFIED As Long = 2
Public Const ERROR_HANDLING_FAILED As Long = -1

' Configuration et état
Private m_logger As Object           ' ILoggerBase
Private m_config As Object           ' IConfigManagerBase
Private m_initialized As Boolean
Private m_rethrowAfterHandling As Boolean
Private m_notificationsEnabled As Object ' Dictionnaire des types d'erreurs avec notifications activées

' Statistiques d'erreurs
Private m_errorCount As Long
Private m_lastErrorCode As Long
Private m_lastErrorDescription As String
Private m_lastErrorSource As String
Private m_errorStats As Object ' Dictionnaire pour les statistiques d'erreurs

' =========================================================
' Initialisation et configuration
' =========================================================
Private Sub Class_Initialize()
    ' Initialiser l'état
    m_initialized = False
    m_rethrowAfterHandling = False
    m_errorCount = 0
    
    ' Initialiser les dictionnaires
    Set m_notificationsEnabled = CreateObject("Scripting.Dictionary")
    Set m_errorStats = CreateObject("Scripting.Dictionary")
    
    ' Activer les notifications par défaut pour les erreurs critiques
    m_notificationsEnabled.Add ERROR_SEVERITY_CRITICAL, True
End Sub

' =========================================================
' Implémentation de IErrorHandlerBase
' =========================================================

Private Function IErrorHandlerBase_Initialize(Optional ByVal options As Object = Nothing) As Boolean
    On Error GoTo ErrorHandler
    
    ' Configurer le logger et la configuration si fournis dans les options
    If Not options Is Nothing Then
        On Error Resume Next
        Set m_logger = options("Logger")
        Set m_config = options("Config")
        On Error GoTo ErrorHandler
    End If
    
    ' Charger la configuration si disponible
    If Not m_config Is Nothing Then
        ' Charger les paramètres depuis la configuration
        m_rethrowAfterHandling = m_config.GetBoolean("error.rethrowAfterHandling", False)
        
        ' Charger les notifications activées depuis la configuration
        Dim notificationSettings As Object
        Set notificationSettings = m_config.GetSection("error.notifications")
        
        If Not notificationSettings Is Nothing Then
            Dim key As Variant
            For Each key In notificationSettings.Keys
                Dim severityOrType As Long
                severityOrType = CLng(key)
                m_notificationsEnabled(severityOrType) = CBool(notificationSettings(key))
            Next
        End If
    End If
    
    m_initialized = True
    IErrorHandlerBase_Initialize = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur lors de l'initialisation du ErrorHandler: " & Err.Description
    IErrorHandlerBase_Initialize = False
End Function

' =========================================================
' Gestion des erreurs
' =========================================================

Private Function IErrorHandlerBase_HandleError(ByVal errNumber As Long, _
                                             ByVal errDescription As String, _
                                             Optional ByVal errSource As String = "", _
                                             Optional ByVal errModuleName As String = "", _
                                             Optional ByVal errProcedureName As String = "", _
                                             Optional ByVal errLineNumber As Long = 0, _
                                             Optional ByVal errAdditionalInfo As String = "", _
                                             Optional ByVal errSeverity As Long = ERROR_SEVERITY_ERROR) As Long
    On Error GoTo ErrorHandler
    
    ' Incrémenter le compteur d'erreurs
    m_errorCount = m_errorCount + 1
    
    ' Stocker les informations de la dernière erreur
    m_lastErrorCode = errNumber
    m_lastErrorDescription = errDescription
    m_lastErrorSource = errSource
    
    ' Mettre à jour les statistiques
    UpdateErrorStats errNumber, errSeverity
    
    ' Construction du message d'erreur complet
    Dim fullErrorMessage As String
    fullErrorMessage = BuildErrorMessage(errNumber, errDescription, errSource, errModuleName, _
                                       errProcedureName, errLineNumber, errAdditionalInfo, errSeverity)
    
    ' Journaliser l'erreur si un logger est disponible
    Dim isLogged As Boolean
    isLogged = LogErrorInternal(errNumber, fullErrorMessage, errSource, errSeverity)
    
    ' Envoyer une notification si activée pour ce type d'erreur
    Dim isNotified As Boolean
    isNotified = False
    
    If IsNotificationEnabledInternal(errSeverity) Then
        isNotified = SendErrorNotification(errNumber, fullErrorMessage, errSeverity)
    End If
    
    ' Déterminer le code de retour en fonction des actions effectuées
    If isLogged And isNotified Then
        IErrorHandlerBase_HandleError = ERROR_HANDLED_NOTIFIED
    ElseIf isLogged Then
        IErrorHandlerBase_HandleError = ERROR_HANDLED_LOGGED
    Else
        IErrorHandlerBase_HandleError = ERROR_HANDLED_SUCCESS
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur lors du traitement de l'erreur: " & Err.Description
    IErrorHandlerBase_HandleError = ERROR_HANDLING_FAILED
End Function

Private Function IErrorHandlerBase_HandleCurrentError(Optional ByVal errSource As String = "", _
                                                   Optional ByVal errModuleName As String = "", _
                                                   Optional ByVal errProcedureName As String = "", _
                                                   Optional ByVal errAdditionalInfo As String = "", _
                                                   Optional ByVal errSeverity As Long = ERROR_SEVERITY_ERROR) As Long
    On Error GoTo ErrorHandler
    
    ' Utiliser HandleError avec les informations de l'erreur en cours
    IErrorHandlerBase_HandleCurrentError = IErrorHandlerBase_HandleError( _
        Err.Number, _
        Err.Description, _
        IIf(Len(errSource) > 0, errSource, Err.Source), _
        errModuleName, _
        errProcedureName, _
        0, ' VBA ne fournit pas de numéro de ligne dans Err
        errAdditionalInfo, _
        errSeverity)
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur lors du traitement de l'erreur en cours: " & Err.Description
    IErrorHandlerBase_HandleCurrentError = ERROR_HANDLING_FAILED
End Function

Private Function IErrorHandlerBase_RaiseError(ByVal errNumber As Long, _
                                           ByVal errDescription As String, _
                                           Optional ByVal errSource As String = "") As Boolean
    On Error Resume Next
    
    ' Enregistrer l'erreur dans les statistiques avant de la lever
    m_errorCount = m_errorCount + 1
    m_lastErrorCode = errNumber
    m_lastErrorDescription = errDescription
    m_lastErrorSource = errSource
    
    ' Lever l'erreur
    Err.Raise errNumber, errSource, errDescription
    
    ' Toujours retourner False pour faciliter l'utilisation dans les gestionnaires d'erreurs
    IErrorHandlerBase_RaiseError = False
End Function

Private Function IErrorHandlerBase_RaiseTypedError(ByVal errDescription As String, _
                                                Optional ByVal errSource As String = "", _
                                                Optional ByVal errType As Long = ERROR_TYPE_SYSTEM) As Boolean
    On Error Resume Next
    
    ' Utiliser RaiseError avec le type d'erreur spécifié
    IErrorHandlerBase_RaiseTypedError = IErrorHandlerBase_RaiseError(errType, errDescription, errSource)
End Function

Private Property Let IErrorHandlerBase_RethrowAfterHandling(ByVal shouldRethrow As Boolean)
    m_rethrowAfterHandling = shouldRethrow
End Property

Private Property Get IErrorHandlerBase_RethrowAfterHandling() As Boolean
    IErrorHandlerBase_RethrowAfterHandling = m_rethrowAfterHandling
End Property

' =========================================================
' Journalisation des erreurs
' =========================================================

Private Function IErrorHandlerBase_LogError(ByVal errNumber As Long, _
                                         ByVal errDescription As String, _
                                         Optional ByVal errSource As String = "", _
                                         Optional ByVal errSeverity As Long = ERROR_SEVERITY_ERROR) As Boolean
    On Error GoTo ErrorHandler
    
    ' Incrémenter le compteur d'erreurs
    m_errorCount = m_errorCount + 1
    
    ' Stocker les informations de la dernière erreur
    m_lastErrorCode = errNumber
    m_lastErrorDescription = errDescription
    m_lastErrorSource = errSource
    
    ' Mettre à jour les statistiques
    UpdateErrorStats errNumber, errSeverity
    
    ' Journaliser l'erreur
    IErrorHandlerBase_LogError = LogErrorInternal(errNumber, errDescription, errSource, errSeverity)
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur lors de la journalisation de l'erreur: " & Err.Description
    IErrorHandlerBase_LogError = False
End Function

Private Function IErrorHandlerBase_LogCurrentError(Optional ByVal errSource As String = "", _
                                                Optional ByVal errSeverity As Long = ERROR_SEVERITY_ERROR) As Boolean
    On Error GoTo ErrorHandler
    
    ' Utiliser LogError avec les informations de l'erreur en cours
    IErrorHandlerBase_LogCurrentError = IErrorHandlerBase_LogError( _
        Err.Number, _
        Err.Description, _
        IIf(Len(errSource) > 0, errSource, Err.Source), _
        errSeverity)
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur lors de la journalisation de l'erreur en cours: " & Err.Description
    IErrorHandlerBase_LogCurrentError = False
End Function

' =========================================================
' Interrogation et statistiques
' =========================================================

Private Property Get IErrorHandlerBase_ErrorCount() As Long
    IErrorHandlerBase_ErrorCount = m_errorCount
End Property

Private Property Get IErrorHandlerBase_LastErrorCode() As Long
    IErrorHandlerBase_LastErrorCode = m_lastErrorCode
End Property

Private Property Get IErrorHandlerBase_LastErrorDescription() As String
    IErrorHandlerBase_LastErrorDescription = m_lastErrorDescription
End Property

Private Property Get IErrorHandlerBase_LastErrorSource() As String
    IErrorHandlerBase_LastErrorSource = m_lastErrorSource
End Property

Private Function IErrorHandlerBase_GetErrorStats() As Object
    ' Retourner les statistiques d'erreurs
    Set IErrorHandlerBase_GetErrorStats = m_errorStats
End Function

' =========================================================
' Notifications et alertes
' =========================================================

Private Sub IErrorHandlerBase_SetNotificationEnabled(ByVal errTypeOrSeverity As Long, ByVal enableNotification As Boolean)
    ' Activer ou désactiver les notifications pour un type d'erreur
    If m_notificationsEnabled.Exists(errTypeOrSeverity) Then
        m_notificationsEnabled(errTypeOrSeverity) = enableNotification
    Else
        m_notificationsEnabled.Add errTypeOrSeverity, enableNotification
    End If
End Sub

Private Function IErrorHandlerBase_IsNotificationEnabled(ByVal errTypeOrSeverity As Long) As Boolean
    ' Vérifier si les notifications sont activées pour un type d'erreur
    IErrorHandlerBase_IsNotificationEnabled = IsNotificationEnabledInternal(errTypeOrSeverity)
End Function

' =========================================================
' Méthodes privées
' =========================================================

' Construction d'un message d'erreur complet et formaté
Private Function BuildErrorMessage(ByVal errNumber As Long, _
                                 ByVal errDescription As String, _
                                 ByVal errSource As String, _
                                 ByVal errModuleName As String, _
                                 ByVal errProcedureName As String, _
                                 ByVal errLineNumber As Long, _
                                 ByVal errAdditionalInfo As String, _
                                 ByVal errSeverity As Long) As String
    On Error Resume Next
    
    Dim severity As String
    Select Case errSeverity
        Case ERROR_SEVERITY_CRITICAL: severity = "CRITIQUE"
        Case ERROR_SEVERITY_ERROR: severity = "ERREUR"
        Case ERROR_SEVERITY_WARNING: severity = "AVERTISSEMENT"
        Case ERROR_SEVERITY_INFO: severity = "INFO"
        Case Else: severity = "INCONNU"
    End Select
    
    Dim message As String
    message = "[" & severity & "] Erreur " & errNumber & ": " & errDescription
    
    ' Ajouter la source si disponible
    If Len(errSource) > 0 Then
        message = message & vbCrLf & "Source: " & errSource
    End If
    
    ' Ajouter les informations de localisation si disponibles
    If Len(errModuleName) > 0 Or Len(errProcedureName) > 0 Then
        message = message & vbCrLf & "Emplacement: "
        
        If Len(errModuleName) > 0 Then
            message = message & errModuleName
            
            If Len(errProcedureName) > 0 Then
                message = message & "." & errProcedureName
            End If
        ElseIf Len(errProcedureName) > 0 Then
            message = message & errProcedureName
        End If
        
        If errLineNumber > 0 Then
            message = message & " (ligne " & errLineNumber & ")"
        End If
    End If
    
    ' Ajouter les informations supplémentaires si disponibles
    If Len(errAdditionalInfo) > 0 Then
        message = message & vbCrLf & "Détails: " & errAdditionalInfo
    End If
    
    BuildErrorMessage = message
End Function

' Journalisation interne de l'erreur
Private Function LogErrorInternal(ByVal errNumber As Long, _
                               ByVal errDescription As String, _
                               ByVal errSource As String, _
                               ByVal errSeverity As Long) As Boolean
    On Error Resume Next
    
    ' Si aucun logger n'est configuré, utiliser Debug.Print
    If m_logger Is Nothing Then
        Debug.Print "[" & Now & "] " & errDescription
        LogErrorInternal = True
        Exit Function
    End If
    
    ' Utiliser le logger approprié selon la gravité
    Dim logResult As Boolean
    
    Select Case errSeverity
        Case ERROR_SEVERITY_CRITICAL
            logResult = m_logger.Critical(errDescription, errSource)
        Case ERROR_SEVERITY_ERROR
            logResult = m_logger.Error(errDescription, errSource)
        Case ERROR_SEVERITY_WARNING
            logResult = m_logger.Warning(errDescription, errSource)
        Case ERROR_SEVERITY_INFO
            logResult = m_logger.Info(errDescription, errSource)
        Case Else
            logResult = m_logger.Debug(errDescription, errSource)
    End Select
    
    LogErrorInternal = logResult
End Function

' Envoi de notification d'erreur
Private Function SendErrorNotification(ByVal errNumber As Long, _
                                     ByVal errMessage As String, _
                                     ByVal errSeverity As Long) As Boolean
    On Error Resume Next
    
    ' Version simple affichant une boîte de message
    ' Dans une implémentation réelle, pourrait envoyer un email, une notification système, etc.
    Dim icon As Long
    
    Select Case errSeverity
        Case ERROR_SEVERITY_CRITICAL
            icon = vbCritical
        Case ERROR_SEVERITY_ERROR
            icon = vbCritical
        Case ERROR_SEVERITY_WARNING
            icon = vbExclamation
        Case Else
            icon = vbInformation
    End Select
    
    ' Vérifier si nous sommes dans un contexte d'interface utilisateur
    ' Si ce n'est pas le cas (par exemple, en mode batch), ne pas afficher de MsgBox
    Dim isUIContext As Boolean
    isUIContext = (Application.Name = "Microsoft Excel")
    
    If isUIContext Then
        MsgBox errMessage, icon, "APEX Framework - Erreur"
    Else
        ' Alternative pour les contextes non-UI
        Debug.Print "[NOTIFICATION] " & errMessage
    End If
    
    SendErrorNotification = True
End Function

' Vérification interne des notifications
Private Function IsNotificationEnabledInternal(ByVal errTypeOrSeverity As Long) As Boolean
    On Error Resume Next
    
    ' Vérifier si le type d'erreur a une configuration spécifique
    If m_notificationsEnabled.Exists(errTypeOrSeverity) Then
        IsNotificationEnabledInternal = m_notificationsEnabled(errTypeOrSeverity)
    Else
        ' Par défaut, activer les notifications uniquement pour les erreurs critiques
        IsNotificationEnabledInternal = (errTypeOrSeverity = ERROR_SEVERITY_CRITICAL)
    End If
End Function

' Mise à jour des statistiques d'erreurs
Private Sub UpdateErrorStats(ByVal errNumber As Long, ByVal errSeverity As Long)
    On Error Resume Next
    
    ' Mettre à jour le compteur pour ce code d'erreur
    Dim errorKey As String
    errorKey = CStr(errNumber)
    
    If m_errorStats.Exists(errorKey) Then
        m_errorStats(errorKey) = CLng(m_errorStats(errorKey)) + 1
    Else
        m_errorStats.Add errorKey, 1
    End If
    
    ' Mettre à jour le compteur pour ce niveau de gravité
    Dim severityKey As String
    severityKey = "SEVERITY_" & CStr(errSeverity)
    
    If m_errorStats.Exists(severityKey) Then
        m_errorStats(severityKey) = CLng(m_errorStats(severityKey)) + 1
    Else
        m_errorStats.Add severityKey, 1
    End If
End Sub

' Fonction utilitaire pour choisir entre deux valeurs
Private Function IIf(condition As Boolean, trueValue As Variant, falseValue As Variant) As Variant
    If condition Then
        IIf = trueValue
    Else
        IIf = falseValue
    End If
End Function 