Attribute VB_Name = "modProcedureUtils"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

'@Module: modProcedureUtils
'@Folder: Apex.Core.Utils
'@Description: Module contenant des procédures utilitaires pour le framework APEX
'@Version: 1.0
'@Date: 12/04/2025
'@Author: APEX Framework Team
'@Dependencies: modConfigManager, clsLogger

Option Explicit

'@Description: Vérifie et prépare un fichier pour le traitement de données
'@Param filePath Le chemin du fichier à traiter
'@Param logLevel Le niveau de log à utiliser (optionnel, par défaut INFO)
'@Param configSection La section de configuration à utiliser (optionnel)
'@Returns Boolean True si le fichier est prêt à être traité, False sinon
Public Function PrepareFileForProcessing(ByVal filePath As String, _
                                       Optional ByVal logLevel As String = "INFO", _
                                       Optional ByVal configSection As String = "") As Boolean
    
    Dim logger As Object
    Set logger = CreateLoggerInstance(logLevel)
    
    ' Journalisation du début de la procédure
    logger.Log "Début de la préparation du fichier: " & filePath, "INFO"
    
    On Error GoTo ErrorHandler
    
    ' Vérifier si le fichier existe
    If Not FileExists(filePath) Then
        logger.Log "ERREUR: Le fichier n'existe pas: " & filePath, "ERROR"
        PrepareFileForProcessing = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Charger la configuration si nécessaire
    Dim config As Variant
    If configSection <> "" Then
        config = modConfigManager.GetConfig(configSection)
        logger.Log "Configuration chargée depuis la section: " & configSection, "DEBUG"
    End If
    
    ' Vérifier les droits d'accès au fichier
    If Not CheckFileAccess(filePath) Then
        logger.Log "ERREUR: Permissions insuffisantes pour le fichier: " & filePath, "ERROR"
        PrepareFileForProcessing = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Vérifier si le fichier est déjà ouvert
    If IsFileOpen(filePath) Then
        logger.Log "AVERTISSEMENT: Le fichier est déjà ouvert: " & filePath, "WARN"
        ' Possibilité d'ajouter une logique pour gérer les fichiers ouverts
    End If
    
    ' Créer une sauvegarde si nécessaire
    If config("createBackup") = True Then
        Dim backupPath As String
        backupPath = CreateBackup(filePath)
        logger.Log "Sauvegarde créée: " & backupPath, "INFO"
    End If
    
    ' Marquer le traitement comme réussi
    PrepareFileForProcessing = True
    logger.Log "Fichier prêt pour le traitement: " & filePath, "INFO"
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    logger.Log "ERREUR lors de la préparation du fichier: " & Err.Description & " (" & Err.Number & ")", "ERROR"
    PrepareFileForProcessing = False
End Function

'@Description: Crée une instance de logger appropriée selon le niveau de log
'@Param logLevel Le niveau de log à utiliser
'@Returns Object L'instance de logger
Private Function CreateLoggerInstance(ByVal logLevel As String) As Object
    Dim logger As Object
    
    ' Utiliser le Factory Pattern pour créer l'instance de logger appropriée
    Select Case UCase(logLevel)
        Case "DEBUG"
            Set logger = New clsDebugLogger
        Case "SHEET"
            Set logger = New clsSheetLogger
        Case Else
            Set logger = New clsLogger
    End Select
    
    Set CreateLoggerInstance = logger
End Function

'@Description: Vérifie si un fichier existe
'@Param filePath Le chemin du fichier à vérifier
'@Returns Boolean True si le fichier existe, False sinon
Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

'@Description: Vérifie si l'utilisateur a les droits d'accès au fichier
'@Param filePath Le chemin du fichier à vérifier
'@Returns Boolean True si l'accès est autorisé, False sinon
Private Function CheckFileAccess(ByVal filePath As String) As Boolean
    On Error Resume Next
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Binary Access Read Write Lock Read Write As #fileNum
    CheckFileAccess = (Err.Number = 0)
    Close #fileNum
    
    On Error GoTo 0
End Function

'@Description: Vérifie si un fichier est déjà ouvert
'@Param filePath Le chemin du fichier à vérifier
'@Returns Boolean True si le fichier est ouvert, False sinon
Private Function IsFileOpen(ByVal filePath As String) As Boolean
    On Error Resume Next
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Input Lock Read As #fileNum
    Close #fileNum
    
    IsFileOpen = (Err.Number <> 0)
    On Error GoTo 0
End Function

'@Description: Crée une sauvegarde du fichier
'@Param filePath Le chemin du fichier à sauvegarder
'@Returns String Le chemin de la sauvegarde
Private Function CreateBackup(ByVal filePath As String) As String
    Dim backupPath As String
    Dim fileName As String
    Dim fileExt As String
    Dim dateStamp As String
    
    ' Extraire le nom et l'extension du fichier
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
    fileExt = Mid(fileName, InStrRev(fileName, "."))
    fileName = Left(fileName, InStrRev(fileName, ".") - 1)
    
    ' Créer un horodatage
    dateStamp = Format(Now, "yyyymmdd_hhnnss")
    
    ' Construire le chemin de sauvegarde
    backupPath = Left(filePath, InStrRev(filePath, "\")) & fileName & "_backup_" & dateStamp & fileExt
    
    ' Copier le fichier
    FileCopy filePath, backupPath
    
    CreateBackup = backupPath
End Function