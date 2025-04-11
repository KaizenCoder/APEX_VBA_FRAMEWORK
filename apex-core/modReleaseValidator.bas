' Migrated to apex-core - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Attribute VB_Name = "modReleaseValidator"
Option Explicit
' ==========================================================================
' Module : modReleaseValidator
' Version : 1.0
' Purpose : Validation des versions de déploiement du framework APEX
' Date    : 10/04/2025
' ==========================================================================

' --- Constantes ---
Private Const MODULE_NAME As String = "modReleaseValidator"

Private Const ESSENTIAL_MODULES() As String = Array( _
    "src\core\clsLogger.cls", _
    "src\core\modConfigManager.bas", _
    "src\core\modVersionInfo.bas", _
    "src\utils\modFileUtils.bas", _
    "src\utils\modTextUtils.bas", _
    "src\utils\modDateUtils.bas" _
)

Private Const ESSENTIAL_CONFIG_FILES() As String = Array( _
    "config\logger_config.ini", _
    "config\test_config.ini" _
)

Private Const MIN_MODULE_COUNT As Long = 10
Private Const VALIDATION_LOG_FILE As String = "logs\validation.log"

' --- Types ---
Private Type ValidationResult
    Success As Boolean
    Message As String
    Details As String
    ErrorCount As Long
    WarningCount As Long
    MissingFiles() As String
    MissingFilesCount As Long
End Type

' --- Variables ---
Private m_logger As Object
Private m_result As ValidationResult
Private m_currentDir As String
Private m_fso As Object

' --- Initialisation ---
Private Sub Initialize()
    ' Créer le gestionnaire de fichiers
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    
    ' Réinitialiser le résultat
    m_result.Success = True
    m_result.Message = ""
    m_result.Details = ""
    m_result.ErrorCount = 0
    m_result.WarningCount = 0
    ReDim m_result.MissingFiles(0)
    m_result.MissingFilesCount = 0
    
    ' Récupérer le répertoire courant
    m_currentDir = CurDir()
    If Right(m_currentDir, 1) <> "\" Then m_currentDir = m_currentDir & "\"
    
    ' Essayer de récupérer le logger
    On Error Resume Next
    Set m_logger = CreateObject("APEX.Logger")
    If Err.Number <> 0 Then Set m_logger = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

' --- Validation publique ---
Public Function ValidateRelease(Optional ByVal releasePath As String = "") As Boolean
    ' Initialiser
    Initialize
    
    ' Si aucun chemin n'est spécifié, utiliser le répertoire courant
    If releasePath = "" Then releasePath = m_currentDir
    
    ' S'assurer que le chemin se termine par "\"
    If Right(releasePath, 1) <> "\" Then releasePath = releasePath & "\"
    
    LogMessage "Début de la validation de la release dans '" & releasePath & "'", "info"
    
    ' Vérifier l'existence du répertoire
    If Not m_fso.FolderExists(releasePath) Then
        m_result.Success = False
        m_result.Message = "Le répertoire de release n'existe pas"
        m_result.ErrorCount = 1
        
        LogMessage "ERREUR: " & m_result.Message, "error"
        ValidateRelease = False
        Exit Function
    End If
    
    ' Effectuer les validations
    Dim isValid As Boolean
    
    isValid = True
    isValid = ValidateEssentialModules(releasePath) And isValid
    isValid = ValidateConfigFiles(releasePath) And isValid
    isValid = ValidateStructure(releasePath) And isValid
    isValid = ValidateVersionFile(releasePath) And isValid
    
    ' Définir le message récapitulatif
    If isValid Then
        m_result.Message = "La validation a réussi"
        LogMessage "La validation a réussi", "info"
    Else
        m_result.Message = "La validation a échoué avec " & m_result.ErrorCount & " erreur(s) et " & m_result.WarningCount & " avertissement(s)"
        LogMessage "La validation a échoué avec " & m_result.ErrorCount & " erreur(s) et " & m_result.WarningCount & " avertissement(s)", "error"
    End If
    
    ' Générer les détails
    m_result.Details = GenerateValidationReport()
    
    ' Écrire le rapport dans un fichier log
    WriteValidationReport
    
    ValidateRelease = isValid
End Function

Public Function GetValidationMessage() As String
    GetValidationMessage = m_result.Message
End Function

Public Function GetValidationDetails() As String
    GetValidationDetails = m_result.Details
End Function

Public Function GetMissingFiles() As Variant
    ' Convertir le tableau en variant pour le retour
    If m_result.MissingFilesCount = 0 Then
        GetMissingFiles = Array()
    Else
        Dim result() As String
        ReDim result(m_result.MissingFilesCount - 1)
        
        Dim i As Long
        For i = 0 To m_result.MissingFilesCount - 1
            result(i) = m_result.MissingFiles(i)
        Next i
        
        GetMissingFiles = result
    End If
End Function

Public Function GetErrorCount() As Long
    GetErrorCount = m_result.ErrorCount
End Function

Public Function GetWarningCount() As Long
    GetWarningCount = m_result.WarningCount
End Function

Public Sub ShowValidationReport()
    Dim title As String
    Dim message As String
    Dim style As VbMsgBoxStyle
    
    title = "Validation de la Release APEX"
    message = m_result.Message & vbCrLf & vbCrLf
    
    If m_result.ErrorCount > 0 Or m_result.WarningCount > 0 Then
        message = message & "Problèmes détectés:" & vbCrLf
        
        If m_result.MissingFilesCount > 0 Then
            message = message & "- " & m_result.MissingFilesCount & " fichier(s) requis manquant(s)" & vbCrLf
            
            ' Limiter le nombre de fichiers affichés
            Dim maxFiles As Long
            maxFiles = 5 ' Maximum de fichiers à afficher
            
            message = message & vbCrLf & "Fichiers manquants:" & vbCrLf
            
            Dim i As Long
            For i = 0 To IIf(m_result.MissingFilesCount > maxFiles, maxFiles - 1, m_result.MissingFilesCount - 1)
                message = message & "  " & m_result.MissingFiles(i) & vbCrLf
            Next i
            
            If m_result.MissingFilesCount > maxFiles Then
                message = message & "  ... et " & (m_result.MissingFilesCount - maxFiles) & " autres fichiers" & vbCrLf
            End If
        End If
        
        message = message & vbCrLf & "Consultez le fichier '" & VALIDATION_LOG_FILE & "' pour plus de détails."
        style = IIf(m_result.ErrorCount > 0, vbCritical, vbExclamation)
    Else
        message = message & "Tous les contrôles ont été validés avec succès."
        style = vbInformation
    End If
    
    MsgBox message, style, title
End Sub

' --- Validations spécifiques ---
Private Function ValidateEssentialModules(ByVal basePath As String) As Boolean
    LogMessage "Vérification des modules essentiels...", "info"
    
    Dim i As Long
    Dim missingCount As Long
    missingCount = 0
    ValidateEssentialModules = True
    
    ' Vérifier chaque module essentiel
    For i = LBound(ESSENTIAL_MODULES) To UBound(ESSENTIAL_MODULES)
        Dim modulePath As String
        modulePath = basePath & ESSENTIAL_MODULES(i)
        
        ' Remplacer les barres obliques si nécessaire
        modulePath = Replace(modulePath, "/", "\")
        
        If Not m_fso.FileExists(modulePath) Then
            ' Ajouter à la liste des fichiers manquants
            m_result.MissingFilesCount = m_result.MissingFilesCount + 1
            ReDim Preserve m_result.MissingFiles(m_result.MissingFilesCount)
            m_result.MissingFiles(m_result.MissingFilesCount - 1) = ESSENTIAL_MODULES(i)
            
            LogMessage "ERREUR: Module essentiel manquant - " & ESSENTIAL_MODULES(i), "error"
            missingCount = missingCount + 1
            ValidateEssentialModules = False
        End If
    Next i
    
    If missingCount > 0 Then
        m_result.ErrorCount = m_result.ErrorCount + missingCount
    Else
        LogMessage "Tous les modules essentiels sont présents", "info"
    End If
End Function

Private Function ValidateConfigFiles(ByVal basePath As String) As Boolean
    LogMessage "Vérification des fichiers de configuration...", "info"
    
    Dim i As Long
    Dim missingCount As Long
    missingCount = 0
    ValidateConfigFiles = True
    
    ' Vérifier chaque fichier de configuration essentiel
    For i = LBound(ESSENTIAL_CONFIG_FILES) To UBound(ESSENTIAL_CONFIG_FILES)
        Dim configPath As String
        configPath = basePath & ESSENTIAL_CONFIG_FILES(i)
        
        ' Remplacer les barres obliques si nécessaire
        configPath = Replace(configPath, "/", "\")
        
        If Not m_fso.FileExists(configPath) Then
            ' Ajouter à la liste des fichiers manquants
            m_result.MissingFilesCount = m_result.MissingFilesCount + 1
            ReDim Preserve m_result.MissingFiles(m_result.MissingFilesCount)
            m_result.MissingFiles(m_result.MissingFilesCount - 1) = ESSENTIAL_CONFIG_FILES(i)
            
            LogMessage "ERREUR: Fichier de configuration manquant - " & ESSENTIAL_CONFIG_FILES(i), "error"
            missingCount = missingCount + 1
            ValidateConfigFiles = False
        End If
    Next i
    
    If missingCount > 0 Then
        m_result.ErrorCount = m_result.ErrorCount + missingCount
    Else
        LogMessage "Tous les fichiers de configuration sont présents", "info"
    End If
End Function

Private Function ValidateStructure(ByVal basePath As String) As Boolean
    LogMessage "Vérification de la structure du framework...", "info"
    
    ValidateStructure = True
    
    ' Vérifier les répertoires principaux
    Dim mainDirs() As String
    mainDirs = Array("src", "config", "docs")
    
    Dim i As Long
    For i = LBound(mainDirs) To UBound(mainDirs)
        Dim dirPath As String
        dirPath = basePath & mainDirs(i)
        
        If Not m_fso.FolderExists(dirPath) Then
            LogMessage "ERREUR: Répertoire principal manquant - " & mainDirs(i), "error"
            m_result.ErrorCount = m_result.ErrorCount + 1
            ValidateStructure = False
        End If
    Next i
    
    ' Vérifier les sous-répertoires src
    If m_fso.FolderExists(basePath & "src") Then
        Dim srcDirs() As String
        srcDirs = Array("core", "utils", "architecture", "xml", "recette")
        
        For i = LBound(srcDirs) To UBound(srcDirs)
            Dim srcSubDir As String
            srcSubDir = basePath & "src\" & srcDirs(i)
            
            If Not m_fso.FolderExists(srcSubDir) Then
                LogMessage "AVERTISSEMENT: Sous-répertoire src manquant - " & srcDirs(i), "warning"
                m_result.WarningCount = m_result.WarningCount + 1
            End If
        Next i
    End If
    
    ' Vérifier le nombre minimal de modules
    Dim moduleCount As Long
    moduleCount = CountFilesWithExtension(basePath & "src", "bas") + CountFilesWithExtension(basePath & "src", "cls")
    
    If moduleCount < MIN_MODULE_COUNT Then
        LogMessage "ERREUR: Nombre insuffisant de modules (" & moduleCount & " trouvés, minimum " & MIN_MODULE_COUNT & " requis)", "error"
        m_result.ErrorCount = m_result.ErrorCount + 1
        ValidateStructure = False
    Else
        LogMessage "Nombre de modules suffisant: " & moduleCount, "info"
    End If
    
    If ValidateStructure Then
        LogMessage "La structure du framework est valide", "info"
    End If
End Function

Private Function ValidateVersionFile(ByVal basePath As String) As Boolean
    LogMessage "Vérification du fichier de version...", "info"
    
    Dim versionFile As String
    versionFile = basePath & "VERSION.txt"
    
    ' Vérifier si le fichier existe
    If Not m_fso.FileExists(versionFile) Then
        LogMessage "ERREUR: Fichier VERSION.txt manquant", "error"
        m_result.ErrorCount = m_result.ErrorCount + 1
        
        ' Ajouter à la liste des fichiers manquants
        m_result.MissingFilesCount = m_result.MissingFilesCount + 1
        ReDim Preserve m_result.MissingFiles(m_result.MissingFilesCount)
        m_result.MissingFiles(m_result.MissingFilesCount - 1) = "VERSION.txt"
        
        ValidateVersionFile = False
        Exit Function
    End If
    
    ' Lire le fichier
    Dim ts As Object
    Dim content As String
    Dim lines() As String
    Dim versionFound As Boolean
    Dim dateFound As Boolean
    Dim versionValue As String
    
    Set ts = m_fso.OpenTextFile(versionFile, 1) ' ForReading = 1
    content = ts.ReadAll
    ts.Close
    
    lines = Split(content, vbCrLf)
    versionFound = False
    dateFound = False
    
    ' Analyser le contenu
    Dim i As Long
    For i = 0 To UBound(lines)
        If Left(Trim(lines(i)), 8) = "Version:" Then
            versionFound = True
            versionValue = Trim(Mid(lines(i), 9))
        ElseIf Left(Trim(lines(i)), 16) = "Date de création:" Then
            dateFound = True
        End If
    Next i
    
    ' Vérifier si les informations essentielles sont présentes
    Dim isValid As Boolean
    isValid = True
    
    If Not versionFound Then
        LogMessage "ERREUR: Information de version manquante dans VERSION.txt", "error"
        m_result.ErrorCount = m_result.ErrorCount + 1
        isValid = False
    End If
    
    If Not dateFound Then
        LogMessage "AVERTISSEMENT: Date de création manquante dans VERSION.txt", "warning"
        m_result.WarningCount = m_result.WarningCount + 1
    End If
    
    ' Vérifier la cohérence de la version
    If versionFound And Len(versionValue) > 0 Then
        ' Vérifier si la version est au format correct (ex: 1.0.0)
        Dim regex As Object
        Set regex = CreateObject("VBScript.RegExp")
        
        With regex
            .Global = False
            .MultiLine = False
            .IgnoreCase = True
            .Pattern = "^\d+\.\d+\.\d+(?:\-[a-zA-Z0-9\.]+)?$"
        End With
        
        If Not regex.Test(versionValue) Then
            LogMessage "ERREUR: Format de version invalide dans VERSION.txt: " & versionValue, "error"
            m_result.ErrorCount = m_result.ErrorCount + 1
            isValid = False
        Else
            LogMessage "Version valide trouvée: " & versionValue, "info"
        End If
        
        ' Vérifier la cohérence avec la version du module
        If modVersionInfo.FRAMEWORK_VERSION <> versionValue Then
            LogMessage "AVERTISSEMENT: Incohérence de version - VERSION.txt: " & versionValue & ", modVersionInfo: " & modVersionInfo.FRAMEWORK_VERSION, "warning"
            m_result.WarningCount = m_result.WarningCount + 1
        End If
    End If
    
    ValidateVersionFile = isValid
    
    If isValid Then
        LogMessage "La validation du fichier de version est réussie", "info"
    End If
End Function

' --- Fonctions utilitaires ---
Private Function CountFilesWithExtension(ByVal folderPath As String, ByVal extension As String) As Long
    Dim count As Long
    Dim folder As Object
    Dim file As Object
    Dim subfolder As Object
    
    count = 0
    
    ' S'assurer que l'extension ne contient pas de point
    If Left(extension, 1) = "." Then extension = Mid(extension, 2)
    
    ' Vérifier si le dossier existe
    If Not m_fso.FolderExists(folderPath) Then
        CountFilesWithExtension = 0
        Exit Function
    End If
    
    Set folder = m_fso.GetFolder(folderPath)
    
    ' Compter les fichiers dans ce dossier
    For Each file In folder.Files
        If LCase(m_fso.GetExtensionName(file.Name)) = LCase(extension) Then
            count = count + 1
        End If
    Next file
    
    ' Compter les fichiers dans les sous-dossiers
    For Each subfolder In folder.SubFolders
        count = count + CountFilesWithExtension(subfolder.Path, extension)
    Next subfolder
    
    CountFilesWithExtension = count
End Function

Private Sub LogMessage(ByVal message As String, ByVal logLevel As String)
    ' Ajouter le message aux détails de validation
    m_result.Details = m_result.Details & Now & " - " & message & vbCrLf
    
    ' Écrire dans le journal si disponible
    On Error Resume Next
    If Not m_logger Is Nothing Then
        Select Case LCase(logLevel)
            Case "debug"
                m_logger.LogDebug MODULE_NAME, message
            Case "info"
                m_logger.LogInfo MODULE_NAME, message
            Case "warning"
                m_logger.LogWarning MODULE_NAME, message
            Case "error"
                m_logger.LogError MODULE_NAME, message
            Case Else
                m_logger.LogInfo MODULE_NAME, message
        End Select
    Else
        ' Écrire dans la fenêtre de débogage
        Debug.Print Now & " - " & logLevel & " - " & message
    End If
    On Error GoTo 0
End Sub

Private Function GenerateValidationReport() As String
    Dim report As String
    
    report = "RAPPORT DE VALIDATION DU FRAMEWORK APEX" & vbCrLf
    report = report & "----------------------------------------" & vbCrLf
    report = report & "Date: " & Format(Now, "yyyy-mm-dd hh:mm:ss") & vbCrLf
    report = report & "Version du framework: " & modVersionInfo.FRAMEWORK_VERSION & vbCrLf & vbCrLf
    
    report = report & "RÉSULTAT: " & IIf(m_result.Success, "SUCCÈS", "ÉCHEC") & vbCrLf
    report = report & "- Erreurs: " & m_result.ErrorCount & vbCrLf
    report = report & "- Avertissements: " & m_result.WarningCount & vbCrLf
    
    ' Ajouter la liste des fichiers manquants
    If m_result.MissingFilesCount > 0 Then
        report = report & vbCrLf & "FICHIERS MANQUANTS:" & vbCrLf
        
        Dim i As Long
        For i = 0 To m_result.MissingFilesCount - 1
            report = report & "- " & m_result.MissingFiles(i) & vbCrLf
        Next i
    End If
    
    report = report & vbCrLf & "DÉTAILS:" & vbCrLf
    report = report & m_result.Details
    
    GenerateValidationReport = report
End Function

Private Sub WriteValidationReport()
    On Error Resume Next
    
    ' Créer le dossier logs si nécessaire
    Dim logsFolder As String
    logsFolder = m_currentDir & "logs"
    
    If Not m_fso.FolderExists(logsFolder) Then
        m_fso.CreateFolder logsFolder
    End If
    
    ' Écrire le rapport
    Dim ts As Object
    Set ts = m_fso.CreateTextFile(m_currentDir & VALIDATION_LOG_FILE, True)
    
    If Err.Number = 0 Then
        ts.Write m_result.Details
        ts.Close
    End If
    
    On Error GoTo 0
End Sub
