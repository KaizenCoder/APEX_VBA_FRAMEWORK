' Migrated to apex-core/utils - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit
' ==========================================================================
' Module : modFileUtils
' Version : 1.0
' Purpose : Utilitaires pour la gestion des fichiers et répertoires
' Date : 10/04/2025
' ==========================================================================

' --- API Windows pour la gestion des fichiers ---
#If VBA7 Then
    Private Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
        (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    
    Private Declare PtrSafe Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" _
        (ByVal lpszPath As String, ByVal lpPrefixString As String, _
        ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
    
    Private Declare PtrSafe Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" _
        (ByVal lpFileName As String) As Long
        
    Private Declare PtrSafe Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
        (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
#Else
    Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
        (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    
    Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" _
        (ByVal lpszPath As String, ByVal lpPrefixString As String, _
        ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
    
    Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" _
        (ByVal lpFileName As String) As Long
        
    Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
        (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
#End If

' --- Constantes pour les attributs de fichier ---
Private Const FILE_ATTRIBUTE_READONLY As Long = &H1
Private Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Private Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY As Long = &H100
Private Const INVALID_FILE_ATTRIBUTES As Long = -1

' --- Variables privées ---
Private m_Logger As Object ' ILoggerBase
Private m_LastError As String

' --- Initialisation ---
Public Sub Initialize(Optional ByVal logger As Object = Nothing)
    ' Initialise le module avec un logger optionnel
    Set m_Logger = logger
    m_LastError = ""
    
    If Not m_Logger Is Nothing Then
        ' TODO: Logger l'initialisation
        ' m_Logger.LogInfo "modFileUtils initialisé", "FILE"
    End If
End Sub

' --- Fonctions publiques pour la vérification d'existence ---
Public Function FileExists(ByVal filePath As String) As Boolean
    ' Vérifie si un fichier existe
    On Error GoTo ErrorHandler
    
    ' Utilisation de Dir pour vérifier l'existence du fichier
    FileExists = (Dir(filePath, vbNormal) <> "")
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la vérification de l'existence du fichier: " & Err.Description
    LogError m_LastError
    FileExists = False
End Function

Public Function DirectoryExists(ByVal folderPath As String) As Boolean
    ' Vérifie si un répertoire existe
    On Error GoTo ErrorHandler
    
    ' Vérifier si le chemin se termine par un séparateur
    If Right(folderPath, 1) = "\" Then
        folderPath = Left(folderPath, Len(folderPath) - 1)
    End If
    
    ' Utilisation de Dir pour vérifier l'existence du répertoire
    DirectoryExists = (Dir(folderPath, vbDirectory) <> "")
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la vérification de l'existence du répertoire: " & Err.Description
    LogError m_LastError
    DirectoryExists = False
End Function

Public Function IsDirectory(ByVal path As String) As Boolean
    ' Vérifie si le chemin est un répertoire
    Dim attributes As Long
    
    On Error GoTo ErrorHandler
    
    ' Obtenir les attributs du fichier
    attributes = GetFileAttributes(path)
    
    ' Vérifier si le chemin est un répertoire
    If attributes <> INVALID_FILE_ATTRIBUTES Then
        IsDirectory = ((attributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
    Else
        IsDirectory = False
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la vérification si le chemin est un répertoire: " & Err.Description
    LogError m_LastError
    IsDirectory = False
End Function

Public Function IsReadOnly(ByVal filePath As String) As Boolean
    ' Vérifie si un fichier est en lecture seule
    Dim attributes As Long
    
    On Error GoTo ErrorHandler
    
    ' Obtenir les attributs du fichier
    attributes = GetFileAttributes(filePath)
    
    ' Vérifier si le fichier est en lecture seule
    If attributes <> INVALID_FILE_ATTRIBUTES Then
        IsReadOnly = ((attributes And FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY)
    Else
        IsReadOnly = False
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la vérification si le fichier est en lecture seule: " & Err.Description
    LogError m_LastError
    IsReadOnly = False
End Function

' --- Fonctions publiques pour la création ---
Public Function CreateDirectory(ByVal folderPath As String) As Boolean
    ' Crée un répertoire (et les répertoires parents si nécessaire)
    Dim parts() As String
    Dim currentPath As String
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' Si le répertoire existe déjà, retourner vrai
    If DirectoryExists(folderPath) Then
        CreateDirectory = True
        Exit Function
    End If
    
    ' Diviser le chemin en ses composants
    parts = Split(folderPath, "\")
    
    ' Construire le chemin progressivement et créer chaque répertoire
    currentPath = parts(0) & "\"
    
    For i = 1 To UBound(parts)
        currentPath = currentPath & parts(i) & "\"
        
        ' Vérifier si le répertoire existe, sinon le créer
        If Not DirectoryExists(currentPath) Then
            MkDir currentPath
        End If
    Next i
    
    CreateDirectory = True
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la création du répertoire: " & Err.Description
    LogError m_LastError
    CreateDirectory = False
End Function

Public Function CreateTemporaryFile(Optional ByVal prefix As String = "APX") As String
    ' Crée un fichier temporaire et retourne son chemin
    Dim tempPath As String
    Dim tempFile As String
    Dim result As Long
    
    On Error GoTo ErrorHandler
    
    ' Obtenir le chemin du répertoire temporaire
    tempPath = Space$(MAX_PATH)
    result = GetTempPath(Len(tempPath), tempPath)
    
    If result = 0 Then
        m_LastError = "Impossible d'obtenir le chemin du répertoire temporaire"
        LogError m_LastError
        CreateTemporaryFile = ""
        Exit Function
    End If
    
    tempPath = Left$(tempPath, result)
    
    ' Obtenir un nom de fichier temporaire
    tempFile = Space$(MAX_PATH)
    result = GetTempFileName(tempPath, prefix, 0, tempFile)
    
    If result = 0 Then
        m_LastError = "Impossible de créer un fichier temporaire"
        LogError m_LastError
        CreateTemporaryFile = ""
        Exit Function
    End If
    
    CreateTemporaryFile = Left$(tempFile, InStr(tempFile, vbNullChar) - 1)
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la création du fichier temporaire: " & Err.Description
    LogError m_LastError
    CreateTemporaryFile = ""
End Function

' --- Fonctions publiques pour la copie et le déplacement ---
Public Function CopyFile(ByVal sourcePath As String, ByVal destPath As String, _
                         Optional ByVal overwrite As Boolean = False) As Boolean
    ' Copie un fichier
    On Error GoTo ErrorHandler
    
    ' Vérifier si le fichier source existe
    If Not FileExists(sourcePath) Then
        m_LastError = "Le fichier source n'existe pas: " & sourcePath
        LogError m_LastError
        CopyFile = False
        Exit Function
    End If
    
    ' Vérifier si le fichier de destination existe déjà
    If FileExists(destPath) And Not overwrite Then
        m_LastError = "Le fichier de destination existe déjà: " & destPath
        LogError m_LastError
        CopyFile = False
        Exit Function
    End If
    
    ' Créer le répertoire de destination si nécessaire
    Dim destFolder As String
    destFolder = Left$(destPath, InStrRev(destPath, "\") - 1)
    
    If Not DirectoryExists(destFolder) Then
        If Not CreateDirectory(destFolder) Then
            CopyFile = False
            Exit Function
        End If
    End If
    
    ' Copier le fichier
    FileCopy sourcePath, destPath
    CopyFile = True
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la copie du fichier: " & Err.Description
    LogError m_LastError
    CopyFile = False
End Function

Public Function MoveFile(ByVal sourcePath As String, ByVal destPath As String, _
                         Optional ByVal overwrite As Boolean = False) As Boolean
    ' Déplace un fichier
    On Error GoTo ErrorHandler
    
    ' Vérifier si le fichier source existe
    If Not FileExists(sourcePath) Then
        m_LastError = "Le fichier source n'existe pas: " & sourcePath
        LogError m_LastError
        MoveFile = False
        Exit Function
    End If
    
    ' Vérifier si le fichier de destination existe déjà
    If FileExists(destPath) Then
        If overwrite Then
            ' Supprimer le fichier de destination existant
            If Not DeleteFile(destPath) Then
                MoveFile = False
                Exit Function
            End If
        Else
            m_LastError = "Le fichier de destination existe déjà: " & destPath
            LogError m_LastError
            MoveFile = False
            Exit Function
        End If
    End If
    
    ' Créer le répertoire de destination si nécessaire
    Dim destFolder As String
    destFolder = Left$(destPath, InStrRev(destPath, "\") - 1)
    
    If Not DirectoryExists(destFolder) Then
        If Not CreateDirectory(destFolder) Then
            MoveFile = False
            Exit Function
        End If
    End If
    
    ' Déplacer le fichier
    Name sourcePath As destPath
    MoveFile = True
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors du déplacement du fichier: " & Err.Description
    LogError m_LastError
    MoveFile = False
End Function

' --- Fonctions publiques pour la suppression ---
Public Function DeleteFile(ByVal filePath As String) As Boolean
    ' Supprime un fichier
    On Error GoTo ErrorHandler
    
    ' Vérifier si le fichier existe
    If Not FileExists(filePath) Then
        ' Si le fichier n'existe pas, considérer comme réussi
        DeleteFile = True
        Exit Function
    End If
    
    ' Vérifier si le fichier est en lecture seule
    If IsReadOnly(filePath) Then
        ' Enlever l'attribut lecture seule
        SetFileAttributes filePath, GetFileAttributes(filePath) And Not FILE_ATTRIBUTE_READONLY
    End If
    
    ' Supprimer le fichier
    Kill filePath
    DeleteFile = True
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la suppression du fichier: " & Err.Description
    LogError m_LastError
    DeleteFile = False
End Function

Public Function DeleteDirectory(ByVal folderPath As String, Optional ByVal recursive As Boolean = False) As Boolean
    ' Supprime un répertoire
    Dim file As String
    Dim subfolder As String
    
    On Error GoTo ErrorHandler
    
    ' Vérifier si le répertoire existe
    If Not DirectoryExists(folderPath) Then
        ' Si le répertoire n'existe pas, considérer comme réussi
        DeleteDirectory = True
        Exit Function
    End If
    
    ' Ajouter un séparateur de chemin si nécessaire
    If Right$(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    If recursive Then
        ' Supprimer d'abord tous les fichiers du répertoire
        file = Dir(folderPath & "*.*", vbNormal)
        
        Do While file <> ""
            If Not DeleteFile(folderPath & file) Then
                DeleteDirectory = False
                Exit Function
            End If
            
            file = Dir()
        Loop
        
        ' Puis supprimer tous les sous-répertoires
        subfolder = Dir(folderPath & "*.*", vbDirectory)
        
        Do While subfolder <> ""
            ' Ignorer "." et ".."
            If subfolder <> "." And subfolder <> ".." Then
                If (GetFileAttributes(folderPath & subfolder) And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
                    If Not DeleteDirectory(folderPath & subfolder, True) Then
                        DeleteDirectory = False
                        Exit Function
                    End If
                End If
            End If
            
            subfolder = Dir()
        Loop
    End If
    
    ' Supprimer le répertoire
    RmDir folderPath
    DeleteDirectory = True
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la suppression du répertoire: " & Err.Description
    LogError m_LastError
    DeleteDirectory = False
End Function

' --- Fonctions publiques pour les attributs de fichier ---
Public Function SetReadOnly(ByVal filePath As String, ByVal readOnly As Boolean) As Boolean
    ' Définit l'attribut lecture seule d'un fichier
    Dim attributes As Long
    
    On Error GoTo ErrorHandler
    
    ' Vérifier si le fichier existe
    If Not FileExists(filePath) Then
        m_LastError = "Le fichier n'existe pas: " & filePath
        LogError m_LastError
        SetReadOnly = False
        Exit Function
    End If
    
    ' Obtenir les attributs actuels du fichier
    attributes = GetFileAttributes(filePath)
    
    If attributes = INVALID_FILE_ATTRIBUTES Then
        m_LastError = "Impossible d'obtenir les attributs du fichier: " & filePath
        LogError m_LastError
        SetReadOnly = False
        Exit Function
    End If
    
    ' Modifier l'attribut lecture seule
    If readOnly Then
        attributes = attributes Or FILE_ATTRIBUTE_READONLY
    Else
        attributes = attributes And Not FILE_ATTRIBUTE_READONLY
    End If
    
    ' Appliquer les nouveaux attributs
    SetReadOnly = (SetFileAttributes(filePath, attributes) <> 0)
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la modification de l'attribut lecture seule: " & Err.Description
    LogError m_LastError
    SetReadOnly = False
End Function

Public Function SetHidden(ByVal filePath As String, ByVal hidden As Boolean) As Boolean
    ' Définit l'attribut caché d'un fichier
    Dim attributes As Long
    
    On Error GoTo ErrorHandler
    
    ' Vérifier si le fichier existe
    If Not FileExists(filePath) And Not DirectoryExists(filePath) Then
        m_LastError = "Le fichier ou répertoire n'existe pas: " & filePath
        LogError m_LastError
        SetHidden = False
        Exit Function
    End If
    
    ' Obtenir les attributs actuels du fichier
    attributes = GetFileAttributes(filePath)
    
    If attributes = INVALID_FILE_ATTRIBUTES Then
        m_LastError = "Impossible d'obtenir les attributs du fichier: " & filePath
        LogError m_LastError
        SetHidden = False
        Exit Function
    End If
    
    ' Modifier l'attribut caché
    If hidden Then
        attributes = attributes Or FILE_ATTRIBUTE_HIDDEN
    Else
        attributes = attributes And Not FILE_ATTRIBUTE_HIDDEN
    End If
    
    ' Appliquer les nouveaux attributs
    SetHidden = (SetFileAttributes(filePath, attributes) <> 0)
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la modification de l'attribut caché: " & Err.Description
    LogError m_LastError
    SetHidden = False
End Function

' --- Fonctions publiques pour la gestion des chemins ---
Public Function GetFileName(ByVal filePath As String) As String
    ' Retourne le nom du fichier à partir d'un chemin complet
    Dim pos As Long
    
    On Error GoTo ErrorHandler
    
    pos = InStrRev(filePath, "\")
    
    If pos > 0 Then
        GetFileName = Mid$(filePath, pos + 1)
    Else
        GetFileName = filePath
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'extraction du nom de fichier: " & Err.Description
    LogError m_LastError
    GetFileName = ""
End Function

Public Function GetFileExtension(ByVal filePath As String) As String
    ' Retourne l'extension d'un fichier
    Dim fileName As String
    Dim pos As Long
    
    On Error GoTo ErrorHandler
    
    fileName = GetFileName(filePath)
    pos = InStrRev(fileName, ".")
    
    If pos > 0 Then
        GetFileExtension = Mid$(fileName, pos + 1)
    Else
        GetFileExtension = ""
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'extraction de l'extension: " & Err.Description
    LogError m_LastError
    GetFileExtension = ""
End Function

Public Function GetFileNameWithoutExtension(ByVal filePath As String) As String
    ' Retourne le nom du fichier sans son extension
    Dim fileName As String
    Dim pos As Long
    
    On Error GoTo ErrorHandler
    
    fileName = GetFileName(filePath)
    pos = InStrRev(fileName, ".")
    
    If pos > 0 Then
        GetFileNameWithoutExtension = Left$(fileName, pos - 1)
    Else
        GetFileNameWithoutExtension = fileName
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'extraction du nom sans extension: " & Err.Description
    LogError m_LastError
    GetFileNameWithoutExtension = ""
End Function

Public Function GetDirectoryPath(ByVal filePath As String) As String
    ' Retourne le chemin du répertoire à partir d'un chemin complet
    Dim pos As Long
    
    On Error GoTo ErrorHandler
    
    pos = InStrRev(filePath, "\")
    
    If pos > 0 Then
        GetDirectoryPath = Left$(filePath, pos)
    Else
        GetDirectoryPath = ""
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'extraction du chemin du répertoire: " & Err.Description
    LogError m_LastError
    GetDirectoryPath = ""
End Function

Public Function CombinePaths(ByVal path1 As String, ByVal path2 As String) As String
    ' Combine deux chemins
    On Error GoTo ErrorHandler
    
    ' Supprimer les séparateurs de fin du premier chemin
    If Right$(path1, 1) = "\" Then
        path1 = Left$(path1, Len(path1) - 1)
    End If
    
    ' Supprimer les séparateurs de début du second chemin
    If Left$(path2, 1) = "\" Then
        path2 = Mid$(path2, 2)
    End If
    
    ' Combiner les chemins
    CombinePaths = path1 & "\" & path2
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la combinaison des chemins: " & Err.Description
    LogError m_LastError
    CombinePaths = ""
End Function

' --- Fonctions publiques pour la lecture/écriture de fichiers ---
Public Function ReadTextFile(ByVal filePath As String, Optional ByVal encoding As String = "utf-8") As String
    ' Lit le contenu d'un fichier texte
    Dim fileNum As Integer
    Dim content As String
    
    On Error GoTo ErrorHandler
    
    ' Vérifier si le fichier existe
    If Not FileExists(filePath) Then
        m_LastError = "Le fichier n'existe pas: " & filePath
        LogError m_LastError
        ReadTextFile = ""
        Exit Function
    End If
    
    ' Ouvrir et lire le fichier
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    content = Input$(LOF(fileNum), #fileNum)
    Close #fileNum
    
    ReadTextFile = content
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la lecture du fichier: " & Err.Description
    LogError m_LastError
    
    ' Fermer le fichier si ouvert
    If fileNum > 0 Then
        Close #fileNum
    End If
    
    ReadTextFile = ""
End Function

Public Function WriteTextFile(ByVal filePath As String, ByVal content As String, _
                              Optional ByVal overwrite As Boolean = True, _
                              Optional ByVal encoding As String = "utf-8") As Boolean
    ' Écrit du contenu dans un fichier texte
    Dim fileNum As Integer
    
    On Error GoTo ErrorHandler
    
    ' Vérifier si le fichier existe déjà
    If FileExists(filePath) And Not overwrite Then
        m_LastError = "Le fichier existe déjà: " & filePath
        LogError m_LastError
        WriteTextFile = False
        Exit Function
    End If
    
    ' Créer le répertoire si nécessaire
    Dim folderPath As String
    folderPath = GetDirectoryPath(filePath)
    
    If folderPath <> "" And Not DirectoryExists(folderPath) Then
        If Not CreateDirectory(folderPath) Then
            WriteTextFile = False
            Exit Function
        End If
    End If
    
    ' Écrire dans le fichier
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, content;
    Close #fileNum
    
    WriteTextFile = True
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'écriture dans le fichier: " & Err.Description
    LogError m_LastError
    
    ' Fermer le fichier si ouvert
    If fileNum > 0 Then
        Close #fileNum
    End If
    
    WriteTextFile = False
End Function

Public Function AppendToTextFile(ByVal filePath As String, ByVal content As String) As Boolean
    ' Ajoute du contenu à un fichier texte existant
    Dim fileNum As Integer
    
    On Error GoTo ErrorHandler
    
    ' Créer le fichier s'il n'existe pas
    If Not FileExists(filePath) Then
        ' Créer le répertoire si nécessaire
        Dim folderPath As String
        folderPath = GetDirectoryPath(filePath)
        
        If folderPath <> "" And Not DirectoryExists(folderPath) Then
            If Not CreateDirectory(folderPath) Then
                AppendToTextFile = False
                Exit Function
            End If
        End If
    End If
    
    ' Ajouter au fichier
    fileNum = FreeFile
    Open filePath For Append As #fileNum
    Print #fileNum, content;
    Close #fileNum
    
    AppendToTextFile = True
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'ajout au fichier: " & Err.Description
    LogError m_LastError
    
    ' Fermer le fichier si ouvert
    If fileNum > 0 Then
        Close #fileNum
    End If
    
    AppendToTextFile = False
End Function

' --- Fonctions publiques pour la recherche de fichiers ---
Public Function FindFiles(ByVal folderPath As String, ByVal pattern As String, _
                          Optional ByVal includeSubfolders As Boolean = False) As Variant
    ' Recherche des fichiers dans un répertoire
    Dim files() As String
    Dim fileCount As Long
    Dim file As String
    Dim subfolder As String
    
    On Error GoTo ErrorHandler
    
    ' Initialiser le tableau
    ReDim files(0 To 99)
    fileCount = 0
    
    ' Vérifier si le répertoire existe
    If Not DirectoryExists(folderPath) Then
        m_LastError = "Le répertoire n'existe pas: " & folderPath
        LogError m_LastError
        FindFiles = Array()
        Exit Function
    End If
    
    ' Ajouter un séparateur de chemin si nécessaire
    If Right$(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    ' Rechercher les fichiers correspondant au motif
    file = Dir(folderPath & pattern, vbNormal)
    
    Do While file <> ""
        ' Ajouter le fichier au tableau
        If fileCount > UBound(files) Then
            ReDim Preserve files(0 To UBound(files) * 2)
        End If
        
        files(fileCount) = folderPath & file
        fileCount = fileCount + 1
        
        file = Dir()
    Loop
    
    ' Rechercher dans les sous-répertoires si demandé
    If includeSubfolders Then
        subfolder = Dir(folderPath & "*.*", vbDirectory)
        
        Do While subfolder <> ""
            ' Ignorer "." et ".."
            If subfolder <> "." And subfolder <> ".." Then
                If (GetFileAttributes(folderPath & subfolder) And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
                    ' Rechercher récursivement dans le sous-répertoire
                    Dim subfiles As Variant
                    subfiles = FindFiles(folderPath & subfolder, pattern, True)
                    
                    ' Ajouter les fichiers trouvés au tableau
                    If IsArray(subfiles) Then
                        Dim i As Long
                        For i = LBound(subfiles) To UBound(subfiles)
                            If fileCount > UBound(files) Then
                                ReDim Preserve files(0 To UBound(files) * 2)
                            End If
                            
                            files(fileCount) = subfiles(i)
                            fileCount = fileCount + 1
                        Next i
                    End If
                End If
            End If
            
            subfolder = Dir()
        Loop
    End If
    
    ' Redimensionner le tableau au nombre exact de fichiers trouvés
    If fileCount > 0 Then
        ReDim Preserve files(0 To fileCount - 1)
        FindFiles = files
    Else
        FindFiles = Array()
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la recherche de fichiers: " & Err.Description
    LogError m_LastError
    FindFiles = Array()
End Function

' --- Constantes utilitaires ---
Private Const MAX_PATH As Long = 260

' --- Propriétés ---
Public Property Get LastError() As String
    ' Retourne la dernière erreur survenue
    LastError = m_LastError
End Property

' --- Fonctions privées ---
Private Sub LogError(ByVal errorMessage As String)
    ' Log les erreurs si un logger est disponible
    If Not m_Logger Is Nothing Then
        ' TODO: Utiliser le logger pour enregistrer l'erreur
        ' m_Logger.LogError errorMessage, "FILE"
    End If
End Sub
