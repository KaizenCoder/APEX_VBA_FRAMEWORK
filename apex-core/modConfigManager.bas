' Migrated to apex-core - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

' ==========================================================================
' Module : modConfigManager
' Version : 1.0
' Purpose : Gestion centralisée des configurations du framework
' Date : 10/04/2025
' ==========================================================================

' --- API Windows pour accéder aux fichiers INI ---
#If VBA7 Then
    Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    
    Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, _
        ByVal lpFileName As String) As Long
    
    Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" _
        (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
        ByVal lpFileName As String) As Long
    
    Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" _
        (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#Else
    Private Declare'@Description: 
'@Param: 
'@Returns: 

 Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    
    Private Declare'@Description: 
'@Param: 
'@Returns: 

 Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, _
        ByVal lpFileName As String) As Long
    
    Private Declare'@Description: 
'@Param: 
'@Returns: 

 Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" _
        (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
        ByVal lpFileName As String) As Long
    
    Private Declare'@Description: 
'@Param: 
'@Returns: 

 Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" _
        (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#End If

' --- Constantes ---
Private Const MAX_BUFFER_SIZE As Long = 32767
Private Const DEFAULT_CONFIG_DIR As String = "config\"
Private Const DEFAULT_CONFIG_FILE As String = "apex_config.ini"

' --- Variables privées ---
Private m_DefaultConfigPath As String
Private m_Logger As Object ' ILoggerBase
Private m_LastError As String

' --- Initialisation ---
'@Description: 
'@Param: 
'@Returns: 

Public Sub Initialize(Optional ByVal logger As Object = Nothing, Optional ByVal defaultConfigPath As String = "")
    ' Initialise le module avec un logger et un chemin de configuration par défaut
    Set m_Logger = logger
    
    ' Définir le chemin de configuration par défaut
    If defaultConfigPath = "" Then
        m_DefaultConfigPath = DEFAULT_CONFIG_DIR & DEFAULT_CONFIG_FILE
    Else
        m_DefaultConfigPath = defaultConfigPath
    End If
    
    m_LastError = ""
    
    If Not m_Logger Is Nothing Then
        ' TODO: Logger l'initialisation
        ' m_Logger.LogInfo "modConfigManager initialisé avec config par défaut: " & m_DefaultConfigPath, "CONFIG"
    End If
End Sub

' --- Fonctions publiques pour la lecture ---
'@Description: 
'@Param: 
'@Returns: 

Public Function ReadString(ByVal section As String, ByVal key As String, Optional ByVal defaultValue As String = "", _
                           Optional ByVal configPath As String = "") As String
    ' Lit une valeur de type chaîne depuis un fichier INI
    Dim buffer As String
    Dim length As Long
    Dim filePath As String
    
    On Error GoTo ErrorHandler
    
    ' Utiliser le chemin par défaut si aucun n'est spécifié
    If configPath = "" Then
        filePath = m_DefaultConfigPath
    Else
        filePath = configPath
    End If
    
    ' Vérifier si le fichier existe
    If Dir(filePath) = "" Then
        m_LastError = "Fichier de configuration introuvable: " & filePath
        LogError m_LastError
        ReadString = defaultValue
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Lire la valeur
    buffer = Space$(MAX_BUFFER_SIZE)
    length = GetPrivateProfileString(section, key, defaultValue, buffer, Len(buffer), filePath)
    
    If length > 0 Then
        ReadString = Left$(buffer, length)
    Else
        ReadString = defaultValue
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la lecture de la valeur: " & Err.Description
    LogError m_LastError
    ReadString = defaultValue
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function ReadInteger(ByVal section As String, ByVal key As String, Optional ByVal defaultValue As Long = 0, _
                            Optional ByVal configPath As String = "") As Long
    ' Lit une valeur de type entier depuis un fichier INI
    Dim value As String
    
    On Error GoTo ErrorHandler
    
    ' Lire la valeur comme une chaîne
    value = ReadString(section, key, CStr(defaultValue), configPath)
    
    ' Convertir en entier
    If IsNumeric(value) Then
        ReadInteger = CLng(value)
    Else
        ReadInteger = defaultValue
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la lecture de la valeur entière: " & Err.Description
    LogError m_LastError
    ReadInteger = defaultValue
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function ReadBoolean(ByVal section As String, ByVal key As String, Optional ByVal defaultValue As Boolean = False, _
                            Optional ByVal configPath As String = "") As Boolean
    ' Lit une valeur de type booléen depuis un fichier INI
    Dim value As String
    
    On Error GoTo ErrorHandler
    
    ' Lire la valeur comme une chaîne
    value = ReadString(section, key, IIf(defaultValue, "True", "False"), configPath)
    
    ' Convertir en booléen
    Select Case UCase(value)
        Case "TRUE", "YES", "1", "OUI", "VRAI"
            ReadBoolean = True
        Case "FALSE", "NO", "0", "NON", "FAUX"
            ReadBoolean = False
        Case Else
            ReadBoolean = defaultValue
    End Select
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la lecture de la valeur booléenne: " & Err.Description
    LogError m_LastError
    ReadBoolean = defaultValue
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function ReadDouble(ByVal section As String, ByVal key As String, Optional ByVal defaultValue As Double = 0, _
                           Optional ByVal configPath As String = "") As Double
    ' Lit une valeur de type double depuis un fichier INI
    Dim value As String
    
    On Error GoTo ErrorHandler
    
    ' Lire la valeur comme une chaîne
    value = ReadString(section, key, CStr(defaultValue), configPath)
    
    ' Convertir en double
    If IsNumeric(value) Then
        ReadDouble = CDbl(value)
    Else
        ReadDouble = defaultValue
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la lecture de la valeur double: " & Err.Description
    LogError m_LastError
    ReadDouble = defaultValue
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function ReadSection(ByVal section As String, Optional ByVal configPath As String = "") As Variant
    ' Lit toutes les entrées d'une section et les retourne dans un tableau
    Dim buffer As String
    Dim length As Long
    Dim filePath As String
    Dim entries() As String
    
    On Error GoTo ErrorHandler
    
    ' Utiliser le chemin par défaut si aucun n'est spécifié
    If configPath = "" Then
        filePath = m_DefaultConfigPath
    Else
        filePath = configPath
    End If
    
    ' Vérifier si le fichier existe
    If Dir(filePath) = "" Then
        m_LastError = "Fichier de configuration introuvable: " & filePath
        LogError m_LastError
        ReadSection = Array()
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Lire la section
    buffer = Space$(MAX_BUFFER_SIZE)
    length = GetPrivateProfileSection(section, buffer, Len(buffer), filePath)
    
    If length > 0 Then
        ' Extraire les entrées
        buffer = Left$(buffer, length)
        
        ' Séparer les entrées (chaque entrée est terminée par un caractère nul)
        entries = Split(Replace(buffer, Chr(0), vbLf), vbLf)
        
        ' Supprimer la dernière entrée si elle est vide
        If entries(UBound(entries)) = "" Then
            ReDim Preserve entries(UBound(entries) - 1)
        End If
        
        ReadSection = entries
    Else
        ReadSection = Array()
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la lecture de la section: " & Err.Description
    LogError m_LastError
    ReadSection = Array()
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetSectionNames(Optional ByVal configPath As String = "") As Variant
    ' Récupère les noms de toutes les sections du fichier INI
    Dim buffer As String
    Dim length As Long
    Dim filePath As String
    Dim sections() As String
    
    On Error GoTo ErrorHandler
    
    ' Utiliser le chemin par défaut si aucun n'est spécifié
    If configPath = "" Then
        filePath = m_DefaultConfigPath
    Else
        filePath = configPath
    End If
    
    ' Vérifier si le fichier existe
    If Dir(filePath) = "" Then
        m_LastError = "Fichier de configuration introuvable: " & filePath
        LogError m_LastError
        GetSectionNames = Array()
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Lire les noms de sections
    buffer = Space$(MAX_BUFFER_SIZE)
    length = GetPrivateProfileSectionNames(buffer, Len(buffer), filePath)
    
    If length > 0 Then
        ' Extraire les noms de sections
        buffer = Left$(buffer, length)
        
        ' Séparer les noms de sections (chaque nom est terminé par un caractère nul)
        sections = Split(Replace(buffer, Chr(0), vbLf), vbLf)
        
        ' Supprimer la dernière entrée si elle est vide
        If sections(UBound(sections)) = "" Then
            ReDim Preserve sections(UBound(sections) - 1)
        End If
        
        GetSectionNames = sections
    Else
        GetSectionNames = Array()
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la récupération des noms de sections: " & Err.Description
    LogError m_LastError
    GetSectionNames = Array()
End Function

' --- Fonctions publiques pour l'écriture ---
'@Description: 
'@Param: 
'@Returns: 

Public Function WriteString(ByVal section As String, ByVal key As String, ByVal value As String, _
                            Optional ByVal configPath As String = "") As Boolean
    ' Écrit une valeur de type chaîne dans un fichier INI
    Dim result As Long
    Dim filePath As String
    
    On Error GoTo ErrorHandler
    
    ' Utiliser le chemin par défaut si aucun n'est spécifié
    If configPath = "" Then
        filePath = m_DefaultConfigPath
    Else
        filePath = configPath
    End If
    
    ' Créer le répertoire si nécessaire
    CreateConfigDirectory filePath
    
    ' Écrire la valeur
    result = WritePrivateProfileString(section, key, value, filePath)
    WriteString = (result <> 0)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'écriture de la valeur: " & Err.Description
    LogError m_LastError
    WriteString = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function WriteInteger(ByVal section As String, ByVal key As String, ByVal value As Long, _
                             Optional ByVal configPath As String = "") As Boolean
    ' Écrit une valeur de type entier dans un fichier INI
    WriteInteger = WriteString(section, key, CStr(value), configPath)
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function WriteBoolean(ByVal section As String, ByVal key As String, ByVal value As Boolean, _
                             Optional ByVal configPath As String = "") As Boolean
    ' Écrit une valeur de type booléen dans un fichier INI
    WriteBoolean = WriteString(section, key, IIf(value, "True", "False"), configPath)
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function WriteDouble(ByVal section As String, ByVal key As String, ByVal value As Double, _
                            Optional ByVal configPath As String = "") As Boolean
    ' Écrit une valeur de type double dans un fichier INI
    WriteDouble = WriteString(section, key, CStr(value), configPath)
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function DeleteKey(ByVal section As String, ByVal key As String, _
                          Optional ByVal configPath As String = "") As Boolean
    ' Supprime une clé d'un fichier INI
    Dim result As Long
    Dim filePath As String
    
    On Error GoTo ErrorHandler
    
    ' Utiliser le chemin par défaut si aucun n'est spécifié
    If configPath = "" Then
        filePath = m_DefaultConfigPath
    Else
        filePath = configPath
    End If
    
    ' Vérifier si le fichier existe
    If Dir(filePath) = "" Then
        m_LastError = "Fichier de configuration introuvable: " & filePath
        LogError m_LastError
        DeleteKey = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Supprimer la clé (en écrivant une valeur NULL)
    result = WritePrivateProfileString(section, key, vbNullString, filePath)
    DeleteKey = (result <> 0)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la suppression de la clé: " & Err.Description
    LogError m_LastError
    DeleteKey = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function DeleteSection(ByVal section As String, Optional ByVal configPath As String = "") As Boolean
    ' Supprime une section entière d'un fichier INI
    Dim result As Long
    Dim filePath As String
    
    On Error GoTo ErrorHandler
    
    ' Utiliser le chemin par défaut si aucun n'est spécifié
    If configPath = "" Then
        filePath = m_DefaultConfigPath
    Else
        filePath = configPath
    End If
    
    ' Vérifier si le fichier existe
    If Dir(filePath) = "" Then
        m_LastError = "Fichier de configuration introuvable: " & filePath
        LogError m_LastError
        DeleteSection = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Supprimer la section (en écrivant une section NULL)
    result = WritePrivateProfileString(section, vbNullString, vbNullString, filePath)
    DeleteSection = (result <> 0)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la suppression de la section: " & Err.Description
    LogError m_LastError
    DeleteSection = False
End Function

' --- Autres fonctions publiques ---
'@Description: 
'@Param: 
'@Returns: 

Public Function KeyExists(ByVal section As String, ByVal key As String, _
                         Optional ByVal configPath As String = "") As Boolean
    ' Vérifie si une clé existe dans un fichier INI
    Dim value As String
    Dim defaultValue As String
    Dim buffer As String
    Dim length As Long
    Dim filePath As String
    
    On Error GoTo ErrorHandler
    
    ' Utiliser le chemin par défaut si aucun n'est spécifié
    If configPath = "" Then
        filePath = m_DefaultConfigPath
    Else
        filePath = configPath
    End If
    
    ' Vérifier si le fichier existe
    If Dir(filePath) = "" Then
        m_LastError = "Fichier de configuration introuvable: " & filePath
        LogError m_LastError
        KeyExists = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Utiliser une valeur par défaut unique pour détecter si la clé existe
    defaultValue = "@@KEY_NOT_FOUND@@" & Format(Now, "yyyymmddhhnnss") & "@@"
    
    ' Lire la valeur
    buffer = Space$(MAX_BUFFER_SIZE)
    length = GetPrivateProfileString(section, key, defaultValue, buffer, Len(buffer), filePath)
    
    value = Left$(buffer, length)
    
    ' Si la valeur lue est différente de la valeur par défaut, la clé existe
    KeyExists = (value <> defaultValue)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la vérification de l'existence de la clé: " & Err.Description
    LogError m_LastError
    KeyExists = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function SectionExists(ByVal section As String, Optional ByVal configPath As String = "") As Boolean
    ' Vérifie si une section existe dans un fichier INI
    Dim sections As Variant
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' Récupérer toutes les sections
    sections = GetSectionNames(configPath)
    
    ' Vérifier si la section existe
    For i = LBound(sections) To UBound(sections)
        If UCase(sections(i)) = UCase(section) Then
            SectionExists = True
            Exit'@Description: 
'@Param: 
'@Returns: 

 Function
        End If
    Next i
    
    SectionExists = False
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la vérification de l'existence de la section: " & Err.Description
    LogError m_LastError
    SectionExists = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function CreateConfigFile(ByVal configPath As String) As Boolean
    ' Crée un nouveau fichier de configuration
    Dim fileNum As Integer
    
    On Error GoTo ErrorHandler
    
    ' Créer le répertoire si nécessaire
    CreateConfigDirectory configPath
    
    ' Créer le fichier
    fileNum = FreeFile
    Open configPath For Output As #fileNum
    Close #fileNum
    
    CreateConfigFile = True
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la création du fichier de configuration: " & Err.Description
    LogError m_LastError
    CreateConfigFile = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function ConfigFileExists(Optional ByVal configPath As String = "") As Boolean
    ' Vérifie si un fichier de configuration existe
    Dim filePath As String
    
    On Error GoTo ErrorHandler
    
    ' Utiliser le chemin par défaut si aucun n'est spécifié
    If configPath = "" Then
        filePath = m_DefaultConfigPath
    Else
        filePath = configPath
    End If
    
    ConfigFileExists = (Dir(filePath) <> "")
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la vérification de l'existence du fichier: " & Err.Description
    LogError m_LastError
    ConfigFileExists = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Property Get LastError() As String
    ' Retourne la dernière erreur survenue
    LastError = m_LastError
End Property

Public Property Get DefaultConfigPath() As String
    ' Retourne le chemin de configuration par défaut
    DefaultConfigPath = m_DefaultConfigPath
End Property

Public Property Let DefaultConfigPath(ByVal value As String)
    ' Définit le chemin de configuration par défaut
    m_DefaultConfigPath = value
End Property

' --- Fonctions privées ---
'@Description: 
'@Param: 
'@Returns: 

Private Sub CreateConfigDirectory(ByVal filePath As String)
    ' Crée le répertoire de configuration si nécessaire
    Dim folderPath As String
    Dim pos As Long
    
    On Error Resume Next
    
    ' Extraire le chemin du dossier
    pos = InStrRev(filePath, "\")
    If pos > 0 Then
        folderPath = Left$(filePath, pos - 1)
        
        ' Créer le dossier s'il n'existe pas
        If folderPath <> "" And Dir(folderPath, vbDirectory) = "" Then
            MkDir folderPath
        End If
    End If
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub LogError(ByVal errorMessage As String)
    ' Log les erreurs si un logger est disponible
    If Not m_Logger Is Nothing Then
        ' TODO: Utiliser le logger pour enregistrer l'erreur
        ' m_Logger.LogError errorMessage, "CONFIG"
    End If
End Sub
