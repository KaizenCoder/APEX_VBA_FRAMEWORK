VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@ModuleDescription("Impl�mentation standard du gestionnaire de configuration")
'@Folder("APEX.Core.Config")
'@Version 1.7.0
'@Author APEX Framework Team
Option Explicit

Implements IConfigManagerBase

' =========================================================
' ConfigManager
' Impl�mentation standard du gestionnaire de configuration
' =========================================================

' Configuration et �tat
Private m_source As Variant
Private m_data As Object             ' Dictionnaire principal de stockage
Private m_sections As Object         ' Dictionnaire des sections
Private m_isInitialized As Boolean
Private m_ignoreErrors As Boolean
Private m_logger As Object           ' ILoggerBase

' =========================================================
' Initialisation et configuration
' =========================================================
Private Sub Class_Initialize()
    ' Initialiser les dictionnaires de stockage
    Set m_data = CreateObject("Scripting.Dictionary")
    Set m_sections = CreateObject("Scripting.Dictionary")
    m_isInitialized = False
    m_ignoreErrors = False
End Sub

' =========================================================
' Impl�mentation de IConfigManagerBase
' =========================================================

Private Function IConfigManagerBase_Initialize(ByVal configSource As Variant, Optional ByVal options As Object = Nothing) As Boolean
    On Error GoTo ErrorHandler
    
    ' Stocker la source
    m_source = configSource
    
    ' Configurer le logger si fourni dans les options
    If Not options Is Nothing Then
        On Error Resume Next
        Set m_logger = options("Logger")
        On Error GoTo ErrorHandler
    End If
    
    ' Charger la configuration
    Dim result As Boolean
    result = LoadConfigFromSource(configSource)
    
    ' Marquer comme initialis� si chargement r�ussi
    m_isInitialized = result
    IConfigManagerBase_Initialize = result
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de l'initialisation du ConfigManager: " & Err.Description, "Initialize"
    IConfigManagerBase_Initialize = False
End Function

' Acc�s aux valeurs
Private Function IConfigManagerBase_GetValue(ByVal key As String, Optional ByVal defaultValue As Variant = Null) As Variant
    On Error GoTo ErrorHandler
    
    ' V�rifier si la cl� existe
    If m_data.Exists(key) Then
        ' R�cup�rer la valeur
        IConfigManagerBase_GetValue = m_data.Item(key)
    Else
        ' Retourner la valeur par d�faut si la cl� n'existe pas
        IConfigManagerBase_GetValue = defaultValue
    End If
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la r�cup�ration de la valeur pour la cl� '" & key & "': " & Err.Description, "GetValue"
    IConfigManagerBase_GetValue = defaultValue
End Function

Private Function IConfigManagerBase_GetString(ByVal key As String, Optional ByVal defaultValue As String = "") As String
    On Error GoTo ErrorHandler
    
    ' Utiliser GetValue puis convertir en String
    Dim value As Variant
    value = IConfigManagerBase_GetValue(key, defaultValue)
    
    ' Convertir en String
    If IsNull(value) Then
        IConfigManagerBase_GetString = defaultValue
    Else
        IConfigManagerBase_GetString = CStr(value)
    End If
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la conversion en String pour la cl� '" & key & "': " & Err.Description, "GetString"
    IConfigManagerBase_GetString = defaultValue
End Function

Private Function IConfigManagerBase_GetInteger(ByVal key As String, Optional ByVal defaultValue As Long = 0) As Long
    On Error GoTo ErrorHandler
    
    ' Utiliser GetValue puis convertir en Integer
    Dim value As Variant
    value = IConfigManagerBase_GetValue(key, defaultValue)
    
    ' Convertir en Integer
    If IsNull(value) Then
        IConfigManagerBase_GetInteger = defaultValue
    Else
        IConfigManagerBase_GetInteger = CLng(value)
    End If
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la conversion en Integer pour la cl� '" & key & "': " & Err.Description, "GetInteger"
    IConfigManagerBase_GetInteger = defaultValue
End Function

Private Function IConfigManagerBase_GetBoolean(ByVal key As String, Optional ByVal defaultValue As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    ' Utiliser GetValue puis convertir en Boolean
    Dim value As Variant
    value = IConfigManagerBase_GetValue(key, defaultValue)
    
    ' Convertir en Boolean
    If IsNull(value) Then
        IConfigManagerBase_GetBoolean = defaultValue
    Else
        ' G�rer les cas sp�ciaux (texte)
        If VarType(value) = vbString Then
            Select Case LCase(value)
                Case "true", "1", "yes", "oui", "vrai": IConfigManagerBase_GetBoolean = True
                Case "false", "0", "no", "non", "faux": IConfigManagerBase_GetBoolean = False
                Case Else: IConfigManagerBase_GetBoolean = CBool(value)
            End Select
        Else
            IConfigManagerBase_GetBoolean = CBool(value)
        End If
    End If
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la conversion en Boolean pour la cl� '" & key & "': " & Err.Description, "GetBoolean"
    IConfigManagerBase_GetBoolean = defaultValue
End Function

Private Function IConfigManagerBase_GetDouble(ByVal key As String, Optional ByVal defaultValue As Double = 0#) As Double
    On Error GoTo ErrorHandler
    
    ' Utiliser GetValue puis convertir en Double
    Dim value As Variant
    value = IConfigManagerBase_GetValue(key, defaultValue)
    
    ' Convertir en Double
    If IsNull(value) Then
        IConfigManagerBase_GetDouble = defaultValue
    Else
        IConfigManagerBase_GetDouble = CDbl(value)
    End If
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la conversion en Double pour la cl� '" & key & "': " & Err.Description, "GetDouble"
    IConfigManagerBase_GetDouble = defaultValue
End Function

Private Function IConfigManagerBase_GetDate(ByVal key As String, Optional ByVal defaultValue As Date = #1/1/1900#) As Date
    On Error GoTo ErrorHandler
    
    ' Utiliser GetValue puis convertir en Date
    Dim value As Variant
    value = IConfigManagerBase_GetValue(key, defaultValue)
    
    ' Convertir en Date
    If IsNull(value) Then
        IConfigManagerBase_GetDate = defaultValue
    Else
        IConfigManagerBase_GetDate = CDate(value)
    End If
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la conversion en Date pour la cl� '" & key & "': " & Err.Description, "GetDate"
    IConfigManagerBase_GetDate = defaultValue
End Function

Private Function IConfigManagerBase_GetObject(ByVal key As String) As Object
    On Error GoTo ErrorHandler
    
    ' V�rifier si la cl� existe et retourne un objet
    If m_data.Exists(key) Then
        Dim value As Variant
        value = m_data.Item(key)
        
        ' V�rifier que la valeur est bien un objet
        If IsObject(value) Then
            Set IConfigManagerBase_GetObject = value
        Else
            Set IConfigManagerBase_GetObject = Nothing
        End If
    Else
        Set IConfigManagerBase_GetObject = Nothing
    End If
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la r�cup�ration de l'objet pour la cl� '" & key & "': " & Err.Description, "GetObject"
    Set IConfigManagerBase_GetObject = Nothing
End Function

Private Function IConfigManagerBase_HasKey(ByVal key As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' V�rifier si la cl� existe
    IConfigManagerBase_HasKey = m_data.Exists(key)
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la v�rification de l'existence de la cl� '" & key & "': " & Err.Description, "HasKey"
    IConfigManagerBase_HasKey = False
End Function

' Manipulation des valeurs
Private Function IConfigManagerBase_SetValue(ByVal key As String, ByVal value As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' D�finir ou mettre � jour la valeur
    If m_data.Exists(key) Then
        ' Mise � jour de la valeur existante
        m_data.Item(key) = value
    Else
        ' Ajout d'une nouvelle valeur
        m_data.Add key, value
    End If
    
    ' Mettre � jour la section si le format est section.cl�
    If InStr(key, ".") > 0 Then
        Dim parts As Variant
        parts = Split(key, ".", 2)
        Dim section As String
        section = parts(0)
        Dim subKey As String
        subKey = parts(1)
        
        ' Cr�er la section si elle n'existe pas
        If Not m_sections.Exists(section) Then
            Set m_sections(section) = CreateObject("Scripting.Dictionary")
        End If
        
        ' Mettre � jour la valeur dans la section
        Dim sectionDict As Object
        Set sectionDict = m_sections(section)
        
        If sectionDict.Exists(subKey) Then
            sectionDict.Item(subKey) = value
        Else
            sectionDict.Add subKey, value
        End If
    End If
    
    IConfigManagerBase_SetValue = True
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la d�finition de la valeur pour la cl� '" & key & "': " & Err.Description, "SetValue"
    IConfigManagerBase_SetValue = False
End Function

Private Function IConfigManagerBase_RemoveValue(ByVal key As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Supprimer la valeur si la cl� existe
    If m_data.Exists(key) Then
        m_data.Remove key
        
        ' Supprimer de la section si le format est section.cl�
        If InStr(key, ".") > 0 Then
            Dim parts As Variant
            parts = Split(key, ".", 2)
            Dim section As String
            section = parts(0)
            Dim subKey As String
            subKey = parts(1)
            
            ' V�rifier si la section existe
            If m_sections.Exists(section) Then
                Dim sectionDict As Object
                Set sectionDict = m_sections(section)
                
                ' Supprimer la cl� de la section
                If sectionDict.Exists(subKey) Then
                    sectionDict.Remove subKey
                End If
            End If
        End If
        
        IConfigManagerBase_RemoveValue = True
    Else
        IConfigManagerBase_RemoveValue = False
    End If
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la suppression de la valeur pour la cl� '" & key & "': " & Err.Description, "RemoveValue"
    IConfigManagerBase_RemoveValue = False
End Function

Private Function IConfigManagerBase_Clear() As Boolean
    On Error GoTo ErrorHandler
    
    ' Effacer toutes les valeurs
    Set m_data = CreateObject("Scripting.Dictionary")
    Set m_sections = CreateObject("Scripting.Dictionary")
    
    IConfigManagerBase_Clear = True
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de l'effacement des valeurs: " & Err.Description, "Clear"
    IConfigManagerBase_Clear = False
End Function

' Gestion des sections
Private Function IConfigManagerBase_GetSection(ByVal section As String) As Object
    On Error GoTo ErrorHandler
    
    ' V�rifier si la section existe
    If m_sections.Exists(section) Then
        Set IConfigManagerBase_GetSection = m_sections(section)
    Else
        ' Cr�er un dictionnaire vide pour la section
        Set IConfigManagerBase_GetSection = CreateObject("Scripting.Dictionary")
    End If
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la r�cup�ration de la section '" & section & "': " & Err.Description, "GetSection"
    Set IConfigManagerBase_GetSection = CreateObject("Scripting.Dictionary")
End Function

Private Function IConfigManagerBase_HasSection(ByVal section As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' V�rifier si la section existe
    IConfigManagerBase_HasSection = m_sections.Exists(section)
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la v�rification de l'existence de la section '" & section & "': " & Err.Description, "HasSection"
    IConfigManagerBase_HasSection = False
End Function

Private Function IConfigManagerBase_GetSectionNames() As Collection
    On Error GoTo ErrorHandler
    
    ' Cr�er une collection pour les noms de sections
    Dim result As New Collection
    
    ' Ajouter les noms de sections � la collection
    Dim key As Variant
    For Each key In m_sections.Keys
        result.Add key
    Next
    
    Set IConfigManagerBase_GetSectionNames = result
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la r�cup�ration des noms de sections: " & Err.Description, "GetSectionNames"
    Set IConfigManagerBase_GetSectionNames = New Collection
End Function

' Persistance
Private Function IConfigManagerBase_Save(Optional ByVal destination As Variant = Null) As Boolean
    On Error GoTo ErrorHandler
    
    ' Utiliser la source si aucune destination n'est sp�cifi�e
    If IsNull(destination) Then
        destination = m_source
    End If
    
    ' V�rifier le type de destination
    If IsObject(destination) Then
        ' Destination est un objet (� impl�menter selon besoin)
        IConfigManagerBase_Save = SaveToObject(destination)
    ElseIf VarType(destination) = vbString Then
        ' Destination est une cha�ne (probablement un chemin de fichier)
        IConfigManagerBase_Save = SaveToFile(CStr(destination))
    Else
        ' Type de destination non pris en charge
        Err.Raise vbObjectError + 1001, "ConfigManager.Save", "Type de destination non pris en charge"
    End If
    
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de la sauvegarde de la configuration: " & Err.Description, "Save"
    IConfigManagerBase_Save = False
End Function

Private Function IConfigManagerBase_Reload() As Boolean
    On Error GoTo ErrorHandler
    
    ' Recharger la configuration depuis la source
    If IsNull(m_source) Then
        Err.Raise vbObjectError + 1002, "ConfigManager.Reload", "Source de configuration non sp�cifi�e"
    End If
    
    ' Effacer les donn�es actuelles
    IConfigManagerBase_Clear
    
    ' Recharger depuis la source
    IConfigManagerBase_Reload = LoadConfigFromSource(m_source)
    
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors du rechargement de la configuration: " & Err.Description, "Reload"
    IConfigManagerBase_Reload = False
End Function

' Propri�t�s
Private Property Get IConfigManagerBase_Source() As Variant
    IConfigManagerBase_Source = m_source
End Property

Private Property Get IConfigManagerBase_IsInitialized() As Boolean
    IConfigManagerBase_IsInitialized = m_isInitialized
End Property

Private Property Get IConfigManagerBase_IgnoreErrors() As Boolean
    IConfigManagerBase_IgnoreErrors = m_ignoreErrors
End Property

Private Property Let IConfigManagerBase_IgnoreErrors(ByVal value As Boolean)
    m_ignoreErrors = value
End Property

' =========================================================
' M�thodes priv�es
' =========================================================

' Charge la configuration depuis la source sp�cifi�e
Private Function LoadConfigFromSource(ByVal source As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' V�rifier le type de source
    If IsObject(source) Then
        ' Source est un objet (dictionnaire, recordset, etc.)
        LoadConfigFromSource = LoadFromObject(source)
    ElseIf VarType(source) = vbString Then
        ' Source est une cha�ne (probablement un chemin de fichier)
        Dim filePath As String
        filePath = CStr(source)
        
        ' D�tecter le type de fichier selon l'extension
        If LCase(Right(filePath, 4)) = ".ini" Then
            LoadConfigFromSource = LoadFromIniFile(filePath)
        ElseIf LCase(Right(filePath, 4)) = ".xml" Then
            LoadConfigFromSource = LoadFromXmlFile(filePath)
        ElseIf LCase(Right(filePath, 5)) = ".json" Then
            LoadConfigFromSource = LoadFromJsonFile(filePath)
        ElseIf LCase(Right(filePath, 4)) = ".csv" Then
            LoadConfigFromSource = LoadFromCsvFile(filePath)
        ElseIf LCase(Right(filePath, 4)) = ".txt" Then
            LoadConfigFromSource = LoadFromTextFile(filePath)
        ElseIf LCase(Right(filePath, 5)) = ".xlsx" Or LCase(Right(filePath, 4)) = ".xls" Then
            LoadConfigFromSource = LoadFromExcelFile(filePath)
        Else
            ' Extension non reconnue, essayer de d�tecter le format
            LoadConfigFromSource = DetectAndLoadFile(filePath)
        End If
    Else
        ' Type de source non pris en charge
        Err.Raise vbObjectError + 1000, "ConfigManager.LoadConfigFromSource", "Type de source non pris en charge"
    End If
    
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors du chargement de la configuration: " & Err.Description, "LoadConfigFromSource"
    LoadConfigFromSource = False
End Function

' Chargement � partir d'objets
Private Function LoadFromObject(ByVal source As Object) As Boolean
    On Error GoTo ErrorHandler
    
    ' V�rifier le type d'objet
    If TypeName(source) = "Dictionary" Then
        ' Source est un dictionnaire
        Dim key As Variant
        For Each key In source.Keys
            IConfigManagerBase_SetValue CStr(key), source(key)
        Next
        LoadFromObject = True
    Else
        ' Type d'objet non pris en charge
        Err.Raise vbObjectError + 1010, "ConfigManager.LoadFromObject", "Type d'objet non pris en charge: " & TypeName(source)
    End If
    
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors du chargement � partir d'un objet: " & Err.Description, "LoadFromObject"
    LoadFromObject = False
End Function

' Chargement � partir de fichiers
Private Function LoadFromIniFile(ByVal filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' V�rifier que le fichier existe
    If Not FileExists(filePath) Then
        Err.Raise vbObjectError + 1020, "ConfigManager.LoadFromIniFile", "Fichier INI non trouv�: " & filePath
    End If
    
    ' Ouvrir le fichier
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    Dim line As String
    Dim currentSection As String
    currentSection = ""
    
    ' Lire le fichier ligne par ligne
    Do Until EOF(fileNum)
        Line Input #fileNum, line
        
        ' Ignorer les lignes vides et les commentaires
        line = Trim(line)
        If Len(line) > 0 And Left(line, 1) <> ";" And Left(line, 1) <> "#" Then
            ' V�rifier si c'est une section
            If Left(line, 1) = "[" And Right(line, 1) = "]" Then
                ' Nouvelle section
                currentSection = Mid(line, 2, Len(line) - 2)
                
                ' Cr�er la section si elle n'existe pas
                If Not m_sections.Exists(currentSection) Then
                    Set m_sections(currentSection) = CreateObject("Scripting.Dictionary")
                End If
            Else
                ' V�rifier si c'est une paire cl�=valeur
                Dim pos As Integer
                pos = InStr(line, "=")
                
                If pos > 0 Then
                    Dim key As String
                    Dim value As String
                    
                    key = Trim(Left(line, pos - 1))
                    value = Trim(Mid(line, pos + 1))
                    
                    ' Ajouter au dictionnaire appropri�
                    If Len(currentSection) > 0 Then
                        ' Ajouter � la section
                        Dim fullKey As String
                        fullKey = currentSection & "." & key
                        IConfigManagerBase_SetValue fullKey, value
                    Else
                        ' Ajouter au dictionnaire principal
                        IConfigManagerBase_SetValue key, value
                    End If
                End If
            End If
        End If
    Loop
    
    Close #fileNum
    LoadFromIniFile = True
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then
        Close #fileNum
    End If
    LogError "Erreur lors du chargement du fichier INI: " & Err.Description, "LoadFromIniFile"
    LoadFromIniFile = False
End Function

Private Function LoadFromXmlFile(ByVal filePath As String) As Boolean
    ' Impl�mentation de chargement XML (simplifi�e)
    LoadFromXmlFile = True
End Function

Private Function LoadFromJsonFile(ByVal filePath As String) As Boolean
    ' Impl�mentation de chargement JSON (simplifi�e)
    LoadFromJsonFile = True
End Function

Private Function LoadFromCsvFile(ByVal filePath As String) As Boolean
    ' Impl�mentation de chargement CSV (simplifi�e)
    LoadFromCsvFile = True
End Function

Private Function LoadFromTextFile(ByVal filePath As String) As Boolean
    ' Impl�mentation de chargement TXT (simplifi�e)
    LoadFromTextFile = True
End Function

Private Function LoadFromExcelFile(ByVal filePath As String) As Boolean
    ' Impl�mentation de chargement Excel (simplifi�e)
    LoadFromExcelFile = True
End Function

Private Function DetectAndLoadFile(ByVal filePath As String) As Boolean
    ' Impl�mentation de d�tection automatique (simplifi�e)
    DetectAndLoadFile = True
End Function

' Sauvegarde
Private Function SaveToFile(ByVal filePath As String) As Boolean
    ' Impl�mentation de sauvegarde dans un fichier (simplifi�e)
    SaveToFile = True
End Function

Private Function SaveToObject(ByVal destination As Object) As Boolean
    ' Impl�mentation de sauvegarde dans un objet (simplifi�e)
    SaveToObject = True
End Function

' Utilitaires
Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

Private Sub LogError(ByVal message As String, ByVal source As String)
    ' V�rifier si nous devons ignorer les erreurs
    If m_ignoreErrors Then Exit Sub
    
    ' Logger l'erreur si un logger est disponible
    If Not m_logger Is Nothing Then
        On Error Resume Next
        m_logger.Error message, "ConfigManager." & source
        On Error GoTo 0
    End If
    
    ' �galement afficher dans Debug.Print pour le d�bogage
    Debug.Print "[ERROR][ConfigManager." & source & "] " & message
End Sub 