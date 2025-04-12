Attribute VB_Name = "modApexBridge"
'@Folder("Interoperabilite.Integration")
'@ModuleDescription("Module d'int�gration entre le framework APEX et l'interop�rabilit� Excel")
Option Explicit
' ==========================================================================
' Module  : modApexBridge
' Version : 1.0
' Purpose : Pont d'int�gration entre le framework APEX et l'infrastructure Excel/VBA
' Author  : APEX Framework Team
' Date    : 2024-04-11
' ==========================================================================

' Constantes globales
Private Const MODULE_NAME As String = "modApexBridge"
Private Const ERR_CONTAINER_INIT As Long = vbObjectError + 3001
Private Const ERR_ADAPTER_INIT As Long = vbObjectError + 3002
Private Const ERR_CONVERTER_INIT As Long = vbObjectError + 3003
Private Const ERR_CONFIG_MISSING As Long = vbObjectError + 3301
Private Const ERR_FACTORY_FAILED As Long = vbObjectError + 3302
Private Const ERR_ADAPTER_FAILED As Long = vbObjectError + 3303
Private Const CONFIG_SECTION As String = "INTEROP"
Private Const CONFIG_SHEET_NAME As String = "Config"
Private Const DEFAULT_LOG_LEVEL As String = "INFO"

' Cl�s du conteneur pour les services standards
Public Const SVC_LOGGER As String = "ILogger"
Public Const SVC_CONFIG As String = "IConfigLoader"
Public Const SVC_DB_CONNECTION As String = "IDbConnection"
Public Const SVC_EXCEL_APP As String = "ExcelApplication"
Public Const SVC_EXCEL_WORKBOOK_ACCESS As String = "IWorkbookAccessor"
Public Const SVC_EXCEL_SHEET_ACCESS As String = "ISheetAccessor"
Public Const SVC_EXCEL_RANGE_ACCESS As String = "IRangeAccessor"
Public Const SVC_UNIT_OF_WORK As String = "IUnitOfWork"

' Types de composants APEX
Public Enum ApexComponentType
    DbComponent = 1
    UiComponent = 2
    LoggingComponent = 3
    ConfigComponent = 4
    UtilityComponent = 5
    CustomComponent = 10
End Enum

' Structure pour mapper des interfaces
Private Type InterfaceMapping
    SourceInterface As String
    TargetInterface As String
    AdapterFactory As String
    ConfigKey As String
End Type

' Variables globales
Private g_container As Object          ' clsDependencyContainer
Private g_logger As Object             ' ILoggerBase
Private g_configMappings As Object     ' Scripting.Dictionary
Private g_interfaceMappings As Collection ' Collection de mappings d'interfaces

' �tat d'initialisation
Private m_initialized As Boolean

' Variables priv�es
Private m_container As clsDependencyContainer
Private m_logger As Object
Private m_unitOfWork As clsUnitOfWork
Private m_configLoader As Object
Private m_isInitialized As Boolean
Private m_excelApp As Object
Private m_lastError As String

' ==========================================================================
' INITIALISATION ET CONFIGURATION
' ==========================================================================

'@Description("Initialise le syst�me d'int�gration APEX-Excel")
Public Function InitializeSystem(Optional ByVal useMocks As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    ' Si d�j� initialis�, renvoyer vrai
    If m_isInitialized Then
        Debug.Print "[INFO] [" & MODULE_NAME & "] Le syst�me est d�j� initialis�"
        InitializeSystem = True
        Exit Function
    End If
    
    ' Cr�er et initialiser le conteneur de d�pendances
    Set m_container = New clsDependencyContainer
    If Not m_container.Initialize() Then
        m_lastError = "�chec de l'initialisation du conteneur de d�pendances: " & m_container.LastError
        Debug.Print "[ERROR] [" & MODULE_NAME & "] " & m_lastError
        InitializeSystem = False
        Exit Function
    End If
    
    ' Enregistrer les services de base
    If Not RegisterCoreServices(useMocks) Then
        m_lastError = "�chec de l'enregistrement des services de base"
        Debug.Print "[ERROR] [" & MODULE_NAME & "] " & m_lastError
        InitializeSystem = False
        Exit Function
    End If
    
    ' Enregistrer les adaptateurs pour APEX
    If Not RegisterApexAdapters(useMocks) Then
        m_lastError = "�chec de l'enregistrement des adaptateurs APEX"
        Debug.Print "[ERROR] [" & MODULE_NAME & "] " & m_lastError
        InitializeSystem = False
        Exit Function
    End If
    
    ' Enregistrer les adaptateurs pour Excel
    If Not RegisterExcelAdapters(useMocks) Then
        m_lastError = "�chec de l'enregistrement des adaptateurs Excel"
        Debug.Print "[ERROR] [" & MODULE_NAME & "] " & m_lastError
        InitializeSystem = False
        Exit Function
    End If
    
    ' Marquer comme initialis�
    m_isInitialized = True
    Debug.Print "[INFO] [" & MODULE_NAME & "] Syst�me d'int�gration initialis� avec succ�s"
    
    InitializeSystem = True
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de l'initialisation du syst�me: " & Err.Description
    Debug.Print "[ERROR] [" & MODULE_NAME & "] " & m_lastError
    InitializeSystem = False
End Function

'@Description("Enregistre les services de base dans le conteneur")
Private Function RegisterCoreServices(ByVal useMocks As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    ' Logger - cr�er selon le mode (mock ou r�el)
    If useMocks Then
        ' Utiliser un logger mock
        Dim mockLogger As New clsMockLogger
        If Not mockLogger.Initialize() Then
            m_lastError = "�chec de l'initialisation du logger mock"
            RegisterCoreServices = False
            Exit Function
        End If
        If Not m_container.RegisterInstance(SVC_LOGGER, mockLogger) Then
            m_lastError = "�chec de l'enregistrement du logger mock: " & m_container.LastError
            RegisterCoreServices = False
            Exit Function
        End If
    Else
        ' Utiliser le logger APEX r�el
        ' Cr�er une fabrique pour le logger APEX
        If Not m_container.RegisterFactory(SVC_LOGGER, "modLoggerFactory", "CreateLogger") Then
            m_lastError = "�chec de l'enregistrement de la fabrique de logger: " & m_container.LastError
            RegisterCoreServices = False
            Exit Function
        End If
    End If
    
    ' Chargeur de configuration
    If useMocks Then
        ' Utiliser un chargeur de configuration mock
        Dim mockConfig As New clsMockConfigLoader
        If Not mockConfig.Initialize() Then
            m_lastError = "�chec de l'initialisation du chargeur de configuration mock"
            RegisterCoreServices = False
            Exit Function
        End If
        If Not m_container.RegisterInstance(SVC_CONFIG, mockConfig) Then
            m_lastError = "�chec de l'enregistrement du chargeur de configuration mock: " & m_container.LastError
            RegisterCoreServices = False
            Exit Function
        End If
    Else
        ' Utiliser le chargeur de configuration APEX r�el
        ' Cr�er une fabrique pour le chargeur de configuration APEX
        If Not m_container.RegisterFactory(SVC_CONFIG, "modConfigFactory", "CreateConfigLoader") Then
            m_lastError = "�chec de l'enregistrement de la fabrique de configuration: " & m_container.LastError
            RegisterCoreServices = False
            Exit Function
        End If
    End If
    
    ' Unit� de travail
    Dim unitOfWork As New clsUnitOfWork
    ' R�cup�rer le logger pour l'initialisation
    Dim logger As Object
    Set logger = m_container.ResolveService(SVC_LOGGER)
    If logger Is Nothing Then
        m_lastError = "Impossible de r�soudre le service logger pour l'unit� de travail"
        RegisterCoreServices = False
        Exit Function
    End If
    
    ' Initialiser l'unit� de travail avec le logger
    If Not unitOfWork.Initialize(logger) Then
        m_lastError = "�chec de l'initialisation de l'unit� de travail: " & unitOfWork.LastError
        RegisterCoreServices = False
        Exit Function
    End If
    
    ' Enregistrer l'unit� de travail
    If Not m_container.RegisterInstance(SVC_UNIT_OF_WORK, unitOfWork) Then
        m_lastError = "�chec de l'enregistrement de l'unit� de travail: " & m_container.LastError
        RegisterCoreServices = False
        Exit Function
    End If
    
    RegisterCoreServices = True
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de l'enregistrement des services de base: " & Err.Description
    RegisterCoreServices = False
End Function

'@Description("Enregistre les adaptateurs pour APEX dans le conteneur")
Private Function RegisterApexAdapters(ByVal useMocks As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    ' Connexion � la base de donn�es
    If useMocks Then
        ' Utiliser une connexion DB mock
        Dim mockDbConnection As New clsMockDbConnection
        If Not mockDbConnection.Initialize() Then
            m_lastError = "�chec de l'initialisation de la connexion DB mock"
            RegisterApexAdapters = False
            Exit Function
        End If
        If Not m_container.RegisterInstance(SVC_DB_CONNECTION, mockDbConnection) Then
            m_lastError = "�chec de l'enregistrement de la connexion DB mock: " & m_container.LastError
            RegisterApexAdapters = False
            Exit Function
        End If
    Else
        ' Utiliser la connexion DB APEX r�elle
        ' Cr�er une fabrique pour la connexion DB APEX
        If Not m_container.RegisterFactory(SVC_DB_CONNECTION, "modDbConnFactory", "CreateDbConnection") Then
            m_lastError = "�chec de l'enregistrement de la fabrique de connexion DB: " & m_container.LastError
            RegisterApexAdapters = False
            Exit Function
        End If
    End If
    
    ' Ajouter d'autres adaptateurs APEX selon les besoins...
    
    RegisterApexAdapters = True
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de l'enregistrement des adaptateurs APEX: " & Err.Description
    RegisterApexAdapters = False
End Function

'@Description("Enregistre les adaptateurs pour Excel dans le conteneur")
Private Function RegisterExcelAdapters(ByVal useMocks As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    ' Application Excel
    If useMocks Then
        ' Utiliser une application Excel mock
        If Not m_container.RegisterFactory(SVC_EXCEL_APP, "modExcelMockFactory", "CreateExcelApplication") Then
            m_lastError = "�chec de l'enregistrement de la fabrique d'application Excel mock: " & m_container.LastError
            RegisterExcelAdapters = False
            Exit Function
        End If
    Else
        ' Utiliser l'application Excel r�elle
        If Not m_container.RegisterFactory(SVC_EXCEL_APP, "modExcelFactory", "CreateExcelApplication") Then
            m_lastError = "�chec de l'enregistrement de la fabrique d'application Excel: " & m_container.LastError
            RegisterExcelAdapters = False
            Exit Function
        End If
    End If
    
    ' Accesseur de classeur
    If useMocks Then
        ' Utiliser un accesseur de classeur mock
        If Not m_container.RegisterFactory(SVC_EXCEL_WORKBOOK_ACCESS, "modExcelMockFactory", "CreateWorkbookAccessor") Then
            m_lastError = "�chec de l'enregistrement de la fabrique d'accesseur de classeur mock: " & m_container.LastError
            RegisterExcelAdapters = False
            Exit Function
        End If
    Else
        ' Utiliser l'accesseur de classeur r�el
        If Not m_container.RegisterFactory(SVC_EXCEL_WORKBOOK_ACCESS, "modExcelFactory", "CreateWorkbookAccessor") Then
            m_lastError = "�chec de l'enregistrement de la fabrique d'accesseur de classeur: " & m_container.LastError
            RegisterExcelAdapters = False
            Exit Function
        End If
    End If
    
    ' Accesseur de feuille
    If useMocks Then
        ' Utiliser un accesseur de feuille mock
        If Not m_container.RegisterFactory(SVC_EXCEL_SHEET_ACCESS, "modExcelMockFactory", "CreateSheetAccessor") Then
            m_lastError = "�chec de l'enregistrement de la fabrique d'accesseur de feuille mock: " & m_container.LastError
            RegisterExcelAdapters = False
            Exit Function
        End If
    Else
        ' Utiliser l'accesseur de feuille r�el
        ' Utiliser l'accesseur avec cache si disponible
        If Not m_container.RegisterFactory(SVC_EXCEL_SHEET_ACCESS, "modExcelFactory", "CreateCachedSheetAccessor") Then
            m_lastError = "�chec de l'enregistrement de la fabrique d'accesseur de feuille: " & m_container.LastError
            RegisterExcelAdapters = False
            Exit Function
        End If
    End If
    
    ' Accesseur de plage
    If useMocks Then
        ' Utiliser un accesseur de plage mock
        If Not m_container.RegisterFactory(SVC_EXCEL_RANGE_ACCESS, "modExcelMockFactory", "CreateRangeAccessor") Then
            m_lastError = "�chec de l'enregistrement de la fabrique d'accesseur de plage mock: " & m_container.LastError
            RegisterExcelAdapters = False
            Exit Function
        End If
    Else
        ' Utiliser l'accesseur de plage r�el
        If Not m_container.RegisterFactory(SVC_EXCEL_RANGE_ACCESS, "modExcelFactory", "CreateRangeAccessor") Then
            m_lastError = "�chec de l'enregistrement de la fabrique d'accesseur de plage: " & m_container.LastError
            RegisterExcelAdapters = False
            Exit Function
        End If
    End If
    
    RegisterExcelAdapters = True
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de l'enregistrement des adaptateurs Excel: " & Err.Description
    RegisterExcelAdapters = False
End Function

'----------------------------------------------------------------------------------------
' Acc�s aux services
'----------------------------------------------------------------------------------------

'@Description("Renvoie une r�f�rence au conteneur de d�pendances")
Public Function GetContainer() As clsDependencyContainer
    If Not m_isInitialized Then
        Debug.Print "[WARNING] [" & MODULE_NAME & "] Tentative d'acc�s au conteneur non initialis�"
        InitializeSystem False
    End If
    
    Set GetContainer = m_container
End Function

'@Description("R�sout un service � partir de sa cl�")
Public Function GetService(ByVal serviceKey As String) As Object
    On Error GoTo ErrorHandler
    
    If Not m_isInitialized Then
        Debug.Print "[WARNING] [" & MODULE_NAME & "] Tentative de r�solution de service sans initialisation"
        If Not InitializeSystem(False) Then
            m_lastError = "Impossible d'initialiser le syst�me pour r�soudre le service"
            Set GetService = Nothing
            Exit Function
        End If
    End If
    
    ' R�soudre le service
    Dim service As Object
    Set service = m_container.ResolveService(serviceKey)
    
    If service Is Nothing Then
        m_lastError = "Service non trouv�: " & serviceKey & " - " & m_container.LastError
        Debug.Print "[WARNING] [" & MODULE_NAME & "] " & m_lastError
    End If
    
    Set GetService = service
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de la r�solution du service " & serviceKey & ": " & Err.Description
    Debug.Print "[ERROR] [" & MODULE_NAME & "] " & m_lastError
    Set GetService = Nothing
End Function

'@Description("Obtient le logger")
Public Function GetLogger() As Object
    Set GetLogger = GetService(SVC_LOGGER)
End Function

'@Description("Obtient le chargeur de configuration")
Public Function GetConfigLoader() As Object
    Set GetConfigLoader = GetService(SVC_CONFIG)
End Function

'@Description("Obtient la connexion � la base de donn�es")
Public Function GetDbConnection() As Object
    Set GetDbConnection = GetService(SVC_DB_CONNECTION)
End Function

'@Description("Obtient l'accesseur d'application Excel")
Public Function GetExcelApplication() As Object
    Set GetExcelApplication = GetService(SVC_EXCEL_APP)
End Function

'@Description("Obtient l'accesseur de classeur Excel")
Public Function GetWorkbookAccessor() As Object
    Set GetWorkbookAccessor = GetService(SVC_EXCEL_WORKBOOK_ACCESS)
End Function

'@Description("Obtient l'accesseur de feuille Excel")
Public Function GetSheetAccessor() As Object
    Set GetSheetAccessor = GetService(SVC_EXCEL_SHEET_ACCESS)
End Function

'@Description("Obtient l'accesseur de plage Excel")
Public Function GetRangeAccessor() As Object
    Set GetRangeAccessor = GetService(SVC_EXCEL_RANGE_ACCESS)
End Function

'@Description("Obtient l'unit� de travail")
Public Function GetUnitOfWork() As clsUnitOfWork
    Set GetUnitOfWork = GetService(SVC_UNIT_OF_WORK)
End Function

'----------------------------------------------------------------------------------------
' Gestion des transactions
'----------------------------------------------------------------------------------------

'@Description("D�marre une nouvelle transaction")
Public Function BeginTransaction() As Boolean
    On Error GoTo ErrorHandler
    
    ' V�rifier l'initialisation
    If Not m_isInitialized Then
        If Not InitializeSystem(False) Then
            m_lastError = "Impossible d'initialiser le syst�me pour d�marrer une transaction"
            BeginTransaction = False
            Exit Function
        End If
    End If
    
    ' Obtenir l'unit� de travail
    Dim unitOfWork As clsUnitOfWork
    Set unitOfWork = GetUnitOfWork()
    
    If unitOfWork Is Nothing Then
        m_lastError = "Unit� de travail non disponible"
        BeginTransaction = False
        Exit Function
    End If
    
    ' D�marrer la transaction
    BeginTransaction = unitOfWork.BeginTransaction()
    If Not BeginTransaction Then
        m_lastError = "�chec du d�marrage de la transaction: " & unitOfWork.LastError
    End If
    
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors du d�marrage de la transaction: " & Err.Description
    BeginTransaction = False
End Function

'@Description("Valide la transaction courante")
Public Function CommitTransaction() As Boolean
    On Error GoTo ErrorHandler
    
    ' V�rifier l'initialisation
    If Not m_isInitialized Then
        m_lastError = "Le syst�me n'est pas initialis�"
        CommitTransaction = False
        Exit Function
    End If
    
    ' Obtenir l'unit� de travail
    Dim unitOfWork As clsUnitOfWork
    Set unitOfWork = GetUnitOfWork()
    
    If unitOfWork Is Nothing Then
        m_lastError = "Unit� de travail non disponible"
        CommitTransaction = False
        Exit Function
    End If
    
    ' Valider la transaction
    CommitTransaction = unitOfWork.CommitTransaction()
    If Not CommitTransaction Then
        m_lastError = "�chec de la validation de la transaction: " & unitOfWork.LastError
    End If
    
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de la validation de la transaction: " & Err.Description
    CommitTransaction = False
End Function

'@Description("Annule la transaction courante")
Public Function RollbackTransaction() As Boolean
    On Error GoTo ErrorHandler
    
    ' V�rifier l'initialisation
    If Not m_isInitialized Then
        m_lastError = "Le syst�me n'est pas initialis�"
        RollbackTransaction = False
        Exit Function
    End If
    
    ' Obtenir l'unit� de travail
    Dim unitOfWork As clsUnitOfWork
    Set unitOfWork = GetUnitOfWork()
    
    If unitOfWork Is Nothing Then
        m_lastError = "Unit� de travail non disponible"
        RollbackTransaction = False
        Exit Function
    End If
    
    ' Annuler la transaction
    RollbackTransaction = unitOfWork.RollbackTransaction()
    If Not RollbackTransaction Then
        m_lastError = "�chec de l'annulation de la transaction: " & unitOfWork.LastError
    End If
    
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de l'annulation de la transaction: " & Err.Description
    RollbackTransaction = False
End Function

'----------------------------------------------------------------------------------------
' Utilitaires
'----------------------------------------------------------------------------------------

'@Description("Retourne la derni�re erreur survenue")
Public Property Get LastError() As String
    LastError = m_lastError
End Property

'@Description("Indique si le syst�me est initialis�")
Public Property Get IsInitialized() As Boolean
    IsInitialized = m_isInitialized
End Property

'@Description("R�initialise le syst�me (� utiliser en d�veloppement)")
Public Sub ResetSystem()
    On Error Resume Next
    
    If Not m_container Is Nothing Then
        m_container.ClearAllServices
    End If
    
    Set m_container = Nothing
    m_isInitialized = False
    m_lastError = ""
    
    Debug.Print "[INFO] [" & MODULE_NAME & "] Syst�me r�initialis�"
End Sub

'----------------------------------------------------------------------------------------
' Conversion de donn�es APEX <-> Excel
'----------------------------------------------------------------------------------------

'@Description("Convertit une plage Excel en recordset APEX")
Public Function ConvertRangeToRecordset(ByVal worksheetName As String, ByVal rangeAddress As String, _
                                       Optional ByVal hasHeaders As Boolean = True) As Object
    On Error GoTo ErrorHandler
    
    ' V�rifier l'initialisation
    If Not m_isInitialized Then
        If Not InitializeSystem(False) Then
            m_lastError = "Impossible d'initialiser le syst�me pour la conversion"
            Set ConvertRangeToRecordset = Nothing
            Exit Function
        End If
    End If
    
    ' Obtenir l'accesseur de plage
    Dim rangeAccessor As Object
    Set rangeAccessor = GetRangeAccessor()
    
    If rangeAccessor Is Nothing Then
        m_lastError = "Accesseur de plage non disponible"
        Set ConvertRangeToRecordset = Nothing
        Exit Function
    End If
    
    ' Obtenir les donn�es de la plage
    Dim rangeData As Variant
    rangeData = rangeAccessor.GetRangeData(worksheetName, rangeAddress)
    
    If IsEmpty(rangeData) Then
        m_lastError = "Aucune donn�e trouv�e dans la plage sp�cifi�e"
        Set ConvertRangeToRecordset = Nothing
        Exit Function
    End If
    
    ' Cr�er un recordset ADODB
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    ' D�terminer les en-t�tes
    Dim fieldNames() As String
    Dim startRow As Long
    
    If hasHeaders Then
        ' Utiliser la premi�re ligne comme en-t�tes
        startRow = 2 ' Commencer � la deuxi�me ligne pour les donn�es
        Dim colCount As Long
        colCount = UBound(rangeData, 2)
        ReDim fieldNames(1 To colCount)
        
        Dim i As Long
        For i = 1 To colCount
            If IsEmpty(rangeData(1, i)) Or IsNull(rangeData(1, i)) Then
                fieldNames(i) = "Column" & i
            Else
                fieldNames(i) = CStr(rangeData(1, i))
            End If
        Next i
    Else
        ' Pas d'en-t�tes, utiliser des noms g�n�riques
        startRow = 1 ' Commencer � la premi�re ligne pour les donn�es
        Dim colCount2 As Long
        colCount2 = UBound(rangeData, 2)
        ReDim fieldNames(1 To colCount2)
        
        Dim j As Long
        For j = 1 To colCount2
            fieldNames(j) = "Column" & j
        Next j
    End If
    
    ' Cr�er les champs dans le recordset
    Dim colIdx As Long
    For colIdx = 1 To UBound(fieldNames)
        rs.Fields.Append fieldNames(colIdx), adVariant
    Next colIdx
    
    ' Ouvrir le recordset
    rs.Open
    
    ' Ajouter les donn�es au recordset
    Dim rowIdx As Long
    For rowIdx = startRow To UBound(rangeData, 1)
        rs.AddNew
        
        For colIdx = 1 To UBound(fieldNames)
            rs.Fields(colIdx - 1).Value = rangeData(rowIdx, colIdx)
        Next colIdx
        
        rs.Update
    Next rowIdx
    
    ' D�placer au premier enregistrement
    If Not rs.EOF Then
        rs.MoveFirst
    End If
    
    Set ConvertRangeToRecordset = rs
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de la conversion de la plage en recordset: " & Err.Description
    Set ConvertRangeToRecordset = Nothing
End Function

'@Description("Convertit un recordset APEX en donn�es pour une plage Excel")
Public Function ConvertRecordsetToRange(ByVal rs As Object, Optional ByVal includeHeaders As Boolean = True) As Variant
    On Error GoTo ErrorHandler
    
    ' V�rifier si le recordset est valide
    If rs Is Nothing Then
        m_lastError = "Recordset non valide"
        ConvertRecordsetToRange = Empty
        Exit Function
    End If
    
    ' V�rifier si le recordset est ouvert
    If rs.State <> 1 Then ' adStateOpen
        m_lastError = "Le recordset n'est pas ouvert"
        ConvertRecordsetToRange = Empty
        Exit Function
    End If
    
    ' Compter le nombre de lignes
    Dim rowCount As Long
    rowCount = 0
    
    ' Sauvegarder la position actuelle
    Dim currentPosition As Variant
    If Not rs.BOF And Not rs.EOF Then
        currentPosition = rs.Bookmark
    End If
    
    ' Aller au d�but
    If Not rs.BOF Then
        rs.MoveFirst
    End If
    
    ' Compter les lignes
    Do Until rs.EOF
        rowCount = rowCount + 1
        rs.MoveNext
    Loop
    
    ' Retourner � la position d'origine
    If rowCount > 0 And Not IsEmpty(currentPosition) Then
        rs.Bookmark = currentPosition
    End If
    
    ' Si le recordset est vide, renvoyer un tableau vide
    If rowCount = 0 Then
        Dim emptyArray() As Variant
        ReDim emptyArray(1 To 1, 1 To 1)
        emptyArray(1, 1) = ""
        ConvertRecordsetToRange = emptyArray
        Exit Function
    End If
    
    ' Obtenir le nombre de colonnes
    Dim colCount As Long
    colCount = rs.Fields.Count
    
    ' Cr�er le tableau pour les donn�es
    Dim dataArray() As Variant
    
    If includeHeaders Then
        ' Inclure une ligne suppl�mentaire pour les en-t�tes
        ReDim dataArray(1 To rowCount + 1, 1 To colCount)
        
        ' Ajouter les en-t�tes
        Dim colIdx As Long
        For colIdx = 1 To colCount
            dataArray(1, colIdx) = rs.Fields(colIdx - 1).Name
        Next colIdx
        
        ' Ajouter les donn�es
        rs.MoveFirst
        Dim rowIdx As Long
        rowIdx = 2 ' Commencer � la deuxi�me ligne (apr�s les en-t�tes)
        
        Do Until rs.EOF
            For colIdx = 1 To colCount
                ' V�rifier si la valeur est Null
                If IsNull(rs.Fields(colIdx - 1).Value) Then
                    dataArray(rowIdx, colIdx) = ""
                Else
                    dataArray(rowIdx, colIdx) = rs.Fields(colIdx - 1).Value
                End If
            Next colIdx
            
            rowIdx = rowIdx + 1
            rs.MoveNext
        Loop
    Else
        ' Pas d'en-t�tes, juste les donn�es
        ReDim dataArray(1 To rowCount, 1 To colCount)
        
        ' Ajouter les donn�es
        rs.MoveFirst
        Dim rowIdx2 As Long
        rowIdx2 = 1
        
        Do Until rs.EOF
            Dim colIdx2 As Long
            For colIdx2 = 1 To colCount
                ' V�rifier si la valeur est Null
                If IsNull(rs.Fields(colIdx2 - 1).Value) Then
                    dataArray(rowIdx2, colIdx2) = ""
                Else
                    dataArray(rowIdx2, colIdx2) = rs.Fields(colIdx2 - 1).Value
                End If
            Next colIdx2
            
            rowIdx2 = rowIdx2 + 1
            rs.MoveNext
        Loop
    End If
    
    ' Retourner le tableau
    ConvertRecordsetToRange = dataArray
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de la conversion du recordset en plage: " & Err.Description
    ConvertRecordsetToRange = Empty
End Function

'@Description("Convertit un dictionnaire APEX en tableau Excel")
Public Function ConvertDictionaryToArray(ByVal dict As Object, Optional ByVal includeKeys As Boolean = True) As Variant
    On Error GoTo ErrorHandler
    
    ' V�rifier si le dictionnaire est valide
    If dict Is Nothing Then
        m_lastError = "Dictionnaire non valide"
        ConvertDictionaryToArray = Empty
        Exit Function
    End If
    
    ' V�rifier s'il y a des �l�ments
    If dict.Count = 0 Then
        Dim emptyArray() As Variant
        ReDim emptyArray(1 To 1, 1 To 1)
        emptyArray(1, 1) = ""
        ConvertDictionaryToArray = emptyArray
        Exit Function
    End If
    
    ' Cr�er le tableau pour les donn�es
    Dim dataArray() As Variant
    
    If includeKeys Then
        ' Deux colonnes : cl� et valeur
        ReDim dataArray(1 To dict.Count + 1, 1 To 2)
        
        ' Ajouter les en-t�tes
        dataArray(1, 1) = "Key"
        dataArray(1, 2) = "Value"
        
        ' Ajouter les donn�es
        Dim keys As Variant
        keys = dict.keys
        
        Dim i As Long
        For i = 0 To dict.Count - 1
            dataArray(i + 2, 1) = keys(i)
            
            ' V�rifier si la valeur est Null
            If IsNull(dict(keys(i))) Then
                dataArray(i + 2, 2) = ""
            Else
                dataArray(i + 2, 2) = dict(keys(i))
            End If
        Next i
    Else
        ' Une seule colonne pour les valeurs
        ReDim dataArray(1 To dict.Count, 1 To 1)
        
        ' Ajouter les donn�es
        Dim items As Variant
        items = dict.items
        
        Dim j As Long
        For j = 0 To dict.Count - 1
            ' V�rifier si la valeur est Null
            If IsNull(items(j)) Then
                dataArray(j + 1, 1) = ""
            Else
                dataArray(j + 1, 1) = items(j)
            End If
        Next j
    End If
    
    ' Retourner le tableau
    ConvertDictionaryToArray = dataArray
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de la conversion du dictionnaire en tableau: " & Err.Description
    ConvertDictionaryToArray = Empty
End Function

'@Description("Convertit un tableau Excel en dictionnaire APEX")
Public Function ConvertArrayToDictionary(ByVal dataArray As Variant, _
                                        Optional ByVal keyColumnIndex As Long = 1, _
                                        Optional ByVal valueColumnIndex As Long = 2, _
                                        Optional ByVal hasHeaders As Boolean = True) As Object
    On Error GoTo ErrorHandler
    
    ' V�rifier si le tableau est valide
    If Not IsArray(dataArray) Then
        m_lastError = "Tableau non valide"
        Set ConvertArrayToDictionary = Nothing
        Exit Function
    End If
    
    ' Cr�er un nouveau dictionnaire
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' D�terminer la premi�re ligne de donn�es
    Dim startRow As Long
    If hasHeaders Then
        startRow = 2 ' Commencer � la deuxi�me ligne (apr�s les en-t�tes)
    Else
        startRow = 1 ' Commencer � la premi�re ligne
    End If
    
    ' Ajouter les donn�es au dictionnaire
    Dim i As Long
    For i = startRow To UBound(dataArray, 1)
        ' V�rifier si les indices sont valides
        If keyColumnIndex > 0 And keyColumnIndex <= UBound(dataArray, 2) And _
           valueColumnIndex > 0 And valueColumnIndex <= UBound(dataArray, 2) Then
            
            ' V�rifier si la cl� existe d�j�
            If Not dict.Exists(dataArray(i, keyColumnIndex)) Then
                dict.Add dataArray(i, keyColumnIndex), dataArray(i, valueColumnIndex)
            End If
        Else
            m_lastError = "Indices de colonnes non valides"
            Set dict = Nothing
            Set ConvertArrayToDictionary = Nothing
            Exit Function
        End If
    Next i
    
    ' Retourner le dictionnaire
    Set ConvertArrayToDictionary = dict
    Exit Function
    
ErrorHandler:
    m_lastError = "Erreur lors de la conversion du tableau en dictionnaire: " & Err.Description
    Set ConvertArrayToDictionary = Nothing
End Function 