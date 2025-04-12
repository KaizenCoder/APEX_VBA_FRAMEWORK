VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestIntegration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Interoperabilite.Tests")
'@ModuleDescription("Tests d'intégration pour la couche d'interopérabilité")
Option Explicit

'----------------------------------------------------------------------------------------
' Constantes
'----------------------------------------------------------------------------------------
Private Const MODULE_NAME As String = "clsTestIntegration"

'----------------------------------------------------------------------------------------
' Variables privées
'----------------------------------------------------------------------------------------
Private m_logger As Object
Private m_testWorkbook As Object
Private m_testSheet As Object
Private m_testData As Variant

'----------------------------------------------------------------------------------------
' Initialisation
'----------------------------------------------------------------------------------------

'@Description("Initialise la classe de test")
Public Function Initialize(Optional ByVal logger As Object = Nothing) As Boolean
    On Error GoTo ErrorHandler
    
    ' Stocker le logger
    Set m_logger = logger
    
    ' Créer un classeur de test
    Set m_testWorkbook = Application.Workbooks.Add
    Set m_testSheet = m_testWorkbook.Sheets(1)
    
    ' Initialiser les données de test
    InitializeTestData
    
    Initialize = True
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de l'initialisation des tests: " & Err.Description
    Initialize = False
End Function

'@Description("Initialise les données de test")
Private Sub InitializeTestData()
    ' Créer un jeu de données de test
    ReDim m_testData(1 To 5, 1 To 3)
    
    ' En-têtes
    m_testData(1, 1) = "ID"
    m_testData(1, 2) = "Nom"
    m_testData(1, 3) = "Valeur"
    
    ' Données
    m_testData(2, 1) = 1
    m_testData(2, 2) = "Test1"
    m_testData(2, 3) = 100
    
    m_testData(3, 1) = 2
    m_testData(3, 2) = "Test2"
    m_testData(3, 3) = 200
    
    m_testData(4, 1) = 3
    m_testData(4, 2) = "Test3"
    m_testData(4, 3) = 300
    
    m_testData(5, 1) = 4
    m_testData(5, 2) = "Test4"
    m_testData(5, 3) = 400
End Sub

'----------------------------------------------------------------------------------------
' Tests d'intégration
'----------------------------------------------------------------------------------------

'@Description("Exécute tous les tests d'intégration")
Public Function RunAllTests() As Boolean
    On Error GoTo ErrorHandler
    
    ' Initialiser le système
    If Not modApexBridge.InitializeSystem(True) Then
        LogError "Échec de l'initialisation du système: " & modApexBridge.LastError
        RunAllTests = False
        Exit Function
    End If
    
    ' Exécuter les tests
    If Not TestDependencyContainer() Then GoTo TestFailed
    If Not TestUnitOfWork() Then GoTo TestFailed
    If Not TestExcelAccess() Then GoTo TestFailed
    If Not TestDataConversion() Then GoTo TestFailed
    If Not TestTransactionManagement() Then GoTo TestFailed
    
    ' Tous les tests ont réussi
    LogInfo "Tous les tests d'intégration ont réussi"
    RunAllTests = True
    Exit Function
    
TestFailed:
    LogError "Un ou plusieurs tests ont échoué"
    RunAllTests = False
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors de l'exécution des tests: " & Err.Description
    RunAllTests = False
End Function

'@Description("Teste le conteneur de dépendances")
Private Function TestDependencyContainer() As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir le conteneur
    Dim container As clsDependencyContainer
    Set container = modApexBridge.GetContainer()
    
    ' Vérifier que le conteneur est valide
    If container Is Nothing Then
        LogError "Le conteneur de dépendances est invalide"
        TestDependencyContainer = False
        Exit Function
    End If
    
    ' Vérifier la résolution des services principaux
    Dim logger As Object
    Set logger = container.ResolveService(SVC_LOGGER)
    If logger Is Nothing Then
        LogError "Échec de la résolution du logger"
        TestDependencyContainer = False
        Exit Function
    End If
    
    Dim config As Object
    Set config = container.ResolveService(SVC_CONFIG)
    If config Is Nothing Then
        LogError "Échec de la résolution du chargeur de configuration"
        TestDependencyContainer = False
        Exit Function
    End If
    
    Dim unitOfWork As Object
    Set unitOfWork = container.ResolveService(SVC_UNIT_OF_WORK)
    If unitOfWork Is Nothing Then
        LogError "Échec de la résolution de l'unité de travail"
        TestDependencyContainer = False
        Exit Function
    End If
    
    LogInfo "Test du conteneur de dépendances réussi"
    TestDependencyContainer = True
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors du test du conteneur: " & Err.Description
    TestDependencyContainer = False
End Function

'@Description("Teste l'unité de travail")
Private Function TestUnitOfWork() As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir l'unité de travail
    Dim unitOfWork As clsUnitOfWork
    Set unitOfWork = modApexBridge.GetUnitOfWork()
    
    ' Vérifier que l'unité de travail est valide
    If unitOfWork Is Nothing Then
        LogError "L'unité de travail est invalide"
        TestUnitOfWork = False
        Exit Function
    End If
    
    ' Tester le démarrage d'une transaction
    If Not unitOfWork.BeginTransaction() Then
        LogError "Échec du démarrage de la transaction: " & unitOfWork.LastError
        TestUnitOfWork = False
        Exit Function
    End If
    
    ' Vérifier l'état de la transaction
    If Not unitOfWork.HasActiveTransaction Then
        LogError "La transaction n'est pas active"
        TestUnitOfWork = False
        Exit Function
    End If
    
    ' Enregistrer une opération
    If Not unitOfWork.RegisterInsert("TestEntity", 1, "TestData") Then
        LogError "Échec de l'enregistrement de l'opération: " & unitOfWork.LastError
        TestUnitOfWork = False
        Exit Function
    End If
    
    ' Valider la transaction
    If Not unitOfWork.CommitTransaction() Then
        LogError "Échec de la validation de la transaction: " & unitOfWork.LastError
        TestUnitOfWork = False
        Exit Function
    End If
    
    LogInfo "Test de l'unité de travail réussi"
    TestUnitOfWork = True
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors du test de l'unité de travail: " & Err.Description
    TestUnitOfWork = False
End Function

'@Description("Teste l'accès à Excel")
Private Function TestExcelAccess() As Boolean
    On Error GoTo ErrorHandler
    
    ' Obtenir les accesseurs Excel
    Dim excelApp As Object
    Set excelApp = modApexBridge.GetExcelApplication()
    
    Dim workbookAccessor As Object
    Set workbookAccessor = modApexBridge.GetWorkbookAccessor()
    
    Dim sheetAccessor As Object
    Set sheetAccessor = modApexBridge.GetSheetAccessor()
    
    Dim rangeAccessor As Object
    Set rangeAccessor = modApexBridge.GetRangeAccessor()
    
    ' Vérifier que les accesseurs sont valides
    If excelApp Is Nothing Or workbookAccessor Is Nothing Or _
       sheetAccessor Is Nothing Or rangeAccessor Is Nothing Then
        LogError "Un ou plusieurs accesseurs Excel sont invalides"
        TestExcelAccess = False
        Exit Function
    End If
    
    ' Écrire les données de test dans la feuille
    Dim testRange As Range
    Set testRange = m_testSheet.Range("A1").Resize(UBound(m_testData, 1), UBound(m_testData, 2))
    testRange.Value = m_testData
    
    ' Lire les données via l'accesseur de plage
    Dim readData As Variant
    readData = rangeAccessor.GetRangeData(m_testSheet.Name, testRange.Address)
    
    ' Vérifier que les données lues correspondent aux données écrites
    If Not CompareArrays(m_testData, readData) Then
        LogError "Les données lues ne correspondent pas aux données écrites"
        TestExcelAccess = False
        Exit Function
    End If
    
    LogInfo "Test de l'accès Excel réussi"
    TestExcelAccess = True
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors du test de l'accès Excel: " & Err.Description
    TestExcelAccess = False
End Function

'@Description("Teste la conversion des données")
Private Function TestDataConversion() As Boolean
    On Error GoTo ErrorHandler
    
    ' Convertir les données de test en recordset
    Dim rs As Object
    Set rs = modApexBridge.ConvertRangeToRecordset(m_testSheet.Name, "A1:C5", True)
    
    ' Vérifier que le recordset est valide
    If rs Is Nothing Then
        LogError "Échec de la conversion en recordset: " & modApexBridge.LastError
        TestDataConversion = False
        Exit Function
    End If
    
    ' Convertir le recordset en tableau
    Dim convertedData As Variant
    convertedData = modApexBridge.ConvertRecordsetToRange(rs, True)
    
    ' Vérifier que les données converties correspondent aux données originales
    If Not CompareArrays(m_testData, convertedData) Then
        LogError "Les données converties ne correspondent pas aux données originales"
        TestDataConversion = False
        Exit Function
    End If
    
    ' Créer un dictionnaire de test
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Key1", "Value1"
    dict.Add "Key2", "Value2"
    
    ' Convertir le dictionnaire en tableau
    Dim dictArray As Variant
    dictArray = modApexBridge.ConvertDictionaryToArray(dict, True)
    
    ' Convertir le tableau en dictionnaire
    Dim convertedDict As Object
    Set convertedDict = modApexBridge.ConvertArrayToDictionary(dictArray, 1, 2, True)
    
    ' Vérifier que les dictionnaires correspondent
    If dict.Count <> convertedDict.Count Then
        LogError "Les dictionnaires convertis ne correspondent pas"
        TestDataConversion = False
        Exit Function
    End If
    
    LogInfo "Test de conversion des données réussi"
    TestDataConversion = True
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors du test de conversion des données: " & Err.Description
    TestDataConversion = False
End Function

'@Description("Teste la gestion des transactions")
Private Function TestTransactionManagement() As Boolean
    On Error GoTo ErrorHandler
    
    ' Démarrer une transaction
    If Not modApexBridge.BeginTransaction() Then
        LogError "Échec du démarrage de la transaction: " & modApexBridge.LastError
        TestTransactionManagement = False
        Exit Function
    End If
    
    ' Effectuer quelques opérations
    Dim unitOfWork As clsUnitOfWork
    Set unitOfWork = modApexBridge.GetUnitOfWork()
    
    If Not unitOfWork.RegisterInsert("TestEntity", 1, "TestData1") Then
        LogError "Échec de l'enregistrement de l'opération 1: " & unitOfWork.LastError
        TestTransactionManagement = False
        Exit Function
    End If
    
    If Not unitOfWork.RegisterUpdate("TestEntity", 2, "OldData", "NewData") Then
        LogError "Échec de l'enregistrement de l'opération 2: " & unitOfWork.LastError
        TestTransactionManagement = False
        Exit Function
    End If
    
    ' Valider la transaction
    If Not modApexBridge.CommitTransaction() Then
        LogError "Échec de la validation de la transaction: " & modApexBridge.LastError
        TestTransactionManagement = False
        Exit Function
    End If
    
    ' Tester le rollback
    If Not modApexBridge.BeginTransaction() Then
        LogError "Échec du démarrage de la seconde transaction: " & modApexBridge.LastError
        TestTransactionManagement = False
        Exit Function
    End If
    
    If Not unitOfWork.RegisterDelete("TestEntity", 3, "DataToDelete") Then
        LogError "Échec de l'enregistrement de l'opération 3: " & unitOfWork.LastError
        TestTransactionManagement = False
        Exit Function
    End If
    
    ' Annuler la transaction
    If Not modApexBridge.RollbackTransaction() Then
        LogError "Échec de l'annulation de la transaction: " & modApexBridge.LastError
        TestTransactionManagement = False
        Exit Function
    End If
    
    LogInfo "Test de gestion des transactions réussi"
    TestTransactionManagement = True
    Exit Function
    
ErrorHandler:
    LogError "Erreur lors du test de gestion des transactions: " & Err.Description
    TestTransactionManagement = False
End Function

'----------------------------------------------------------------------------------------
' Utilitaires
'----------------------------------------------------------------------------------------

'@Description("Compare deux tableaux 2D")
Private Function CompareArrays(ByVal arr1 As Variant, ByVal arr2 As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier les dimensions
    If UBound(arr1, 1) <> UBound(arr2, 1) Or _
       UBound(arr1, 2) <> UBound(arr2, 2) Then
        CompareArrays = False
        Exit Function
    End If
    
    ' Comparer les éléments
    Dim i As Long, j As Long
    For i = LBound(arr1, 1) To UBound(arr1, 1)
        For j = LBound(arr1, 2) To UBound(arr1, 2)
            If arr1(i, j) <> arr2(i, j) Then
                CompareArrays = False
                Exit Function
            End If
        Next j
    Next i
    
    CompareArrays = True
    Exit Function
    
ErrorHandler:
    CompareArrays = False
End Function

'----------------------------------------------------------------------------------------
' Journalisation
'----------------------------------------------------------------------------------------

'@Description("Enregistre un message d'information")
Private Sub LogInfo(ByVal message As String)
    On Error Resume Next
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo MODULE_NAME, message
    Else
        Debug.Print "[INFO] [" & MODULE_NAME & "] " & message
    End If
End Sub

'@Description("Enregistre une erreur")
Private Sub LogError(ByVal message As String)
    On Error Resume Next
    
    If Not m_logger Is Nothing Then
        m_logger.LogError MODULE_NAME, message
    Else
        Debug.Print "[ERROR] [" & MODULE_NAME & "] " & message
    End If
End Sub

'----------------------------------------------------------------------------------------
' Nettoyage
'----------------------------------------------------------------------------------------

Private Sub Class_Terminate()
    On Error Resume Next
    
    ' Fermer et supprimer le classeur de test
    If Not m_testWorkbook Is Nothing Then
        m_testWorkbook.Close SaveChanges:=False
    End If
    
    ' Nettoyer les références
    Set m_testWorkbook = Nothing
    Set m_testSheet = Nothing
    Set m_logger = Nothing
End Sub 