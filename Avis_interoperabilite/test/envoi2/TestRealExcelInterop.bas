Attribute VB_Name = "TestRealExcelInterop"
Option Explicit

' ============================================================================
' Module de Test pour la Validation Post-Refactoring APEX Framework
' Référence: APEX-VAL-RUBY-001
' Date: 2025-04-30
' ============================================================================

' Variables privées pour ce module de test
Private mAppContext As IAppContext
Private mTestSheetName As String

' ============================================================================
' Fonctions de test appelées par le script Ruby de validation
' ============================================================================

' Test de l'initialisation d'IAppContext
Public Function TestIAppContextInitialization() As Boolean
    On Error GoTo ErrorHandler
    
    ' Initialiser le contexte applicatif
    Set mAppContext = CreateDefaultAppContext()
    mTestSheetName = "TestSheet"
    
    ' Vérifier que l'initialisation est correcte
    TestIAppContextInitialization = Not (mAppContext Is Nothing)
    
    ' Configuration de base pour les tests
    If Not mAppContext Is Nothing Then
        ' Désactiver les mocks pour utiliser les implémentations réelles
        mAppContext.SetConfigValue "ExcelFactory.UseMocks", False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestIAppContextInitialization: " & Err.Description
    TestIAppContextInitialization = False
End Function

' Test de l'accès à ExcelFactory via IAppContext
Public Function TestExcelFactoryAccess() As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que le contexte est initialisé
    If mAppContext Is Nothing Then
        If Not TestIAppContextInitialization() Then
            TestExcelFactoryAccess = False
            Exit Function
        End If
    End If
    
    ' Obtenir une référence à la factory Excel via IAppContext
    Dim excelFactory As clsExcelFactory
    Set excelFactory = mAppContext.GetExcelFactory()
    
    ' Vérifier que la factory a été correctement obtenue
    If excelFactory Is Nothing Then
        TestExcelFactoryAccess = False
        Exit Function
    End If
    
    ' Test simple : créer un RangeAccessor sur une cellule
    Dim testRange As Range
    Set testRange = ActiveSheet.Range("A1")
    
    Dim rangeAccessor As IExcelRangeAccessor
    Set rangeAccessor = excelFactory.CreateRangeAccessor(testRange, mAppContext)
    
    ' Vérifier que l'accesseur a été créé correctement
    TestExcelFactoryAccess = Not (rangeAccessor Is Nothing)
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestExcelFactoryAccess: " & Err.Description
    TestExcelFactoryAccess = False
End Function

' Test de l'accès au Logger via IAppContext
Public Function TestLoggerAccess() As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que le contexte est initialisé
    If mAppContext Is Nothing Then
        If Not TestIAppContextInitialization() Then
            TestLoggerAccess = False
            Exit Function
        End If
    End If
    
    ' Obtenir une référence au logger via IAppContext
    Dim logger As ILoggerBase
    Set logger = mAppContext.GetLogger()
    
    ' Vérifier que le logger a été correctement obtenu
    If logger Is Nothing Then
        TestLoggerAccess = False
        Exit Function
    End If
    
    ' Test simple : journaliser un message
    logger.LogInfo "TestRealExcelInterop", "Test de validation du logger via IAppContext"
    
    ' Le test est réussi si aucune exception n'a été levée
    TestLoggerAccess = True
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestLoggerAccess: " & Err.Description
    TestLoggerAccess = False
End Function

' Test minimal du cache
Public Function TestCacheMinimal() As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que le contexte est initialisé
    If mAppContext Is Nothing Then
        If Not TestIAppContextInitialization() Then
            TestCacheMinimal = False
            Exit Function
        End If
    End If
    
    ' Obtenir une référence au cache via IAppContext
    Dim cache As ICacheProvider
    Set cache = mAppContext.GetCacheProvider()
    
    ' Vérifier que le cache a été correctement obtenu
    If cache Is Nothing Then
        TestCacheMinimal = False
        Exit Function
    End If
    
    ' Test simple : stocker et récupérer une valeur
    Dim testKey As String
    Dim testValue As String
    
    testKey = "TestKey_" & Format$(Now(), "yyyymmddhhnnss")
    testValue = "TestValue_" & Format$(Now(), "yyyymmddhhnnss")
    
    ' Stocker la valeur
    cache.Set testKey, testValue
    
    ' Récupérer la valeur
    Dim retrievedValue As Variant
    retrievedValue = cache.Get(testKey)
    
    ' Vérifier que la valeur récupérée correspond à la valeur stockée
    TestCacheMinimal = (CStr(retrievedValue) = testValue)
    
    ' Nettoyer
    cache.Remove testKey
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestCacheMinimal: " & Err.Description
    TestCacheMinimal = False
End Function

' ============================================================================
' Fonctions utilitaires
' ============================================================================

' Crée une instance par défaut de IAppContext
Private Function CreateDefaultAppContext() As IAppContext
    On Error GoTo ErrorHandler
    
    ' Créer l'instance de contexte applicatif
    Dim appContext As New clsStandardAppContext
    
    ' Initialiser avec la configuration par défaut
    appContext.Initialize
    
    ' Retourner l'instance
    Set CreateDefaultAppContext = appContext
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur dans CreateDefaultAppContext: " & Err.Description
    Set CreateDefaultAppContext = Nothing
End Function 