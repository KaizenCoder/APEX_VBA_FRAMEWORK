Attribute VB_Name = "TestRealExcelInterop"
Option Explicit

' ============================================================================
' Module de Test pour la Validation Post-Refactoring APEX Framework
' R�f�rence: APEX-VAL-RUBY-001
' Date: 2025-04-30
' ============================================================================

' Variables priv�es pour ce module de test
Private mAppContext As IAppContext
Private mTestSheetName As String

' ============================================================================
' Fonctions de test appel�es par le script Ruby de validation
' ============================================================================

' Test de l'initialisation d'IAppContext
Public Function TestIAppContextInitialization() As Boolean
    On Error GoTo ErrorHandler
    
    ' Initialiser le contexte applicatif
    Set mAppContext = CreateDefaultAppContext()
    mTestSheetName = "TestSheet"
    
    ' V�rifier que l'initialisation est correcte
    TestIAppContextInitialization = Not (mAppContext Is Nothing)
    
    ' Configuration de base pour les tests
    If Not mAppContext Is Nothing Then
        ' D�sactiver les mocks pour utiliser les impl�mentations r�elles
        mAppContext.SetConfigValue "ExcelFactory.UseMocks", False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestIAppContextInitialization: " & Err.Description
    TestIAppContextInitialization = False
End Function

' Test de l'acc�s � ExcelFactory via IAppContext
Public Function TestExcelFactoryAccess() As Boolean
    On Error GoTo ErrorHandler
    
    ' V�rifier que le contexte est initialis�
    If mAppContext Is Nothing Then
        If Not TestIAppContextInitialization() Then
            TestExcelFactoryAccess = False
            Exit Function
        End If
    End If
    
    ' Obtenir une r�f�rence � la factory Excel via IAppContext
    Dim excelFactory As clsExcelFactory
    Set excelFactory = mAppContext.GetExcelFactory()
    
    ' V�rifier que la factory a �t� correctement obtenue
    If excelFactory Is Nothing Then
        TestExcelFactoryAccess = False
        Exit Function
    End If
    
    ' Test simple : cr�er un RangeAccessor sur une cellule
    Dim testRange As Range
    Set testRange = ActiveSheet.Range("A1")
    
    Dim rangeAccessor As IExcelRangeAccessor
    Set rangeAccessor = excelFactory.CreateRangeAccessor(testRange, mAppContext)
    
    ' V�rifier que l'accesseur a �t� cr�� correctement
    TestExcelFactoryAccess = Not (rangeAccessor Is Nothing)
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestExcelFactoryAccess: " & Err.Description
    TestExcelFactoryAccess = False
End Function

' Test de l'acc�s au Logger via IAppContext
Public Function TestLoggerAccess() As Boolean
    On Error GoTo ErrorHandler
    
    ' V�rifier que le contexte est initialis�
    If mAppContext Is Nothing Then
        If Not TestIAppContextInitialization() Then
            TestLoggerAccess = False
            Exit Function
        End If
    End If
    
    ' Obtenir une r�f�rence au logger via IAppContext
    Dim logger As ILoggerBase
    Set logger = mAppContext.GetLogger()
    
    ' V�rifier que le logger a �t� correctement obtenu
    If logger Is Nothing Then
        TestLoggerAccess = False
        Exit Function
    End If
    
    ' Test simple : journaliser un message
    logger.LogInfo "TestRealExcelInterop", "Test de validation du logger via IAppContext"
    
    ' Le test est r�ussi si aucune exception n'a �t� lev�e
    TestLoggerAccess = True
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur dans TestLoggerAccess: " & Err.Description
    TestLoggerAccess = False
End Function

' Test minimal du cache
Public Function TestCacheMinimal() As Boolean
    On Error GoTo ErrorHandler
    
    ' V�rifier que le contexte est initialis�
    If mAppContext Is Nothing Then
        If Not TestIAppContextInitialization() Then
            TestCacheMinimal = False
            Exit Function
        End If
    End If
    
    ' Obtenir une r�f�rence au cache via IAppContext
    Dim cache As ICacheProvider
    Set cache = mAppContext.GetCacheProvider()
    
    ' V�rifier que le cache a �t� correctement obtenu
    If cache Is Nothing Then
        TestCacheMinimal = False
        Exit Function
    End If
    
    ' Test simple : stocker et r�cup�rer une valeur
    Dim testKey As String
    Dim testValue As String
    
    testKey = "TestKey_" & Format$(Now(), "yyyymmddhhnnss")
    testValue = "TestValue_" & Format$(Now(), "yyyymmddhhnnss")
    
    ' Stocker la valeur
    cache.Set testKey, testValue
    
    ' R�cup�rer la valeur
    Dim retrievedValue As Variant
    retrievedValue = cache.Get(testKey)
    
    ' V�rifier que la valeur r�cup�r�e correspond � la valeur stock�e
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

' Cr�e une instance par d�faut de IAppContext
Private Function CreateDefaultAppContext() As IAppContext
    On Error GoTo ErrorHandler
    
    ' Cr�er l'instance de contexte applicatif
    Dim appContext As New clsStandardAppContext
    
    ' Initialiser avec la configuration par d�faut
    appContext.Initialize
    
    ' Retourner l'instance
    Set CreateDefaultAppContext = appContext
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur dans CreateDefaultAppContext: " & Err.Description
    Set CreateDefaultAppContext = Nothing
End Function 