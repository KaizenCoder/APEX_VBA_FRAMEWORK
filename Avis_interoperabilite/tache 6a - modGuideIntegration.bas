' modGuideIntegration.bas
' Description: Guide d'intégration et d'utilisation de l'architecture d'interopérabilité Apex-Excel
Option Explicit

' Ce module contient des exemples et instructions pour intégrer l'architecture
' d'interopérabilité dans un projet VBA existant. Il peut être consulté comme
' référence mais n'est pas destiné à être exécuté directement.

' ==========================================================================
' GUIDE D'INTÉGRATION DE L'ARCHITECTURE APEX-EXCEL
' ==========================================================================
'
' Ce guide vous explique comment intégrer et utiliser l'architecture
' d'interopérabilité Apex-Excel dans vos projets VBA existants.
'
' ÉTAPE 1: IMPORT DES FICHIERS
' ----------------------------
'
' Importez les fichiers suivants dans votre projet VBA:
'
' a) Interfaces:
'    - ILoggerBase.cls
'    - IWorkbookAccessor.cls
'    - ISheetAccessor.cls
'    - ICellAccessor.cls
'    - IAppContext.cls
'
' b) Implémentations de logging:
'    - clsDebugLogger.cls
'    - clsSheetLogger.cls
'    - clsFileLogger.cls
'    - clsTestLogger.cls
'    - clsCompositeLogger.cls
'    - modLogFactory.bas
'
' c) Implémentations d'accès Excel:
'    - clsExcelWorkbookAccessor.cls
'    - clsExcelSheetAccessor.cls
'    - clsExcelCellAccessor.cls
'
' d) Implémentations pour tests:
'    - clsMockWorkbookAccessor.cls
'    - clsMockSheetAccessor.cls
'    - clsMockCellAccessor.cls
'
' e) Contexte d'application:
'    - clsAppContext.cls
'
' ÉTAPE 2: CONFIGURATION DU CONTEXTE
' ----------------------------------
'
' Exemple de configuration du contexte d'application:

Public Sub ExempleConfigurationContexte()
    ' Créer et initialiser le contexte
    Dim ctx As New clsAppContext
    ctx.Init LOGGER_DEV ' Ou LOGGER_TEST, LOGGER_PROD selon l'environnement
    
    ' Pour accéder au logger
    ctx.Logger.Info "Initialisation du contexte terminée"
    
    ' Pour accéder à la configuration
    Dim defaultSheet As String
    defaultSheet = ctx.Config("DefaultSheet")
    
    ' Utiliser le contexte dans vos modules métier
    ' ModuleA.ExecuterTraitement ctx
End Sub

' ÉTAPE 3: UTILISATION DANS UN MODULE MÉTIER
' ------------------------------------------
'
' Exemple de structure d'un module métier:

Public Sub ExempleModuleMétier()
    ' Ce code est destiné à être placé dans votre module métier
    
    Private ctx As IAppContext
    
    Public Sub ExecuterTraitement(ByVal injectedCtx As IAppContext)
        Set ctx = injectedCtx
        On Error GoTo GestionErreur
        
        ctx.Logger.Info "Début du traitement"
        
        ' Accès aux données Excel via abstraction
        Dim workbook As IWorkbookAccessor
        Set workbook = ctx.GetWorkbookAccessor(ThisWorkbook)
        
        Dim sheet As ISheetAccessor
        Set sheet = workbook.GetSheet("MaFeuille")
        
        ' Lecture de données
        Dim data As Variant
        data = sheet.ReadRange(2, 1, 10, 5)
        
        ' Traitement des données
        ' ...
        
        ' Écriture des résultats
        sheet.GetCell(1, 1).Value = "Résultat"
        
        ctx.Logger.Info "Fin du traitement"
        Exit Sub
        
    GestionErreur:
        ctx.ReportException "ExecuterTraitement"
    End Sub
End Sub

' ÉTAPE 4: TESTS UNITAIRES
' ------------------------
'
' Exemple de test unitaire avec mocks:

Public Sub ExempleTestUnitaire()
    ' Créer un logger de test
    Dim testLogger As New clsTestLogger
    
    ' Créer un contexte de test
    Dim ctx As New clsAppContext
    SetLogger testLogger
    
    ' Créer un mock workbook pour les tests
    Dim mockWb As New clsMockWorkbookAccessor
    mockWb.AddMockSheet "TestSheet"
    
    Dim sheet As ISheetAccessor
    Set sheet = mockWb.GetSheet("TestSheet")
    
    ' Préparer des données de test
    sheet.GetCell(1, 1).Value = "Test"
    
    ' Exécuter le traitement à tester
    ' ModuleA.ExecuterTraitement ctx
    
    ' Vérifier les résultats
    If testLogger.Contains("Erreur") Then
        Debug.Print "Le test a échoué"
    Else
        Debug.Print "Le test a réussi"
    End If
End Sub

' ÉTAPE 5: ASTUCES ET BONNES PRATIQUES
' ------------------------------------
'
' 1. Injection de dépendances
'    Toujours injecter le contexte dans vos modules plutôt que de le créer
'    directement, afin de faciliter les tests et la réutilisation.
'
' 2. Gestion des erreurs
'    Utilisez systématiquement le pattern with On Error et ctx.ReportException
'    pour capturer et journaliser les erreurs.
'
' 3. Pattern d'architecture
'    Structurez vos modules selon le pattern présenté dans modTraitementStandard.bas
'    avec des étapes clairement définies.
'
' 4. Nommage
'    Suivez les conventions de nommage: 
'    - cls pour les classes
'    - I pour les interfaces
'    - mod pour les modules
'
' 5. Tests
'    Écrivez des tests unitaires pour chaque module métier en utilisant les
'    implémentations mock fournies.

' ==========================================================================
' EXEMPLES D'INTÉGRATION AVANCÉE
' ==========================================================================

' Exemple d'intégration avec une application existante
Public Sub IntegrerDansApplicationExistante()
    ' 1. Créer ou récupérer une instance du contexte
    Dim ctx As IAppContext
    
    ' Option 1: Créer une nouvelle instance
    Set ctx = New clsAppContext
    
    ' Option 2: Récupérer depuis un module global
    ' Set ctx = GlobalContext.GetContext()
    
    ' 2. Configurer un logger adapté à l'environnement actuel
    Dim env As LoggerEnvironment
    
    ' Déterminer l'environnement (exemple)
    If InStr(ThisWorkbook.Name, "DEV") > 0 Then
        env = LOGGER_DEV
    ElseIf InStr(ThisWorkbook.Name, "TEST") > 0 Then
        env = LOGGER_TEST
    Else
        env = LOGGER_PROD
    End If
    
    ctx.Init env
    
    ' 3. Exécuter les traitements en injectant le contexte
    ' ModuleA.ExecuterTraitement ctx
    ' ModuleB.ExecuterTraitement ctx
End Sub

' Exemple d'extension de l'architecture
Public Sub ExtensionArchitecture()
    ' Ce code montre comment étendre l'architecture
    
    ' Exemple: Créer un nouveau type de logger personnalisé
    ' (Implémentation simplifiée pour l'exemple)
    '
    ' Class clsEmailLogger
    '     Implements ILoggerBase
    '     
    '     Private Sub ILoggerBase_Log(ByVal level As String, ByVal message As String)
    '         If level = "ERROR" Then
    '             ' Envoyer un email avec la librairie d'email de votre choix
    '             SendEmail "admin@example.com", "Erreur application", message
    '         End If
    '     End Sub
    '     
    '     ' Autres méthodes ILoggerBase...
    ' End Class
    '
    ' Function CreateEmailLogger() As ILoggerBase
    '     Dim logger As New clsEmailLogger
    '     Set CreateEmailLogger = logger
    ' End Function
End Sub 