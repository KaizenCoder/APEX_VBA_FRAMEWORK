' modGuideIntegration.bas
' Description: Guide d'int�gration et d'utilisation de l'architecture d'interop�rabilit� Apex-Excel
Option Explicit

' Ce module contient des exemples et instructions pour int�grer l'architecture
' d'interop�rabilit� dans un projet VBA existant. Il peut �tre consult� comme
' r�f�rence mais n'est pas destin� � �tre ex�cut� directement.

' ==========================================================================
' GUIDE D'INT�GRATION DE L'ARCHITECTURE APEX-EXCEL
' ==========================================================================
'
' Ce guide vous explique comment int�grer et utiliser l'architecture
' d'interop�rabilit� Apex-Excel dans vos projets VBA existants.
'
' �TAPE 1: IMPORT DES FICHIERS
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
' b) Impl�mentations de logging:
'    - clsDebugLogger.cls
'    - clsSheetLogger.cls
'    - clsFileLogger.cls
'    - clsTestLogger.cls
'    - clsCompositeLogger.cls
'    - modLogFactory.bas
'
' c) Impl�mentations d'acc�s Excel:
'    - clsExcelWorkbookAccessor.cls
'    - clsExcelSheetAccessor.cls
'    - clsExcelCellAccessor.cls
'
' d) Impl�mentations pour tests:
'    - clsMockWorkbookAccessor.cls
'    - clsMockSheetAccessor.cls
'    - clsMockCellAccessor.cls
'
' e) Contexte d'application:
'    - clsAppContext.cls
'
' �TAPE 2: CONFIGURATION DU CONTEXTE
' ----------------------------------
'
' Exemple de configuration du contexte d'application:

Public Sub ExempleConfigurationContexte()
    ' Cr�er et initialiser le contexte
    Dim ctx As New clsAppContext
    ctx.Init LOGGER_DEV ' Ou LOGGER_TEST, LOGGER_PROD selon l'environnement
    
    ' Pour acc�der au logger
    ctx.Logger.Info "Initialisation du contexte termin�e"
    
    ' Pour acc�der � la configuration
    Dim defaultSheet As String
    defaultSheet = ctx.Config("DefaultSheet")
    
    ' Utiliser le contexte dans vos modules m�tier
    ' ModuleA.ExecuterTraitement ctx
End Sub

' �TAPE 3: UTILISATION DANS UN MODULE M�TIER
' ------------------------------------------
'
' Exemple de structure d'un module m�tier:

Public Sub ExempleModuleM�tier()
    ' Ce code est destin� � �tre plac� dans votre module m�tier
    
    Private ctx As IAppContext
    
    Public Sub ExecuterTraitement(ByVal injectedCtx As IAppContext)
        Set ctx = injectedCtx
        On Error GoTo GestionErreur
        
        ctx.Logger.Info "D�but du traitement"
        
        ' Acc�s aux donn�es Excel via abstraction
        Dim workbook As IWorkbookAccessor
        Set workbook = ctx.GetWorkbookAccessor(ThisWorkbook)
        
        Dim sheet As ISheetAccessor
        Set sheet = workbook.GetSheet("MaFeuille")
        
        ' Lecture de donn�es
        Dim data As Variant
        data = sheet.ReadRange(2, 1, 10, 5)
        
        ' Traitement des donn�es
        ' ...
        
        ' �criture des r�sultats
        sheet.GetCell(1, 1).Value = "R�sultat"
        
        ctx.Logger.Info "Fin du traitement"
        Exit Sub
        
    GestionErreur:
        ctx.ReportException "ExecuterTraitement"
    End Sub
End Sub

' �TAPE 4: TESTS UNITAIRES
' ------------------------
'
' Exemple de test unitaire avec mocks:

Public Sub ExempleTestUnitaire()
    ' Cr�er un logger de test
    Dim testLogger As New clsTestLogger
    
    ' Cr�er un contexte de test
    Dim ctx As New clsAppContext
    SetLogger testLogger
    
    ' Cr�er un mock workbook pour les tests
    Dim mockWb As New clsMockWorkbookAccessor
    mockWb.AddMockSheet "TestSheet"
    
    Dim sheet As ISheetAccessor
    Set sheet = mockWb.GetSheet("TestSheet")
    
    ' Pr�parer des donn�es de test
    sheet.GetCell(1, 1).Value = "Test"
    
    ' Ex�cuter le traitement � tester
    ' ModuleA.ExecuterTraitement ctx
    
    ' V�rifier les r�sultats
    If testLogger.Contains("Erreur") Then
        Debug.Print "Le test a �chou�"
    Else
        Debug.Print "Le test a r�ussi"
    End If
End Sub

' �TAPE 5: ASTUCES ET BONNES PRATIQUES
' ------------------------------------
'
' 1. Injection de d�pendances
'    Toujours injecter le contexte dans vos modules plut�t que de le cr�er
'    directement, afin de faciliter les tests et la r�utilisation.
'
' 2. Gestion des erreurs
'    Utilisez syst�matiquement le pattern with On Error et ctx.ReportException
'    pour capturer et journaliser les erreurs.
'
' 3. Pattern d'architecture
'    Structurez vos modules selon le pattern pr�sent� dans modTraitementStandard.bas
'    avec des �tapes clairement d�finies.
'
' 4. Nommage
'    Suivez les conventions de nommage: 
'    - cls pour les classes
'    - I pour les interfaces
'    - mod pour les modules
'
' 5. Tests
'    �crivez des tests unitaires pour chaque module m�tier en utilisant les
'    impl�mentations mock fournies.

' ==========================================================================
' EXEMPLES D'INT�GRATION AVANC�E
' ==========================================================================

' Exemple d'int�gration avec une application existante
Public Sub IntegrerDansApplicationExistante()
    ' 1. Cr�er ou r�cup�rer une instance du contexte
    Dim ctx As IAppContext
    
    ' Option 1: Cr�er une nouvelle instance
    Set ctx = New clsAppContext
    
    ' Option 2: R�cup�rer depuis un module global
    ' Set ctx = GlobalContext.GetContext()
    
    ' 2. Configurer un logger adapt� � l'environnement actuel
    Dim env As LoggerEnvironment
    
    ' D�terminer l'environnement (exemple)
    If InStr(ThisWorkbook.Name, "DEV") > 0 Then
        env = LOGGER_DEV
    ElseIf InStr(ThisWorkbook.Name, "TEST") > 0 Then
        env = LOGGER_TEST
    Else
        env = LOGGER_PROD
    End If
    
    ctx.Init env
    
    ' 3. Ex�cuter les traitements en injectant le contexte
    ' ModuleA.ExecuterTraitement ctx
    ' ModuleB.ExecuterTraitement ctx
End Sub

' Exemple d'extension de l'architecture
Public Sub ExtensionArchitecture()
    ' Ce code montre comment �tendre l'architecture
    
    ' Exemple: Cr�er un nouveau type de logger personnalis�
    ' (Impl�mentation simplifi�e pour l'exemple)
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
    '     ' Autres m�thodes ILoggerBase...
    ' End Class
    '
    ' Function CreateEmailLogger() As ILoggerBase
    '     Dim logger As New clsEmailLogger
    '     Set CreateEmailLogger = logger
    ' End Function
End Sub 