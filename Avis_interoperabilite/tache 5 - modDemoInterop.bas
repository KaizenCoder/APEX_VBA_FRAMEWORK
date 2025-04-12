' modDemoInterop.bas
' Description: Module démonstratif de l'interopérabilité Apex-Excel
Option Explicit

' Démonstration complète de l'architecture
Public Sub DemonstrationInteroperabilite()
    On Error GoTo GestionErreur
    
    ' 1. Initialisation du contexte avec un logger de test
    Dim testLogger As ILoggerBase
    Set testLogger = CreateTestLogger()
    
    Dim ctx As New clsAppContext
    SetLogger testLogger ' Injecter le logger de test
    ctx.Init LOGGER_DEV
    
    ' 2. Journalisation
    ctx.Logger.Info "Début de la démonstration"
    ctx.Logger.Debug "Mode de développement activé"
    
    ' 3. Création et utilisation d'un mock workbook
    DemoMockWorkbook ctx
    
    ' 4. Utilisation avec Excel réel
    DemoRealExcel ctx
    
    ' 5. Vérifier les résultats de journalisation
    VerifyLogs testLogger
    
    MsgBox "Démonstration terminée avec succès !", vbInformation, "Démo Interopérabilité"
    
    Exit Sub
    
GestionErreur:
    MsgBox "Erreur lors de la démonstration: " & Err.Description, vbCritical, "Erreur"
End Sub

' Démonstration avec un mock workbook
Private Sub DemoMockWorkbook(ByVal ctx As IAppContext)
    ctx.Logger.Info "Test avec mock workbook"
    
    ' Créer un mock workbook
    Dim mockWb As New clsMockWorkbookAccessor
    
    ' Ajouter des feuilles
    mockWb.AddMockSheet "Feuil1"
    mockWb.AddMockSheet "Données"
    
    ' Récupérer une feuille
    Dim sheet As ISheetAccessor
    Set sheet = mockWb.GetSheet("Données")
    
    ' Écrire des données
    sheet.GetCell(1, 1).Value = "En-tête 1"
    sheet.GetCell(1, 2).Value = "En-tête 2"
    sheet.GetCell(1, 3).Value = "En-tête 3"
    
    Dim data(1 To 3, 1 To 3) As Variant
    data(1, 1) = "A1"
    data(1, 2) = "B1"
    data(1, 3) = "C1"
    data(2, 1) = "A2"
    data(2, 2) = "B2"
    data(2, 3) = "C2"
    data(3, 1) = "A3"
    data(3, 2) = "B3"
    data(3, 3) = "C3"
    
    sheet.WriteRange 2, 1, data
    
    ' Lire des données
    Dim readData As Variant
    readData = sheet.ReadRange(2, 1, 4, 3)
    
    ' Vérifier une cellule spécifique
    Dim cell As ICellAccessor
    Set cell = sheet.GetCell(3, 2)
    
    If cell.Value = "B2" Then
        ctx.Logger.Info "Test mock workbook réussi: valeur correcte lue"
    Else
        ctx.Logger.Error "Test mock workbook échoué: valeur incorrecte"
    End If
    
    ctx.Logger.Debug "Test mock workbook terminé"
End Sub

' Démonstration avec Excel réel
Private Sub DemoRealExcel(ByVal ctx As IAppContext)
    On Error GoTo GestionErreur
    
    ctx.Logger.Info "Test avec Excel réel"
    
    ' Utiliser le classeur actif
    Dim workbook As IWorkbookAccessor
    Set workbook = ctx.GetWorkbookAccessor(ThisWorkbook)
    
    ' Vérifier si une feuille de test existe, sinon la créer
    On Error Resume Next
    Dim testSheet As Worksheet
    Set testSheet = ThisWorkbook.Sheets("Test_Interop")
    
    If testSheet Is Nothing Then
        Set testSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        testSheet.Name = "Test_Interop"
    End If
    On Error GoTo GestionErreur
    
    ' Utiliser l'abstraction pour accéder à la feuille
    Dim sheet As ISheetAccessor
    Set sheet = workbook.GetSheet("Test_Interop")
    
    If sheet Is Nothing Then
        ctx.Logger.Error "Impossible d'accéder à la feuille de test"
        Exit Sub
    End If
    
    ' Écrire des données
    sheet.GetCell(1, 1).Value = "Test Interopérabilité"
    sheet.GetCell(1, 2).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    ' Exemple d'utilisation avancée: tableau de données
    Dim headerRow(1 To 1, 1 To 3) As Variant
    headerRow(1, 1) = "ID"
    headerRow(1, 2) = "Nom"
    headerRow(1, 3) = "Valeur"
    
    sheet.WriteRange 3, 1, headerRow
    
    Dim dataRows(1 To 3, 1 To 3) As Variant
    dataRows(1, 1) = 1
    dataRows(1, 2) = "Produit A"
    dataRows(1, 3) = 100
    dataRows(2, 1) = 2
    dataRows(2, 2) = "Produit B"
    dataRows(2, 3) = 200
    dataRows(3, 1) = 3
    dataRows(3, 2) = "Produit C"
    dataRows(3, 3) = 300
    
    sheet.WriteRange 4, 1, dataRows
    
    ' Lire des données
    Dim readData As Variant
    readData = sheet.ReadRange(4, 1, 6, 3)
    
    ' Vérifier les données lues
    If readData(4, 2) = "Produit A" Then
        ctx.Logger.Info "Test Excel réel réussi: valeur correcte lue"
    Else
        ctx.Logger.Error "Test Excel réel échoué: valeur incorrecte"
    End If
    
    ctx.Logger.Debug "Test Excel réel terminé"
    Exit Sub
    
GestionErreur:
    ctx.ReportException "DemoRealExcel"
End Sub

' Vérification des logs générés
Private Sub VerifyLogs(ByVal testLogger As ILoggerBase)
    ' Vérification spécifique aux tests
    If TypeOf testLogger Is clsTestLogger Then
        Dim logs As Collection
        Set logs = testLogger.GetLogs()
        
        ' Afficher un résumé
        Debug.Print "---- Résumé des logs ----"
        Debug.Print "Nombre de messages: " & logs.Count
        
        ' Vérifier la présence de messages spécifiques
        If testLogger.Contains("Test mock workbook réussi") And _
           testLogger.Contains("Test Excel réel réussi") Then
            Debug.Print "Tous les tests ont réussi!"
        Else
            Debug.Print "Certains tests ont échoué, vérifier les logs."
        End If
    End If
End Sub

' Exemple de pattern Command
Public Sub ExecuterTraitementAvecLogs()
    ' Créer et configurer le contexte
    Dim ctx As New clsAppContext
    
    ' Créer une feuille de logs si nécessaire
    On Error Resume Next
    Dim logSheet As Worksheet
    Set logSheet = ThisWorkbook.Sheets("Logs")
    
    If logSheet Is Nothing Then
        Set logSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        logSheet.Name = "Logs"
    End If
    On Error GoTo 0
    
    ' Configurer un logger composite (console + feuille)
    Dim sheetLogger As ILoggerBase
    Set sheetLogger = CreateSheetLogger(logSheet)
    
    Dim logger As ILoggerBase
    Set logger = CreateCompositeLogger(CreateDebugLogger(), sheetLogger)
    
    SetLogger logger
    
    ' Exécuter le traitement avec le contexte configuré
    Dim moduleStandard As New modTraitementStandard
    moduleStandard.RunTraitementStandard ctx
End Sub 