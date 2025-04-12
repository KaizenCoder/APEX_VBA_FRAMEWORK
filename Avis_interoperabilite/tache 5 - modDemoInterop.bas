' modDemoInterop.bas
' Description: Module d�monstratif de l'interop�rabilit� Apex-Excel
Option Explicit

' D�monstration compl�te de l'architecture
Public Sub DemonstrationInteroperabilite()
    On Error GoTo GestionErreur
    
    ' 1. Initialisation du contexte avec un logger de test
    Dim testLogger As ILoggerBase
    Set testLogger = CreateTestLogger()
    
    Dim ctx As New clsAppContext
    SetLogger testLogger ' Injecter le logger de test
    ctx.Init LOGGER_DEV
    
    ' 2. Journalisation
    ctx.Logger.Info "D�but de la d�monstration"
    ctx.Logger.Debug "Mode de d�veloppement activ�"
    
    ' 3. Cr�ation et utilisation d'un mock workbook
    DemoMockWorkbook ctx
    
    ' 4. Utilisation avec Excel r�el
    DemoRealExcel ctx
    
    ' 5. V�rifier les r�sultats de journalisation
    VerifyLogs testLogger
    
    MsgBox "D�monstration termin�e avec succ�s !", vbInformation, "D�mo Interop�rabilit�"
    
    Exit Sub
    
GestionErreur:
    MsgBox "Erreur lors de la d�monstration: " & Err.Description, vbCritical, "Erreur"
End Sub

' D�monstration avec un mock workbook
Private Sub DemoMockWorkbook(ByVal ctx As IAppContext)
    ctx.Logger.Info "Test avec mock workbook"
    
    ' Cr�er un mock workbook
    Dim mockWb As New clsMockWorkbookAccessor
    
    ' Ajouter des feuilles
    mockWb.AddMockSheet "Feuil1"
    mockWb.AddMockSheet "Donn�es"
    
    ' R�cup�rer une feuille
    Dim sheet As ISheetAccessor
    Set sheet = mockWb.GetSheet("Donn�es")
    
    ' �crire des donn�es
    sheet.GetCell(1, 1).Value = "En-t�te 1"
    sheet.GetCell(1, 2).Value = "En-t�te 2"
    sheet.GetCell(1, 3).Value = "En-t�te 3"
    
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
    
    ' Lire des donn�es
    Dim readData As Variant
    readData = sheet.ReadRange(2, 1, 4, 3)
    
    ' V�rifier une cellule sp�cifique
    Dim cell As ICellAccessor
    Set cell = sheet.GetCell(3, 2)
    
    If cell.Value = "B2" Then
        ctx.Logger.Info "Test mock workbook r�ussi: valeur correcte lue"
    Else
        ctx.Logger.Error "Test mock workbook �chou�: valeur incorrecte"
    End If
    
    ctx.Logger.Debug "Test mock workbook termin�"
End Sub

' D�monstration avec Excel r�el
Private Sub DemoRealExcel(ByVal ctx As IAppContext)
    On Error GoTo GestionErreur
    
    ctx.Logger.Info "Test avec Excel r�el"
    
    ' Utiliser le classeur actif
    Dim workbook As IWorkbookAccessor
    Set workbook = ctx.GetWorkbookAccessor(ThisWorkbook)
    
    ' V�rifier si une feuille de test existe, sinon la cr�er
    On Error Resume Next
    Dim testSheet As Worksheet
    Set testSheet = ThisWorkbook.Sheets("Test_Interop")
    
    If testSheet Is Nothing Then
        Set testSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        testSheet.Name = "Test_Interop"
    End If
    On Error GoTo GestionErreur
    
    ' Utiliser l'abstraction pour acc�der � la feuille
    Dim sheet As ISheetAccessor
    Set sheet = workbook.GetSheet("Test_Interop")
    
    If sheet Is Nothing Then
        ctx.Logger.Error "Impossible d'acc�der � la feuille de test"
        Exit Sub
    End If
    
    ' �crire des donn�es
    sheet.GetCell(1, 1).Value = "Test Interop�rabilit�"
    sheet.GetCell(1, 2).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    ' Exemple d'utilisation avanc�e: tableau de donn�es
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
    
    ' Lire des donn�es
    Dim readData As Variant
    readData = sheet.ReadRange(4, 1, 6, 3)
    
    ' V�rifier les donn�es lues
    If readData(4, 2) = "Produit A" Then
        ctx.Logger.Info "Test Excel r�el r�ussi: valeur correcte lue"
    Else
        ctx.Logger.Error "Test Excel r�el �chou�: valeur incorrecte"
    End If
    
    ctx.Logger.Debug "Test Excel r�el termin�"
    Exit Sub
    
GestionErreur:
    ctx.ReportException "DemoRealExcel"
End Sub

' V�rification des logs g�n�r�s
Private Sub VerifyLogs(ByVal testLogger As ILoggerBase)
    ' V�rification sp�cifique aux tests
    If TypeOf testLogger Is clsTestLogger Then
        Dim logs As Collection
        Set logs = testLogger.GetLogs()
        
        ' Afficher un r�sum�
        Debug.Print "---- R�sum� des logs ----"
        Debug.Print "Nombre de messages: " & logs.Count
        
        ' V�rifier la pr�sence de messages sp�cifiques
        If testLogger.Contains("Test mock workbook r�ussi") And _
           testLogger.Contains("Test Excel r�el r�ussi") Then
            Debug.Print "Tous les tests ont r�ussi!"
        Else
            Debug.Print "Certains tests ont �chou�, v�rifier les logs."
        End If
    End If
End Sub

' Exemple de pattern Command
Public Sub ExecuterTraitementAvecLogs()
    ' Cr�er et configurer le contexte
    Dim ctx As New clsAppContext
    
    ' Cr�er une feuille de logs si n�cessaire
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
    
    ' Ex�cuter le traitement avec le contexte configur�
    Dim moduleStandard As New modTraitementStandard
    moduleStandard.RunTraitementStandard ctx
End Sub 