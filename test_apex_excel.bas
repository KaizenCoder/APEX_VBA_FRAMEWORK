Attribute VB_Name = "test_apex_excel"
Option Explicit

' Script de test pour créer un classeur Excel à partir du plan de situation APEX
' Utilise les accesseurs Excel du framework APEX
' Date: 13/04/2025

Public Sub CopyPlanSituationToExcel()
    ' Déclaration des variables APEX
    Dim workbookAccessor As IWorkbookAccessor
    Dim sheetAccessor As ISheetAccessor
    Dim tableAccessor As ITableAccessor
    
    ' Création d'un nouveau classeur Excel
    Dim newWorkbook As Workbook
    Set newWorkbook = Workbooks.Add
    
    ' Renommer la première feuille
    newWorkbook.Sheets(1).Name = "Plan de Situation"
    
    ' Initialiser l'accesseur de classeur APEX via la factory
    Dim factory As New ModExcelFactory
    Set workbookAccessor = factory.CreateWorkbookAccessor(newWorkbook)
    
    ' Obtenir l'accesseur de feuille pour la première feuille
    Set sheetAccessor = workbookAccessor.GetSheet("Plan de Situation")
    
    ' Écrire le titre
    With sheetAccessor
        .WriteValue 1, 1, "Plan de Situation APEX Framework - 2024-04-14"
        .GetCell(1, 1).FontBold = True
        .GetCell(1, 1).BackColor = RGB(230, 230, 230)
    End With
    
    ' Position courante pour l'écriture
    Dim currentRow As Long
    currentRow = 3
    
    ' Écrire les sections
    WriteSection "Composants Database", currentRow, sheetAccessor
    WriteComponentsTable "1. Interfaces", Array("IDbDriver", "IQueryBuilder", "IDBAccessorBase", "IEntityMapping"), _
                        Array("?", "?", "?", "?"), _
                        Array("?? Cursor", "?? Cursor", "?? VSCode", "?? Cursor"), _
                        currentRow, sheetAccessor
                        
    WriteComponentsTable "2. Implémentations", Array("clsDBAccessor", "clsSqlQueryBuilder", "ClsOrmBase", "clsEntityMappingFactory"), _
                        Array("?", "?", "?", "?"), _
                        Array("?? VSCode", "?? Cursor", "?? Cursor", "?? Cursor"), _
                        currentRow, sheetAccessor
                        
    WriteComponentsTable "3. Tests", Array("TestQueryBuilder", "TestQueryBuilderIntegration", "TestDbAccessorIntegration", _
                                         "TestDBAccessorAdvanced", "TestEntityMappingFactory", "TestOrmIntegration", "TestOrmPerformance"), _
                        Array("?", "?", "?", "?", "?", "?", "?"), _
                        Array("?? Cursor", "?? Cursor", "?? VSCode", "?? Cursor", "?? Cursor", "?? Cursor", "?? Cursor"), _
                        currentRow, sheetAccessor
                        
    WriteSection "Composants Excel", currentRow, sheetAccessor
    WriteComponentsTable "1. Interfaces", Array("IWorkbookAccessor", "ISheetAccessor", "ITableAccessor", "IRangeAccessor", "ICellAccessor"), _
                        Array("?", "?", "?", "?", "?"), _
                        Array("?? VSCode", "?? Cursor", "?? Cursor", "?? VSCode", "?? Cursor"), _
                        currentRow, sheetAccessor
                        
    WriteComponentsTable "2. Implémentations", Array("clsExcelWorkbookAccessor", "clsExcelSheetAccessor", "clsExcelTableAccessor", _
                                                    "clsExcelRangeAccessor", "clsExcelCellAccessor"), _
                        Array("?", "?", "?", "?", "?"), _
                        Array("?? VSCode", "?? Cursor", "?? Cursor", "?? VSCode", "?? Cursor"), _
                        currentRow, sheetAccessor
                        
    ' Écrire les statistiques
    WriteSection "Statistiques", currentRow, sheetAccessor
    With sheetAccessor
        .WriteValue currentRow, 1, "Couverture des tests:"
        currentRow = currentRow + 1
        .WriteValue currentRow, 1, "Tests unitaires: 95%"
        currentRow = currentRow + 1
        .WriteValue currentRow, 1, "Tests d'intégration: 90%"
        currentRow = currentRow + 1
        .WriteValue currentRow, 1, "Tests de performance: 95%"
        currentRow = currentRow + 1
        .WriteValue currentRow, 1, "Tests de sécurité: 75%"
        currentRow = currentRow + 1
        .WriteValue currentRow, 1, "Tests ORM: 85%"
        currentRow = currentRow + 1
        .WriteValue currentRow, 1, "Documentation: 100%"
    End With
    
    ' Formater le classeur
    FormatWorkbook sheetAccessor
    
    ' Sauvegarder le classeur
    workbookAccessor.SaveAs "D:\Dev\Apex_VBA_FRAMEWORK\docs\implementation\PLAN_SITUATION_2024_04_14.xlsx"
End Sub

Private Sub WriteSection(ByVal sectionName As String, ByRef currentRow As Long, ByVal sheetAccessor As ISheetAccessor)
    currentRow = currentRow + 1
    With sheetAccessor
        .WriteValue currentRow, 1, sectionName
        .GetCell(currentRow, 1).FontBold = True
        .GetCell(currentRow, 1).BackColor = RGB(200, 200, 200)
    End With
    currentRow = currentRow + 1
End Sub

Private Sub WriteComponentsTable(ByVal tableName As String, ByVal components As Variant, _
                               ByVal statuses As Variant, ByVal contributors As Variant, _
                               ByRef currentRow As Long, ByVal sheetAccessor As ISheetAccessor)
    Dim i As Long
    
    ' Écrire le nom de la table
    With sheetAccessor
        .WriteValue currentRow, 2, tableName
        .GetCell(currentRow, 2).FontBold = True
    End With
    currentRow = currentRow + 1
    
    ' Écrire les en-têtes
    With sheetAccessor
        .WriteValue currentRow, 2, "Composant"
        .WriteValue currentRow, 3, "État"
        .WriteValue currentRow, 4, "Contributeur"
        .GetRange("B" & currentRow & ":D" & currentRow).BackColor = RGB(240, 240, 240)
    End With
    currentRow = currentRow + 1
    
    ' Écrire les données
    For i = LBound(components) To UBound(components)
        With sheetAccessor
            .WriteValue currentRow, 2, components(i)
            .WriteValue currentRow, 3, statuses(i)
            .WriteValue currentRow, 4, contributors(i)
        End With
        currentRow = currentRow + 1
    Next i
    
    currentRow = currentRow + 1
End Sub

Private Sub FormatWorkbook(ByVal sheetAccessor As ISheetAccessor)
    ' Ajuster les colonnes
    With sheetAccessor
        .GetRange("A:A").ColumnWidth = 5
        .GetRange("B:B").ColumnWidth = 30
        .GetRange("C:C").ColumnWidth = 10
        .GetRange("D:D").ColumnWidth = 15
    End With
End Sub

' Fonction pour lire le contenu du fichier MD
Private Function ReadMdFile(filePath As String) As String
    Dim fileNum As Integer
    Dim fileContent As String
    Dim tempLine As String
    
    fileNum = FreeFile
    
    ' Ouvrir le fichier en lecture
    Open filePath For Input As #fileNum
    
    ' Lire tout le contenu
    While Not EOF(fileNum)
        Line Input #fileNum, tempLine
        fileContent = fileContent & tempLine & vbCrLf
    Wend
    
    ' Fermer le fichier
    Close #fileNum
    
    ReadMdFile = fileContent
End Function

' Création d'un tableau Excel avancé pour les dernières mises à jour
Private Sub CreateUpdatesTable(ws As Worksheet, startRow As Integer)
    ' Titre du tableau
    ws.Range("A" & startRow).Value = "Dernières Mises à Jour"
    ws.Range("A" & startRow).Font.Bold = True
    ws.Range("A" & startRow).Font.Size = 12
    
    ' En-têtes du tableau
    ws.Range("A" & (startRow + 1)).Value = "Date"
    ws.Range("B" & (startRow + 1)).Value = "Description"
    ws.Range("C" & (startRow + 1)).Value = "Contributeur"
    ws.Range("A" & (startRow + 1) & ":C" & (startRow + 1)).Font.Bold = True
    
    ' Données du tableau
    ws.Range("A" & (startRow + 2)).Value = "2024-04-14"
    ws.Range("B" & (startRow + 2)).Value = "Tests d'intégration ORM"
    ws.Range("C" & (startRow + 2)).Value = "Cursor"
    
    ws.Range("A" & (startRow + 3)).Value = "2024-04-14"
    ws.Range("B" & (startRow + 3)).Value = "Factory des mappings d'entités"
    ws.Range("C" & (startRow + 3)).Value = "Cursor"
    
    ws.Range("A" & (startRow + 4)).Value = "2024-04-14"
    ws.Range("B" & (startRow + 4)).Value = "Tests avancés DBAccessor"
    ws.Range("C" & (startRow + 4)).Value = "Cursor"
    
    ws.Range("A" & (startRow + 5)).Value = "2024-04-13"
    ws.Range("B" & (startRow + 5)).Value = "Tests d'intégration QueryBuilder"
    ws.Range("C" & (startRow + 5)).Value = "Cursor"
    
    ws.Range("A" & (startRow + 6)).Value = "2024-04-12"
    ws.Range("B" & (startRow + 6)).Value = "Accesseurs Excel"
    ws.Range("C" & (startRow + 6)).Value = "VSCode"
    
    ' Formatage du tableau
    ws.Range("A" & (startRow + 1) & ":C" & (startRow + 6)).BorderAround xlContinuous
    ws.Range("A" & (startRow + 1) & ":C" & (startRow + 1)).BorderAround xlContinuous
    
    ' Ajout d'une ligne de version
    ws.Range("A" & (startRow + 8)).Value = "Version: 2.3"
    ws.Range("A" & (startRow + 9)).Value = "Dernière mise à jour: 2024-04-14"
End Sub

' Fonction d'entrée pour le test
Public Sub TestApexExcel()
    CopyPlanSituationToExcel
End Sub