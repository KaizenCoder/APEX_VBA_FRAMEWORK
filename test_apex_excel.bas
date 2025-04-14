Attribute VB_Name = "test_apex_excel"
Option Explicit

' Script de test pour cr�er un classeur Excel � partir du plan de situation APEX
' Utilise les accesseurs Excel du framework APEX
' Date: 13/04/2025

Public Sub CopyPlanSituationToExcel()
    ' D�claration des variables APEX
    Dim workbookAccessor As Object ' IWorkbookAccessor
    Dim sheetAccessor As Object ' ISheetAccessor
    Dim tableAccessor As Object ' ITableAccessor
    
    ' Cr�ation d'un nouveau classeur Excel
    Dim newWorkbook As Workbook
    Set newWorkbook = Workbooks.Add
    
    ' Renommer la premi�re feuille
    newWorkbook.Sheets(1).Name = "Plan de Situation"
    
    ' Initialiser l'accesseur de classeur APEX
    Set workbookAccessor = CreateObject("clsExcelWorkbookAccessor")
    workbookAccessor.Init newWorkbook
    
    ' Obtenir l'accesseur de feuille pour la premi�re feuille
    Set sheetAccessor = workbookAccessor.GetSheet("Plan de Situation")
    
    ' Lire le contenu du plan de situation
    Dim planContent As String
    planContent = ReadMdFile("D:\Dev\Apex_VBA_FRAMEWORK\docs\implementation\PLAN_SITUATION_2024_04_14.md")
    
    ' Formater le titre et les sections principales
    With newWorkbook.Sheets("Plan de Situation")
        ' Titre principal
        .Range("A1").Value = "Plan de Situation APEX Framework - 2024-04-14"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        
        ' Sections principales (Composants Database)
        .Range("A3").Value = "Composants Database"
        .Range("A3").Font.Bold = True
        .Range("A3").Font.Size = 14
        
        ' Tableau des interfaces
        .Range("A5").Value = "1. Interfaces"
        .Range("A5").Font.Bold = True
        
        ' En-t�tes du tableau
        .Range("A6").Value = "Composant"
        .Range("B6").Value = "�tat"
        .Range("C6").Value = "Contributeur"
        .Range("A6:C6").Font.Bold = True
        
        ' Remplissage du tableau des interfaces
        .Range("A7").Value = "IDbDriver"
        .Range("B7").Value = "Compl�t�"
        .Range("C7").Value = "Cursor"
        
        .Range("A8").Value = "IQueryBuilder"
        .Range("B8").Value = "Compl�t�"
        .Range("C8").Value = "Cursor"
        
        .Range("A9").Value = "IDBAccessorBase"
        .Range("B9").Value = "Compl�t�"
        .Range("C9").Value = "VSCode"
        
        .Range("A10").Value = "IEntityMapping"
        .Range("B10").Value = "Compl�t�"
        .Range("C10").Value = "Cursor"
        
        ' Tableau des impl�mentations
        .Range("A12").Value = "2. Impl�mentations"
        .Range("A12").Font.Bold = True
        
        ' En-t�tes du tableau
        .Range("A13").Value = "Composant"
        .Range("B13").Value = "�tat"
        .Range("C13").Value = "Contributeur"
        .Range("A13:C13").Font.Bold = True
        
        ' Remplissage du tableau des impl�mentations
        .Range("A14").Value = "clsDBAccessor"
        .Range("B14").Value = "Compl�t�"
        .Range("C14").Value = "VSCode"
        
        .Range("A15").Value = "clsSqlQueryBuilder"
        .Range("B15").Value = "Compl�t�"
        .Range("C15").Value = "Cursor"
        
        .Range("A16").Value = "ClsOrmBase"
        .Range("B16").Value = "Compl�t�"
        .Range("C16").Value = "Cursor"
        
        .Range("A17").Value = "clsEntityMappingFactory"
        .Range("B17").Value = "Compl�t�"
        .Range("C17").Value = "Cursor"
        
        ' Cr�er un tableau pour les composants Excel (seconde section)
        .Range("A19").Value = "Composants Excel"
        .Range("A19").Font.Bold = True
        .Range("A19").Font.Size = 14
        
        ' Format de tous les tableaux
        .Range("A6:C10").BorderAround xlContinuous
        .Range("A6:C6").BorderAround xlContinuous
        .Range("A13:C17").BorderAround xlContinuous
        .Range("A13:C13").BorderAround xlContinuous
        
        ' Ajout d'un pied de page avec statistiques
        .Range("A30").Value = "Couverture de Tests"
        .Range("A30").Font.Bold = True
        
        .Range("A31").Value = "Tests unitaires:"
        .Range("B31").Value = "95%"
        
        .Range("A32").Value = "Tests d'int�gration:"
        .Range("B32").Value = "90%"
        
        .Range("A33").Value = "Tests de performance:"
        .Range("B33").Value = "60%"
        
        .Range("A34").Value = "Tests de s�curit�:"
        .Range("B34").Value = "75%"
        
        .Range("A35").Value = "Tests ORM:"
        .Range("B35").Value = "85%"
    End With
    
    ' Cr�er un tableau Excel avanc� pour les derni�res mises � jour
    CreateUpdatesTable newWorkbook.Sheets("Plan de Situation"), 37
    
    ' Ajustement automatique des colonnes
    newWorkbook.Sheets("Plan de Situation").Columns("A:D").AutoFit
    
    ' Sauvegarde du nouveau classeur
    Dim savePath As String
    savePath = "D:\Dev\Apex_VBA_FRAMEWORK\APEX_PLAN_SITUATION.xlsx"
    newWorkbook.SaveAs savePath
    
    MsgBox "Le plan de situation a �t� export� avec succ�s vers " & savePath, vbInformation, "APEX Framework"
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

' Cr�ation d'un tableau Excel avanc� pour les derni�res mises � jour
Private Sub CreateUpdatesTable(ws As Worksheet, startRow As Integer)
    ' Titre du tableau
    ws.Range("A" & startRow).Value = "Derni�res Mises � Jour"
    ws.Range("A" & startRow).Font.Bold = True
    ws.Range("A" & startRow).Font.Size = 12
    
    ' En-t�tes du tableau
    ws.Range("A" & (startRow + 1)).Value = "Date"
    ws.Range("B" & (startRow + 1)).Value = "Description"
    ws.Range("C" & (startRow + 1)).Value = "Contributeur"
    ws.Range("A" & (startRow + 1) & ":C" & (startRow + 1)).Font.Bold = True
    
    ' Donn�es du tableau
    ws.Range("A" & (startRow + 2)).Value = "2024-04-14"
    ws.Range("B" & (startRow + 2)).Value = "Tests d'int�gration ORM"
    ws.Range("C" & (startRow + 2)).Value = "Cursor"
    
    ws.Range("A" & (startRow + 3)).Value = "2024-04-14"
    ws.Range("B" & (startRow + 3)).Value = "Factory des mappings d'entit�s"
    ws.Range("C" & (startRow + 3)).Value = "Cursor"
    
    ws.Range("A" & (startRow + 4)).Value = "2024-04-14"
    ws.Range("B" & (startRow + 4)).Value = "Tests avanc�s DBAccessor"
    ws.Range("C" & (startRow + 4)).Value = "Cursor"
    
    ws.Range("A" & (startRow + 5)).Value = "2024-04-13"
    ws.Range("B" & (startRow + 5)).Value = "Tests d'int�gration QueryBuilder"
    ws.Range("C" & (startRow + 5)).Value = "Cursor"
    
    ws.Range("A" & (startRow + 6)).Value = "2024-04-12"
    ws.Range("B" & (startRow + 6)).Value = "Accesseurs Excel"
    ws.Range("C" & (startRow + 6)).Value = "VSCode"
    
    ' Formatage du tableau
    ws.Range("A" & (startRow + 1) & ":C" & (startRow + 6)).BorderAround xlContinuous
    ws.Range("A" & (startRow + 1) & ":C" & (startRow + 1)).BorderAround xlContinuous
    
    ' Ajout d'une ligne de version
    ws.Range("A" & (startRow + 8)).Value = "Version: 2.3"
    ws.Range("A" & (startRow + 9)).Value = "Derni�re mise � jour: 2024-04-14"
End Sub

' Fonction d'entr�e pour le test
Public Sub TestApexExcel()
    CopyPlanSituationToExcel
End Sub