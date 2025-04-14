Attribute VB_Name = "InitializeApexExcel"
Option Explicit

' ==========================================================================
' Module    : InitializeApexExcel
' Purpose   : Module d'initialisation du framework APEX Excel
' Author    : APEX Framework Team
' Date      : 2024-04-14
' ==========================================================================

Public Sub InitializeFramework()
    ' V�rifier si Excel est en mode Automation
    If Not Application.Interactive Then
        MsgBox "Excel doit �tre en mode interactif pour initialiser le framework", vbCritical
        Exit Sub
    End If
    
    ' Activer l'acc�s au mod�le VBA
    Application.AutomationSecurity = msoAutomationSecurityLow
    
    ' Activer les r�f�rences n�cessaires
    AddRequiredReferences
    
    ' Cr�er une instance de la factory
    Dim factory As New ModExcelFactory
    
    ' Cr�er un classeur de test pour v�rifier les composants
    Dim wb As Workbook
    Set wb = Application.Workbooks.Add
    
    ' Tester les composants
    On Error GoTo ErrorHandler
    
    ' Test WorkbookAccessor
    Dim workbookAccessor As IWorkbookAccessor
    Set workbookAccessor = factory.CreateWorkbookAccessor(wb)
    
    ' Test SheetAccessor
    Dim sheetAccessor As ISheetAccessor
    Set sheetAccessor = workbookAccessor.GetSheet(wb.Sheets(1).Name)
    
    ' Test CellAccessor
    Dim cellAccessor As ICellAccessor
    Set cellAccessor = sheetAccessor.GetCell(1, 1)
    
    ' Test RangeAccessor
    Dim rangeAccessor As IRangeAccessor
    Set rangeAccessor = sheetAccessor.GetRange("A1:B2")
    
    ' Nettoyage
    wb.Close False
    
    MsgBox "Framework APEX Excel initialis� avec succ�s", vbInformation
    Exit Sub
    
ErrorHandler:
    If Not wb Is Nothing Then wb.Close False
    MsgBox "Erreur lors de l'initialisation du framework : " & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub AddRequiredReferences()
    Dim ref As Object
    
    ' V�rifier si les r�f�rences sont d�j� pr�sentes
    For Each ref In ThisWorkbook.VBProject.References
        Select Case ref.Name
            Case "Excel"
                ' D�j� pr�sente
            Case "VBA"
                ' D�j� pr�sente
            Case "stdole"
                ' D�j� pr�sente
            Case Else
                ' Autres r�f�rences � ajouter si n�cessaire
        End Select
    Next ref
End Sub 