Attribute VB_Name = "InitializeApexExcel"
Option Explicit

' ==========================================================================
' Module    : InitializeApexExcel
' Purpose   : Module d'initialisation du framework APEX Excel
' Author    : APEX Framework Team
' Date      : 2024-04-14
' ==========================================================================

Public Sub InitializeFramework()
    ' Vérifier si Excel est en mode Automation
    If Not Application.Interactive Then
        MsgBox "Excel doit être en mode interactif pour initialiser le framework", vbCritical
        Exit Sub
    End If
    
    ' Activer l'accès au modèle VBA
    Application.AutomationSecurity = msoAutomationSecurityLow
    
    ' Activer les références nécessaires
    AddRequiredReferences
    
    ' Créer une instance de la factory
    Dim factory As New ModExcelFactory
    
    ' Créer un classeur de test pour vérifier les composants
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
    
    MsgBox "Framework APEX Excel initialisé avec succès", vbInformation
    Exit Sub
    
ErrorHandler:
    If Not wb Is Nothing Then wb.Close False
    MsgBox "Erreur lors de l'initialisation du framework : " & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub AddRequiredReferences()
    Dim ref As Object
    
    ' Vérifier si les références sont déjà présentes
    For Each ref In ThisWorkbook.VBProject.References
        Select Case ref.Name
            Case "Excel"
                ' Déjà présente
            Case "VBA"
                ' Déjà présente
            Case "stdole"
                ' Déjà présente
            Case Else
                ' Autres références à ajouter si nécessaire
        End Select
    Next ref
End Sub 