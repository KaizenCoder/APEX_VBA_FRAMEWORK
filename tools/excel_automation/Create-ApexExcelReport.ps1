# -----------------------------------------------------------------------------
# Script: Create-ApexExcelReport.ps1
# Description: Generation d'un classeur Excel a partir du framework APEX sans xlwings
# Author: APEX Framework Team
# Date: 2024-04-14
# Version: 1.0
# -----------------------------------------------------------------------------

# Respect des consignes d'encodage APEX (voir docs/requirements/powershell_encoding.md)
# Ce script est encode en UTF-8 sans BOM

# Configuration des constantes
$WORKSPACE_ROOT = "D:\Dev\Apex_VBA_FRAMEWORK"
$PLAN_SITUATION_PATH = Join-Path $WORKSPACE_ROOT "docs\implementation\PLAN_SITUATION_2024_04_14.md"
$OUTPUT_PATH = Join-Path $WORKSPACE_ROOT "APEX_PLAN_SITUATION_NATIVE.xlsx"
$VBA_MODULE_PATH = Join-Path $WORKSPACE_ROOT "tools\excel_automation\ApexExcelGenerator.bas"

Write-Host "[...] Initialisation du rapport Excel avec le framework APEX..." -ForegroundColor Cyan

# Verification des prerequis
if (-not (Test-Path $PLAN_SITUATION_PATH)) {
    Write-Host "[X] Le fichier du plan de situation est introuvable: $PLAN_SITUATION_PATH" -ForegroundColor Red
    exit 1
}

# Creation du module VBA APEX qui sera injecte dans Excel
# Version modifiée pour fonctionner sans dépendance directe à l'add-in APEX
$VBA_MODULE_CONTENT = @'
Attribute VB_Name = "ApexExcelGenerator"
Option Explicit

' Module de generation automatique de rapport Excel
' Version autonome sans dépendance à l'add-in APEX
' Date: 14/04/2025

' Point d''entree principal
Public Sub Main()
    GenerateSituationReport
End Sub

' Genere le rapport du plan de situation
Public Sub GenerateSituationReport()
    On Error Resume Next
    
    ' Renommer la première feuille
    ThisWorkbook.Sheets(1).Name = "Plan de Situation"
    
    ' Configuration de la feuille
    ConfigureWorksheet
    
    ' Insertion des donnees structurees
    InsertDatabaseComponents
    InsertExcelComponents
    InsertMetrics
    InsertHistory
    
    ' Auto-ajustement des colonnes
    ThisWorkbook.Sheets("Plan de Situation").Columns("A:D").AutoFit
    
    ' Message de confirmation
    MsgBox "Rapport APEX genere avec succes!", vbInformation, "APEX Framework"
End Sub

' Configure la mise en page de la feuille
Private Sub ConfigureWorksheet()
    ' En-tete principal
    Dim headerRange As Range
    Set headerRange = ThisWorkbook.Sheets("Plan de Situation").Range("A1")
    headerRange.Value = "Plan de Situation APEX Framework - 2024-04-14"
    headerRange.Font.Size = 16
    headerRange.Font.Bold = True
    
    ' Description
    Dim descRange As Range
    Set descRange = ThisWorkbook.Sheets("Plan de Situation").Range("A2")
    descRange.Value = "Etat d''avancement du framework APEX - Document genere"
    descRange.Font.Italic = True
End Sub

' Insere la section des composants de base de donnees
Private Sub InsertDatabaseComponents()
    With ThisWorkbook.Sheets("Plan de Situation")
        ' Titre de section
        .Range("A4").Value = "Composants Database"
        .Range("A4").Font.Size = 14
        .Range("A4").Font.Bold = True
        
        ' Sous-section Interfaces
        .Range("A6").Value = "1. Interfaces"
        .Range("A6").Font.Bold = True
        
        ' En-tetes du tableau
        .Range("A7").Value = "Composant"
        .Range("B7").Value = "Etat"
        .Range("C7").Value = "Contributeur"
        .Range("A7:C7").Font.Bold = True
        
        ' Donnees
        .Range("A8").Value = "IDbDriver"
        .Range("B8").Value = "Complete"
        .Range("C8").Value = "Cursor"
        
        .Range("A9").Value = "IQueryBuilder"
        .Range("B9").Value = "Complete"
        .Range("C9").Value = "Cursor"
        
        .Range("A10").Value = "IDBAccessorBase"
        .Range("B10").Value = "Complete"
        .Range("C10").Value = "VSCode"
        
        .Range("A11").Value = "IEntityMapping"
        .Range("B11").Value = "Complete"
        .Range("C11").Value = "Cursor"
        
        ' Formatage du tableau
        .Range("A7:C11").Borders.LineStyle = xlContinuous
        .Range("A7:C7").Borders.Weight = xlMedium
        ' Centrage de la colonne Contributeur pour un meilleur alignement visuel
        .Range("C7:C11").HorizontalAlignment = xlCenter
        
        ' Sous-section Implementations
        .Range("A13").Value = "2. Implementations"
        .Range("A13").Font.Bold = True
        
        ' En-tetes du tableau
        .Range("A14").Value = "Composant"
        .Range("B14").Value = "Etat"
        .Range("C14").Value = "Contributeur"
        .Range("A14:C14").Font.Bold = True
        
        ' Donnees
        .Range("A15").Value = "clsDBAccessor"
        .Range("B15").Value = "Complete"
        .Range("C15").Value = "VSCode"
        
        .Range("A16").Value = "clsSqlQueryBuilder"
        .Range("B16").Value = "Complete"
        .Range("C16").Value = "Cursor"
        
        .Range("A17").Value = "ClsOrmBase"
        .Range("B17").Value = "Complete"
        .Range("C17").Value = "Cursor"
        
        .Range("A18").Value = "clsEntityMappingFactory"
        .Range("B18").Value = "Complete"
        .Range("C18").Value = "Cursor"
        
        ' Formatage du tableau
        .Range("A14:C18").Borders.LineStyle = xlContinuous
        .Range("A14:C14").Borders.Weight = xlMedium
        ' Centrage de la colonne Contributeur pour un meilleur alignement visuel
        .Range("C14:C18").HorizontalAlignment = xlCenter
    End With
End Sub

' Insere la section des composants Excel
Private Sub InsertExcelComponents()
    With ThisWorkbook.Sheets("Plan de Situation")
        ' Titre de section
        .Range("A20").Value = "Composants Excel"
        .Range("A20").Font.Size = 14
        .Range("A20").Font.Bold = True
        
        ' Sous-section Interfaces
        .Range("A22").Value = "1. Interfaces"
        .Range("A22").Font.Bold = True
        
        ' En-tetes du tableau
        .Range("A23").Value = "Composant"
        .Range("B23").Value = "Etat"
        .Range("C23").Value = "Contributeur"
        .Range("A23:C23").Font.Bold = True
        
        ' Donnees
        .Range("A24").Value = "IWorkbookAccessor"
        .Range("B24").Value = "Complete"
        .Range("C24").Value = "VSCode"
        
        .Range("A25").Value = "ISheetAccessor"
        .Range("B25").Value = "Complete"
        .Range("C25").Value = "Cursor"
        
        .Range("A26").Value = "ITableAccessor"
        .Range("B26").Value = "Complete"
        .Range("C26").Value = "Cursor"
        
        .Range("A27").Value = "IRangeAccessor"
        .Range("B27").Value = "Complete"
        .Range("C27").Value = "VSCode"
        
        .Range("A28").Value = "ICellAccessor"
        .Range("B28").Value = "Complete"
        .Range("C28").Value = "Cursor"
        
        ' Formatage du tableau
        .Range("A23:C28").Borders.LineStyle = xlContinuous
        .Range("A23:C23").Borders.Weight = xlMedium
        ' Centrage de la colonne Contributeur pour un meilleur alignement visuel
        .Range("C23:C28").HorizontalAlignment = xlCenter
        
        ' Sous-section Implementations
        .Range("A30").Value = "2. Implementations"
        .Range("A30").Font.Bold = True
        
        ' En-tetes du tableau
        .Range("A31").Value = "Composant"
        .Range("B31").Value = "Etat"
        .Range("C31").Value = "Contributeur"
        .Range("A31:C31").Font.Bold = True
        
        ' Donnees
        .Range("A32").Value = "clsExcelWorkbookAccessor"
        .Range("B32").Value = "Complete"
        .Range("C32").Value = "VSCode"
        
        .Range("A33").Value = "clsExcelSheetAccessor"
        .Range("B33").Value = "Complete"
        .Range("C33").Value = "Cursor"
        
        .Range("A34").Value = "clsExcelTableAccessor"
        .Range("B34").Value = "Complete"
        .Range("C34").Value = "Cursor"
        
        .Range("A35").Value = "clsExcelRangeAccessor"
        .Range("B35").Value = "Complete"
        .Range("C35").Value = "VSCode"
        
        .Range("A36").Value = "clsExcelCellAccessor"
        .Range("B36").Value = "Complete"
        .Range("C36").Value = "Cursor"
        
        ' Formatage du tableau
        .Range("A31:C36").Borders.LineStyle = xlContinuous
        .Range("A31:C31").Borders.Weight = xlMedium
        ' Centrage de la colonne Contributeur pour un meilleur alignement visuel
        .Range("C31:C36").HorizontalAlignment = xlCenter
    End With
End Sub

' Insere les metriques de projet
Private Sub InsertMetrics()
    With ThisWorkbook.Sheets("Plan de Situation")
        ' Titre de section
        .Range("A38").Value = "Metriques du projet"
        .Range("A38").Font.Size = 14
        .Range("A38").Font.Bold = True
        
        ' Sous-section Tests
        .Range("A40").Value = "Couverture de Tests"
        .Range("A40").Font.Bold = True
        
        ' Donnees
        .Range("A41").Value = "Tests unitaires:"
        .Range("B41").Value = "95%"
        
        .Range("A42").Value = "Tests d''integration:"
        .Range("B42").Value = "87%"
        
        .Range("A43").Value = "Tests de performance:"
        .Range("B43").Value = "65%"
        
        .Range("A44").Value = "Tests de securite:"
        .Range("B44").Value = "76%"
        
        ' Formatage
        .Range("A41:B44").Borders.LineStyle = xlContinuous
    End With
End Sub

' Insere l''historique des mises a jour
Private Sub InsertHistory()
    With ThisWorkbook.Sheets("Plan de Situation")
        ' Titre de section
        .Range("A46").Value = "Historique des mises a jour"
        .Range("A46").Font.Size = 14
        .Range("A46").Font.Bold = True
        
        ' En-tetes du tableau
        .Range("A48").Value = "Date"
        .Range("B48").Value = "Description"
        .Range("C48").Value = "Contributeur"
        .Range("A48:C48").Font.Bold = True
        
        ' Donnees
        .Range("A49").Value = "2024-04-14"
        .Range("B49").Value = "Mise a jour du module ORM"
        .Range("C49").Value = "Cursor"
        
        .Range("A50").Value = "2024-04-14"
        .Range("B50").Value = "Integration des accesseurs Excel"
        .Range("C50").Value = "VSCode"
        
        .Range("A51").Value = "2024-04-13"
        .Range("B51").Value = "Ajout des tests d''integration"
        .Range("C51").Value = "Cursor"
        
        .Range("A52").Value = "2024-04-12"
        .Range("B52").Value = "Definition des interfaces principales"
        .Range("C52").Value = "VSCode"
        
        ' Formatage du tableau
        .Range("A48:C52").Borders.LineStyle = xlContinuous
        .Range("A48:C48").Borders.Weight = xlMedium
        
        ' Insertion de la version
        .Range("A54").Value = "Version du framework: 2.3"
        .Range("A54").Font.Bold = True
        
        .Range("A55").Value = "Genere le: " & Format(Now, "yyyy-mm-dd hh:mm")
    End With
End Sub
'@

# Fonction pour ecrire le module VBA avec encodage UTF-8
function Set-FileContent {
    param(
        [string]$Path,
        [string]$Content
    )
    
    try {
        $utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBomEncoding)
        Write-Host "[+] Fichier cree avec succes: $Path" -ForegroundColor Green
    }
    catch {
        Write-Host "[X] Erreur lors de la creation du fichier: $_" -ForegroundColor Red
    }
}

# Creer le module VBA temporaire
Set-FileContent -Path $VBA_MODULE_PATH -Content $VBA_MODULE_CONTENT

# Executer Excel et injecter automatiquement le code VBA
try {
    Write-Host "[...] Creation d'une instance Excel..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    Write-Host "[...] Creation d'un nouveau classeur..." -ForegroundColor Yellow
    $workbook = $excel.Workbooks.Add()
    
    Write-Host "[...] Importation du module VBA APEX..." -ForegroundColor Yellow
    $vbProject = $workbook.VBProject
    
    # Verification de l'acces au VBA Project
    try {
        $null = $vbProject.Name
    }
    catch {
        Write-Host "[X] Acces au VBA Project refuse. Veuillez activer 'Acces approuve au modele d'objet VBA' dans les parametres de securite Excel." -ForegroundColor Red
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        exit 1
    }
    
    # Import du module VBA
    $vbComponent = $vbProject.VBComponents.Import($VBA_MODULE_PATH)
    
    # Execution de la macro principale
    Write-Host "[...] Generation du rapport Excel avec APEX..." -ForegroundColor Yellow
    $excel.Run("Main")
    
    # Sauvegarde du classeur
    Write-Host "[...] Sauvegarde du classeur..." -ForegroundColor Yellow
    $workbook.SaveAs($OUTPUT_PATH)
    
    # Nettoyage
    $workbook.Close($true)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    Remove-Item -Path $VBA_MODULE_PATH -Force
    
    Write-Host "[+] Classeur Excel genere avec succes: $OUTPUT_PATH" -ForegroundColor Green
    
    # Ouverture du classeur genere
    Write-Host "[...] Ouverture du classeur..." -ForegroundColor Yellow
    Start-Process $OUTPUT_PATH
}
catch {
    Write-Host "[X] Erreur lors de la generation du classeur: $_" -ForegroundColor Red
}
finally {
    # Nettoyage supplementaire
    if (Test-Path $VBA_MODULE_PATH) {
        Remove-Item -Path $VBA_MODULE_PATH -Force
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}