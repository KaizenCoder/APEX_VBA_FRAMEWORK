# Script de création du classeur Excel avec macros
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# Créer un nouveau classeur
$workbook = $excel.Workbooks.Add()

# Activer les macros
$excel.EnableEvents = $true

# Chemin du projet
$projectPath = "D:\Dev\Apex_VBA_FRAMEWORK"

# Importer les modules VBA
$vbaProject = $workbook.VBProject
$moduleFiles = @(
    "$projectPath\src\APEX_FRAMEWORK.bas",
    "$projectPath\src\Scripts\GeneratePlanSituationRunner.bas",
    "$projectPath\src\Scripts\GeneratePlanSituation.cls",
    "$projectPath\src\Interfaces\IExcelHandlerBase.cls",
    "$projectPath\src\Implementations\clsExcelHandler.cls"
)

foreach ($moduleFile in $moduleFiles) {
    if (Test-Path $moduleFile) {
        $vbaProject.VBComponents.Import($moduleFile)
        Write-Host "Module importé: $moduleFile"
    }
    else {
        Write-Host "Module non trouvé: $moduleFile"
    }
}

# Sauvegarder le classeur avec macros
$workbook.SaveAs("$projectPath\src\APEX_FRAMEWORK.xlsm", 52) # 52 = xlOpenXMLWorkbookMacroEnabled

# Fermer Excel
$workbook.Close($false)
$excel.Quit()

# Libérer les ressources COM
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "Création du classeur APEX_FRAMEWORK.xlsm terminée" 