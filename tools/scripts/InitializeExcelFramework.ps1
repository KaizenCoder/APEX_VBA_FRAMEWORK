# Script d'initialisation du Framework APEX dans Excel
try {
    # Créer une instance d'Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    
    # Chemin du projet
    $projectPath = "D:\Dev\Apex_VBA_FRAMEWORK"
    
    # Créer un nouveau classeur
    $workbook = $excel.Workbooks.Add()
    
    # Activer les macros et l'accès au VBA
    $excel.EnableEvents = $true
    
    # Accéder au projet VBA
    $vbaProject = $workbook.VBProject
    
    # Définir l'ordre d'importation des modules
    $moduleGroups = @{
        # Core Framework
        Core            = @(
            "src\Interfaces\ILoggerBase.cls",
            "src\Interfaces\IFileAccessorBase.cls",
            "src\Interfaces\IDbAccessorBase.cls",
            "src\Interfaces\IQueryBuilder.cls",
            "src\Interfaces\IEntityMapping.cls",
            "src\Core\Constants.bas",
            "src\Core\Utilities.bas"
        )
        
        # Implementations de base
        BaseImpl        = @(
            "src\Implementations\clsLogger.cls",
            "src\Implementations\clsFileAccessor.cls",
            "src\Implementations\clsDbAccessor.cls"
        )
        
        # Interfaces Excel
        ExcelInterfaces = @(
            "src\Interfaces\IExcelHandlerBase.cls"
        )
        
        # Implementations Excel
        ExcelImpl       = @(
            "src\Implementations\clsExcelHandler.cls"
        )
        
        # Scripts spécifiques
        Scripts         = @(
            "src\Scripts\GeneratePlanSituation.cls",
            "src\Scripts\GeneratePlanSituationRunner.bas"
        )
    }
    
    # Fonction pour importer un groupe de modules
    function Import-ModuleGroup {
        param (
            [string]$groupName,
            [string[]]$modules
        )
        
        Write-Host "`nImportation du groupe: $groupName"
        foreach ($module in $modules) {
            $fullPath = Join-Path $projectPath $module
            if (Test-Path $fullPath) {
                try {
                    $vbaProject.VBComponents.Import($fullPath)
                    Write-Host "✓ Module importé: $module"
                }
                catch {
                    Write-Host "✗ Erreur lors de l'importation de $module : $_"
                }
            }
            else {
                Write-Host "! Module non trouvé: $module"
            }
        }
    }
    
    # Importer les modules dans l'ordre
    foreach ($group in $moduleGroups.GetEnumerator()) {
        Import-ModuleGroup -groupName $group.Key -modules $group.Value
    }
    
    # Sauvegarder le classeur avec macros
    $excelFile = "$projectPath\src\APEX_FRAMEWORK.xlsm"
    $workbook.SaveAs($excelFile, 52) # 52 = xlOpenXMLWorkbookMacroEnabled
    
    Write-Host "`nFramework APEX initialisé avec succès dans: $excelFile"
}
catch {
    Write-Host "Erreur lors de l'initialisation: $_"
}
finally {
    if ($workbook) {
        $workbook.Close($false)
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
} 