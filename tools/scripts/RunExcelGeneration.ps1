# Script d'exécution de la macro VBA
try {
    # Créer une instance d'Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    
    # Chemin du projet
    $projectPath = "D:\Dev\Apex_VBA_FRAMEWORK"
    
    # Vérifier si le fichier Markdown existe
    $mdFile = "$projectPath\docs\implementation\PLAN_SITUATION_2024_04_14.md"
    if (-not (Test-Path $mdFile)) {
        throw "Fichier Markdown non trouvé: $mdFile"
    }
    
    # Lire le contenu du fichier Markdown
    $mdContent = Get-Content $mdFile -Raw
    
    # Créer un nouveau classeur
    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = "Plan Situation"
    
    # En-têtes
    $worksheet.Cells(1, 1) = "Date"
    $worksheet.Cells(1, 2) = "Description"
    $worksheet.Cells(1, 3) = "Contributeur"
    
    # Formatage des en-têtes
    $headerRange = $worksheet.Range("A1:C1")
    $headerRange.Font.Bold = $true
    $headerRange.Interior.ColorIndex = 15
    
    # Traiter le contenu Markdown
    $lines = $mdContent -split "`n"
    $row = 2
    
    foreach ($line in $lines) {
        if ($line -match "\|\s*(.*?)\s*\|\s*(.*?)\s*\|\s*(.*?)\s*\|") {
            if (-not ($line -match "---" -or $line -match "Date")) {
                $worksheet.Cells($row, 1) = $matches[1].Trim()
                $worksheet.Cells($row, 2) = $matches[2].Trim()
                $worksheet.Cells($row, 3) = $matches[3].Trim()
                $row++
            }
        }
    }
    
    # Formatage final
    $usedRange = $worksheet.UsedRange
    $usedRange.Columns.AutoFit()
    $usedRange.Borders.LineStyle = 1
    
    # Sauvegarder
    $excelFile = "$projectPath\docs\implementation\PLAN_SITUATION_2024_04_14.xlsx"
    $workbook.SaveAs($excelFile)
    
    Write-Host "Plan de situation généré avec succès: $excelFile"
}
catch {
    Write-Host "Erreur: $_"
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