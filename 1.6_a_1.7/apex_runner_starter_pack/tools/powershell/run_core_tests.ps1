# Script to launch core test suite
$ErrorActionPreference = "Stop"
$logPath = "D:\Dev\Apex_VBA_FRAMEWORK\logs\core_tests_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$mdFallbackPath = "D:\Dev\Apex_VBA_FRAMEWORK\logs\test_fallback_$(Get-Date -Format 'yyyyMMdd_HHmmss').md"

function Write-Log {
    param($Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Tee-Object -FilePath $logPath -Append
}

Write-Log "Starting Apex Core test suite..."

try {
    # Tentative de création d'un objet Excel
    $excel = New-Object -ComObject Excel.Application
    if ($null -eq $excel) {
        throw "Excel COM object could not be created"
    }
    
    Write-Log "Excel COM object successfully created"
    $excel.Visible = $false
    
    # Exécution des tests via COM
    Write-Log "Running core tests..."
    $result = $excel.Run("Test_CoreRunner_Execution")
    
    Write-Log "Tests completed successfully"
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
}
catch {
    Write-Log "ERROR: $_"
    
    # Création du rapport de fallback en Markdown
    @"
# Apex Core Tests Fallback Report
## Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
### Error Details
\`\`\`
$_
\`\`\`
### Status
- ❌ Excel COM object unavailable
- ⚠️ Manual verification required
### Next Steps
1. Verify Excel installation
2. Check COM server registration
3. Run tests manually through Excel VBA IDE
"@ | Out-File -FilePath $mdFallbackPath -Encoding UTF8
    
    Write-Log "Fallback report generated at: $mdFallbackPath"
    exit 1
}
