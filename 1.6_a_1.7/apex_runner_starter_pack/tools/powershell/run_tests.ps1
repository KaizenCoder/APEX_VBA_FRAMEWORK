# Execute Runner unit tests
Write-Host "=== Exécution des tests unitaires Runner ===" -ForegroundColor Cyan

$ErrorActionPreference = "Stop"
$baseDir = Split-Path -Parent $PSScriptRoot | Split-Path -Parent

# Find Excel path
$excelPath = "${env:ProgramFiles}\Microsoft Office\root\Office16\EXCEL.EXE"
if (-not (Test-Path $excelPath)) {
    $excelPath = "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE"
}
if (-not (Test-Path $excelPath)) {
    throw "Could not find Excel installation"
}

# Ensure required directories exist
$requiredDirs = @(
    (Join-Path $baseDir "logs"),
    (Join-Path $baseDir "output")
)

foreach ($dir in $requiredDirs) {
    if (-not (Test-Path $dir)) {
        New-Item -Path $dir -ItemType Directory | Out-Null
        Write-Host "Created directory: $dir"
    }
}

# Run the tests
try {
    Write-Host "`nLancement des tests..."
    & "$excelPath" "/r `"$baseDir\tests\Core\modTestRunner_ApexCore.bas`""
    
    # Wait for log files to be created
    Start-Sleep -Seconds 2
    
    # Validate test results
    $logPath = Join-Path $baseDir "logs\CoreDemoRunner.log"
    $statePath = Join-Path $baseDir "output\demo_runner_state.json"
    
    $success = $true
    
    if (Test-Path $logPath) {
        Write-Host "Log file created: " -NoNewline
        Write-Host "OK" -ForegroundColor Green
        
        $logContent = Get-Content $logPath -Raw
        if ($logContent -match "Execution terminée sans erreur") {
            Write-Host "Test execution successful: " -NoNewline
            Write-Host "OK" -ForegroundColor Green
        }
        else {
            Write-Host "Test execution failed: " -NoNewline
            Write-Host "ERROR" -ForegroundColor Red
            $success = $false
        }
    }
    else {
        Write-Host "Log file missing: " -NoNewline
        Write-Host "ERROR" -ForegroundColor Red
        $success = $false
    }
    
    if (Test-Path $statePath) {
        Write-Host "State file created: " -NoNewline
        Write-Host "OK" -ForegroundColor Green
    }
    else {
        Write-Host "State file missing: " -NoNewline
        Write-Host "ERROR" -ForegroundColor Red
        $success = $false
    }
    
    if ($success) {
        Write-Host "`nAll tests completed successfully!" -ForegroundColor Green
        exit 0
    }
    else {
        Write-Host "`nTests failed!" -ForegroundColor Red
        exit 1
    }
}
catch {
    Write-Host "Error executing tests: $_" -ForegroundColor Red
    exit 1
}