# Validate that Apex runner executes correctly
Write-Host "=== Validation Runner Apex ===" -ForegroundColor Cyan

$baseDir = Split-Path -Parent $PSScriptRoot | Split-Path -Parent
$runnerPath = Join-Path $baseDir "src\Core\Runners\clsCoreDemoRunner.cls"
$logPath = Join-Path $baseDir "logs\CoreDemoRunner.log"
$statePath = Join-Path $baseDir "output\demo_runner_state.json"

$allValid = $true

# Vérification des fichiers requis
@(
    @{Path = $runnerPath; Desc = "Runner class file" },
    @{Path = $logPath; Desc = "Log file" },
    @{Path = $statePath; Desc = "State output file" }
) | ForEach-Object {
    if (Test-Path $_.Path) {
        Write-Host "$($_.Desc) exists: " -NoNewline
        Write-Host "OK" -ForegroundColor Green
    }
    else {
        Write-Host "$($_.Desc) missing: " -NoNewline
        Write-Host "ERROR" -ForegroundColor Red
        $allValid = $false
    }
}

# Vérification du contenu des logs
if (Test-Path $logPath) {
    $logContent = Get-Content $logPath -Raw
    if ($logContent -match "Execution terminée sans erreur") {
        Write-Host "Log contains success message: " -NoNewline
        Write-Host "OK" -ForegroundColor Green
    }
    else {
        Write-Host "Log missing success message: " -NoNewline
        Write-Host "ERROR" -ForegroundColor Red
        $allValid = $false
    }
}

# Vérification du fichier d'état JSON
if (Test-Path $statePath) {
    try {
        $state = Get-Content $statePath | ConvertFrom-Json
        if ($state.status -eq "success") {
            Write-Host "State file valid: " -NoNewline
            Write-Host "OK" -ForegroundColor Green
        }
        else {
            Write-Host "State file invalid: " -NoNewline
            Write-Host "ERROR" -ForegroundColor Red
            $allValid = $false
        }
    }
    catch {
        Write-Host "State file parse error: " -NoNewline
        Write-Host "ERROR" -ForegroundColor Red
        $allValid = $false
    }
}

if ($allValid) {
    Write-Host "`nValidation completed successfully!" -ForegroundColor Green
    exit 0
}
else {
    Write-Host "`nValidation failed!" -ForegroundColor Red
    exit 1
}
