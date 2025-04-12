function Get-EncodingPatterns {
    $patternsFile = Join-Path $PSScriptRoot "..\config\patterns.json"
    $patterns = Get-Content $patternsFile -Raw | ConvertFrom-Json
    
    $result = @()
    foreach ($category in $patterns.PSObject.Properties) {
        foreach ($pattern in $category.Value.PSObject.Properties) {
            $result += @{
                Pattern = $pattern.Name
                Replacement = $pattern.Value
            }
        }
    }
    return $result
}

function Test-PowerShellSyntax {
    param([string]$Content)
    
    $errors = $null
    $null = [System.Management.Automation.PSParser]::Tokenize($Content, [ref]$errors)
    return $errors.Count -eq 0
}

function Backup-File {
    param([string]$Path)
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupPath = "$Path.$timestamp.bak"
    Copy-Item $Path $backupPath
    return $backupPath
}

Export-ModuleMember -Function Get-EncodingPatterns, Test-PowerShellSyntax, Backup-File 