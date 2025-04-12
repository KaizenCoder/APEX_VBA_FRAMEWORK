# Module ApexWSLBridge
# Référence: chat_051 (2024-04-11 17:00)
# Source: chat_050 (Pipeline validation)

# Fonction de test d'encodage des fichiers
function Test-FileEncoding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ProjectRoot
    )

    # Vérifier si le chemin existe
    if (-not (Test-Path $ProjectRoot)) {
        Write-Error "Le chemin spécifié n'existe pas : $ProjectRoot"
        return @{
            HasErrors = $true
            InvalidFiles = @()
            Error = "Chemin invalide"
        }
    }

    # Structure pour les résultats
    $results = @{
        HasErrors = $false
        InvalidFiles = @()
        Error = $null
    }

    try {
        # Récupérer tous les fichiers du projet
        $files = Get-ChildItem -Path $ProjectRoot -Recurse -File |
            Where-Object { $_.Extension -match '\.(ps1|psm1|md|txt|py|sh)$' }

        foreach ($file in $files) {
            try {
                # Vérifier si le fichier est accessible
                if ((Get-Item $file.FullName).IsReadOnly) {
                    $results.HasErrors = $true
                    $results.InvalidFiles += @{
                        Path = $file.FullName
                        Encoding = "Fichier en lecture seule"
                    }
                    continue
                }

                # Lire les premiers octets pour détecter le BOM
                $stream = [System.IO.File]::OpenRead($file.FullName)
                $bytes = New-Object byte[] 3
                $read = $stream.Read($bytes, 0, 3)
                $stream.Close()

                $hasBOM = $read -eq 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF

                if ($hasBOM) {
                    $results.HasErrors = $true
                    $results.InvalidFiles += @{
                        Path = $file.FullName
                        Encoding = "UTF-8 with BOM"
                    }
                }
            }
            catch {
                $results.HasErrors = $true
                $results.InvalidFiles += @{
                    Path = $file.FullName
                    Encoding = "Erreur d'accès: $_"
                }
            }
        }
    }
    catch {
        Write-Error "Erreur lors de la validation : $_"
        $results.HasErrors = $true
        $results.Error = $_.Exception.Message
    }

    return $results
}

# Exporter les fonctions
Export-ModuleMember -Function Test-FileEncoding 