@{
    # Informations de base du module
    RootModule = 'Apex.SessionManager.psm1'
    ModuleVersion = '1.0.0'
    GUID = '12345678-1234-1234-1234-123456789012'  # Générer un nouveau GUID pour la production
    Author = 'APEX Framework Team'
    CompanyName = 'APEX Framework'
    Copyright = '(c) 2025 APEX Framework. Tous droits réservés.'
    Description = 'Module de gestion des sessions de développement pour APEX VBA Framework'
    
    # Version minimale de PowerShell requise
    PowerShellVersion = '5.1'
    
    # Fonctions à exporter
    FunctionsToExport = @(
        'New-ApexSession',
        'Add-TaskToSession',
        'Complete-ApexSession',
        'Get-CurrentSession'
    )
    
    # Alias à exporter
    AliasesToExport = @()
    
    # Variables à exporter
    VariablesToExport = @()
    
    # Cmdlets à exporter
    CmdletsToExport = @()
    
    # Tags pour la découverte du module
    PrivateData = @{
        PSData = @{
            Tags = @('APEX', 'VBA', 'Framework', 'Session', 'Development')
            LicenseUri = 'https://github.com/votre-repo/LICENSE'
            ProjectUri = 'https://github.com/votre-repo'
            ReleaseNotes = 'Version initiale du module de gestion des sessions APEX'
        }
    }
} 