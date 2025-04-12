# DEPRECATED: Ce script est obsolÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨te. Utilisez 'tools/python/generate_apex_addin.py' ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  la place.
# Voir [Guide de migration](docs/MIGRATION_GUIDE.md#scripts-de-build)

#==========================================================================
# Script : CreateApexAddIn.ps1
# Version : 1.0
# Purpose : CrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation automatisÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e du fichier ApexVBAFramework.xlam (OBSOLÃƒÆ’Ã†â€™Ãƒâ€¹Ã¢â‚¬Â TE)
# Date : 12/04/2025
#==========================================================================

Write-Host "[AVERTISSEMENT] Ce script (CreateApexAddIn.ps1) est obsolÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨te et sera supprimÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© prochainement." -ForegroundColor Yellow
Write-Host "Utilisez le script Python unifiÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© : python tools/python/generate_apex_addin.py" -ForegroundColor Yellow
Write-Host ""
Start-Sleep -Seconds 3

# Affichage du titre
Write-Host "===== CRÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â°ATION DU FICHIER ApexVBAFramework.xlam =====" -ForegroundColor Cyan
Write-Host ""

# --- Variables ---
$LogFile = "create_addin_log.txt"
$AddinName = "ApexVBAFramework.xlam"
$ProjectName = "ApexVbaFramework"
$OutputDir = "$env:APPDATA\Microsoft\AddIns"  # Dossier par dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©faut des add-ins Excel
$TempDir = "temp_addin"
$SourceFolders = @(
    "apex-core",
    "apex-metier",
    "apex-ui"
)
$CoreFiles = @(
    "apex-core\clsLogger.cls", 
    "apex-core\modConfigManager.bas",
    "apex-core\modVersionInfo.bas",
    "apex-core\utils\modFileUtils.bas",
    "apex-core\utils\modTextUtils.bas",
    "apex-core\utils\modDateUtils.bas"
)

# --- Initialisation ---
Write-Host "[INFO] Initialisation de la crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation..." -ForegroundColor Green
if (Test-Path $LogFile) { Remove-Item $LogFile -Force }
"DÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©but de la crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation: $(Get-Date)" | Out-File $LogFile

# --- VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification des prÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©requis ---
Write-Host "[INFO] VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification des prÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©requis..." -ForegroundColor Green

# VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier que les dossiers source existent
foreach ($folder in $SourceFolders) {
    if (-not (Test-Path $folder)) {
        Write-Host "[ERREUR] Dossier source non trouvÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $folder" -ForegroundColor Red
        "ERREUR: Dossier source non trouvÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $folder" | Out-File $LogFile -Append
        exit 1
    }
}

# VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier que les fichiers essentiels existent
foreach ($file in $CoreFiles) {
    if (-not (Test-Path $file)) {
        Write-Host "[ERREUR] Fichier essentiel non trouvÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $file" -ForegroundColor Red
        "ERREUR: Fichier essentiel non trouvÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $file" | Out-File $LogFile -Append
        exit 1
    }
}

# VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier que les fichiers essentiels ne sont pas vides
foreach ($file in $CoreFiles) {
    $fileContent = Get-Content -Path $file -ErrorAction SilentlyContinue
    if ($null -eq $fileContent -or $fileContent.Count -eq 0) {
        Write-Host "[ERREUR] Fichier essentiel vide: $file" -ForegroundColor Red
        "ERREUR: Fichier essentiel vide: $file" | Out-File $LogFile -Append
        exit 1
    }
}

# Tester les fichiers avec xlwings si disponible
$XlwingsExe = ".\xlwings.exe"
if (Test-Path $XlwingsExe) {
    Write-Host "[INFO] Test des fichiers avec xlwings..." -ForegroundColor Green
    
    try {
        # CrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©er un script Python temporaire pour tester les fichiers
        $PythonTestScript = @"
import sys
import os

def check_vba_files(file_paths):
    missing_files = []
    empty_files = []
    
    for path in file_paths:
        if not os.path.exists(path):
            missing_files.append(path)
        elif os.path.getsize(path) == 0:
            empty_files.append(path)
    
    return missing_files, empty_files

# Liste des fichiers ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  vÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier
files = [
$(foreach ($file in $CoreFiles) { "    '$file'," }) 
]

missing, empty = check_vba_files(files)

if missing or empty:
    print("ERREUR: ProblÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨mes dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©tectÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©s avec les fichiers VBA:")
    if missing:
        print("Fichiers manquants:")
        for f in missing:
            print(f"  - {f}")
    if empty:
        print("Fichiers vides:")
        for f in empty:
            print(f"  - {f}")
    sys.exit(1)
else:
    print("Tous les fichiers VBA sont prÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©sents et non vides.")
    sys.exit(0)
"@
        $PythonTestPath = "$TempDir\check_vba_files.py"
        $PythonTestScript | Out-File -FilePath $PythonTestPath -Encoding utf8
        
        # ExÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cuter le script via xlwings
        $XlwingsOutput = & $XlwingsExe run "$PythonTestPath" 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Host "[ERREUR] Test xlwings des fichiers VBA ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©chouÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©:" -ForegroundColor Red
            $XlwingsOutput | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
            "ERREUR: Test xlwings ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©chouÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $XlwingsOutput" | Out-File $LogFile -Append
            exit 1
        }
        else {
            Write-Host "[INFO] Test xlwings des fichiers VBA rÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ussi." -ForegroundColor Green
            "Test xlwings rÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ussi: $XlwingsOutput" | Out-File $LogFile -Append
        }
    }
    catch {
        Write-Host "[AVERTISSEMENT] Erreur lors du test avec xlwings: $($_.Exception.Message)" -ForegroundColor Yellow
        "AVERTISSEMENT: Test xlwings ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©chouÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $($_.Exception.Message)" | Out-File $LogFile -Append
        # Ne pas quitter en cas d'erreur xlwings, c'est un test supplÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©mentaire
    }
}

# --- PrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©paration du dossier temporaire ---
Write-Host "[INFO] PrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©paration du dossier temporaire..." -ForegroundColor Green
if (Test-Path $TempDir) { 
    Remove-Item $TempDir -Recurse -Force 
    Start-Sleep -Seconds 1  # Attendre que la suppression soit terminÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e
}
New-Item -Path $TempDir -ItemType Directory -Force | Out-Null
Start-Sleep -Seconds 1  # Attendre que la crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation soit terminÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e

# VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier que le dossier temporaire a bien ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©tÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©
if (-not (Test-Path $TempDir)) {
    Write-Host "[ERREUR] Impossible de crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©er le dossier temporaire: $TempDir" -ForegroundColor Red
    "ERREUR: Impossible de crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©er le dossier temporaire: $TempDir" | Out-File $LogFile -Append
    exit 1
}

# --- PrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©paration du dossier de sortie ---
Write-Host "[INFO] PrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©paration du dossier de sortie..." -ForegroundColor Green
if (-not (Test-Path $OutputDir)) {
    try {
        New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
        "Dossier de sortie crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $OutputDir" | Out-File $LogFile -Append
    }
    catch {
        Write-Host "[ERREUR] Impossible de crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©er le dossier de sortie: $OutputDir" -ForegroundColor Red
        "ERREUR: Impossible de crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©er le dossier: $($_.Exception.Message)" | Out-File $LogFile -Append
        exit 1
    }
}

# --- Obtenir la liste des fichiers VBA ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  importer ---
Write-Host "[INFO] RÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cupÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ration de la liste des fichiers VBA..." -ForegroundColor Green
$ClassFiles = Get-ChildItem -Path $SourceFolders -Filter "*.cls" -Recurse | Select-Object -ExpandProperty FullName
$ModuleFiles = Get-ChildItem -Path $SourceFolders -Filter "*.bas" -Recurse | Select-Object -ExpandProperty FullName
$FormFiles = Get-ChildItem -Path $SourceFolders -Filter "*.frm" -Recurse | Select-Object -ExpandProperty FullName

$TotalFiles = $ClassFiles.Count + $ModuleFiles.Count + $FormFiles.Count
Write-Host "[INFO] Total de fichiers ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  importer: $TotalFiles" -ForegroundColor Green
"Fichiers ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  importer: $TotalFiles (Classes: $($ClassFiles.Count), Modules: $($ModuleFiles.Count), Formulaires: $($FormFiles.Count))" | Out-File $LogFile -Append

# --- CrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©er le module de dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©marrage ---
Write-Host "[INFO] CrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation du module de dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©marrage..." -ForegroundColor Green
$StartupModuleContent = @"
Attribute VB_Name = "modAddInStartup"
Option Explicit

Public Sub Auto_Open()
    ' Cette procÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©dure s'exÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cute ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  l'ouverture de l'add-in
    Debug.Print "APEX Framework Add-In initialisÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©"
    ' Vous pouvez ajouter ici d'autres opÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rations d'initialisation
End Sub

Public Sub RegisterAddIn()
    ' Cette procÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©dure peut ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Âªtre appelÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e pour enregistrer l'add-in
    MsgBox "APEX Framework enregistrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© avec succÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s.", vbInformation, "APEX Framework"
End Sub
"@
$StartupModulePath = Join-Path -Path (Get-Location) -ChildPath "$TempDir\modAddInStartup.bas"
$StartupModuleContent | Out-File -FilePath $StartupModulePath -Encoding utf8

# VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rifier que le fichier de dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©marrage a bien ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©tÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©
if (-not (Test-Path $StartupModulePath)) {
    Write-Host "[ERREUR] Impossible de crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©er le module de dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©marrage: $StartupModulePath" -ForegroundColor Red
    "ERREUR: Impossible de crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©er le module de dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©marrage: $StartupModulePath" | Out-File $LogFile -Append
    exit 1
}

# --- CrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation du fichier Add-In via Excel COM ---
Write-Host "[INFO] Lancement d'Excel pour crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©er l'add-in..." -ForegroundColor Green
try {
    # CrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©er une instance Excel
    $Excel = New-Object -ComObject Excel.Application
    $Excel.DisplayAlerts = $false
    $Excel.Visible = $false
    "Excel dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©marrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© (Invisible)" | Out-File $LogFile -Append

    # CrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©er un nouveau classeur
    $Workbook = $Excel.Workbooks.Add()
    "Nouveau classeur crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©" | Out-File $LogFile -Append

    # RÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©fÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rencer le projet VBA
    $VBProject = $Workbook.VBProject
    
    # DÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©finir le nom du projet
    try {
        $VBProject.Name = $ProjectName
        "Projet VBA renommÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $ProjectName" | Out-File $LogFile -Append
    }
    catch {
        Write-Host "[AVERTISSEMENT] Impossible de renommer le projet VBA. Assurez-vous que l'accÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s au modÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨le d'objet VBA est activÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© dans les paramÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨tres de macro Excel." -ForegroundColor Yellow
        "AVERTISSEMENT: Impossible de renommer le projet VBA: $($_.Exception.Message)" | Out-File $LogFile -Append
    }

    # Importer le module de dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©marrage
    Write-Host "[INFO] Importation du module de dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©marrage..." -ForegroundColor Green
    try {
        $VBProject.VBComponents.Import($StartupModulePath)
        "Module de dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©marrage importÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: modAddInStartup.bas" | Out-File $LogFile -Append
    }
    catch {
        Write-Host "[ERREUR] Impossible d'importer le module de dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©marrage: $($_.Exception.Message)" -ForegroundColor Red
        "ERREUR: Impossible d'importer le module de dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©marrage: $($_.Exception.Message)" | Out-File $LogFile -Append
    }

    # Importer les modules, classes et formulaires
    Write-Host "[INFO] Importation des modules, classes et formulaires..." -ForegroundColor Green
    $ImportCount = 0

    # Importer les modules
    foreach ($file in $ModuleFiles) {
        try {
            $VBComponent = $VBProject.VBComponents.Import($file)
            $ImportCount++
            $FileName = Split-Path -Path $file -Leaf
            $ModuleName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
            
            # Renommer le module pour utiliser le nom du fichier au lieu du nom gÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©nÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rique
            try {
                $VBComponent.Name = $ModuleName
                "Module importÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© et renommÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $FileName -> $ModuleName" | Out-File $LogFile -Append
            }
            catch {
                Write-Host "`n[AVERTISSEMENT] Impossible de renommer le module: $ModuleName - $($_.Exception.Message)" -ForegroundColor Yellow
                "AVERTISSEMENT: Impossible de renommer le module: $ModuleName - $($_.Exception.Message)" | Out-File $LogFile -Append
            }
            
            # Afficher la progression
            $Progress = [Math]::Round(($ImportCount / $TotalFiles) * 100)
            Write-Host "`rProgression: $Progress% ($ImportCount/$TotalFiles)" -NoNewline -ForegroundColor Green
        }
        catch {
            Write-Host "`n[AVERTISSEMENT] ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â°chec de l'importation: $file - $($_.Exception.Message)" -ForegroundColor Yellow
            "AVERTISSEMENT: ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â°chec de l'importation: $file - $($_.Exception.Message)" | Out-File $LogFile -Append
        }
    }

    # Importer les classes
    foreach ($file in $ClassFiles) {
        try {
            $VBComponent = $VBProject.VBComponents.Import($file)
            $ImportCount++
            $FileName = Split-Path -Path $file -Leaf
            $ClassName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
            
            # Essayer de renommer si nÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cessaire
            try {
                if ($VBComponent.Name -ne $ClassName) {
                    $VBComponent.Name = $ClassName
                    "Classe importÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e et renommÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e: $FileName -> $ClassName" | Out-File $LogFile -Append
                } else {
                    "Classe importÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©e: $FileName" | Out-File $LogFile -Append
                }
            }
            catch {
                Write-Host "`n[AVERTISSEMENT] Impossible de renommer la classe: $ClassName - $($_.Exception.Message)" -ForegroundColor Yellow
                "AVERTISSEMENT: Impossible de renommer la classe: $ClassName - $($_.Exception.Message)" | Out-File $LogFile -Append
            }
            
            # Afficher la progression
            $Progress = [Math]::Round(($ImportCount / $TotalFiles) * 100)
            Write-Host "`rProgression: $Progress% ($ImportCount/$TotalFiles)" -NoNewline -ForegroundColor Green
        }
        catch {
            Write-Host "`n[AVERTISSEMENT] ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â°chec de l'importation: $file - $($_.Exception.Message)" -ForegroundColor Yellow
            "AVERTISSEMENT: ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â°chec de l'importation: $file - $($_.Exception.Message)" | Out-File $LogFile -Append
        }
    }

    # Importer les formulaires
    foreach ($file in $FormFiles) {
        try {
            $VBComponent = $VBProject.VBComponents.Import($file)
            $ImportCount++
            $FileName = Split-Path -Path $file -Leaf
            $FormName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
            
            # Essayer de renommer si nÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©cessaire
            try {
                if ($VBComponent.Name -ne $FormName) {
                    $VBComponent.Name = $FormName
                    "Formulaire importÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© et renommÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $FileName -> $FormName" | Out-File $LogFile -Append
                } else {
                    "Formulaire importÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $FileName" | Out-File $LogFile -Append
                }
            }
            catch {
                Write-Host "`n[AVERTISSEMENT] Impossible de renommer le formulaire: $FormName - $($_.Exception.Message)" -ForegroundColor Yellow
                "AVERTISSEMENT: Impossible de renommer le formulaire: $FormName - $($_.Exception.Message)" | Out-File $LogFile -Append
            }
            
            # Afficher la progression
            $Progress = [Math]::Round(($ImportCount / $TotalFiles) * 100)
            Write-Host "`rProgression: $Progress% ($ImportCount/$TotalFiles)" -NoNewline -ForegroundColor Green
        }
        catch {
            Write-Host "`n[AVERTISSEMENT] ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â°chec de l'importation: $file - $($_.Exception.Message)" -ForegroundColor Yellow
            "AVERTISSEMENT: ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â°chec de l'importation: $file - $($_.Exception.Message)" | Out-File $LogFile -Append
        }
    }

    Write-Host "`n[INFO] $ImportCount fichiers importÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©s sur $TotalFiles" -ForegroundColor Green

    # Enregistrer le classeur en tant qu'add-in
    $OutputPath = Join-Path -Path $OutputDir -ChildPath $AddinName
    Write-Host "[INFO] Enregistrement de l'add-in: $OutputPath" -ForegroundColor Green
    $Workbook.SaveAs($OutputPath, 55)  # 55 = xlOpenXMLAddIn (Excel Add-In format)
    "Add-in enregistrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©: $OutputPath" | Out-File $LogFile -Append

    # Fermer Excel
    $Workbook.Close($false)
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    "Excel fermÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© proprement" | Out-File $LogFile -Append
}
catch {
    Write-Host "[ERREUR] Une erreur s'est produite lors de la crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation de l'add-in: $($_.Exception.Message)" -ForegroundColor Red
    "ERREUR: CrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation de l'add-in: $($_.Exception.Message)" | Out-File $LogFile -Append
    
    # S'assurer qu'Excel est fermÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© en cas d'erreur
    try {
        if ($Excel) {
            $Excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
        }
    }
    catch { }
    
    exit 1
}

# --- Nettoyage ---
Write-Host "[INFO] Nettoyage..." -ForegroundColor Green
if (Test-Path $TempDir) { Remove-Item $TempDir -Recurse -Force }

# --- VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification finale ---
if (Test-Path $OutputPath) {
    Write-Host "[SUCCÃƒÆ’Ã†â€™Ãƒâ€¹Ã¢â‚¬Â S] Le fichier Add-In a ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©tÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© avec succÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s: $OutputPath" -ForegroundColor Green
    
    # Ajouter une note sur la configuration des rÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©fÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rences
    Write-Host "`n[IMPORTANT] Configuration des rÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©fÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rences VBA requise:" -ForegroundColor Yellow
    Write-Host "1. Ouvrez Excel et allez dans Fichier > Options > ComplÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ments" -ForegroundColor White
    Write-Host "2. SÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©lectionnez 'ComplÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ments Excel' dans la liste dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©roulante 'GÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rer' et cliquez sur 'Atteindre...'" -ForegroundColor White
    Write-Host "3. Cochez la case pour 'ApexVBAFramework' et cliquez sur OK" -ForegroundColor White
    Write-Host "4. Ouvrez l'ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©diteur VBA (Alt+F11) et allez dans Outils > RÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©fÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rences..." -ForegroundColor White
    Write-Host "5. Assurez-vous que les rÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©fÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rences suivantes sont cochÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©es:" -ForegroundColor White
    Write-Host "   - Microsoft Scripting Runtime" -ForegroundColor White
    Write-Host "   - Microsoft ActiveX Data Objects (derniÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨re version)" -ForegroundColor White
    Write-Host "   - Microsoft VBScript Regular Expressions 5.5" -ForegroundColor White
    
    "Add-in crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© avec succÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¨s" | Out-File $LogFile -Append
}
else {
    Write-Host "[ERREUR] La crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation de l'add-in a ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©chouÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©. Consultez le fichier journal pour plus de dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©tails: $LogFile" -ForegroundColor Red
    "ERREUR: VÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rification finale - Add-in non trouvÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â© ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â  l'emplacement attendu" | Out-File $LogFile -Append
    exit 1
}

# --- Rapport final ---
Write-Host "[INFO] Finalisation..." -ForegroundColor Green
"Fin de la crÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©ation: $(Get-Date)" | Out-File $LogFile -Append

Write-Host ""
Write-Host "===== FIN DE LA CRÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â°ATION DE L'ADD-IN =====" -ForegroundColor Cyan 