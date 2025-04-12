#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
==========================================================
DEPRECATED - NE PAS UTILISER - CE SCRIPT EST OBSOLÈTE
==========================================================
Utilisez 'tools/python/generate_apex_addin.py' à la place.
Voir [Guide de migration](docs/MIGRATION_GUIDE.md#scripts-de-build)

Script pour créer l'add-in APEX Framework sous WSL (OBSOLÈTE)
"""
# DEPRECATED: Script obsolète, remplacé par generate_apex_addin.py
import os
import sys
import shutil
import subprocess
import time
import glob
import logging
from datetime import datetime
import json

# Configuration du journal avec horodatage
timestamp = datetime.now().strftime('%Y%m%d%H%M')
log_file = f"create_addin_log_{timestamp}.txt"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, mode='a', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("CreateApexAddin")

# Constantes
ADDIN_NAME = "ApexVBAFramework.xlam"
PROJECT_NAME = "ApexVbaFramework"
TEMP_DIR = "temp_addin"
SOURCE_FOLDERS = ["apex-core", "apex-metier", "apex-ui"]
CORE_FILES = [
    "apex-core/clsLogger.cls",
    "apex-core/modConfigManager.bas",
    "apex-core/modVersionInfo.bas",
    "apex-core/utils/modFileUtils.bas",
    "apex-core/utils/modTextUtils.bas",
    "apex-core/utils/modDateUtils.bas"
]

# Récupérer le nom d'utilisateur Windows
def get_windows_username():
    try:
        result = subprocess.run(['cmd.exe', '/c', 'echo %USERNAME%'], 
                                stdout=subprocess.PIPE, 
                                stderr=subprocess.PIPE,
                                text=True,
                                check=True)
        return result.stdout.strip()
    except subprocess.CalledProcessError as e:
        logger.error(f"Erreur lors de la récupération du nom d'utilisateur Windows: {str(e)}")
        return None

def get_addins_path():
    """Retourne le chemin du dossier pour sauvegarder l'add-in Excel."""
    # Utiliser le dossier dist du projet pour une meilleure compatibilité WSL
    current_dir = os.getcwd()
    dist_dir = os.path.join(current_dir, "dist")
    
    # Créer le dossier dist s'il n'existe pas
    if not os.path.exists(dist_dir):
        try:
            os.makedirs(dist_dir, exist_ok=True)
            logger.info(f"Dossier dist créé: {dist_dir}")
        except Exception as e:
            logger.error(f"Erreur lors de la création du dossier dist: {str(e)}")
            return None
    
    logger.info(f"Utilisation du dossier dist pour l'add-in: {dist_dir}")
    return dist_dir

def backup_existing_addin(addins_path, addin_name):
    """Sauvegarde une version existante de l'add-in avec un numéro de version."""
    full_path = os.path.join(addins_path, addin_name)
    if os.path.exists(full_path):
        # Récupérer tous les fichiers de sauvegarde existants
        pattern = os.path.join(addins_path, f"{os.path.splitext(addin_name)[0]}_*.xlam")
        existing_backups = glob.glob(pattern)
        
        # Déterminer le prochain numéro de version
        max_version = 0
        for backup in existing_backups:
            try:
                version = int(os.path.basename(backup).split('_')[1].split('.')[0])
                max_version = max(max_version, version)
            except (ValueError, IndexError):
                pass
        
        next_version = max_version + 1
        backup_name = f"{os.path.splitext(addin_name)[0]}_{next_version}.xlam"
        backup_path = os.path.join(addins_path, backup_name)
        
        try:
            shutil.copy2(full_path, backup_path)
            logger.info(f"Version existante sauvegardée: {backup_path}")
            return backup_path
        except Exception as e:
            logger.error(f"Erreur lors de la sauvegarde de l'add-in existant: {str(e)}")
    
    return None

def create_startup_module():
    """Crée le module de démarrage pour l'add-in."""
    if not os.path.exists(TEMP_DIR):
        os.makedirs(TEMP_DIR, exist_ok=True)
    
    startup_code = """Attribute VB_Name = "modAddInStartup"
Option Explicit

Public Sub Auto_Open()
    ' Cette procédure s'exécute à l'ouverture de l'add-in
    Debug.Print "APEX Framework Add-In initialisé"
    ' Vous pouvez ajouter ici d'autres opérations d'initialisation
End Sub

Public Sub RegisterAddIn()
    ' Cette procédure peut être appelée pour enregistrer l'add-in
    MsgBox "APEX Framework enregistré avec succès.", vbInformation, "APEX Framework"
End Sub
"""
    startup_path = os.path.join(TEMP_DIR, "modAddInStartup.bas")
    with open(startup_path, 'w', encoding='utf-8') as f:
        f.write(startup_code)
    
    return startup_path

def check_prerequisites():
    """Vérifie que les prérequis sont satisfaits."""
    # Vérifier les dossiers sources
    for folder in SOURCE_FOLDERS:
        if not os.path.exists(folder):
            logger.error(f"Dossier source non trouvé: {folder}")
            return False
    
    # Vérifier les fichiers essentiels
    for file in CORE_FILES:
        file_path = file.replace('/', os.path.sep)
        if not os.path.exists(file_path):
            logger.error(f"Fichier essentiel non trouvé: {file_path}")
            return False
        if os.path.getsize(file_path) == 0:
            logger.error(f"Fichier essentiel vide: {file_path}")
            return False
    
    return True

def get_vba_files():
    """Récupère tous les fichiers VBA à importer."""
    class_files = []
    module_files = []
    form_files = []
    
    for folder in SOURCE_FOLDERS:
        for root, _, files in os.walk(folder):
            for file in files:
                full_path = os.path.join(root, file)
                if file.endswith('.cls'):
                    class_files.append(full_path)
                elif file.endswith('.bas'):
                    module_files.append(full_path)
                elif file.endswith('.frm'):
                    form_files.append(full_path)
    
    logger.info(f"Fichiers à importer: Classes={len(class_files)}, Modules={len(module_files)}, Formulaires={len(form_files)}")
    return class_files, module_files, form_files

def prepare_files():
    """Prépare les fichiers pour la création d'un script PowerShell."""
    # Vérifier les prérequis
    if not check_prerequisites():
        logger.error("Échec de la vérification des prérequis")
        return False
    
    # Créer le module de démarrage
    startup_module = create_startup_module()
    if not os.path.exists(startup_module):
        logger.error(f"Échec de la création du module de démarrage: {startup_module}")
        return False
    
    # Récupérer les fichiers VBA
    class_files, module_files, form_files = get_vba_files()
    
    # Créer un fichier JSON contenant les informations nécessaires
    info = {
        "startup_module": startup_module,
        "class_files": class_files,
        "module_files": module_files,
        "form_files": form_files,
        "addin_name": ADDIN_NAME,
        "project_name": PROJECT_NAME
    }
    
    json_file = os.path.join(TEMP_DIR, "addin_info.json")
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(info, f, indent=2)
    
    logger.info(f"Informations sauvegardées dans {json_file}")
    return True

def create_powershell_script():
    """Crée un script PowerShell qui s'exécutera sur Windows."""
    # Chemin du fichier JSON
    json_file = os.path.join(TEMP_DIR, "addin_info.json")
    if not os.path.exists(json_file):
        logger.error("Fichier d'informations non trouvé")
        return False
    
    # Lire les informations
    with open(json_file, 'r', encoding='utf-8') as f:
        info = json.load(f)
    
    # Chemin de sortie pour l'add-in
    dist_path = get_addins_path()
    if not dist_path:
        logger.error("Impossible de déterminer le chemin du dossier dist")
        return False
    
    output_path = os.path.join(dist_path, info["addin_name"])
    
    # Créer le script PowerShell
    ps_script = os.path.join(TEMP_DIR, "create_addin.ps1")
    
    # Préparer le contenu du script PowerShell - version simplifiée pour éviter les problèmes d'accolades
    ps_content = []
    
    # En-tête du script
    ps_content.append("""
# Script pour créer l'add-in APEX VBA Framework
# Généré automatiquement le {0}

# Variables
$LogFile = "create_addin_log.txt"
$AddinName = "{1}"
$ProjectName = "{2}"
$OutputPath = "{3}"

# Journal
"Script PowerShell démarré: $(Get-Date)" | Out-File $LogFile -Append

# Fonction pour importer et renommer un composant VBA
function Import-VbaComponent($Path, $Type) {{
    try {{
        $VBComponent = $VBProject.VBComponents.Import($Path)
        $ImportCount++
        $Name = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        
        # Tenter de renommer le composant
        try {{
            if ($VBComponent.Name -ne $Name) {{
                $VBComponent.Name = $Name
                "$Type importé et renommé: $(Split-Path $Path -Leaf) -> $Name" | Out-File $LogFile -Append
            }} else {{
                "$Type importé: $(Split-Path $Path -Leaf)" | Out-File $LogFile -Append
            }}
        }} catch {{
            "AVERTISSEMENT: Impossible de renommer le $Type $Name" | Out-File $LogFile -Append
        }}
        return $true
    }} catch {{
        "AVERTISSEMENT: Échec de l'importation de $Path" | Out-File $LogFile -Append
        return $false
    }}
}}

# Bloc principal
try {{
    # Créer Excel
    $Excel = New-Object -ComObject Excel.Application
    $Excel.DisplayAlerts = $false
    $Excel.Visible = $false
    "Excel démarré (Invisible)" | Out-File $LogFile -Append

    # Classeur
    $Workbook = $Excel.Workbooks.Add()
    "Nouveau classeur créé" | Out-File $LogFile -Append

    # Projet VBA
    $VBProject = $Workbook.VBProject
    $ImportCount = 0
    
    # Renommer le projet
    try {{ $VBProject.Name = $ProjectName }} catch {{ "AVERTISSEMENT: Impossible de renommer le projet" | Out-File $LogFile -Append }}

    # Importer le module de démarrage
    Import-VbaComponent -Path "{4}" -Type "Module de démarrage"
""".format(
        datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
        info['addin_name'],
        info['project_name'],
        output_path.replace('/', '\\'),
        info['startup_module'].replace('/', '\\')
    ))
    
    # Modules
    for file in info["module_files"]:
        windows_path = file.replace('/', '\\')
        ps_content.append('''
    # Module: {0}
    Import-VbaComponent -Path "{1}" -Type "Module"'''.format(
            os.path.basename(file),
            windows_path
        ))
    
    # Classes
    for file in info["class_files"]:
        windows_path = file.replace('/', '\\')
        ps_content.append('''
    # Classe: {0}
    Import-VbaComponent -Path "{1}" -Type "Classe"'''.format(
            os.path.basename(file),
            windows_path
        ))
    
    # Formulaires
    for file in info["form_files"]:
        windows_path = file.replace('/', '\\')
        ps_content.append('''
    # Formulaire: {0}
    Import-VbaComponent -Path "{1}" -Type "Formulaire"'''.format(
            os.path.basename(file),
            windows_path
        ))
    
    # Finalisation
    ps_content.append("""
    # Bilan
    "Total: $ImportCount fichiers importés" | Out-File $LogFile -Append

    # Enregistrer l'add-in
    $Workbook.SaveAs($OutputPath, 55)  # 55 = Excel Add-In format
    "Add-in enregistré: $OutputPath" | Out-File $LogFile -Append

    # Message
    cmd.exe /c "echo Add-in créé avec succès: $OutputPath"

    # Fermer Excel
    $Workbook.Close($false)
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    if (Test-Path $OutputPath) {
        "Add-in créé avec succès" | Out-File $LogFile -Append
        exit 0
    } else {
        "ERREUR: Fichier add-in non trouvé" | Out-File $LogFile -Append
        exit 1
    }
} catch {
    # Gérer l'erreur
    "ERREUR: $($_.Exception.Message)" | Out-File $LogFile -Append
    
    # Fermer Excel si ouvert
    if ($Excel) {
        try { 
            $Excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
        } catch { }
    }
    
    # Message d'erreur
    cmd.exe /c "echo ERREUR: $($_.Exception.Message)"
    exit 1
}
""")
    
    # Écrire le script complet
    full_content = ''.join(ps_content)
    
    with open(ps_script, 'w', encoding='utf-16') as f:
        f.write(full_content)
    
    logger.info(f"Script PowerShell créé: {ps_script}")
    return ps_script

def run_powershell_script(ps_script):
    """Exécute le script PowerShell sur Windows en essayant une invocation directe."""
    logger.info("Tentative d'exécution du script PowerShell: {}".format(ps_script))
    
    try:
        # Vérifier que le fichier script PowerShell existe
        if not os.path.exists(ps_script):
            logger.error(f"Script PowerShell non trouvé: {ps_script}")
            return False
            
        # Convertir le chemin Linux du script en chemin Windows
        # Exemple: /mnt/d/Dev/temp_addin/create_addin.ps1 -> D:\Dev\temp_addin\create_addin.ps1
        if ps_script.startswith('/mnt/'):
            drive_letter = ps_script[5]
            windows_path_relative = ps_script[7:].replace('/', '\\')
            windows_path = f"{drive_letter.upper()}:\\{windows_path_relative}"
        else:
            # Si le chemin n'est pas standard /mnt/, on essaie une conversion simple
            # Cela pourrait échouer si le script n'est pas dans un montage Windows
            windows_path = ps_script.replace('/', '\\') 
            logger.warning(f"Conversion de chemin non standard, résultat: {windows_path}")

        logger.info(f"Chemin Windows calculé pour le script PS: {windows_path}")

        # Construire la commande pour powershell.exe
        cmd = [
            'powershell.exe', 
            '-ExecutionPolicy', 'Bypass',
            '-NoProfile',
            '-NonInteractive',
            '-File', windows_path
        ]
        cmd_str = ' '.join(cmd)
        logger.info("Commande PowerShell construite: {}".format(cmd_str))
        
        # Exécuter la commande PowerShell
        result = subprocess.run(
            cmd, 
            capture_output=True, # Utiliser capture_output pour stdout/stderr
            text=True,
            encoding='utf-8',
            errors='replace'
        )
        
        # Journaliser la sortie et les erreurs
        logger.info(f"Code de retour PowerShell: {result.returncode}")
        if result.stdout:
            logger.info(f"Sortie standard PowerShell:\n{result.stdout}")
        if result.stderr:
            # Les messages d'erreur PowerShell (comme l'échec COM) sortent souvent sur stderr
            logger.error(f"Erreur standard PowerShell:\n{result.stderr}") 
        
        # Vérifier le code de retour
        if result.returncode == 0:
            logger.info("Script PowerShell exécuté avec succès (selon le code de retour)")
            return True
        else:
            logger.error(f"Échec de l'exécution du script PowerShell (code de retour non nul: {result.returncode})")
            return False
            
    except FileNotFoundError:
        logger.error("Erreur: 'powershell.exe' non trouvé dans le PATH de WSL.")
        logger.error("Assurez-vous que Windows est accessible depuis WSL et que powershell.exe est dans le PATH.")
        return False
    except Exception as e:
        logger.error("Erreur Python lors de l'exécution du script PowerShell: {}".format(str(e)))
        return False

def cleanup():
    """Nettoie les fichiers temporaires."""
    if os.path.exists(TEMP_DIR):
        try:
            shutil.rmtree(TEMP_DIR)
            logger.info(f"Dossier temporaire supprimé: {TEMP_DIR}")
        except Exception as e:
            logger.warning(f"Impossible de supprimer le dossier temporaire: {str(e)}")

def show_windows_notification(title, message):
    """Affiche une notification Windows via PowerShell (méthode simplifiée)."""
    try:
        # Ne pas utiliser de fonctionnalités avancées, juste le minimum pour afficher une notification
        ps_cmd = f'''
        powershell.exe -Command "Write-Output '{title}: {message}'"
        '''
        
        subprocess.run(ps_cmd, shell=True, check=False)
        logger.info("Notification Windows affichée")
        return True
    except Exception as e:
        logger.error(f"Erreur lors de l'affichage de la notification: {str(e)}")
        return False

def create_notification_file(success, message, addins_path):
    """Crée un fichier de notification qui peut être ouvert pour afficher un message."""
    # Utiliser le dossier dist pour la notification
    dist_dir = os.path.join(os.getcwd(), "dist")
    if not os.path.exists(dist_dir):
        try:
            os.makedirs(dist_dir, exist_ok=True)
        except Exception as e:
            logger.error(f"Erreur lors de la création du dossier dist: {str(e)}")
            return False
    
    notification_file = f"APEX_notification_{timestamp}.txt"
    notification_path = os.path.join(dist_dir, notification_file)
    
    status = "Succès" if success else "Erreur"
    
    # Récupérer le chemin du dossier AddIns standard pour les instructions
    windows_username = get_windows_username()
    standard_addins_path = f"C:\\Users\\{windows_username}\\AppData\\Roaming\\Microsoft\\AddIns"
    
    # Utiliser un fichier texte simple au lieu de HTML pour éviter les problèmes de rendu
    content = f"""
==============================================
{status}: APEX VBA Framework Add-In
==============================================

{message}

INFORMATIONS:
- L'add-in a été créé dans: {addins_path}
- Journal de création: {log_file}
- Date et heure: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}

INSTALLATION MANUELLE REQUISE:
1. Copiez le fichier "{ADDIN_NAME}" du dossier:
   {addins_path}
   vers le dossier des add-ins Excel:
   {standard_addins_path}

2. Ouvrez Excel et allez dans Fichier > Options > Compléments
3. Sélectionnez 'Compléments Excel' dans la liste déroulante 'Gérer'
4. Cliquez sur 'Atteindre...' et cochez la case pour '{ADDIN_NAME}'
5. Vérifiez que les modules ont leurs noms explicites dans l'éditeur VBA (Alt+F11)

CONFIGURATION DES RÉFÉRENCES VBA:
- Microsoft Scripting Runtime
- Microsoft ActiveX Data Objects
- Microsoft VBScript Regular Expressions 5.5

==============================================
"""
    
    try:
        with open(notification_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        logger.info(f"Fichier de notification créé: {notification_path}")
        
        # Ouvrir le fichier texte avec notepad si c'est un succès
        if success:
            windows_path = notification_path.replace('/', '\\')
            # Utiliser notepad.exe pour ouvrir le fichier texte
            cmd = ['cmd.exe', '/c', f'notepad.exe "{windows_path}"']
            subprocess.Popen(cmd, 
                            stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE,
                            text=True,
                            errors='replace')
            logger.info(f"Notification ouverte avec notepad: {windows_path}")
        
        return True
    except Exception as e:
        logger.error(f"Erreur lors de la création du fichier de notification: {str(e)}")
        return False

def main():
    """Fonction principale."""
    print("===== CRÉATION DU FICHIER ApexVBAFramework.xlam SOUS WSL =====")
    
    # Vérifier les droits d'accès aux fichiers Windows
    test_file = "/mnt/c/Windows/notepad.exe"
    if not os.path.exists(test_file):
        logger.warning("Impossible d'accéder aux fichiers Windows. Vérifiez que WSL est correctement configuré.")
        print("\n[AVERTISSEMENT] Accès limité au système de fichiers Windows détecté.")
        print("Cela peut affecter le fonctionnement du script.")
    
    # Récupérer le chemin des add-ins (dossier dist)
    dist_path = get_addins_path()
    if not dist_path:
        logger.error("Impossible de créer ou d'accéder au dossier dist")
        return 1
    
    # Sauvegarder une version existante si nécessaire
    backup_path = backup_existing_addin(dist_path, ADDIN_NAME)
    if backup_path:
        print(f"Add-in existant sauvegardé en tant que: {os.path.basename(backup_path)}")
    
    # Préparer les fichiers
    if not prepare_files():
        logger.error("Échec de la préparation des fichiers")
        return 1
    
    # Créer le script PowerShell
    ps_script = create_powershell_script()
    if not ps_script:
        logger.error("Échec de la création du script PowerShell")
        return 1
    
    # Exécuter le script PowerShell
    success = run_powershell_script(ps_script)
    
    # Ne pas nettoyer les fichiers temporaires pour faciliter le débogage
    # cleanup()
    
    if success:
        # Obtenir le chemin de l'add-in créé
        output_path = os.path.join(dist_path, ADDIN_NAME)
        
        # Vérifier que l'add-in a bien été créé
        if not os.path.exists(output_path):
            logger.warning(f"L'add-in n'est pas trouvé à l'emplacement attendu: {output_path}")
            print("\n[AVERTISSEMENT]")
            print(f"L'add-in n'est pas trouvé à l'emplacement attendu: {output_path}")
            print("Le processus s'est terminé avec succès, mais le fichier pourrait être dans un autre emplacement.")
        
        success_message = f"Add-in créé avec succès à l'emplacement: {output_path}"
        
        # Afficher dans la console
        print("\n[SUCCÈS]")
        print(success_message)
        print("\nL'add-in a été créé dans le dossier 'dist' du projet:")
        print(f"{dist_path}")
        
        # Récupérer le chemin du dossier AddIns standard pour les instructions
        windows_username = get_windows_username()
        standard_addins_path = f"C:\\Users\\{windows_username}\\AppData\\Roaming\\Microsoft\\AddIns"
        
        print("\nINSTALLATION MANUELLE REQUISE:")
        print(f"1. Copiez le fichier '{ADDIN_NAME}' vers: {standard_addins_path}")
        print("2. Ouvrez Excel et allez dans Fichier > Options > Compléments")
        print("3. Sélectionnez 'Compléments Excel' dans la liste déroulante 'Gérer'")
        print("4. Cliquez sur 'Atteindre...' et cochez la case pour l'add-in")
        print("5. Vérifiez que les modules ont leurs noms explicites dans l'éditeur VBA (Alt+F11)")
        print("\nN'oubliez pas de configurer les références VBA:")
        print("- Microsoft Scripting Runtime")
        print("- Microsoft ActiveX Data Objects")
        print("- Microsoft VBScript Regular Expressions 5.5")
        
        # Créer un fichier de notification et l'ouvrir
        create_notification_file(True, success_message, dist_path)
        
        logger.info("Création de l'add-in terminée avec succès")
        return 0
    else:
        error_message = "Échec de la création de l'add-in. Vérifiez le fichier log pour plus de détails."
        
        # Afficher dans la console
        print(f"\n[ERREUR] {error_message}")
        print(f"Journal détaillé: {log_file}")
        
        # Créer un fichier de notification sans l'ouvrir
        create_notification_file(False, error_message, dist_path)
        
        logger.error("Échec de la création de l'add-in")
        return 1

if __name__ == "__main__":
    sys.exit(main()) 