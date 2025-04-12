#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
==========================================================
DEPRECATED - NE PAS UTILISER - CE SCRIPT EST OBSOLÈTE
==========================================================
Utilisez 'tools/python/generate_apex_addin.py' à la place.
Voir [Guide de migration](docs/MIGRATION_GUIDE.md#scripts-de-build)

Script pour créer l'add-in APEX Framework avec xlwings (OBSOLÈTE)
"""
# DEPRECATED: Script obsolète, remplacé par generate_apex_addin.py
import os
import sys
import shutil
import time
import glob
import logging
from datetime import datetime
import xlwings as xw

# Configuration du journal avec horodatage
timestamp = datetime.now().strftime('%Y%m%d%H%M')
log_file = f"create_addin_log_{timestamp}.txt"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, mode='w', encoding='utf-8'),
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

# Obtenir le chemin pour les add-ins Excel
def get_addins_path():
    """Retourne le chemin du dossier des add-ins Excel."""
    if sys.platform == 'win32':
        return os.path.join(os.environ['APPDATA'], 'Microsoft', 'AddIns')
    elif sys.platform == 'darwin':  # macOS
        return os.path.expanduser('~/Library/Group Containers/UBF8T346G9.Office/User Content/Add-ins')
    else:
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

def create_addin():
    """Crée l'add-in APEX Framework avec xlwings."""
    logger.info("Début de la création de l'add-in APEX Framework")
    
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
    all_files = class_files + module_files + form_files
    
    if not all_files:
        logger.error("Aucun fichier VBA trouvé")
        return False
    
    # Chemin de sortie pour l'add-in
    addins_path = get_addins_path()
    if not addins_path:
        logger.error("Impossible de déterminer le chemin des add-ins Excel")
        return False
    
    output_path = os.path.join(addins_path, ADDIN_NAME)
    
    # Sauvegarder une version existante si nécessaire
    backup_path = backup_existing_addin(addins_path, ADDIN_NAME)
    if backup_path:
        logger.info(f"Add-in existant sauvegardé en tant que: {os.path.basename(backup_path)}")
    
    try:
        # Lancer Excel et créer l'add-in
        logger.info("Lancement d'Excel...")
        app = xw.App(visible=False)
        app.display_alerts = False
        
        try:
            # Créer un nouveau classeur
            wb = app.books.add()
            
            # Renommer le projet VBA (nécessite une référence au projet VBA)
            try:
                vba_project = wb.api.VBProject
                vba_project.Name = PROJECT_NAME
                logger.info(f"Projet VBA renommé: {PROJECT_NAME}")
            except Exception as e:
                logger.warning(f"Impossible de renommer le projet VBA: {str(e)}")
                logger.warning("Assurez-vous que l'accès au modèle d'objet VBA est activé dans les paramètres de macro Excel.")
            
            # Importer le module de démarrage
            try:
                vba_project.VBComponents.Import(startup_module)
                logger.info("Module de démarrage importé")
            except Exception as e:
                logger.error(f"Impossible d'importer le module de démarrage: {str(e)}")
                raise
            
            # Importer les modules, classes et formulaires
            import_count = 0
            
            # Fonction pour importer et renommer un fichier VBA
            def import_vba_file(file_path):
                nonlocal import_count
                try:
                    # Importer le composant
                    vb_component = vba_project.VBComponents.Import(file_path)
                    
                    # Obtenir le nom du fichier pour le renommage
                    file_name = os.path.basename(file_path)
                    module_name = os.path.splitext(file_name)[0]
                    
                    # Renommer le composant
                    try:
                        if vb_component.Name != module_name:
                            original_name = vb_component.Name
                            vb_component.Name = module_name
                            logger.info(f"Composant renommé: {original_name} -> {module_name}")
                        else:
                            logger.info(f"Composant importé: {module_name}")
                    except Exception as e:
                        logger.warning(f"Impossible de renommer le composant {module_name}: {str(e)}")
                    
                    import_count += 1
                    return True
                except Exception as e:
                    logger.warning(f"Échec de l'importation de {file_path}: {str(e)}")
                    return False
            
            # Importer les modules
            for file in module_files:
                import_vba_file(file)
            
            # Importer les classes
            for file in class_files:
                import_vba_file(file)
            
            # Importer les formulaires
            for file in form_files:
                import_vba_file(file)
            
            logger.info(f"{import_count} fichiers importés sur {len(all_files)}")
            
            # Enregistrer le classeur en tant qu'add-in
            logger.info(f"Enregistrement de l'add-in: {output_path}")
            wb.save(output_path)
            
            # Fermer le classeur
            wb.close()
            logger.info("Add-in créé avec succès")
            
        finally:
            # Fermer Excel
            app.quit()
            logger.info("Excel fermé proprement")
        
        # Vérifier que l'add-in a été créé
        if os.path.exists(output_path):
            logger.info(f"Vérification réussie: {output_path}")
            return True
        else:
            logger.error(f"L'add-in n'a pas été créé à l'emplacement attendu: {output_path}")
            return False
            
    except Exception as e:
        logger.error(f"Erreur lors de la création de l'add-in: {str(e)}")
        return False

def cleanup():
    """Nettoie les fichiers temporaires."""
    if os.path.exists(TEMP_DIR):
        try:
            shutil.rmtree(TEMP_DIR)
            logger.info(f"Dossier temporaire supprimé: {TEMP_DIR}")
        except Exception as e:
            logger.warning(f"Impossible de supprimer le dossier temporaire: {str(e)}")

def main():
    """Fonction principale."""
    print("===== CRÉATION DU FICHIER ApexVBAFramework.xlam =====")
    
    # Nettoyer les fichiers temporaires existants
    cleanup()
    
    # Créer l'add-in
    success = create_addin()
    
    # Nettoyer les fichiers temporaires
    cleanup()
    
    if success:
        # Obtenir le chemin de l'add-in créé
        addins_path = get_addins_path()
        if addins_path:
            output_path = os.path.join(addins_path, ADDIN_NAME)
            print("\n[SUCCÈS]")
            print(f"Add-in créé avec succès à l'emplacement: {output_path}")
            print("\nL'add-in a été créé directement dans le dossier standard des add-ins Excel:")
            print(f"{addins_path}")
            print("\nPour utiliser cet add-in:")
            print("1. Ouvrez Excel et allez dans Fichier > Options > Compléments")
            print("2. Sélectionnez 'Compléments Excel' dans la liste déroulante 'Gérer'")
            print("3. Cliquez sur 'Atteindre...' et vérifiez que l'add-in est coché")
            print("4. Vérifiez que les modules ont leurs noms explicites dans l'éditeur VBA (Alt+F11)")
            print("\nN'oubliez pas de configurer les références VBA:")
            print("- Microsoft Scripting Runtime")
            print("- Microsoft ActiveX Data Objects")
            print("- Microsoft VBScript Regular Expressions 5.5")
        else:
            print("\n[SUCCÈS AVEC AVERTISSEMENT]")
            print("Add-in créé avec succès mais impossible de déterminer le chemin exact")
        
        logger.info("Création de l'add-in terminée avec succès")
        return 0
    else:
        print("\n[ERREUR]")
        print("Échec de la création de l'add-in. Vérifiez le fichier log pour plus de détails.")
        logger.error("Échec de la création de l'add-in")
        return 1

if __name__ == "__main__":
    sys.exit(main()) 