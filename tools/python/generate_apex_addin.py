import xlwings as xw
import json
import os
import sys
import shutil
import datetime
import logging
from pathlib import Path
import time
import re
import subprocess

CONFIG_FILE = "config.json" # Nom du fichier de configuration attendu

# --- Logging Setup ---
def setup_logging(config):
    """Configure le système de journalisation."""
    log_config = config.get("logging", {})
    log_level = getattr(logging, log_config.get("level", "INFO").upper(), logging.INFO)
    log_format = log_config.get("format", "%(asctime)s - %(levelname)s - %(message)s")
    log_file = log_config.get("file")

    handlers = [logging.StreamHandler(sys.stdout)] # Log vers console par défaut
    if log_file:
        # S'assurer que le dossier du log existe
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(log_path, mode='w', encoding='utf-8')) # 'w' pour écraser le log à chaque exécution

    # Utiliser force=True pour permettre la reconfiguration (utile si la fonction est appelée plusieurs fois)
    logging.basicConfig(level=log_level, format=log_format, handlers=handlers, force=True)
    logging.info("Logging initialized.")

# --- Fonctions pour la gestion des placeholders ---
def is_placeholder(file_path):
    """Vérifie si le fichier est un placeholder (vide, stub ou marqué comme À DÉVELOPPER)."""
    # Vérifier si c'est un fichier .stub
    if str(file_path).endswith('.stub'):
        return True
        
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
            return not content.strip() or "À DÉVELOPPER" in content or "TODO" in content or "(Stub)" in content
    except Exception as e:
        logging.debug(f"Erreur lors de la vérification du fichier {file_path}: {str(e)}")
        # En cas d'erreur, considérer que ce n'est pas un placeholder
        return False

# --- Fonction pour résoudre les modules manquants ---
def run_resolve_missing_modules(check_only=False):
    """Exécute le script resolve_missing.py pour générer des stubs pour les modules manquants."""
    resolve_script = Path("tools/python/resolve_missing.py")
    
    if not resolve_script.exists():
        logging.warning(f"Script {resolve_script} non trouvé. La résolution des modules manquants n'est pas disponible.")
        return False
        
    try:
        cmd = [sys.executable, str(resolve_script)]
        if check_only:
            cmd.append("--check-only")
            
        logging.info(f"Exécution de {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        # Afficher la sortie du script
        if result.stdout:
            for line in result.stdout.splitlines():
                logging.info(f"[resolve_missing] {line}")
                
        if result.stderr:
            for line in result.stderr.splitlines():
                logging.warning(f"[resolve_missing] {line}")
                
        success = result.returncode == 0
        if success:
            logging.info("Vérification/génération des modules manquants terminée avec succès")
        else:
            logging.error(f"Échec de la vérification/génération des modules manquants: code {result.returncode}")
            
        return success
    except Exception as e:
        logging.error(f"Erreur lors de l'exécution de resolve_missing.py: {str(e)}")
        return False

def create_minimal_stub(file_path, workbook):
    """Insère un stub vide dans le classeur sans charger un module cassé."""
    vba_project = workbook.api.VBProject
    module_name = Path(file_path).stem
    # Si le nom se termine par .stub, enlever l'extension .stub
    if module_name.endswith('.cls') or module_name.endswith('.bas'):
        module_name = Path(module_name).stem
    
    # Déterminer le type de composant en fonction de l'extension
    module_type = 1  # vbext_ct_StdModule par défaut
    extension = file_path.suffix.lower()
    
    # Si c'est un fichier .stub, déterminer le type en fonction de l'extension dans le nom
    if extension == '.stub':
        if '.cls.' in file_path.name:
            module_type = 2  # vbext_ct_ClassModule
        elif '.bas.' in file_path.name:
            module_type = 1  # vbext_ct_StdModule
        # Extraire le vrai nom du module sans les extensions
        parts = file_path.name.split('.')
        if len(parts) >= 3:  # format: nom.ext.stub
            module_name = parts[0]
    elif extension == '.cls':
        module_type = 2  # vbext_ct_ClassModule
    
    # Générer le code approprié selon le type de module
    if module_type == 2:  # Classe
        code = f'''Attribute VB_Name = "{module_name}"
Option Explicit

' ==========================================================================
' Module    : {module_name}
' État      : À DÉVELOPPER (Stub généré automatiquement)
' Description : Placeholder
' ==========================================================================

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
End Sub
'''
    else:  # Module standard ou autre
        code = f'''Attribute VB_Name = "{module_name}"
Option Explicit

' ==========================================================================
' Module    : {module_name}
' État      : À DÉVELOPPER (Stub généré automatiquement)
' Description : Placeholder
' ==========================================================================

' Fonctions à implémenter
'''
    
    try:
        new_component = vba_project.VBComponents.Add(module_type)
        new_component.Name = module_name
        new_component.CodeModule.AddFromString(code)
        return True
    except Exception as e:
        logging.error(f"Erreur lors de la création du stub pour {module_name}: {str(e)}")
        return False

def update_placeholders_registry(placeholder_files):
    """Met à jour le fichier MODULES_PLANIFIES.md avec la liste des placeholders détectés."""
    if not placeholder_files:
        logging.info("Pas de placeholders à enregistrer.")
        return
        
    md_path = Path('docs/MODULES_PLANIFIES.md')
    
    # Préparer le contenu du fichier
    header = """# Modules Apex à développer (placeholders)

| Module | Objectif | Dépendances | Priorité | Date prévue | Auteur |
|--------|----------|-------------|----------|-------------|--------|
"""
    
    # Ajouter chaque placeholder au tableau
    rows = []
    for file_path in placeholder_files:
        module_name = file_path.name
        module_type = "Class" if file_path.suffix.lower() == ".cls" else "Module"
        relative_path = file_path.relative_to(Path('.'))
        row = f"| {module_name} | {module_type} à développer | - | - | - | - |"
        rows.append(row)
    
    content = header + "\n".join(rows) + "\n\n---\n\n*Ce fichier est généré automatiquement par generate_apex_addin.py*"
    
    # Créer le dossier docs si nécessaire
    md_path.parent.mkdir(parents=True, exist_ok=True)
    
    # Vérifier si le fichier existe déjà
    if md_path.exists():
        logging.info(f"Le fichier {md_path} existe déjà. Mise à jour sans écraser les informations manuelles.")
        try:
            # Lire le fichier existant
            with open(md_path, 'r', encoding='utf-8') as f:
                existing_content = f.read()
                
            # Extraire l'en-tête et le préserver
            if "# Modules Apex à développer" in existing_content:
                # Garder tout jusqu'au tableau inclus
                parts = existing_content.split("| Module | Objectif |", 1)
                if len(parts) > 1:
                    # Reconstruire avec l'en-tête existant et les nouvelles lignes
                    custom_header = parts[0] + "| Module | Objectif |"
                    content = custom_header + "\n".join(rows) + "\n\n---\n\n*Ce fichier a été mis à jour automatiquement par generate_apex_addin.py*"
        except Exception as e:
            logging.warning(f"Erreur lors de la lecture du fichier existant {md_path}: {str(e)}")
    
    # Écrire le contenu dans le fichier
    try:
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(content)
        logging.info(f"Fichier {md_path} mis à jour avec {len(placeholder_files)} modules à développer.")
    except Exception as e:
        logging.error(f"Erreur lors de l'écriture du fichier {md_path}: {str(e)}")

# --- Configuration Loading ---
def load_config(config_path):
    """Charge et valide la configuration depuis un fichier JSON."""
    path = Path(config_path)
    if not path.is_file():
        # Logger n'est peut-être pas encore initialisé, utiliser print pour cette erreur critique
        print(f"CRITICAL: Configuration file not found: {config_path}", file=sys.stderr)
        raise FileNotFoundError(f"Configuration file not found: {config_path}")
    try:
        with open(path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        # La configuration est chargée, on peut initialiser le logger maintenant si ce n'est pas déjà fait
        # setup_logging(config) # Déplacé dans le bloc main pour éviter initialisation multiple
        logging.info(f"Configuration loaded successfully from {config_path}")
        # Validation basique (peut être étendue)
        required_keys = ["output_dir", "addin_basename", "source_folders"]
        if not all(k in config for k in required_keys):
             raise ValueError(f"Configuration file missing one or more required keys: {required_keys}")
        return config
    except json.JSONDecodeError as e:
        logging.critical(f"Error decoding JSON configuration file: {e}")
        raise
    except Exception as e:
        logging.critical(f"Error loading configuration: {e}")
        raise

# --- File Operations ---
def collect_source_files(source_folders, import_order):
    """Collecte les fichiers source VBA depuis les dossiers spécifiés et les trie."""
    source_files = []
    base_path = Path('.') # Chemins relatifs au script ou au répertoire de travail courant
    logging.info(f"Scanning for source files in folders relative to: {base_path.resolve()}")
    for folder in source_folders:
        folder_path = base_path / folder
        if not folder_path.is_dir():
            logging.warning(f"Source folder not found or not a directory: {folder_path}")
            continue
        logging.debug(f"Scanning folder: {folder_path}")
        # Utiliser rglob pour la recherche récursive
        for ext in ["*.cls", "*.bas", "*.frm", "*.cls.stub", "*.bas.stub"]:  # Ajout des extensions .stub
            found = list(folder_path.rglob(ext))
            if found:
                 logging.debug(f"  Found {len(found)} files with extension {ext}")
                 source_files.extend(found)

    # Trier les fichiers selon l'ordre d'importation spécifié
    if import_order:
        ext_priority = {ext.lower(): i for i, ext in enumerate(import_order)}
        # Trier en priorité par extension, puis par nom de fichier pour un ordre stable
        source_files.sort(key=lambda p: (ext_priority.get(p.suffix.lower(), 99), p.name))
        logging.info(f"Source files sorted by import order: {import_order}")
    else:
         source_files.sort(key=lambda p: p.name) # Tri alphabétique par défaut
         logging.info("Source files sorted alphabetically by name.")

    # Log des fichiers .stub trouvés
    stub_files = [f for f in source_files if str(f).endswith('.stub')]
    if stub_files:
        logging.info(f"Found {len(stub_files)} stub files that will be used as placeholders:")
        for stub in stub_files[:5]:  # Limiter l'affichage aux 5 premiers pour éviter de surcharger le log
            logging.info(f"  - {stub}")
        if len(stub_files) > 5:
            logging.info(f"  ... and {len(stub_files) - 5} more stubs")

    logging.info(f"Found {len(source_files)} total source files including stubs.")
    if not source_files:
        logging.warning("No source files found in specified folders.")
    return source_files

def check_prerequisites(collected_files, essential_files_config):
    """Vérifie si les fichiers essentiels sont présents et non vides."""
    logging.info("Checking prerequisites (essential files presence and content)...")
    # Créer un dictionnaire pour un accès rapide aux chemins par nom de fichier
    collected_file_map = {p.name: p for p in collected_files}
    missing_essentials = []
    empty_essentials = []
    all_ok = True

    for essential_filename in essential_files_config:
        # Normaliser le nom de fichier essentiel (au cas où il contient des / ou \)
        normalized_essential_name = Path(essential_filename).name
        if normalized_essential_name not in collected_file_map:
            missing_essentials.append(essential_filename) # Reporter le nom tel que dans config
            all_ok = False
        else:
            # Le fichier existe, vérifier s'il est vide
            file_path = collected_file_map[normalized_essential_name]
            try:
                # Utiliser stat().st_size pour obtenir la taille du fichier
                if file_path.stat().st_size == 0:
                    empty_essentials.append(essential_filename) # Reporter le nom tel que dans config
                    all_ok = False
            except OSError as e:
                 logging.warning(f"Could not check size of essential file '{essential_filename}': {e}")
                 # Optionnel: Traiter l'incapacité de vérifier comme une erreur ?
                 # all_ok = False

    if missing_essentials:
        logging.error("Missing essential source files:")
        for missing in missing_essentials:
            logging.error(f"  - {missing}")

    if empty_essentials:
        logging.error("Empty essential source files detected:")
        for empty in empty_essentials:
            logging.error(f"  - {empty}")

    if all_ok:
        logging.info("All essential files checks passed.")
    else:
        logging.error("Essential files check failed.")

    return all_ok

def backup_existing_addin(target_addin_path):
    """Sauvegarde le fichier add-in existant avec un numéro de version incrémental."""
    if not target_addin_path.exists():
        logging.info(f"No existing add-in found at '{target_addin_path}'. No backup needed.")
        return True

    # Construire le nom de base et l'extension pour la sauvegarde
    base = target_addin_path.with_suffix('') # Chemin sans l'extension finale
    ext = target_addin_path.suffix # Extension (ex: '.xlam')
    backup_suffix = ".bak" # Suffixe pour les sauvegardes
    version = 1
    # Utiliser le répertoire parent du fichier cible pour stocker les sauvegardes
    backup_dir = target_addin_path.parent
    backup_path = backup_dir / f"{base.name}_v{version}{ext}{backup_suffix}"

    # Trouver le prochain numéro de version disponible
    while backup_path.exists():
        version += 1
        backup_path = backup_dir / f"{base.name}_v{version}{ext}{backup_suffix}"
        if version > 999: # Limite de sécurité pour éviter une boucle infinie
             logging.error(f"Found more than 999 backups for '{target_addin_path.name}'. Please clean up backups.")
             return False

    logging.info(f"Backing up existing add-in '{target_addin_path}' to '{backup_path}'...")
    try:
        # Copier le fichier en préservant les métadonnées si possible
        shutil.copy2(target_addin_path, backup_path)
        logging.info("Backup successful.")
        return True
    except Exception as e:
        logging.error(f"Error backing up file: {e}")
        logging.warning("Ensure you have write permissions in the output directory.")
        return False

# --- Excel Generation ---
def generate_addin(config, source_files):
    """Crée le complément Excel en utilisant xlwings."""
    output_dir = Path(config.get("output_dir", "dist"))
    output_dir.mkdir(parents=True, exist_ok=True) # S'assurer que le dossier de sortie existe
    addin_basename = config.get("addin_basename", "GeneratedAddin")
    addin_name = addin_basename + ".xlam"
    target_addin_path = output_dir / addin_name
    # Obtenir le chemin absolu pour xlwings (plus fiable)
    addin_path_str = str(target_addin_path.resolve())

    options = config.get("options", {})
    startup_config = config.get("startup_module", {})
    vba_comp_types = config.get("vba_component_types", {})
    vba_project_name = config.get("vba_project_name") # Nom souhaité pour le projet VBA
    
    # Liste pour suivre les placeholders détectés
    placeholder_files = []

    # --- Sauvegarde ---
    if options.get("enable_backup", True):
        if not backup_existing_addin(target_addin_path):
            logging.error("Backup failed. Aborting generation.")
            return False # Indiquer l'échec

    # --- Instance Excel ---
    app = None # Initialiser à None pour le bloc finally
    wb = None
    start_time = time.time()
    logging.info("Starting Excel add-in generation process...")
    try:
        # Lancer ou se connecter à une instance Excel invisible
        logging.info("STEP 1: Launching invisible Excel instance (xw.App)...")
        app = xw.App(visible=False, add_book=False) # Instance invisible sans nouveau classeur
        app.display_alerts = False # Désactiver les alertes Excel
        # Capture une référence à l'application Excel lancée
        logging.info("STEP 1: Excel instance launched successfully.")

        # --- Créer le classeur ---
        logging.info("STEP 2: Creating new workbook...")
        # Créer un nouveau classeur Excel
        wb = app.books.add()
        logging.info("STEP 2: New workbook created.")

        # --- Projet VBA ---
        # Obtenir une référence au projet VBA du classeur
        vbp = wb.api.VBProject
        
        # --- Renommer le projet VBA si configuré ---
        if vba_project_name:
            logging.info(f"STEP 3: Attempting to rename VBA project to '{vba_project_name}'...")
            try:
                current_name = vbp.Name
                vbp.Name = vba_project_name
                logging.info(f"Successfully renamed VBA project from '{current_name}' to '{vba_project_name}'.")
            except Exception as e:
                logging.warning(f"Failed to rename VBA project: {e}")
                logging.warning("This might be due to locked VBA project or security settings.")
        else:
            logging.info("STEP 3: VBA project renaming skipped (no name provided in config).")
        
        # --- Supprimer les objets par défaut si configuré ---
        if options.get("delete_default_items", False):
            logging.info("STEP 4: Deleting default workbook items (Sheets, Modules)...")
            # TODO: Implémenter la suppression des éléments par défaut si nécessaire 
            # et pas risqué pour les projets Excel existants
            logging.warning("Delete default items functionality not fully implemented. Skipping.")
        else:
            logging.info("STEP 4: Skipping deletion of default workbook items.")
        
        # --- Importer les fichiers source ---
        logging.info(f"STEP 5: Starting import of {len(source_files)} VBA components...")
        # Garder une trace des noms importés pour éviter les doublons
        imported_component_names = set()
        import_success_count = 0
        import_failure_count = 0
        import_skipped_count = 0 # Compteur pour les placeholders
        
        # Parcourir chaque fichier source
        for file_path in source_files:
            # Vérifier si c'est un placeholder
            if is_placeholder(file_path):
                logging.info(f"  ⏩ Module ignoré (placeholder) : {file_path.name}")
                placeholder_files.append(file_path)
                
                # Créer un stub minimal pour ce placeholder
                if create_minimal_stub(file_path, wb):
                    import_skipped_count += 1
                    imported_component_names.add(file_path.stem)  # Ajouter le nom du stub importé
                else:
                    import_failure_count += 1
                
                continue
                
            # Essayer d'importer le fichier VBA
            try:
                # Obtenir le type de composant en fonction de l'extension
                ext = file_path.suffix.lower()
                comp_type = vba_comp_types.get(ext, 1) # Type 1 = vbext_ct_StdModule par défaut

                # Importer le fichier dans le projet
                new_component = vbp.VBComponents.Import(str(file_path))
                # Enregistrer le nom du composant importé
                imported_component_names.add(new_component.Name)
                
                # Parfois, l'import donne un nom par défaut au composant (ex: Module1)
                # Vérifier si le nom du composant correspond au nom du fichier (sans extension)
                expected_name = file_path.stem
                if new_component.Name != expected_name:
                    logging.info(f"  Renaming component '{new_component.Name}' to '{expected_name}'")
                    new_component.Name = expected_name
                
                import_success_count += 1 # Compteur succès
            except Exception as import_e:
                logging.error(f"Failed to import file '{file_path.name}': {import_e}")
                import_failure_count += 1 # Compteur échec
                # On continue l'import même si un fichier échoue
        
        logging.info(f"STEP 5: Import finished. Success: {import_success_count}, Skipped (placeholders): {import_skipped_count}, Failures: {import_failure_count}")
        
        # Mettre à jour le registre des modules planifiés
        if placeholder_files and options.get("update_placeholders_registry", True):
            update_placeholders_registry(placeholder_files)

        # --- Créer le Module de Démarrage (si configuré) ---
        if options.get("create_startup_module"):
            logging.info("STEP 6: Creating startup module...")
            module_name = startup_config.get("name_in_vba", "modAddInStartup")
            module_code = startup_config.get("default_content", "")
            if not module_code:
                 logging.warning("Startup module creation enabled but no 'default_content' found in config.")
            else:
                 logging.info(f"Creating startup module: {module_name}")
                 try:
                     # Vérifier si un composant avec ce nom existe déjà (importé ou créé précédemment)
                     try:
                         existing_comp = vbp.VBComponents(module_name)
                         logging.warning(f"Startup module '{module_name}' seems to already exist. Skipping creation/overwrite.")
                     except: # Erreur COM attendue si le module n'existe pas
                         # Ajouter le nouveau module standard
                         module_type = vba_comp_types.get(".bas", 1) # Type 1 = vbext_ct_StdModule
                         new_mod = vbp.VBComponents.Add(module_type)
                         new_mod.Name = module_name
                         # Ajouter le code au module
                         new_mod.CodeModule.AddFromString(module_code)
                         logging.info(f"  Added startup code to new module '{module_name}'.")
                         imported_component_names.add(module_name) # Enregistrer le nom
                 except Exception as startup_e:
                     logging.error(f"Failed to create startup module '{module_name}': {startup_e}")
        else:
            logging.info("STEP 6: Skipping startup module creation.")

        # --- Sauvegarder en tant qu'Add-in (.xlam) ---
        logging.info(f"STEP 7: Saving workbook as Add-in: {addin_path_str}...")
        try:
            win_path = addin_path_str.replace('/', '\\')
            xlOpenXMLAddIn = 55 
            wb.api.SaveAs(win_path, FileFormat=xlOpenXMLAddIn)
            logging.info("STEP 7: Add-in saved successfully.")
        except Exception as save_e:
            logging.critical(f"CRITICAL: Failed to save Add-in '{addin_path_str}': {save_e}")
            logging.error("Check if the file is locked, if path is valid, or permissions are sufficient.")
            return False 

        duration = time.time() - start_time
        logging.info(f"Excel add-in generation process completed in {duration:.2f} seconds.")
        return True 

    except Exception as e:
        logging.critical(f"An error occurred during Excel generation: {e}", exc_info=True)
        return False
    finally:
        logging.info("STEP 8: Starting cleanup (closing Excel resources)...")
        # --- Nettoyage de l'instance Excel ---
        if wb is not None:
            try:
                logging.debug("  Closing workbook...")
                wb.close() 
                logging.debug("  Workbook closed.")
            except Exception as e_close:
                 logging.warning(f"  Exception while closing workbook: {e_close}")
        if app is not None:
            try:
                if options.get("close_excel_after_generation", True):
                    logging.debug("  Quitting Excel application...")
                    app.quit()
                    logging.debug("  Excel application quit.")
                else:
                     logging.info("  Excel application left running as per configuration (might be invisible).")
            except Exception as e_quit:
                logging.warning(f"  Exception while quitting Excel app: {e_quit}")
        logging.info("STEP 8: Cleanup finished.")


# --- Exécution Principale ---
if __name__ == "__main__":
    print(f"--- APEX Framework Add-in Generator (using {CONFIG_FILE}) ---") # Titre indiquant le fichier config
    config = None # Définir config à None au cas où load_config échoue
    try:
        # Charger la configuration en premier
        config = load_config(CONFIG_FILE)
        # Initialiser le logging basé sur la config chargée
        setup_logging(config)
        
        # Vérifier et résoudre les modules manquants
        options = config.get("options", {})
        if options.get("check_missing_modules", True):
            logging.info("ÉTAPE PRÉLIMINAIRE: Vérification des modules manquants...")
            check_only = options.get("check_only_missing_modules", False)
            if check_only:
                logging.info("Mode vérification uniquement: les stubs ne seront pas générés")
            else:
                logging.info("Mode génération automatique: les stubs seront générés pour les modules manquants")
            
            run_resolve_missing_modules(check_only=check_only)
            logging.info("Vérification des modules manquants terminée")

        # Collecter les fichiers source
        source_folders = config.get("source_folders", [])
        logging.info(f"Analyse des dossiers source: {', '.join(source_folders)}")
        source_files = collect_source_files(source_folders, config.get("import_order", []))
        if not source_files:
             # Si aucun fichier source n'est trouvé, c'est probablement une erreur de config
             logging.error("No source files found based on configuration. Check 'source_folders' in config.")
             # On peut décider d'arrêter ici ou de continuer pour créer un add-in vide
             # sys.exit(1) # Décommenter pour arrêter si aucun fichier source

        # Vérifier les prérequis
        if not check_prerequisites(source_files, config.get("essential_files", [])):
            logging.error("Prerequisite check failed. Aborting generation.")
            sys.exit(1) # Arrêter si les prérequis ne sont pas remplis

        # Générer l'add-in (contient la sauvegarde, création, import, sauvegarde finale)
        success = generate_addin(config, source_files)

        # --- Affichage Final et Instructions ---
        if success:
            logging.info("Add-in generation process completed successfully.")
            print("\n[SUCCÈS]")

            # Calculer le chemin final pour l'affichage
            output_dir = Path(config.get("output_dir", "dist"))
            addin_basename = config.get("addin_basename", "GeneratedAddin")
            addin_name = addin_basename + ".xlam"
            final_addin_path = (output_dir / addin_name).resolve()

            print(f"\nAdd-in '{addin_name}' généré avec succès à l'emplacement :")
            print(f"{final_addin_path}")

            # Afficher les instructions
            print("\n--- Instructions pour l'activation et la configuration ---")
            print("\nPour activer cet add-in dans Excel :")
            print("1. Ouvrez Excel.")
            print("2. Allez dans Fichier > Options > Compléments.")
            print("3. En bas, dans 'Gérer', sélectionnez 'Compléments Excel' et cliquez sur 'Atteindre...'.")
            print("4. Cliquez sur le bouton 'Parcourir...'.")
            print(f"5. Naviguez jusqu'à l'emplacement ci-dessus et sélectionnez le fichier '{addin_name}'.")
            print(f"6. Assurez-vous que la case à côté de '{addin_basename}' (ou similaire) est cochée, puis cliquez sur OK.")

            print("\nConfiguration des Références VBA (Important) :")
            print("Une fois l'add-in activé :")
            print("1. Ouvrez l'éditeur VBA (Alt+F11).")
            print("2. Allez dans Outils > Références...")
            print("3. Assurez-vous que les références suivantes sont cochées (les versions peuvent varier) :")
            print("   - Visual Basic For Applications")
            print("   - Microsoft Excel XX.X Object Library")
            print("   - OLE Automation")
            print("   - Microsoft Office XX.X Object Library")
            print("   - Microsoft Scripting Runtime")
            print("   - Microsoft ActiveX Data Objects X.X Library (choisir la plus récente, ex: 6.1)")
            print("   - Microsoft VBScript Regular Expressions 5.5")
            print("4. Cliquez sur OK.")
            print("\n---------------------------------------------------------")

            sys.exit(0) # Code de sortie 0 pour succès
        else:
            # L'échec a déjà été logué dans generate_addin ou ailleurs
            print("\n[ERREUR]")
            log_file_path = config.get('logging', {}).get('file', 'log_non_configure.txt')
            print(f"Échec de la génération de l'add-in. Consultez le fichier log '{log_file_path}' pour les détails.")
            sys.exit(1) # Code de sortie non nul pour échec

    except FileNotFoundError as e:
        # Gérer spécifiquement l'absence du fichier config qui empêche même le logging
        print(f"\n[ERREUR CRITIQUE] Fichier de configuration introuvable : {e}")
        print(f"Veuillez créer le fichier '{CONFIG_FILE}' ou vérifier son chemin.")
        sys.exit(1)
    except ValueError as e:
        # Gérer les erreurs de configuration (ex: clé manquante)
        print(f"\n[ERREUR DE CONFIGURATION] {e}")
        print(f"Veuillez vérifier le contenu du fichier '{CONFIG_FILE}'.")
        # Essayer de logger si possible
        try: logging.error(f"Configuration error: {e}", exc_info=True)
        except: pass
        sys.exit(1)
    except Exception as e:
        # Capturer toute autre exception imprévue
        print(f"\n[ERREUR CRITIQUE INATTENDUE] {e}")
        # Essayer de logger l'erreur complète pour le débogage
        try:
            # Reconfigurer le logger au cas où il n'aurait pas été initialisé
            if config: setup_logging(config)
            logging.critical(f"An unexpected critical error occurred: {e}", exc_info=True)
            log_file_path = config.get('logging', {}).get('file', 'log_non_configure.txt') if config else 'log_non_configure.txt'
            print(f"Consultez le fichier log '{log_file_path}' pour les détails techniques.")
        except:
            print("Impossible d'écrire l'erreur détaillée dans le fichier log.")
        sys.exit(1) 