import xlwings as xw
import os
from pathlib import Path
import logging
import sys

# Fichiers identifiés avec l'erreur "L'entrée dépasse la fin de fichier"
lst_fichiers_problem = [
    "apex-metier/outlook/clsAttachmentProcessor.cls",
    "apex-metier/outlook/clsMailBuilder.cls",
    "apex-metier/outlook/clsMailFetcher.cls",
    "apex-metier/outlook/clsOutlookClient.cls",
    "apex-metier/xml/clsXmlConfigManager.cls",
    "apex-metier/xml/clsXmlDiffer.cls",
    "apex-metier/xml/clsXmlFlattener.cls",
    "apex-metier/xml/clsXmlValidator.cls",
    "apex-core/modEnvVars.bas",
]

LOG_FILE = "fix_classes_log.txt"
SUFFIXE_FIX = "_fixed" # Suffixe pour les fichiers potentiellement corrigés

# --- Configuration du Logging ---
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    handlers=[logging.FileHandler(LOG_FILE, mode='w', encoding='utf-8'),
                              logging.StreamHandler(sys.stdout)])

def fix_vba_files(file_list):
    """Tente de corriger les fichiers VBA en les important et exportant via Excel."""
    logging.info("Démarrage du processus de tentative de correction des fichiers VBA...")
    fixed_count = 0
    error_count = 0
    app = None
    wb = None

    try:
        logging.info("Lancement d'une instance Excel invisible...")
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False

        logging.info("Création d'un classeur temporaire...")
        wb = app.books.add()
        vbp = wb.api.VBProject

        for relative_path_str in file_list:
            chemin_original = Path(relative_path_str).resolve()
            # Construire le chemin pour le fichier corrigé
            chemin_fix = chemin_original.with_name(f"{chemin_original.stem}{SUFFIXE_FIX}{chemin_original.suffix}")
            
            logging.info(f"Traitement de : {chemin_original.name}")

            if not chemin_original.exists():
                logging.error(f"  Fichier original non trouvé : {chemin_original}")
                error_count += 1
                continue

            # Supprimer un éventuel fichier corrigé d'une exécution précédente
            if chemin_fix.exists():
                try:
                    chemin_fix.unlink()
                    logging.debug(f"  Ancien fichier corrigé supprimé : {chemin_fix.name}")
                except Exception as e_del:
                    logging.warning(f"  Impossible de supprimer l'ancien fichier corrigé '{chemin_fix.name}': {e_del}")

            composant_a_supprimer = None
            try:
                # Importer le composant
                logging.debug(f"  Importation de '{chemin_original}'...")
                composant = vbp.VBComponents.Import(str(chemin_original))
                nom_composant = composant.Name # Récupérer le nom assigné par Excel
                logging.debug(f"  Importé en tant que '{nom_composant}'.")

                # Exporter immédiatement vers le nouveau fichier
                logging.debug(f"  Exportation vers '{chemin_fix}'...")
                composant.Export(str(chemin_fix))

                # Garder une référence pour la suppression ultérieure
                composant_a_supprimer = composant

                logging.info(f"  Traité et exporté avec succès vers '{chemin_fix.name}'.")
                fixed_count += 1

            except Exception as e:
                logging.error(f"  Échec du traitement du fichier '{chemin_original.name}': {e}")
                error_count += 1
            finally:
                # Essayer de supprimer le composant importé du classeur temporaire
                # que l'export ait réussi ou non, pour éviter les conflits.
                if composant_a_supprimer:
                    try:
                        logging.debug(f"  Suppression du composant temporaire '{composant_a_supprimer.Name}' du classeur.")
                        vbp.VBComponents.Remove(composant_a_supprimer)
                    except Exception as e_rem:
                        logging.warning(f"  Impossible de supprimer le composant temporaire '{composant_a_supprimer.Name}': {e_rem}")

    except Exception as e_main:
        logging.critical(f"Une erreur critique est survenue durant le processus : {e_main}", exc_info=True)
    finally:
        # Nettoyage de l'instance Excel
        if wb is not None:
            try:
                wb.close() # Fermer le classeur temporaire
                logging.debug("Classeur temporaire fermé.")
            except Exception as e_close:
                 logging.warning(f"Exception lors de la fermeture du classeur temporaire : {e_close}")
        if app is not None:
            try:
                app.quit() # Fermer Excel
                logging.debug("Application Excel fermée.")
            except Exception as e_quit:
                logging.warning(f"Exception lors de la fermeture de l'application Excel : {e_quit}")

    logging.info("--- Résumé du Processus de Correction ---")
    logging.info(f"Fichiers traités avec succès (import/export) : {fixed_count}")
    logging.info(f"Erreurs rencontrées : {error_count}")
    if fixed_count > 0:
        logging.info(f"Les fichiers potentiellement corrigés ont été sauvegardés avec le suffixe '{SUFFIXE_FIX}'.")
        logging.info("Veuillez EXAMINER ces fichiers (vérifier l'encodage, le contenu) avant de remplacer les originaux.")
    logging.info("----------------------------------------")

if __name__ == "__main__":
    fix_vba_files(lst_fichiers_problem) 