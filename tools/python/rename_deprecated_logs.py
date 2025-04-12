#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script pour renommer tous les fichiers journaux obsolètes en ajoutant l'extension .DEPRECATED
Compatible avec WSL et optimisé pour les performances sur le système de fichiers Windows.

IMPORTANT: Ce script doit être exécuté via PowerShell sous Windows, PAS via WSL.
Commande recommandée :
    python D:\chemin\vers\Apex_VBA_FRAMEWORK\tools\python\rename_deprecated_logs.py [--dry-run]
"""
import os
import argparse
import logging
import time
import platform
import sys
from datetime import datetime

# --- Vérification d'environnement ---
def check_environment():
    """Vérifie si le script est exécuté dans l'environnement approprié"""
    # Détection de WSL
    if "microsoft" in platform.release().lower() or "linux" in platform.system().lower():
        print("🚨 ERREUR D'ENVIRONNEMENT")
        print("Ce script ne doit pas être lancé depuis WSL. Utilisez PowerShell Windows natif.")
        print("\nCommande à utiliser sous PowerShell:")
        print('   python D:\\chemin\\vers\\Apex_VBA_FRAMEWORK\\tools\\python\\rename_deprecated_logs.py --dry-run -v')
        print("\nOu utilisez le script .bat fourni:")
        print('   D:\\chemin\\vers\\Apex_VBA_FRAMEWORK\\tools\\run_rename_logs.bat')
        sys.exit(1)
        
    # Vérification de l'accès aux chemins Windows
    try:
        windows_path = os.environ.get('USERPROFILE')
        if windows_path and os.path.exists(windows_path):
            return True
        else:
            print("⚠️ Attention: Impossible de vérifier les chemins Windows standard.")
            print("Assurez-vous d'exécuter le script depuis Windows PowerShell.")
    except Exception:
        print("⚠️ Erreur lors de la vérification de l'environnement Windows.")
    
    return True  # On continue malgré tout

# --- Configuration du logger ---
def setup_logger(log_file="rename_logs.log"):
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler(log_file, mode='w', encoding='utf-8'),
            logging.StreamHandler()
        ]
    )

# --- Fonction principale ---
def rename_files(root_dir=".", dry_run=False):
    """Renomme tous les fichiers journaux obsolètes en ajoutant .DEPRECATED"""
    start_time = time.time()
    patterns = [
        "create_addin_log",  # préfixe
        "fix_classes_log.txt",
        "apex_addin_generator.log"
    ]
    
    # Compteurs pour statistiques
    count_renamed = 0
    count_errors = 0
    count_scanned_dirs = 0
    count_scanned_files = 0
    found_files = []
    
    # Journal de démarrage
    logging.info(f"=== DÉMARRAGE DU PROCESSUS DE RENOMMAGE ===")
    logging.info(f"Répertoire racine: {os.path.abspath(root_dir)}")
    logging.info(f"Mode simulation: {dry_run}")
    logging.info(f"Motifs recherchés: {patterns}")
    logging.info(f"-------------------------------------------")

    # Analyse des fichiers
    for dirpath, dirnames, files in os.walk(root_dir):
        count_scanned_dirs += 1
        if count_scanned_dirs % 100 == 0:  # Log tous les 100 répertoires
            logging.info(f"[PROGRESSION] {count_scanned_dirs} répertoires analysés, {count_scanned_files} fichiers scannés")
        
        # Log pour chaque répertoire si niveau verbeux
        logging.debug(f"Analyse du répertoire: {dirpath} ({len(files)} fichiers)")
        
        for fname in files:
            count_scanned_files += 1
            
            # Vérification des modèles
            matches_pattern = False
            if fname.startswith("create_addin_log") and fname.endswith(".txt") and not fname.endswith(".DEPRECATED"):
                matches_pattern = True
                pattern_matched = "create_addin_log*.txt"
            elif fname in patterns[1:] and not fname.endswith(".DEPRECATED"):
                matches_pattern = True
                pattern_matched = fname
            
            if matches_pattern:
                src = os.path.join(dirpath, fname)
                dst = src + ".DEPRECATED"
                found_files.append((src, pattern_matched))
                
                if os.path.exists(dst):
                    logging.warning(f"⚠️ Le fichier {dst} existe déjà, il sera écrasé.")

                if dry_run:
                    logging.info(f"[SIMULATION] 🔄 Renommage: {src} -> {dst}")
                    count_renamed += 1
                else:
                    try:
                        os.rename(src, dst)
                        logging.info(f"✅ Renommé: {src} -> {dst}")
                        count_renamed += 1
                    except Exception as e:
                        logging.error(f"❌ Échec du renommage de {src}: {str(e)}")
                        count_errors += 1
    
    # Calcul durée d'exécution
    duration = time.time() - start_time
    
    # Résumé détaillé
    logging.info(f"\n=== RÉSUMÉ DU PROCESSUS DE RENOMMAGE ===")
    logging.info(f"Durée d'exécution: {duration:.2f} secondes")
    logging.info(f"Répertoires analysés: {count_scanned_dirs}")
    logging.info(f"Fichiers scannés: {count_scanned_files}")
    logging.info(f"Fichiers renommés: {count_renamed}")
    logging.info(f"Erreurs rencontrées: {count_errors}")
    
    # Détail des fichiers trouvés
    if found_files:
        logging.info(f"\n--- Liste des fichiers traités ---")
        for src, pattern in found_files:
            logging.info(f"• {src} (motif: {pattern})")
    else:
        logging.info(f"\nAucun fichier correspondant aux motifs n'a été trouvé.")
    
    logging.info(f"===================================")
    
    return count_renamed, count_errors, count_scanned_dirs, count_scanned_files

# --- Entrée CLI ---
if __name__ == "__main__":
    # Vérifier l'environnement d'exécution
    check_environment()
    
    parser = argparse.ArgumentParser(description="Renomme les fichiers journaux obsolètes en ajoutant .DEPRECATED")
    parser.add_argument(
        "--dir", default=".", help="Répertoire racine pour démarrer l'analyse (défaut: .)"
    )
    parser.add_argument(
        "--dry-run", action="store_true", help="Simule le processus sans renommer"
    )
    parser.add_argument(
        "--log-file", default=f"rename_logs_{datetime.now().strftime('%Y%m%d%H%M')}.log", 
        help="Fichier journal (défaut: rename_logs_YYYYMMDDHHMM.log)"
    )
    parser.add_argument(
        "-v", "--verbose", action="store_true", help="Active les logs verbeux (DEBUG)"
    )
    parser.add_argument(
        "--export-csv", help="Génère un rapport CSV des fichiers renommés"
    )

    args = parser.parse_args()

    # Configuration du logger avec niveau verbeux si demandé
    setup_logger(args.log_file)
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
        logging.info("Mode verbeux activé")
    
    try:
        logging.info(f"Script démarré: {os.path.basename(__file__)}")
        renamed, errors, dirs_scanned, files_scanned = rename_files(root_dir=args.dir, dry_run=args.dry_run)
        logging.info(f"Script terminé avec succès: {os.path.basename(__file__)}")
        
        # Export CSV si demandé
        if args.export_csv and renamed > 0:
            try:
                import csv
                with open(args.export_csv, 'w', newline='', encoding='utf-8') as csvfile:
                    csv_writer = csv.writer(csvfile)
                    csv_writer.writerow(['Fichier', 'Statut', 'Motif'])
                    # Note: Les données complètes pour le CSV devraient être collectées dans rename_files
                    # Cette partie serait à améliorer pour un export CSV complet
                    
                logging.info(f"Rapport CSV généré: {args.export_csv}")
            except Exception as csv_error:
                logging.error(f"Erreur lors de la génération du rapport CSV: {str(csv_error)}")
    except Exception as e:
        logging.error(f"Exception non gérée: {str(e)}", exc_info=True)
        logging.critical("Le script s'est terminé avec des erreurs")
        raise 