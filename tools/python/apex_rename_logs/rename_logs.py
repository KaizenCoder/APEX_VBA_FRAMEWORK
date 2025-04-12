#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Module CLI pour renommer les fichiers journaux obsolètes en ajoutant l'extension .DEPRECATED
Compatible avec Windows et optimisé pour les performances sur les systèmes de fichiers Windows.
"""
import os
import argparse
import logging
import time
import platform
import sys
import csv
from datetime import datetime


# --- Vérification d'environnement ---
def check_environment():
    """Vérifie si le script est exécuté dans l'environnement approprié"""
    # Détection de WSL
    if "microsoft" in platform.release().lower() or "linux" in platform.system().lower():
        print("🚨 ERREUR D'ENVIRONNEMENT")
        print("Ce script ne doit pas être lancé depuis WSL. Utilisez PowerShell Windows natif.")
        print("\nCommande à utiliser sous PowerShell:")
        print('   apex-rename-logs --dry-run -v')
        print("\nOu utilisez le script .bat fourni:")
        print('   D:\\chemin\\vers\\run_rename_logs.bat')
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
def rename_files(root_dir=".", dry_run=False, export_csv=None):
    """Renomme tous les fichiers journaux obsolètes en ajoutant .DEPRECATED"""
    start_time = time.time()
    
    # Patterns pour les noms de fichiers
    file_patterns = [
        "create_addin_log",  # préfixe
        "fix_classes_log.txt",
        "apex_addin_generator.log",
        "CreateApexAddIn.ps1"  # Script PowerShell obsolète
    ]
    
    # Extensions de fichiers à vérifier pour des marqueurs d'obsolescence dans leur contenu
    content_extensions = ['.ps1', '.py', '.bat', '.cmd']
    
    # Marqueurs d'obsolescence dans le contenu des fichiers
    content_markers = [
        "# DEPRECATED", "' DEPRECATED", "REM DEPRECATED",
        "# Ce script est obsolète", "' Ce script est obsolète", "REM Ce script est obsolète"
    ]
    
    # Compteurs pour statistiques
    count_renamed = 0
    count_errors = 0
    count_scanned_dirs = 0
    count_scanned_files = 0
    found_files = []
    csv_data = []  # Pour l'export CSV
    
    # Journal de démarrage
    logging.info(f"=== DÉMARRAGE DU PROCESSUS DE RENOMMAGE ===")
    logging.info(f"Répertoire racine: {os.path.abspath(root_dir)}")
    logging.info(f"Mode simulation: {dry_run}")
    logging.info(f"Motifs de noms recherchés: {file_patterns}")
    logging.info(f"Extensions pour analyse de contenu: {content_extensions}")
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
            file_path = os.path.join(dirpath, fname)
            
            # Ignorer les fichiers déjà marqués comme .DEPRECATED
            if fname.endswith(".DEPRECATED"):
                continue
                
            # Vérification des modèles de noms de fichiers
            matches_pattern = False
            pattern_matched = ""
            
            # Vérification du nom de fichier
            if fname.startswith("create_addin_log") and fname.endswith(".txt"):
                matches_pattern = True
                pattern_matched = "create_addin_log*.txt"
            elif fname in file_patterns[1:]:
                matches_pattern = True
                pattern_matched = fname
            
            # Si le nom ne correspond pas, vérifier le contenu pour certaines extensions
            if not matches_pattern and os.path.splitext(fname)[1] in content_extensions:
                try:
                    # Lire seulement les premières lignes du fichier pour efficacité
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                        # Lire les 10 premières lignes, suffisant pour détecter les commentaires d'en-tête
                        header = ''.join([next(f, '') for _ in range(10)])
                        
                    # Vérifier si un des marqueurs est présent dans l'en-tête
                    for marker in content_markers:
                        if marker in header:
                            matches_pattern = True
                            pattern_matched = f"Contenu: {marker}"
                            break
                except Exception as e:
                    logging.debug(f"Erreur lors de la lecture du fichier {file_path}: {str(e)}")
            
            # Si un motif correspond, renommer le fichier
            if matches_pattern:
                src = file_path
                dst = src + ".DEPRECATED"
                found_files.append((src, pattern_matched))
                
                if os.path.exists(dst):
                    logging.warning(f"⚠️ Le fichier {dst} existe déjà, il sera écrasé.")

                if dry_run:
                    status = "Simulation"
                    logging.info(f"[SIMULATION] 🔄 Renommage: {src} -> {dst}")
                    count_renamed += 1
                else:
                    try:
                        os.rename(src, dst)
                        status = "Renommé"
                        logging.info(f"✅ Renommé: {src} -> {dst}")
                        count_renamed += 1
                    except Exception as e:
                        status = f"Erreur: {str(e)}"
                        logging.error(f"❌ Échec du renommage de {src}: {str(e)}")
                        count_errors += 1
                
                # Ajouter les données pour le CSV si export demandé
                if export_csv is not None:
                    csv_data.append([src, status, pattern_matched])
    
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
    
    # Export CSV si demandé
    if export_csv and csv_data:
        try:
            with open(export_csv, 'w', newline='', encoding='utf-8') as csvfile:
                csv_writer = csv.writer(csvfile)
                csv_writer.writerow(['Fichier', 'Statut', 'Motif'])
                csv_writer.writerows(csv_data)
            logging.info(f"Rapport CSV généré: {export_csv}")
        except Exception as csv_error:
            logging.error(f"Erreur lors de la génération du rapport CSV: {str(csv_error)}")
    
    return count_renamed, count_errors, count_scanned_dirs, count_scanned_files, csv_data


def main():
    """Point d'entrée principal pour le CLI"""
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
        logging.info(f"Script démarré: apex-rename-logs")
        renamed, errors, dirs_scanned, files_scanned, _ = rename_files(
            root_dir=args.dir, 
            dry_run=args.dry_run,
            export_csv=args.export_csv
        )
        logging.info(f"Script terminé avec succès: apex-rename-logs")
        
        # Retourner un code de sortie en fonction du résultat
        if errors > 0:
            return 1  # Erreur
        return 0  # Succès
        
    except Exception as e:
        logging.error(f"Exception non gérée: {str(e)}", exc_info=True)
        logging.critical("Le script s'est terminé avec des erreurs")
        return 1


if __name__ == "__main__":
    sys.exit(main()) 