#!/usr/bin/env python3
"""
Script de migration des sessions de développement APEX vers la nouvelle structure.
Réorganise les fichiers de session selon l'arborescence :
/sessions/YYYY/MM/{active|completed}/
"""

import os
import sys
import shutil
import logging
import re
import json
import argparse
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Set
from dataclasses import dataclass, asdict
import concurrent.futures
import time
import csv

@dataclass
class MigrationConfig:
    """Configuration de la migration."""
    dry_run: bool = False
    backup: bool = True
    force: bool = False
    cleanup: bool = False
    validate: bool = True
    parallel: bool = True  # Nouveau: activation du traitement parallèle
    export_csv: bool = True  # Nouveau: export des statistiques en CSV
    max_workers: int = 4  # Nouveau: nombre maximum de workers pour le traitement parallèle

@dataclass
class SessionMetadata:
    """Métadonnées d'une session."""
    filename: str
    year: str
    month: str
    day: str
    status: str
    tasks_count: int
    has_conclusion: bool
    encoding: str

@dataclass
class AIValidationConfig:
    """Configuration de la validation IA."""
    model: str = "claude"  # Modèle IA à utiliser
    context: Dict[str, str] = None  # Contexte de validation
    rules: List[str] = None  # Règles de validation
    mode: str = "strict"  # Mode de validation

class AISessionValidator:
    """Validateur IA pour les sessions."""
    
    def __init__(self, config: AIValidationConfig):
        self.config = config
        self.validation_results = {}
    
    def validate_content(self, content: str) -> List[str]:
        """Validation du contenu par IA."""
        errors = []
        # TODO: Intégration avec l'API IA
        return errors
    
    def suggest_improvements(self, content: str) -> List[str]:
        """Suggestions d'amélioration."""
        suggestions = []
        # TODO: Intégration avec l'API IA
        return suggestions
    
    def check_consistency(self, content: str, metadata: SessionMetadata) -> List[str]:
        """Vérification de cohérence."""
        issues = []
        # TODO: Intégration avec l'API IA
        return issues

class SessionMigrator:
    """Gestionnaire de migration des sessions."""
    
    def __init__(self, workspace_root: str, config: MigrationConfig):
        self.workspace_root = Path(workspace_root)
        self.workflow_dir = self.workspace_root / 'tools' / 'workflow'
        self.old_sessions_dir = self.workflow_dir / 'sessions'
        self.new_sessions_dir = self.workflow_dir / 'sessions'
        self.backup_dir = self.workflow_dir / 'sessions_backup'
        self.session_pattern = re.compile(r'(\d{4})_(\d{2})_(\d{2})_.*\.md$')
        self.migrated_files: List[Dict[str, str]] = []
        self.config = config
        self.metadata_store: Dict[str, SessionMetadata] = {}
        self.start_time = time.time()
        self.processing_times: Dict[str, float] = {}
        self.failed_files: Set[str] = set()
        self.ai_validator = AISessionValidator(AIValidationConfig())

    def validate_workspace(self) -> bool:
        """Vérifie que l'environnement de travail est valide."""
        if not self.workflow_dir.exists():
            logging.error(f"Répertoire workflow non trouvé: {self.workflow_dir}")
            return False
        if not self.old_sessions_dir.exists():
            logging.error(f"Répertoire sessions non trouvé: {self.old_sessions_dir}")
            return False
        return True

    def create_directory_structure(self) -> None:
        """Crée la nouvelle structure de répertoires si nécessaire."""
        current_year = datetime.now().year
        # Crée les répertoires pour l'année en cours et la précédente
        for year in range(current_year - 1, current_year + 1):
            for month in range(1, 13):
                for status in ['active', 'completed']:
                    new_dir = self.new_sessions_dir / str(year) / f"{month:02d}" / status
                    new_dir.mkdir(parents=True, exist_ok=True)
                    logging.info(f"Création du répertoire: {new_dir}")

    def parse_session_date(self, filename: str) -> Tuple[str, str, str]:
        """Extrait la date d'un nom de fichier de session."""
        match = self.session_pattern.match(filename)
        if not match:
            raise ValueError(f"Format de nom de fichier invalide: {filename}")
        return match.groups()  # (year, month, day)

    def determine_session_status(self, content: str) -> str:
        """Détermine si une session est active ou terminée."""
        # Vérifie la présence d'indicateurs de fin de session
        completed_indicators = [
            "## 🛠️ Bilan de session",
            "Session terminée",
            "## Conclusion"
        ]
        return "completed" if any(indicator in content for indicator in completed_indicators) else "active"

    def backup_sessions(self) -> None:
        """Crée une sauvegarde des sessions avant migration."""
        if not self.config.backup:
            return
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_path = self.backup_dir / f"backup_{timestamp}"
        
        if not self.config.dry_run:
            shutil.copytree(self.old_sessions_dir, backup_path)
            logging.info(f"Sauvegarde créée: {backup_path}")
        else:
            logging.info(f"[DRY RUN] Sauvegarde simulée: {backup_path}")

    def cleanup_old_structure(self) -> None:
        """Nettoie l'ancienne structure après migration réussie."""
        if not self.config.cleanup or self.config.dry_run:
            return
            
        for file in self.migrated_files:
            old_path = Path(file['original'])
            if old_path.exists():
                old_path.unlink()
                logging.info(f"Fichier original supprimé: {old_path}")

    def extract_session_metadata(self, file_path: Path, content: str) -> SessionMetadata:
        """Extrait les métadonnées d'une session."""
        year, month, day = self.parse_session_date(file_path.name)
        status = self.determine_session_status(content)
        
        # Analyse du contenu
        tasks = re.findall(r'- \[ \]|\- \[x\]', content)
        has_conclusion = any(marker in content for marker in [
            "## 🛠️ Bilan de session",
            "## Conclusion"
        ])
        
        # Détection de l'encodage
        try:
            content.encode('ascii')
            encoding = 'ascii'
        except UnicodeEncodeError:
            encoding = 'utf-8'
        
        return SessionMetadata(
            filename=file_path.name,
            year=year,
            month=month,
            day=day,
            status=status,
            tasks_count=len(tasks),
            has_conclusion=has_conclusion,
            encoding=encoding
        )

    def validate_session_content(self, content: str) -> List[str]:
        """Valide le contenu d'une session."""
        errors = []
        
        # Validation standard
        required_sections = [
            "## 🎯 Objectif(s)",
            "## 📌 Suivi des tâches",
            "## 🧪 Tests effectués"
        ]
        
        for section in required_sections:
            if section not in content:
                errors.append(f"Section manquante: {section}")
        
        # Validation des blocs de code
        code_blocks = content.count("```")
        if code_blocks % 2 != 0:
            errors.append("Bloc de code non fermé")
            
        # Validation IA
        if self.config.validate:
            ai_errors = self.ai_validator.validate_content(content)
            errors.extend(ai_errors)
            
        return errors

    def migrate_session_file(self, old_path: Path) -> None:
        """Migre un fichier de session vers la nouvelle structure."""
        try:
            # Lecture et validation du contenu
            content = old_path.read_text(encoding='utf-8')
            validation_errors = self.validate_session_content(content)
            
            if validation_errors and not self.config.force:
                logging.warning(f"Validation échouée pour {old_path.name}:")
                for error in validation_errors:
                    logging.warning(f"  - {error}")
                if not self.config.force:
                    raise ValueError("Validation échouée")
            
            # Extraction des métadonnées
            metadata = self.extract_session_metadata(old_path, content)
            self.metadata_store[old_path.name] = metadata
            
            # Création du nouveau chemin
            new_dir = self.new_sessions_dir / metadata.year / metadata.month / metadata.status
            new_path = new_dir / old_path.name
            
            # Copie du fichier
            if not self.config.dry_run:
                new_dir.mkdir(parents=True, exist_ok=True)
                shutil.copy2(old_path, new_path)
            
            self.migrated_files.append({
                'original': str(old_path),
                'new': str(new_path),
                'status': metadata.status
            })
            
            logging.info(f"{'[DRY RUN] ' if self.config.dry_run else ''}Migration réussie: {old_path.name} -> {new_path}")
            
        except Exception as e:
            logging.error(f"Erreur lors de la migration de {old_path}: {str(e)}")
            raise

    def create_migration_report(self) -> None:
        """Génère un rapport de migration détaillé."""
        report_path = self.workflow_dir / 'migration_report.md'
        metadata_path = self.workflow_dir / 'session_metadata.json'
        
        # Statistiques détaillées
        stats = {
            'total': len(self.migrated_files),
            'active': sum(1 for f in self.migrated_files if f['status'] == 'active'),
            'completed': sum(1 for f in self.migrated_files if f['status'] == 'completed'),
            'with_conclusion': sum(1 for m in self.metadata_store.values() if m.has_conclusion),
            'encodings': {
                'utf8': sum(1 for m in self.metadata_store.values() if m.encoding == 'utf-8'),
                'ascii': sum(1 for m in self.metadata_store.values() if m.encoding == 'ascii')
            },
            'ai_validation': {
                'validated': len(self.ai_validator.validation_results),
                'improved': sum(1 for r in self.ai_validator.validation_results.values() if r.get('improvements'))
            }
        }
        
        report_content = [
            "# 📋 Rapport de Migration des Sessions",
            f"\nDate de migration: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "\n## 📊 Statistiques Détaillées",
            f"- Sessions migrées: {stats['total']}",
            f"- Sessions actives: {stats['active']}",
            f"- Sessions terminées: {stats['completed']}",
            f"- Avec conclusion: {stats['with_conclusion']}",
            f"- Encodage UTF-8: {stats['encodings']['utf8']}",
            f"- Encodage ASCII: {stats['encodings']['ascii']}",
            "\n## 📝 Détails de Migration"
        ]
        
        # Détails par session
        for file in self.migrated_files:
            metadata = self.metadata_store[Path(file['original']).name]
            report_content.extend([
                f"\n### {metadata.filename}",
                f"- Statut: {metadata.status}",
                f"- Tâches: {metadata.tasks_count}",
                f"- Conclusion: {'Oui' if metadata.has_conclusion else 'Non'}",
                f"- Encodage: {metadata.encoding}",
                f"- Ancien chemin: {file['original']}",
                f"- Nouveau chemin: {file['new']}"
            ])
        
        if not self.config.dry_run:
            report_path.write_text('\n'.join(report_content), encoding='utf-8')
            # Sauvegarde des métadonnées en JSON pour analyse future
            with open(metadata_path, 'w', encoding='utf-8') as f:
                json.dump({k: asdict(v) for k, v in self.metadata_store.items()}, f, indent=2)
            
            logging.info(f"Rapport de migration créé: {report_path}")
            logging.info(f"Métadonnées sauvegardées: {metadata_path}")
        else:
            logging.info("[DRY RUN] Rapport et métadonnées non sauvegardés")

    def export_statistics(self) -> None:
        """Exporte les statistiques de migration en CSV."""
        if not self.config.export_csv or self.config.dry_run:
            return

        stats_file = self.workflow_dir / 'migration_statistics.csv'
        
        # Calcul des statistiques
        total_time = time.time() - self.start_time
        avg_time = sum(self.processing_times.values()) / len(self.processing_times) if self.processing_times else 0
        
        with open(stats_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Métrique', 'Valeur'])
            writer.writerow(['Temps total (s)', f'{total_time:.2f}'])
            writer.writerow(['Temps moyen par fichier (s)', f'{avg_time:.2f}'])
            writer.writerow(['Fichiers traités', len(self.migrated_files)])
            writer.writerow(['Fichiers en erreur', len(self.failed_files)])
            writer.writerow([''])
            writer.writerow(['Fichier', 'Temps de traitement (s)', 'Statut'])
            
            for file, process_time in self.processing_times.items():
                status = 'Erreur' if file in self.failed_files else 'Succès'
                writer.writerow([file, f'{process_time:.2f}', status])

        logging.info(f"Statistiques exportées dans : {stats_file}")

    def migrate_session_file_with_timing(self, old_path: Path) -> Optional[Dict[str, str]]:
        """Version instrumentée de migrate_session_file."""
        start_time = time.time()
        try:
            self.migrate_session_file(old_path)
            process_time = time.time() - start_time
            self.processing_times[old_path.name] = process_time
            return self.migrated_files[-1] if self.migrated_files else None
        except Exception as e:
            process_time = time.time() - start_time
            self.processing_times[old_path.name] = process_time
            self.failed_files.add(old_path.name)
            logging.error(f"Erreur lors de la migration de {old_path}: {str(e)}")
            return None

    def migrate(self) -> bool:
        """Version améliorée de la migration avec support du parallélisme."""
        try:
            if not self.validate_workspace():
                return False
            
            logging.info(f"{'[DRY RUN] ' if self.config.dry_run else ''}Début de la migration des sessions")
            
            # Sauvegarde si configurée
            self.backup_sessions()
            
            # Création de la structure
            if not self.config.dry_run:
                self.create_directory_structure()
            
            # Migration des fichiers
            session_files = list(self.old_sessions_dir.glob('*.md'))
            
            if self.config.parallel and len(session_files) > 1:
                logging.info(f"Migration parallèle avec {self.config.max_workers} workers")
                with concurrent.futures.ThreadPoolExecutor(max_workers=self.config.max_workers) as executor:
                    future_to_file = {
                        executor.submit(self.migrate_session_file_with_timing, file_path): file_path
                        for file_path in session_files
                    }
                    
                    for future in concurrent.futures.as_completed(future_to_file):
                        file_path = future_to_file[future]
                        try:
                            result = future.result()
                            if result:
                                logging.info(f"Migration parallèle réussie pour: {file_path.name}")
                        except Exception as e:
                            logging.error(f"Erreur lors de la migration parallèle de {file_path}: {str(e)}")
            else:
                for file_path in session_files:
                    self.migrate_session_file_with_timing(file_path)
            
            # Export des statistiques
            self.export_statistics()
            
            # Génération du rapport
            self.create_migration_report()
            
            # Nettoyage si configuré
            self.cleanup_old_structure()
            
            success = len(self.migrated_files) > 0 and len(self.failed_files) == 0
            logging.info(f"{'[DRY RUN] ' if self.config.dry_run else ''}Migration terminée avec {'succès' if success else 'des erreurs'}")
            return success
            
        except Exception as e:
            logging.error(f"Erreur lors de la migration: {str(e)}")
            return False

def parse_args() -> MigrationConfig:
    """Version améliorée du parsing des arguments."""
    parser = argparse.ArgumentParser(description="Migration des sessions de développement APEX")
    parser.add_argument('--dry-run', action='store_true', help="Simulation sans modification")
    parser.add_argument('--no-backup', action='store_true', help="Désactive la sauvegarde")
    parser.add_argument('--force', action='store_true', help="Force la migration malgré les erreurs")
    parser.add_argument('--cleanup', action='store_true', help="Nettoie l'ancienne structure")
    parser.add_argument('--no-validate', action='store_true', help="Désactive la validation")
    parser.add_argument('--no-parallel', action='store_true', help="Désactive le traitement parallèle")
    parser.add_argument('--no-csv', action='store_true', help="Désactive l'export CSV des statistiques")
    parser.add_argument('--max-workers', type=int, default=4, help="Nombre maximum de workers parallèles")
    
    args = parser.parse_args()
    return MigrationConfig(
        dry_run=args.dry_run,
        backup=not args.no_backup,
        force=args.force,
        cleanup=args.cleanup,
        validate=not args.no_validate,
        parallel=not args.no_parallel,
        export_csv=not args.no_csv,
        max_workers=args.max_workers
    )

def main():
    """Point d'entrée principal."""
    try:
        # Configuration du logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('session_migration.log'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        
        # Parsing des arguments
        config = parse_args()
        
        # Détection du chemin du workspace
        workspace_root = os.path.abspath(os.path.join(
            os.path.dirname(__file__),
            "..",
            "..",
            ".."
        ))
        
        migrator = SessionMigrator(workspace_root, config)
        success = migrator.migrate()
        
        if success:
            print("\n✅ Migration terminée avec succès")
            if not config.dry_run:
                print("📝 Consultez session_migration.log et migration_report.md pour plus de détails")
            else:
                print("📝 [DRY RUN] Simulation terminée, consultez session_migration.log")
        else:
            print("\n❌ La migration a échoué")
            print("📝 Consultez session_migration.log pour plus de détails")
        
        sys.exit(0 if success else 1)
        
    except Exception as e:
        logging.error(f"Erreur fatale: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main() 