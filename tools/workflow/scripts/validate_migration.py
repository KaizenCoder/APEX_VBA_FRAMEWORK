#!/usr/bin/env python3
"""
Script de validation post-migration des sessions APEX.
Vérifie l'intégrité de la nouvelle structure et des fichiers migrés.
"""

import os
import sys
import json
import logging
import hashlib
from pathlib import Path
from typing import Dict, List, Set
from dataclasses import dataclass
from datetime import datetime

@dataclass
class ValidationResult:
    """Résultat de validation d'une session."""
    filename: str
    path: Path
    is_valid: bool
    errors: List[str]
    warnings: List[str]

class MigrationValidator:
    """Validateur de la migration des sessions."""
    
    def __init__(self, workspace_root: str):
        self.workspace_root = Path(workspace_root)
        self.workflow_dir = self.workspace_root / 'tools' / 'workflow'
        self.sessions_dir = self.workflow_dir / 'sessions'
        self.metadata_file = self.workflow_dir / 'session_metadata.json'
        self.results: Dict[str, ValidationResult] = {}
        self.original_files: Set[str] = set()
        
    def load_metadata(self) -> Dict[str, dict]:
        """Charge les métadonnées de migration."""
        if not self.metadata_file.exists():
            logging.error(f"Fichier de métadonnées non trouvé: {self.metadata_file}")
            return {}
        
        with open(self.metadata_file, 'r', encoding='utf-8') as f:
            return json.load(f)
            
    def calculate_file_hash(self, file_path: Path) -> str:
        """Calcule le hash SHA-256 d'un fichier."""
        sha256_hash = hashlib.sha256()
        with open(file_path, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
        
    def validate_file_structure(self, file_path: Path, metadata: dict) -> List[str]:
        """Valide la structure d'un fichier de session."""
        errors = []
        
        # Vérification du chemin
        expected_path = self.sessions_dir / metadata['year'] / metadata['month'] / metadata['status'] / metadata['filename']
        if file_path != expected_path:
            errors.append(f"Chemin incorrect: {file_path} (attendu: {expected_path})")
            
        # Vérification du contenu
        try:
            content = file_path.read_text(encoding='utf-8')
            
            # Sections requises
            required_sections = [
                "## 🎯 Objectif(s)",
                "## 📌 Suivi des tâches",
                "## 🧪 Tests effectués"
            ]
            
            for section in required_sections:
                if section not in content:
                    errors.append(f"Section manquante: {section}")
            
            # Validation de l'encodage
            try:
                content.encode('ascii')
                if metadata['encoding'] != 'ascii':
                    errors.append("Incohérence d'encodage: fichier ASCII mais métadonnées UTF-8")
            except UnicodeEncodeError:
                if metadata['encoding'] != 'utf-8':
                    errors.append("Incohérence d'encodage: fichier UTF-8 mais métadonnées ASCII")
                    
        except Exception as e:
            errors.append(f"Erreur de lecture: {str(e)}")
            
        return errors
        
    def validate_session(self, file_path: Path, metadata: dict) -> ValidationResult:
        """Valide une session migrée."""
        errors = []
        warnings = []
        
        # Validation de base
        if not file_path.exists():
            errors.append("Fichier non trouvé")
            return ValidationResult(
                filename=metadata['filename'],
                path=file_path,
                is_valid=False,
                errors=errors,
                warnings=warnings
            )
            
        # Validation de la structure
        structure_errors = self.validate_file_structure(file_path, metadata)
        errors.extend(structure_errors)
        
        # Validation des tâches
        try:
            content = file_path.read_text(encoding='utf-8')
            task_count = len([line for line in content.split('\n') if line.strip().startswith('- [')])
            if task_count != metadata['tasks_count']:
                warnings.append(f"Nombre de tâches incohérent: {task_count} (métadonnées: {metadata['tasks_count']})")
                
            # Validation de la conclusion
            has_conclusion = any(marker in content for marker in ["## 🛠️ Bilan de session", "## Conclusion"])
            if has_conclusion != metadata['has_conclusion']:
                warnings.append("État de conclusion incohérent avec les métadonnées")
                
        except Exception as e:
            errors.append(f"Erreur lors de la validation du contenu: {str(e)}")
            
        return ValidationResult(
            filename=metadata['filename'],
            path=file_path,
            is_valid=len(errors) == 0,
            errors=errors,
            warnings=warnings
        )
        
    def generate_validation_report(self) -> None:
        """Génère un rapport de validation."""
        report_path = self.workflow_dir / 'validation_report.md'
        
        # Statistiques
        total_files = len(self.results)
        valid_files = sum(1 for r in self.results.values() if r.is_valid)
        files_with_warnings = sum(1 for r in self.results.values() if r.warnings)
        
        report_content = [
            "# 🔍 Rapport de Validation Post-Migration",
            f"\nDate de validation: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "\n## 📊 Résumé",
            f"- Fichiers validés: {total_files}",
            f"- Fichiers valides: {valid_files}",
            f"- Fichiers avec avertissements: {files_with_warnings}",
            f"- Fichiers avec erreurs: {total_files - valid_files}",
            "\n## 📝 Détails par Fichier"
        ]
        
        # Tri par statut puis par nom
        sorted_results = sorted(
            self.results.items(),
            key=lambda x: (not x[1].is_valid, bool(x[1].warnings), x[0])
        )
        
        for filename, result in sorted_results:
            status = "✅" if result.is_valid and not result.warnings else "⚠️" if result.is_valid else "❌"
            report_content.extend([
                f"\n### {status} {filename}",
                f"- Chemin: {result.path}",
                f"- Statut: {'Valide' if result.is_valid else 'Non valide'}"
            ])
            
            if result.errors:
                report_content.append("- Erreurs:")
                for error in result.errors:
                    report_content.append(f"  * {error}")
                    
            if result.warnings:
                report_content.append("- Avertissements:")
                for warning in result.warnings:
                    report_content.append(f"  * {warning}")
                    
        report_path.write_text('\n'.join(report_content), encoding='utf-8')
        logging.info(f"Rapport de validation créé: {report_path}")
        
    def validate(self) -> bool:
        """Exécute la validation complète."""
        try:
            logging.info("Début de la validation post-migration")
            
            # Chargement des métadonnées
            metadata = self.load_metadata()
            if not metadata:
                return False
                
            # Validation de chaque session
            for filename, session_metadata in metadata.items():
                file_path = self.sessions_dir / session_metadata['year'] / session_metadata['month'] / session_metadata['status'] / filename
                result = self.validate_session(file_path, session_metadata)
                self.results[filename] = result
                
                if not result.is_valid:
                    logging.warning(f"Validation échouée pour {filename}")
                    for error in result.errors:
                        logging.error(f"  - {error}")
                elif result.warnings:
                    logging.warning(f"Avertissements pour {filename}")
                    for warning in result.warnings:
                        logging.warning(f"  - {warning}")
                else:
                    logging.info(f"Validation réussie pour {filename}")
                    
            # Génération du rapport
            self.generate_validation_report()
            
            # Vérification finale
            success = all(result.is_valid for result in self.results.values())
            logging.info(f"Validation terminée: {'Succès' if success else 'Échec'}")
            return success
            
        except Exception as e:
            logging.error(f"Erreur lors de la validation: {str(e)}")
            return False

def main():
    """Point d'entrée principal."""
    try:
        # Configuration du logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('session_validation.log'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        
        # Détection du chemin du workspace
        workspace_root = os.path.abspath(os.path.join(
            os.path.dirname(__file__),
            "..",
            "..",
            ".."
        ))
        
        validator = MigrationValidator(workspace_root)
        success = validator.validate()
        
        if success:
            print("\n✅ Validation terminée avec succès")
            print("📝 Consultez validation_report.md pour plus de détails")
        else:
            print("\n❌ La validation a échoué")
            print("📝 Consultez session_validation.log et validation_report.md pour plus de détails")
            
        sys.exit(0 if success else 1)
        
    except Exception as e:
        logging.error(f"Erreur fatale: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main() 