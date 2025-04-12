#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Agent Documentaire APEX Framework
---------------------------------

Ce script analyse les fichiers du projet APEX VBA Framework pour vérifier
la conformité des commentaires et de la documentation avec les guidelines définies.

Usage:
    python doc_agent.py [--target=<folder>] [--fix] [--report=<path>]
    
Options:
    --target=<folder>   Dossier à analyser [default: .]
    --fix               Tenter de corriger automatiquement les problèmes
    --report=<path>     Chemin pour le rapport [default: ./reports/doc_compliance.md]
"""

import os
import re
import json
import datetime
import argparse
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any


class DocumentationRule:
    """Classe de base pour les règles de documentation"""
    
    def __init__(self, name: str, description: str, severity: str = "WARNING"):
        self.name = name
        self.description = description
        self.severity = severity
    
    def check(self, file_path: Path, content: str) -> List[Dict[str, Any]]:
        """Vérifie si le contenu respecte la règle"""
        # À implémenter dans les classes dérivées
        return []
    
    def fix(self, content: str) -> Tuple[str, bool]:
        """Tente de corriger le contenu pour respecter la règle"""
        # À implémenter dans les classes dérivées
        return content, False


class VBACommentRule(DocumentationRule):
    """Vérifie si les fichiers VBA ont les commentaires conformes aux standards"""
    
    def __init__(self):
        super().__init__(
            "VBAComment", 
            "Les commentaires VBA doivent suivre les standards documentaires",
            "WARNING"
        )
        self.required_module_tags = [
            "@Module",
            "@Description",
            "@Version",
            "@Date",
            "@Author"
        ]
        self.method_tags = [
            "@Description",
            "@Param",
            "@Returns"
        ]
    
    def check(self, file_path: Path, content: str) -> List[Dict[str, Any]]:
        issues = []
        
        # Vérifier la présence des tags module requis
        missing_tags = []
        for tag in self.required_module_tags:
            if not re.search(f"'@{tag}:", content, re.IGNORECASE):
                missing_tags.append(tag)
        
        if missing_tags:
            issues.append({
                "rule": self.name,
                "severity": self.severity,
                "message": f"Tags module manquants: {', '.join(missing_tags)}",
                "file": str(file_path),
                "line": 1,
                "can_fix": True,
                "missing_tags": missing_tags
            })
        
        # Extraction des méthodes/fonctions et vérification des commentaires
        method_pattern = r"(?:Public|Private|Friend)?\s+(?:Sub|Function)\s+([A-Za-z0-9_]+).*?(?:\n|\r\n)"
        methods = re.finditer(method_pattern, content)
        
        for match in methods:
            method_name = match.group(1)
            method_pos = match.start()
            line_num = content[:method_pos].count('\n') + 1
            
            # Vérifier si des commentaires documentaires précèdent
            prev_content = content[:method_pos].strip()
            has_desc = re.search(r"'@Description:.*?(?:\n|\r\n)", prev_content, re.IGNORECASE)
            
            if not has_desc and method_name not in ["Class_Initialize", "Class_Terminate"]:
                issues.append({
                    "rule": self.name,
                    "severity": "INFO",
                    "message": f"Méthode '{method_name}' sans documentation @Description",
                    "file": str(file_path),
                    "line": line_num,
                    "can_fix": True,
                    "method_name": method_name
                })
        
        return issues
    
    def fix(self, content: str) -> Tuple[str, bool]:
        fixed = False
        
        # Ajouter les tags module manquants si nécessaire
        if not any(f"'@{tag}:" in content for tag in self.required_module_tags):
            header = """'@Module: {module_name}
'@Description: 
'@Version: 1.0
'@Date: {date}
'@Author: APEX Framework Team

"""
            today = datetime.datetime.now().strftime("%d/%m/%Y")
            header = header.format(module_name="[NomDuModule]", date=today)
            
            # Insérer après Attribute VB_Name ou Option Explicit
            if "Attribute VB_Name" in content:
                content = re.sub(r"(Attribute VB_Name.*?)(\n|\r\n)", r"\1\2\n" + header, content)
                fixed = True
            elif "Option Explicit" in content:
                content = re.sub(r"(Option Explicit.*?)(\n|\r\n)", r"\1\2\n" + header, content)
                fixed = True
            else:
                content = header + content
                fixed = True
        
        # Vérifier chaque méthode et ajouter des templates de documentation
        method_pattern = r"(?:Public|Private|Friend)?\s+(?:Sub|Function)\s+([A-Za-z0-9_]+).*?(?:\n|\r\n)"
        
        # Créer une liste de tuples (position, méthode, commentaire)
        methods_to_fix = []
        for match in re.finditer(method_pattern, content):
            method_name = match.group(1)
            method_pos = match.start()
            
            # Ignorer les méthodes d'initialisation de classe
            if method_name in ["Class_Initialize", "Class_Terminate"]:
                continue
                
            # Vérifier si la méthode est déjà documentée
            # On cherche dans les 10 lignes précédentes
            prev_lines = content[:method_pos].split('\n')[-10:]
            prev_content = '\n'.join(prev_lines)
            
            has_desc = re.search(r"'@Description:", prev_content, re.IGNORECASE)
            
            if not has_desc:
                # Créer un template de documentation pour cette méthode
                doc_template = f"""'@Description: 
'@Param: 
'@Returns: 

"""
                methods_to_fix.append((method_pos, method_name, doc_template))
        
        # Appliquer les corrections en commençant par la fin pour éviter de décaler les positions
        for pos, method_name, doc_template in sorted(methods_to_fix, reverse=True):
            content = content[:pos] + doc_template + content[pos:]
            fixed = True
            
        return content, fixed


class MarkdownStructureRule(DocumentationRule):
    """Vérifie si les fichiers Markdown ont la structure correcte"""
    
    def __init__(self):
        super().__init__(
            "MarkdownStructure", 
            "Les fichiers Markdown doivent suivre la structure standard",
            "WARNING"
        )
    
    def check(self, file_path: Path, content: str) -> List[Dict[str, Any]]:
        issues = []
        
        # Vérifier le titre principal (H1)
        if not re.match(r'^# ', content.strip()):
            issues.append({
                "rule": self.name,
                "severity": self.severity,
                "message": "Le document doit commencer par un titre H1 (# Titre)",
                "file": str(file_path),
                "line": 1,
                "can_fix": False
            })
        
        # Vérifier les sections obligatoires pour les guides
        if "GUIDE" in file_path.name.upper() or "DOCUMENTATION" in file_path.name.upper():
            required_sections = ["## Objectif", "## Prérequis", "## Utilisation"]
            for section in required_sections:
                if section not in content:
                    issues.append({
                        "rule": self.name,
                        "severity": self.severity,
                        "message": f"Section obligatoire manquante: {section}",
                        "file": str(file_path),
                        "line": 1,
                        "can_fix": True,
                        "missing_section": section
                    })
        
        return issues
    
    def fix(self, content: str) -> Tuple[str, bool]:
        # Si pas de titre H1, en ajouter un
        fixed = False
        if not re.match(r'^# ', content.strip()):
            filename_base = os.path.basename(os.path.dirname(content))
            title = f"# {filename_base}\n\n"
            content = title + content
            fixed = True
            
        # Si c'est un guide et qu'il manque les sections obligatoires
        if "GUIDE" in content or "Documentation" in content:
            template = """
## Objectif

## Prérequis

## Utilisation

## Exemples

"""
            if "## Objectif" not in content:
                # Ajouter après le titre H1 ou à la fin
                h1_match = re.search(r'^# .*(\n|\r\n)', content)
                if h1_match:
                    pos = h1_match.end()
                    content = content[:pos] + "\n" + template + content[pos:]
                else:
                    content += "\n" + template
                fixed = True
                
        return content, fixed


class DocumentationAgent:
    """Agent principal pour vérifier la documentation"""
    
    def __init__(self, target_folder: str, config_path: Optional[str] = None):
        self.target_folder = Path(target_folder)
        self.config_path = config_path
        self.config = self._load_config()
        self.rules = self._init_rules()
        
    def _load_config(self) -> Dict:
        """Charge la configuration depuis un fichier JSON"""
        default_config = {
            "vba_patterns": {
                "module_header": [
                    "@Module",
                    "@Description",
                    "@Version",
                    "@Date",
                    "@Author"
                ],
                "method_header": [
                    "@Description",
                    "@Param",
                    "@Returns"
                ],
                "class_prefixes": ["cls"],
                "module_prefixes": ["mod"],
                "form_prefixes": ["frm"]
            },
            "markdown_patterns": {
                "required_sections": {
                    "guide": ["Objectif", "Prérequis", "Utilisation", "Exemples"],
                    "api": ["Description", "Interface", "Méthodes", "Exemples d'utilisation"],
                    "component": ["Vue d'ensemble", "Architecture", "Dépendances", "Configuration", "Utilisation"]
                }
            },
            "file_patterns": {
                "vba": [".cls", ".bas", ".frm"],
                "markdown": [".md"],
                "config": [".json", ".ini"]
            }
        }
        
        if self.config_path and os.path.exists(self.config_path):
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        return default_config
    
    def _init_rules(self) -> List[DocumentationRule]:
        """Initialise les règles de vérification"""
        return [
            VBACommentRule(),
            MarkdownStructureRule()
        ]
    
    def scan_files(self) -> List[Path]:
        """Parcourt le dossier cible pour trouver les fichiers à analyser"""
        files = []
        vba_extensions = tuple(self.config["file_patterns"]["vba"])
        markdown_extensions = tuple(self.config["file_patterns"]["markdown"])
        
        # Si la cible est un fichier spécifique
        if self.target_folder.is_file():
            if (self.target_folder.suffix.lower() in vba_extensions or 
                self.target_folder.suffix.lower() in markdown_extensions):
                files.append(self.target_folder)
            return files
            
        # Si la cible est un dossier
        for file_path in self.target_folder.glob("**/*"):
            if file_path.is_file():
                # Ignorer les dossiers ignorés courants
                ignore_dirs = [".git", ".vscode", "__pycache__", "venv", "node_modules"]
                if any(ignore_dir in str(file_path) for ignore_dir in ignore_dirs):
                    continue
                    
                if file_path.suffix.lower() in vba_extensions or file_path.suffix.lower() in markdown_extensions:
                    files.append(file_path)
        
        return files
    
    def analyze_file(self, file_path: Path) -> List[Dict[str, Any]]:
        """Analyse un fichier pour vérifier sa conformité"""
        issues = []
        
        # Lecture du contenu
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except UnicodeDecodeError:
            try:
                with open(file_path, 'r', encoding='latin1') as f:
                    content = f.read()
            except Exception as e:
                return [{
                    "rule": "FileEncoding",
                    "severity": "ERROR",
                    "message": f"Impossible de lire le fichier: {str(e)}",
                    "file": str(file_path),
                    "line": 1,
                    "can_fix": False
                }]
        
        # Appliquer les règles appropriées
        for rule in self.rules:
            if file_path.suffix.lower() in self.config["file_patterns"]["vba"] and isinstance(rule, VBACommentRule):
                issues.extend(rule.check(file_path, content))
            elif file_path.suffix.lower() in self.config["file_patterns"]["markdown"] and isinstance(rule, MarkdownStructureRule):
                issues.extend(rule.check(file_path, content))
        
        return issues
    
    def fix_issues(self, issues: List[Dict[str, Any]]) -> int:
        """Tente de corriger les problèmes identifiés"""
        fixed_count = 0
        
        # Regrouper les problèmes par fichier
        files_with_issues = {}
        for issue in issues:
            if issue.get("can_fix", False):
                file_path = issue["file"]
                if file_path not in files_with_issues:
                    files_with_issues[file_path] = []
                files_with_issues[file_path].append(issue)
        
        # Traiter chaque fichier
        for file_path, file_issues in files_with_issues.items():
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
            except UnicodeDecodeError:
                try:
                    with open(file_path, 'r', encoding='latin1') as f:
                        content = f.read()
                except Exception:
                    continue  # Skip if can't read file
            
            original_content = content
            modified = False
            
            # Appliquer les règles pour corriger
            for rule in self.rules:
                if any(i["rule"] == rule.name for i in file_issues):
                    if Path(file_path).suffix.lower() in self.config["file_patterns"]["vba"] and isinstance(rule, VBACommentRule):
                        content, was_fixed = rule.fix(content)
                        if was_fixed:
                            modified = True
                    elif Path(file_path).suffix.lower() in self.config["file_patterns"]["markdown"] and isinstance(rule, MarkdownStructureRule):
                        content, was_fixed = rule.fix(content)
                        if was_fixed:
                            modified = True
            
            # Sauvegarder si modifié
            if modified:
                try:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(content)
                    fixed_count += 1
                except Exception as e:
                    print(f"Erreur lors de la correction de {file_path}: {str(e)}")
        
        return fixed_count
    
    def generate_report(self, issues: List[Dict[str, Any]], report_path: str) -> None:
        """Génère un rapport Markdown des problèmes identifiés"""
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        
        # Créer le dossier de rapport si nécessaire
        os.makedirs(os.path.dirname(report_path), exist_ok=True)
        
        # Regrouper les problèmes par fichier
        files_with_issues = {}
        for issue in issues:
            file_path = issue["file"]
            if file_path not in files_with_issues:
                files_with_issues[file_path] = []
            files_with_issues[file_path].append(issue)
        
        # Compter par sévérité
        errors = sum(1 for i in issues if i["severity"] == "ERROR")
        warnings = sum(1 for i in issues if i["severity"] == "WARNING")
        infos = sum(1 for i in issues if i["severity"] == "INFO")
        
        # Analyser les types de fichiers
        vba_files = len({f for f, i in files_with_issues.items() if Path(f).suffix.lower() in self.config["file_patterns"]["vba"]})
        md_files = len({f for f, i in files_with_issues.items() if Path(f).suffix.lower() in self.config["file_patterns"]["markdown"]})
        
        # Générer le rapport
        report = f"""# Rapport de Conformité Documentaire APEX
Date: {now}

## Résumé
- Fichiers analysés: {len(set(i["file"] for i in issues))}
  - Fichiers VBA: {vba_files}
  - Fichiers Markdown: {md_files}
- Problèmes détectés: {len(issues)}
  - Erreurs: {errors}
  - Avertissements: {warnings}
  - Informations: {infos}

## Détails des problèmes
"""
        
        # Trier par sévérité (ERROR > WARNING > INFO)
        severity_order = {"ERROR": 0, "WARNING": 1, "INFO": 2}
        
        # Ajouter les détails pour chaque fichier
        for file_path, file_issues in sorted(files_with_issues.items(), 
                                            key=lambda x: min(severity_order.get(i["severity"], 3) for i in x[1])):
            report += f"\n### {file_path}\n"
            
            # Trier les problèmes par numéro de ligne puis par sévérité
            sorted_issues = sorted(file_issues, 
                                  key=lambda x: (x.get("line", 0), severity_order.get(x["severity"], 3)))
            
            for issue in sorted_issues:
                severity_icon = "🔴" if issue["severity"] == "ERROR" else "🟡" if issue["severity"] == "WARNING" else "🔵"
                report += f"- {severity_icon} **{issue['rule']}**: {issue['message']} (ligne {issue.get('line', '?')})\n"
                
                # Ajouter une suggestion si disponible
                if "method_name" in issue:
                    report += f"  > Suggestion: Ajouter documentation pour la méthode '{issue['method_name']}'\n"
                elif "missing_tags" in issue:
                    report += f"  > Suggestion: Ajouter les tags manquants: {', '.join(issue['missing_tags'])}\n"
                elif "missing_section" in issue:
                    report += f"  > Suggestion: Ajouter la section '{issue['missing_section']}'\n"
        
        # Écrire le rapport
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report)
    
    def run(self, fix: bool = False, report_path: Optional[str] = None) -> int:
        """Exécute l'analyse complète"""
        print(f"🔍 Analyse de la documentation dans {self.target_folder}")
        
        # Trouver les fichiers à analyser
        files = self.scan_files()
        print(f"📁 {len(files)} fichiers trouvés")
        
        # Analyser chaque fichier
        all_issues = []
        for file_path in files:
            print(f"  Analyse de {file_path.name}...")
            issues = self.analyze_file(file_path)
            if issues:
                all_issues.extend(issues)
        
        # Afficher un résumé
        error_count = sum(1 for i in all_issues if i["severity"] == "ERROR")
        warning_count = sum(1 for i in all_issues if i["severity"] == "WARNING")
        info_count = sum(1 for i in all_issues if i["severity"] == "INFO")
        
        print(f"❗ {error_count} erreurs, {warning_count} avertissements, et {info_count} informations détectés")
        
        # Corriger si demandé
        if fix and all_issues:
            fixed = self.fix_issues([i for i in all_issues if i.get("can_fix", False)])
            print(f"🔧 {fixed} problèmes corrigés automatiquement")
        
        # Générer un rapport si demandé
        if report_path and all_issues:
            self.generate_report(all_issues, report_path)
            print(f"📝 Rapport généré: {report_path}")
        
        return len(all_issues)


def main():
    """Fonction principale"""
    parser = argparse.ArgumentParser(description="Agent de vérification de documentation APEX")
    parser.add_argument("--target", default=".", help="Dossier ou fichier à analyser")
    parser.add_argument("--config", help="Chemin vers la configuration")
    parser.add_argument("--fix", action="store_true", help="Tenter de corriger les problèmes")
    parser.add_argument("--report", help="Chemin pour le rapport")
    parser.add_argument("--verbose", action="store_true", help="Mode verbeux")
    
    args = parser.parse_args()
    
    agent = DocumentationAgent(args.target, args.config)
    issue_count = agent.run(
        fix=args.fix,
        report_path=args.report
    )
    
    # Retourne 0 si pas de problème, sinon le nombre de problèmes
    # Mais uniquement les erreurs et avertissements comptent pour le code de retour
    return min(issue_count, 1) if issue_count > 0 else 0


if __name__ == "__main__":
    exit(main())