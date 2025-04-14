#!/usr/bin/env python3
# -----------------------------------------------------------------------------
# Script: vba_validator.py
# Description: Validateur de code VBA pour APEX Framework
# Author: APEX Framework Team
# Date: 2025-04-13
# Version: 1.0
# -----------------------------------------------------------------------------

import os
import sys
import re
import json
import argparse
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple, Any, Optional, Set


class VBAValidator:
    """Classe principale pour la validation du code VBA"""
    
    # Constantes pour les types de problèmes
    SEVERITY_INFO = "INFO"
    SEVERITY_WARNING = "WARNING"
    SEVERITY_ERROR = "ERROR"
    
    def __init__(self, config_path: Optional[str] = None):
        """Initialisation du validateur avec configuration"""
        self.issues = []
        self.stats = {
            "files_processed": 0,
            "lines_processed": 0,
            "issues_found": 0,
            "error_count": 0,
            "warning_count": 0,
            "info_count": 0
        }
        self.config = self._load_config(config_path)
    
    def _load_config(self, config_path: Optional[str] = None) -> Dict:
        """Charge la configuration depuis un fichier JSON ou utilise les valeurs par défaut"""
        default_config = {
            "naming": {
                "module_prefix": {
                    "standard": "mod",
                    "class": "cls",
                    "form": "frm"
                },
                "variable_prefixes": {
                    "Boolean": "b",
                    "Integer": "i",
                    "Long": "l",
                    "String": "s",
                    "Double": "d",
                    "Object": "obj",
                    "Variant": "v"
                },
                "case": {
                    "functions": "PascalCase",
                    "variables": "camelCase",
                    "constants": "ALL_CAPS"
                }
            },
            "complexity": {
                "max_function_length": 100,
                "max_sub_length": 100,
                "max_line_length": 120,
                "max_params": 7,
                "max_nesting": 5
            },
            "style": {
                "indentation": 4,
                "require_option_explicit": True,
                "require_comments": {
                    "functions": True,
                    "complex_logic": True
                }
            },
            "rules_enabled": {
                "naming": True,
                "complexity": True,
                "style": True,
                "best_practices": True,
                "performance": True
            }
        }
        
        if config_path and os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    user_config = json.load(f)
                    # Fusion des configurations (user_config écrase default_config)
                    self._merge_configs(default_config, user_config)
            except Exception as e:
                print(f"[⚠️] Erreur lors du chargement de la configuration: {e}")
                print(f"[ℹ️] Utilisation de la configuration par défaut.")
        
        return default_config
    
    def _merge_configs(self, default_config: Dict, user_config: Dict) -> None:
        """Fusionne récursivement deux dictionnaires de configuration"""
        for key, value in user_config.items():
            if key in default_config and isinstance(default_config[key], dict) and isinstance(value, dict):
                self._merge_configs(default_config[key], value)
            else:
                default_config[key] = value
    
    def validate_file(self, file_path: str) -> List[Dict]:
        """Valide un fichier VBA et retourne les problèmes détectés"""
        file_issues = []
        
        if not os.path.exists(file_path):
            return [self._create_issue(
                file_path, 0, 
                "Le fichier n'existe pas", 
                self.SEVERITY_ERROR, 
                "file_access"
            )]
        
        # Déterminer le type de fichier
        file_type = self._get_file_type(file_path)
        if not file_type:
            return [self._create_issue(
                file_path, 0, 
                "Type de fichier non supporté", 
                self.SEVERITY_ERROR, 
                "file_type"
            )]
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
                lines = content.splitlines()
                
                # Statistiques
                self.stats["files_processed"] += 1
                self.stats["lines_processed"] += len(lines)
                
                # Validations
                if self.config["rules_enabled"]["naming"]:
                    file_issues.extend(self._check_naming_conventions(file_path, file_type, lines))
                
                if self.config["rules_enabled"]["style"]:
                    file_issues.extend(self._check_style(file_path, lines))
                
                if self.config["rules_enabled"]["complexity"]:
                    file_issues.extend(self._check_complexity(file_path, lines))
                
                if self.config["rules_enabled"]["best_practices"]:
                    file_issues.extend(self._check_best_practices(file_path, lines))
                
                if self.config["rules_enabled"]["performance"]:
                    file_issues.extend(self._check_performance(file_path, lines))
                
                # Mise à jour des statistiques
                self.stats["issues_found"] += len(file_issues)
                for issue in file_issues:
                    if issue["severity"] == self.SEVERITY_ERROR:
                        self.stats["error_count"] += 1
                    elif issue["severity"] == self.SEVERITY_WARNING:
                        self.stats["warning_count"] += 1
                    else:
                        self.stats["info_count"] += 1
                
                return file_issues
        
        except UnicodeDecodeError:
            return [self._create_issue(
                file_path, 0, 
                "Le fichier n'est pas encodé en UTF-8", 
                self.SEVERITY_ERROR, 
                "encoding"
            )]
        except Exception as e:
            return [self._create_issue(
                file_path, 0, 
                f"Erreur lors de la lecture du fichier: {str(e)}", 
                self.SEVERITY_ERROR, 
                "file_access"
            )]
    
    def _get_file_type(self, file_path: str) -> Optional[str]:
        """Détermine le type de fichier VBA en fonction de son extension et de son contenu"""
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.bas':
            return "module"
        elif ext == '.cls':
            return "class"
        elif ext == '.frm':
            return "form"
        return None
    
    def _create_issue(self, file_path: str, line_number: int, message: str, 
                     severity: str, rule_id: str, code_snippet: str = None) -> Dict:
        """Crée une entrée standardisée pour un problème détecté"""
        return {
            "file": file_path,
            "line": line_number,
            "message": message,
            "severity": severity,
            "rule_id": rule_id,
            "code_snippet": code_snippet
        }
    
    def _check_naming_conventions(self, file_path: str, file_type: str, lines: List[str]) -> List[Dict]:
        """Vérifie les conventions de nommage"""
        issues = []
        
        # Vérification du préfixe de fichier
        filename = os.path.basename(file_path)
        expected_prefix = self.config["naming"]["module_prefix"].get(file_type, "")
        
        if expected_prefix and not filename.startswith(expected_prefix):
            issues.append(self._create_issue(
                file_path, 0, 
                f"Le nom du fichier ne commence pas par le préfixe attendu '{expected_prefix}'", 
                self.SEVERITY_WARNING, 
                "naming.file_prefix"
            ))
        
        # Vérification des noms de variables, fonctions, etc.
        function_pattern = re.compile(r'(Public|Private)?\s*(Function|Sub)\s+([A-Za-z0-9_]+)')
        variable_pattern = re.compile(r'(Dim|Public|Private)\s+([A-Za-z0-9_]+)\s+As\s+([A-Za-z0-9_]+)')
        const_pattern = re.compile(r'(Public|Private)?\s*Const\s+([A-Za-z0-9_]+)')
        
        for i, line in enumerate(lines):
            # Vérification des fonctions et procédures
            function_match = function_pattern.search(line)
            if function_match:
                function_name = function_match.group(3)
                if self.config["naming"]["case"]["functions"] == "PascalCase" and not self._is_pascal_case(function_name):
                    issues.append(self._create_issue(
                        file_path, i+1, 
                        f"Le nom de fonction '{function_name}' ne suit pas la convention PascalCase", 
                        self.SEVERITY_WARNING, 
                        "naming.function_case",
                        line.strip()
                    ))
            
            # Vérification des variables
            variable_match = variable_pattern.search(line)
            if variable_match:
                var_name = variable_match.group(2)
                var_type = variable_match.group(3)
                
                # Vérification du préfixe de type
                expected_var_prefix = self.config["naming"]["variable_prefixes"].get(var_type)
                if expected_var_prefix and not var_name.startswith(expected_var_prefix):
                    issues.append(self._create_issue(
                        file_path, i+1, 
                        f"Le nom de variable '{var_name}' de type {var_type} devrait commencer par '{expected_var_prefix}'", 
                        self.SEVERITY_INFO, 
                        "naming.variable_prefix",
                        line.strip()
                    ))
                
                # Vérification du style de casse
                if self.config["naming"]["case"]["variables"] == "camelCase" and not self._is_camel_case(var_name):
                    issues.append(self._create_issue(
                        file_path, i+1, 
                        f"Le nom de variable '{var_name}' ne suit pas la convention camelCase", 
                        self.SEVERITY_INFO, 
                        "naming.variable_case",
                        line.strip()
                    ))
            
            # Vérification des constantes
            const_match = const_pattern.search(line)
            if const_match:
                const_name = const_match.group(2)
                if self.config["naming"]["case"]["constants"] == "ALL_CAPS" and not self._is_all_caps(const_name):
                    issues.append(self._create_issue(
                        file_path, i+1, 
                        f"Le nom de constante '{const_name}' ne suit pas la convention ALL_CAPS", 
                        self.SEVERITY_WARNING, 
                        "naming.constant_case",
                        line.strip()
                    ))
        
        return issues
    
    def _check_style(self, file_path: str, lines: List[str]) -> List[Dict]:
        """Vérifie les règles de style"""
        issues = []
        
        # Vérification de Option Explicit
        has_option_explicit = False
        for line in lines[:10]:  # Vérifier dans les 10 premières lignes
            if re.match(r'^\s*Option\s+Explicit\s*$', line, re.IGNORECASE):
                has_option_explicit = True
                break
        
        if self.config["style"]["require_option_explicit"] and not has_option_explicit:
            issues.append(self._create_issue(
                file_path, 1, 
                "Il manque 'Option Explicit' en début de fichier", 
                self.SEVERITY_ERROR, 
                "style.option_explicit"
            ))
        
        # Vérification de la longueur des lignes
        max_line_length = self.config["complexity"]["max_line_length"]
        for i, line in enumerate(lines):
            if len(line.rstrip()) > max_line_length:
                issues.append(self._create_issue(
                    file_path, i+1, 
                    f"La ligne dépasse la longueur maximale recommandée ({len(line.rstrip())}/{max_line_length})", 
                    self.SEVERITY_INFO, 
                    "style.line_length",
                    line.strip()
                ))
        
        # Vérification de l'indentation
        indentation = self.config["style"]["indentation"]
        in_block = False
        block_level = 0
        for i, line in enumerate(lines):
            stripped = line.strip()
            
            # Ignorer les lignes vides ou les commentaires
            if not stripped or stripped.startswith("'"):
                continue
            
            # Détecter le début et la fin des blocs
            if re.search(r'\b(Function|Sub|If|For|Do|While|Select Case)\b', stripped, re.IGNORECASE):
                if not stripped.endswith("_") and not re.search(r'\bEnd\s+(Function|Sub|If|Select|Property)\b', stripped, re.IGNORECASE):
                    in_block = True
                    if not re.search(r'\bThen\b.*\b(If|Else|ElseIf)\b', stripped, re.IGNORECASE):  # Exclure les If...Then...Else inline
                        block_level += 1
            
            if re.search(r'\bEnd\s+(Function|Sub|If|Select|Property)\b|\bNext\b|\bLoop\b|\bWend\b', stripped, re.IGNORECASE):
                if in_block:
                    block_level = max(0, block_level - 1)  # Éviter les valeurs négatives
            
            # Vérifier l'indentation
            if in_block and block_level > 0:
                expected_indent = block_level * indentation
                actual_indent = len(line) - len(line.lstrip())
                
                # Permettre une certaine flexibilité pour les lignes spéciales
                if not re.search(r'\b(Else|ElseIf|Case|End)\b', stripped, re.IGNORECASE):
                    if actual_indent != expected_indent:
                        issues.append(self._create_issue(
                            file_path, i+1, 
                            f"Indentation incorrecte. Attendu: {expected_indent} espaces, trouvé: {actual_indent}", 
                            self.SEVERITY_INFO, 
                            "style.indentation",
                            line
                        ))
        
        return issues
    
    def _check_complexity(self, file_path: str, lines: List[str]) -> List[Dict]:
        """Vérifie la complexité du code"""
        issues = []
        
        # Variables pour suivre les fonctions/procédures
        in_function = False
        function_name = ""
        function_start_line = 0
        function_lines = 0
        nesting_level = 0
        max_nesting = 0
        
        # Comptage des paramètres
        function_pattern = re.compile(r'(Public|Private)?\s*(Function|Sub|Property)\s+([A-Za-z0-9_]+)[\s\n]*\((.*?)\)', re.DOTALL)
        
        # Analyse ligne par ligne
        for i, line in enumerate(lines):
            stripped = line.strip()
            
            # Détecter le début d'une fonction/procédure
            if re.search(r'\b(Function|Sub|Property\s+[GLS]et)\b', stripped, re.IGNORECASE) and not in_function:
                match = function_pattern.search('\n'.join(lines[i:i+10]))  # Regarder quelques lignes pour gérer les paramètres multi-lignes
                if match:
                    function_name = match.group(3)
                    function_start_line = i + 1
                    function_lines = 0
                    in_function = True
                    nesting_level = 0
                    max_nesting = 0
                    
                    # Vérification du nombre de paramètres
                    params_text = match.group(4)
                    if params_text.strip():
                        params = [p.strip() for p in params_text.split(',')]
                        if len(params) > self.config["complexity"]["max_params"]:
                            issues.append(self._create_issue(
                                file_path, i+1, 
                                f"La fonction '{function_name}' a trop de paramètres ({len(params)}/{self.config['complexity']['max_params']})", 
                                self.SEVERITY_WARNING, 
                                "complexity.too_many_params",
                                stripped
                            ))
            
            # Suivre les niveaux d'imbrication
            if in_function:
                function_lines += 1
                
                # Augmenter le niveau d'imbrication
                if re.search(r'\b(If|For|Do|While|Select Case)\b', stripped, re.IGNORECASE) and not re.search(r'\bEnd\s+If\b|\bNext\b|\bLoop\b|\bWend\b', stripped, re.IGNORECASE):
                    if not re.search(r'\bThen\b.*\b(If|Else|ElseIf)\b', stripped, re.IGNORECASE):  # Exclure les If...Then...Else inline
                        nesting_level += 1
                        max_nesting = max(max_nesting, nesting_level)
                
                # Diminuer le niveau d'imbrication
                if re.search(r'\bEnd\s+(If|Select)\b|\bNext\b|\bLoop\b|\bWend\b', stripped, re.IGNORECASE):
                    nesting_level = max(0, nesting_level - 1)  # Éviter les valeurs négatives
            
            # Détecter la fin d'une fonction/procédure
            if in_function and re.search(r'\bEnd\s+(Function|Sub|Property)\b', stripped, re.IGNORECASE):
                in_function = False
                
                # Vérification de la longueur de la fonction
                max_length = self.config["complexity"]["max_function_length"]
                if function_lines > max_length:
                    issues.append(self._create_issue(
                        file_path, function_start_line, 
                        f"La fonction '{function_name}' est trop longue ({function_lines} lignes, max recommandé: {max_length})", 
                        self.SEVERITY_WARNING, 
                        "complexity.function_length"
                    ))
                
                # Vérification du niveau d'imbrication
                max_nesting_allowed = self.config["complexity"]["max_nesting"]
                if max_nesting > max_nesting_allowed:
                    issues.append(self._create_issue(
                        file_path, function_start_line, 
                        f"La fonction '{function_name}' a un niveau d'imbrication trop élevé ({max_nesting}, max recommandé: {max_nesting_allowed})", 
                        self.SEVERITY_WARNING, 
                        "complexity.nesting_level"
                    ))
        
        return issues
    
    def _check_best_practices(self, file_path: str, lines: List[str]) -> List[Dict]:
        """Vérifie les bonnes pratiques"""
        issues = []
        
        # Vérification des commentaires de documentation
        in_function = False
        has_doc_comment = False
        function_name = ""
        function_line = 0
        
        for i, line in enumerate(lines):
            stripped = line.strip()
            
            # Ignorer les lignes vides
            if not stripped:
                continue
            
            # Détecter le début d'une fonction/procédure
            if re.search(r'\b(Function|Sub|Property\s+[GLS]et)\b', stripped, re.IGNORECASE) and not in_function:
                match = re.search(r'\b(Function|Sub|Property\s+[GLS]et)\s+([A-Za-z0-9_]+)', stripped)
                if match:
                    # Vérifier si la ligne précédente contient un commentaire
                    has_doc_comment = False
                    for j in range(i-1, max(0, i-5), -1):
                        if j >= 0 and re.search(r"^\s*'", lines[j]):
                            has_doc_comment = True
                            break
                    
                    function_name = match.group(2)
                    function_line = i + 1
                    in_function = True
                    
                    if self.config["style"]["require_comments"]["functions"] and not has_doc_comment:
                        issues.append(self._create_issue(
                            file_path, i+1, 
                            f"La fonction '{function_name}' n'a pas de commentaire de documentation", 
                            self.SEVERITY_INFO, 
                            "best_practices.missing_doc",
                            stripped
                        ))
            
            # Détecter la fin d'une fonction/procédure
            if in_function and re.search(r'\bEnd\s+(Function|Sub|Property)\b', stripped, re.IGNORECASE):
                in_function = False
        
        # Vérification des variables non utilisées (analyse simple)
        variables = {}
        variable_pattern = re.compile(r'(Dim|Public|Private)\s+([A-Za-z0-9_]+)')
        
        # Première passe: collecter les variables
        for i, line in enumerate(lines):
            if line.strip().startswith("'"):  # Ignorer les commentaires
                continue
                
            for match in variable_pattern.finditer(line):
                var_name = match.group(2)
                variables[var_name] = {"line": i+1, "used": False}
        
        # Deuxième passe: vérifier l'utilisation
        for i, line in enumerate(lines):
            if line.strip().startswith("'"):  # Ignorer les commentaires
                continue
                
            for var_name in variables:
                # Ne pas compter la ligne de déclaration
                if variables[var_name]["line"] == i+1:
                    continue
                    
                # Recherche simple du nom de variable dans le reste du code
                if re.search(r'\b' + re.escape(var_name) + r'\b', line):
                    variables[var_name]["used"] = True
        
        # Rapporter les variables non utilisées
        for var_name, info in variables.items():
            if not info["used"]:
                issues.append(self._create_issue(
                    file_path, info["line"], 
                    f"La variable '{var_name}' est déclarée mais semble non utilisée", 
                    self.SEVERITY_INFO, 
                    "best_practices.unused_variable"
                ))
        
        return issues
    
    def _check_performance(self, file_path: str, lines: List[str]) -> List[Dict]:
        """Vérifie les problèmes de performance"""
        issues = []
        
        # Variables pour suivre les boucles et le contexte
        in_loop = False
        loop_content = []
        loop_start_line = 0
        
        # Analyser le code ligne par ligne
        for i, line in enumerate(lines):
            stripped = line.strip()
            
            # Détecter le début d'une boucle
            if re.search(r'\b(For\s+Each|For|Do\s+While|Do\s+Until|While)\b', stripped, re.IGNORECASE) and not in_loop:
                in_loop = True
                loop_content = []
                loop_start_line = i + 1
            
            # Collecter le contenu de la boucle
            if in_loop:
                loop_content.append(stripped)
            
            # Détecter la fin d'une boucle
            if in_loop and re.search(r'\b(Next|Loop|Wend)\b', stripped, re.IGNORECASE):
                in_loop = False
                
                # Analyse des problèmes courants dans les boucles
                loop_content_str = '\n'.join(loop_content)
                
                # Vérification des accès fréquents à des objets Office
                if re.search(r'\.Worksheets\(', loop_content_str) or re.search(r'\.Cells\(', loop_content_str) or re.search(r'\.Range\(', loop_content_str):
                    issues.append(self._create_issue(
                        file_path, loop_start_line, 
                        "Accès répétés à des objets Excel dans une boucle. Considérez stocker les références dans des variables", 
                        self.SEVERITY_WARNING, 
                        "performance.excel_in_loop"
                    ))
        
        # Vérification de l'utilisation de Select/Activate
        for i, line in enumerate(lines):
            stripped = line.strip()
            
            # Détecter les appels .Select ou .Activate
            if re.search(r'\.(Select|Activate)\b', stripped, re.IGNORECASE) and not stripped.startswith("'"):
                issues.append(self._create_issue(
                    file_path, i+1, 
                    "Utilisation de .Select ou .Activate, qui peut ralentir le code. Préférez les références directes", 
                    self.SEVERITY_INFO, 
                    "performance.select_activate",
                    stripped
                ))
                
            # Vérifier l'utilisation de With pour les accès multiples
            if re.search(r'(\w+)\.(\w+).*\1\.\2', stripped, re.IGNORECASE) and not stripped.startswith("'"):
                issues.append(self._create_issue(
                    file_path, i+1, 
                    "Accès répétés au même objet. Considérez utiliser 'With...End With' pour améliorer les performances", 
                    self.SEVERITY_INFO, 
                    "performance.repeated_access",
                    stripped
                ))
        
        return issues
    
    def _is_pascal_case(self, name: str) -> bool:
        """Vérifie si un nom suit la convention PascalCase"""
        return re.match(r'^[A-Z][a-zA-Z0-9]*$', name) is not None
    
    def _is_camel_case(self, name: str) -> bool:
        """Vérifie si un nom suit la convention camelCase"""
        return re.match(r'^[a-z][a-zA-Z0-9]*$', name) is not None
    
    def _is_all_caps(self, name: str) -> bool:
        """Vérifie si un nom est en ALL_CAPS"""
        return re.match(r'^[A-Z][A-Z0-9_]*$', name) is not None
    
    def validate_directory(self, directory_path: str, pattern: str = "*.bas,*.cls,*.frm") -> Dict:
        """Valide tous les fichiers VBA dans un répertoire correspondant au pattern"""
        all_issues = []
        patterns = pattern.split(',')
        
        for root, _, files in os.walk(directory_path):
            for file in files:
                for pat in patterns:
                    if Path(file).match(pat.strip()):
                        file_path = os.path.join(root, file)
                        file_issues = self.validate_file(file_path)
                        all_issues.extend(file_issues)
                        break
        
        return {
            "issues": all_issues,
            "stats": self.stats
        }
    
    def generate_report(self, results: Dict, output_format: str = "text", output_file: Optional[str] = None) -> None:
        """Génère un rapport des problèmes détectés au format spécifié"""
        if output_format == "json":
            report = json.dumps(results, indent=2)
        else:  # Format texte par défaut
            report = self._generate_text_report(results)
        
        if output_file:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(report)
            print(f"[✓] Rapport écrit dans {output_file}")
        else:
            print(report)
    
    def _generate_text_report(self, results: Dict) -> str:
        """Génère un rapport au format texte"""
        issues = results["issues"]
        stats = results["stats"]
        
        report = []
        report.append("=== Rapport de validation VBA APEX Framework ===")
        report.append(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        report.append("")
        report.append("--- Statistiques ---")
        report.append(f"Fichiers analysés: {stats['files_processed']}")
        report.append(f"Lignes analysées: {stats['lines_processed']}")
        report.append(f"Problèmes détectés: {stats['issues_found']}")
        report.append(f"  - Erreurs: {stats['error_count']}")
        report.append(f"  - Avertissements: {stats['warning_count']}")
        report.append(f"  - Informations: {stats['info_count']}")
        report.append("")
        
        # Regrouper les problèmes par fichier
        issues_by_file = {}
        for issue in issues:
            file_path = issue["file"]
            if file_path not in issues_by_file:
                issues_by_file[file_path] = []
            issues_by_file[file_path].append(issue)
        
        # Trier les fichiers par nombre de problèmes (du plus au moins)
        sorted_files = sorted(issues_by_file.keys(), 
                             key=lambda x: len(issues_by_file[x]), 
                             reverse=True)
        
        # Génération du rapport détaillé
        if issues:
            report.append("--- Problèmes détectés ---")
            for file_path in sorted_files:
                file_issues = issues_by_file[file_path]
                report.append(f"\nFichier: {file_path}")
                report.append("-" * (len(file_path) + 9))
                
                # Trier les problèmes par numéro de ligne
                file_issues.sort(key=lambda x: x["line"])
                
                for issue in file_issues:
                    severity_marker = "❌" if issue["severity"] == self.SEVERITY_ERROR else "⚠️" if issue["severity"] == self.SEVERITY_WARNING else "ℹ️"
                    report.append(f"{severity_marker} Ligne {issue['line']}: {issue['message']} [{issue['rule_id']}]")
                    if issue.get("code_snippet"):
                        report.append(f"   {issue['code_snippet']}")
        else:
            report.append("Aucun problème détecté! 🎉")
        
        report.append("\n=== Fin du rapport ===")
        return "\n".join(report)


def main():
    parser = argparse.ArgumentParser(description="Validateur de code VBA pour APEX Framework")
    parser.add_argument("target", help="Fichier ou répertoire à valider")
    parser.add_argument("--config", "-c", help="Chemin vers le fichier de configuration JSON")
    parser.add_argument("--pattern", "-p", default="*.bas,*.cls,*.frm", 
                       help="Pattern des fichiers à analyser (séparés par des virgules, par défaut: '*.bas,*.cls,*.frm')")
    parser.add_argument("--output", "-o", help="Fichier de sortie pour le rapport")
    parser.add_argument("--format", "-f", choices=["text", "json"], default="text",
                       help="Format du rapport (text ou json, par défaut: text)")
    parser.add_argument("--verbose", "-v", action="store_true", help="Mode verbeux")
    
    args = parser.parse_args()
    
    validator = VBAValidator(args.config)
    
    start_time = datetime.now()
    
    if args.verbose:
        print(f"Validation de: {args.target}")
        print(f"Configuration: {args.config if args.config else 'par défaut'}")
    
    if os.path.isdir(args.target):
        results = validator.validate_directory(args.target, args.pattern)
    else:
        file_issues = validator.validate_file(args.target)
        results = {
            "issues": file_issues,
            "stats": validator.stats
        }
    
    end_time = datetime.now()
    duration = (end_time - start_time).total_seconds()
    
    if args.verbose:
        print(f"Validation terminée en {duration:.2f} secondes")
    
    validator.generate_report(results, args.format, args.output)


if __name__ == "__main__":
    main()