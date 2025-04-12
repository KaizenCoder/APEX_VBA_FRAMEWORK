#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Script de vérification des fichiers du workflow
Usage: python check_files.py
"""

import os
import sys
import json

def print_color(text, color):
    """Affiche du texte coloré"""
    colors = {
        "red": "\033[91m",
        "green": "\033[92m",
        "yellow": "\033[93m",
        "blue": "\033[94m",
        "end": "\033[0m"
    }
    print(f"{colors.get(color, '')}{text}{colors['end']}")

def check_file_content(filepath, expected_patterns, min_size=100):
    """Vérifie qu'un fichier existe et contient certains patterns"""
    if not os.path.exists(filepath):
        print_color(f"❌ Fichier manquant: {filepath}", "red")
        return False
    
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
            
        if len(content) < min_size:
            print_color(f"⚠️ Fichier trop petit: {filepath} ({len(content)} caractères)", "yellow")
            return False
            
        missing_patterns = []
        for pattern in expected_patterns:
            if pattern not in content:
                missing_patterns.append(pattern)
        
        if missing_patterns:
            print_color(f"⚠️ Patterns manquants dans {filepath}:", "yellow")
            for pattern in missing_patterns:
                print_color(f"   - {pattern}", "yellow")
            return False
        
        print_color(f"✅ Fichier valide: {filepath}", "green")
        return True
    except Exception as e:
        print_color(f"❌ Erreur lors de la lecture de {filepath}: {e}", "red")
        return False

def main():
    """Vérification principale des fichiers"""
    print_color("=== Vérification des fichiers du workflow APEX VBA Framework ===", "blue")
    
    files_to_check = {
        "docs/GIT_COMMIT_CONVENTION.md": [
            "Format standard", 
            "Types de modifications",
            "Portée (scope)",
            "Hooks Git",
            "feat", "fix", "docs", "refactor"
        ],
        "tools/workflow/scripts/commit_with_context.ps1": [
            "Get-Date",
            "Write-Host",
            "Out-File",
            "ConvertTo-Json",
            "git commit",
            "Read-Host"
        ],
        "tools/workflow/templates/session_log_template.md": [
            "Session de travail",
            "Objectif(s)",
            "Suivi des tâches",
            "Prompts IA",
            "Tests effectués",
            "Bilan de session"
        ],
        "tools/workflow/git-hooks/commit-msg": [
            "#!/bin/bash",
            "commit_msg_file",
            "pattern",
            "exit 0"
        ],
        "tools/workflow/scripts/install_hooks.ps1": [
            "Resolve-Path",
            "Write-Host",
            "Copy-Item",
            "git config"
        ],
        "tools/workflow/ci/CHANGELOG_IA.md": [
            "Historique des contributions IA",
            "Format d'entrée",
            "Claude",
            "GPT",
            "Validation"
        ]
    }
    
    results = {}
    for filepath, patterns in files_to_check.items():
        results[filepath] = check_file_content(filepath, patterns)
    
    # Affichage du résumé
    print_color("\n=== Résumé de la vérification ===", "blue")
    valid_count = sum(1 for r in results.values() if r)
    total_count = len(results)
    
    print_color(f"Fichiers valides: {valid_count}/{total_count}", "green" if valid_count == total_count else "yellow")
    
    if valid_count < total_count:
        print_color("\nActions recommandées:", "yellow")
        for filepath, is_valid in results.items():
            if not is_valid:
                print_color(f"- Vérifier le contenu de {filepath}", "yellow")
    
    if valid_count == total_count:
        print_color("\n🎉 Tous les fichiers sont valides et contiennent le contenu attendu.", "green")
        print_color("Le système de gestion des commits et journalisation est prêt à être utilisé.", "green")
    
    return 0 if valid_count == total_count else 1

if __name__ == "__main__":
    sys.exit(main()) 