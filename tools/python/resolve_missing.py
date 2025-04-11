#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Script pour detecter et resoudre les modules manquants dans le framework APEX.
Genere des fichiers stub pour les composants manquants et produit un rapport.
"""

import os
import sys
import argparse
import json
import logging
from pathlib import Path
from datetime import datetime

# Configuration du logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

# Liste des modules essentiels au framework
ESSENTIAL_MODULES = [
    # Core - Interfaces
    ("apex-core/interfaces/ILoggerBase.cls", "Interface de base du systeme de log.", "Haute"),
    ("apex-core/interfaces/IPlugin.cls", "Interface pour les plugins du framework.", "Haute"),
    ("apex-core/interfaces/IQueryBuilder.cls", "Interface pour la construction de requetes SQL.", "Moyenne"),
    
    # Core - Testing
    ("apex-core/testing/clsTestSuite.cls", "Classe principale pour l'execution des suites de tests.", "Moyenne"),
    ("apex-core/testing/modTestAssertions.bas", "Assertions pour les tests unitaires.", "Moyenne"),
    ("apex-core/testing/modTestRegistry.bas", "Registre de declaration des cas de test.", "Moyenne"),
    ("apex-core/testing/modTestRunner.bas", "Execution centralisee des tests.", "Moyenne"),
    
    # Core - Utility
    ("apex-core/modReleaseValidator.bas", "Validation automatisee d'une release.", "Basse"),
    ("apex-core/modVersionInfo.bas", "Informations sur la version du framework.", "Haute"),
    ("apex-core/clsPluginManager.cls", "Gestion des plugins.", "Moyenne"),
    
    # Metier - Database
    ("apex-metier/database/interfaces/IDbAccessorBase.cls", "Interface d'acces a la base de donnees.", "Haute"),
    ("apex-metier/database/interfaces/IDbDriver.cls", "Interface des drivers de base de donnees.", "Haute"),
    ("apex-metier/database/interfaces/IQueryBuilder.cls", "Construction dynamique de requetes SQL.", "Moyenne"),
    ("apex-metier/database/clsAccessDriver.cls", "Driver pour bases Access.", "Moyenne"),
    
    # Metier - ORM
    ("apex-metier/orm/interfaces/IRelationMetadata.cls", "Metadonnees de relations ORM.", "Moyenne"),
    ("apex-metier/orm/interfaces/IRelationalObject.cls", "Objet relationnel.", "Moyenne"),
    ("apex-metier/orm/clsOrmBase.cls", "Classe de base pour ORM.", "Moyenne"),
    ("apex-metier/orm/clsRelationMetadata.cls", "Implementation des metadonnees relationnelles.", "Moyenne"),
    
    # UI
    ("apex-ui/handlers/modRibbonCallbacks.bas", "Callbacks pour le ruban Excel.", "Haute")
]

# Liste des modules deja planifies dans MODULES_PLANIFIES.md
PLANNED_MODULES = [
    # Outlook
    ("apex-metier/outlook/clsAttachmentProcessor.cls", "Traitement des pieces jointes depuis Outlook", "Haute"),
    ("apex-metier/outlook/clsMailBuilder.cls", "Construction d'emails via API", "Haute"),
    ("apex-metier/outlook/clsMailFetcher.cls", "Recuperation d'emails", "Haute"),
    ("apex-metier/outlook/clsOutlookClient.cls", "Client Outlook principal", "Haute"),
    
    # XML
    ("apex-metier/xml/clsXmlConfigManager.cls", "Gestion de configuration via XML", "Moyenne"),
    ("apex-metier/xml/clsXmlDiffer.cls", "Comparaison de structures XML", "Moyenne"),
    ("apex-metier/xml/clsXmlFlattener.cls", "Conversion de XML en structure plate", "Moyenne"),
    ("apex-metier/xml/clsXmlValidator.cls", "Validation de XML selon schema", "Moyenne"),
    
    # Environment Variables
    ("apex-core/modEnvVars.bas", "Acces aux variables d'environnement systeme", "Basse")
]

def generate_class_stub(module_name, description, path):
    """Genere un fichier stub pour une classe"""
    content = f"""VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "{module_name}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Module    : {module_name}
' Etat      : A DEVELOPPER (Stub)
' Description : {description}
' Date de creation : {datetime.now().strftime('%d/%m/%Y')}
' ==========================================================================

' -- Interface stub --
"""
    path.write_text(content, encoding="utf-8")
    return content

def generate_module_stub(module_name, description, path):
    """Genere un fichier stub pour un module standard"""
    content = f"""Attribute VB_Name = "{module_name}"
Option Explicit

' ==========================================================================
' Module    : {module_name}
' Etat      : A DEVELOPPER (Stub)
' Description : {description}
' Date de creation : {datetime.now().strftime('%d/%m/%Y')}
' ==========================================================================

' -- Module stub --
"""
    path.write_text(content, encoding="utf-8")
    return content

def scan_modules(root_dir, check_only=False):
    """Analyse les modules existants et manquants"""
    root = Path(root_dir)
    missing_modules = []
    existing_modules = []
    
    # Verifier les modules essentiels
    for rel_path, desc, priority in ESSENTIAL_MODULES + PLANNED_MODULES:
        path = root / rel_path
        module_name = path.stem
        if not path.exists():
            missing_modules.append((rel_path, module_name, desc, priority))
        else:
            existing_modules.append(rel_path)
    
    return existing_modules, missing_modules

def generate_stubs(root_dir, missing_modules):
    """Genere les fichiers stub pour les modules manquants"""
    root = Path(root_dir)
    generated_stubs = []
    
    for rel_path, module_name, desc, priority in missing_modules:
        target_path = root / rel_path
        target_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Verifier si c'est une classe ou un module standard
        if rel_path.endswith('.cls'):
            generate_class_stub(module_name, desc, target_path)
        else:
            generate_module_stub(module_name, desc, target_path)
        
        generated_stubs.append((rel_path, module_name, desc, priority))
        logging.info(f"Stub genere: {rel_path}")
    
    return generated_stubs

def generate_report(existing_modules, missing_modules, generated_stubs, output_file):
    """Genere un rapport Markdown des modules manquants et des stubs generes"""
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("# 📄 Rapport de modules manquants - Framework APEX\n\n")
        f.write(f"*Genere le {datetime.now().strftime('%d/%m/%Y a %H:%M:%S')}*\n\n")
        
        # Resume
        f.write("## 📊 Resume\n\n")
        f.write(f"- **Modules existants**: {len(existing_modules)}\n")
        f.write(f"- **Modules manquants**: {len(missing_modules)}\n")
        f.write(f"- **Stubs generes**: {len(generated_stubs)}\n\n")
        
        # Modules par priorite
        f.write("## 🎯 Modules manquants par priorite\n\n")
        
        for priority in ["Haute", "Moyenne", "Basse"]:
            f.write(f"### Priorite {priority}\n\n")
            priority_modules = [m for m in missing_modules if m[3] == priority]
            
            if not priority_modules:
                f.write("*Aucun module manquant dans cette categorie*\n\n")
                continue
                
            f.write("| Module | Description | Chemin |\n")
            f.write("|--------|-------------|--------|\n")
            
            for rel_path, module_name, desc, _ in priority_modules:
                f.write(f"| `{module_name}` | {desc} | `{rel_path}` |\n")
            
            f.write("\n")
        
        # Stubs generes
        if generated_stubs:
            f.write("## ✅ Stubs generes\n\n")
            f.write("| Module | Description | Chemin |\n")
            f.write("|--------|-------------|--------|\n")
            
            for rel_path, module_name, desc, _ in generated_stubs:
                f.write(f"| `{module_name}` | {desc} | `{rel_path}` |\n")
            
            f.write("\n")
        
        # Instructions pour la suite
        f.write("## 🔄 Prochaines etapes\n\n")
        f.write("1. Implementer les modules manquants en priorite Haute\n")
        f.write("2. Mettre a jour le fichier `MODULES_PLANIFIES.md`\n")
        f.write("3. Executer `generate_apex_addin.py` pour verifier que tous les modules sont importes correctement\n\n")
        
        f.write("---\n\n")
        f.write("*Ce rapport est genere automatiquement par `resolve_missing.py`*\n")
        
    logging.info(f"Rapport genere: {output_file}")
    return output_file

def main():
    parser = argparse.ArgumentParser(description="Detecte et resout les modules manquants du framework APEX")
    parser.add_argument("--dir", default=".", help="Repertoire racine du framework (defaut: repertoire courant)")
    parser.add_argument("--check-only", action="store_true", help="Verifier sans generer de stubs")
    parser.add_argument("--report", default="missing_modules_report.md", help="Chemin du rapport (defaut: missing_modules_report.md)")
    args = parser.parse_args()
    
    logging.info(f"Analyse des modules dans {args.dir}...")
    existing_modules, missing_modules = scan_modules(args.dir, args.check_only)
    
    if args.check_only:
        logging.info(f"Mode verification uniquement: {len(missing_modules)} modules manquants detectes")
        generated_stubs = []
    else:
        logging.info(f"Generation des stubs pour {len(missing_modules)} modules manquants...")
        generated_stubs = generate_stubs(args.dir, missing_modules)
    
    report_path = generate_report(existing_modules, missing_modules, generated_stubs, args.report)
    
    # Affichage du resume
    print("\n" + "="*50)
    print(f"📊 RESUME - {len(existing_modules)} modules existants, {len(missing_modules)} modules manquants")
    print("="*50)
    
    if missing_modules:
        print("\n🔴 Modules manquants de priorite HAUTE:")
        high_priority = [m for m in missing_modules if m[3] == "Haute"]
        for _, module_name, desc, _ in high_priority:
            print(f"  • {module_name}: {desc}")
    
    if generated_stubs:
        print(f"\n✅ {len(generated_stubs)} stubs generes avec succes")
    
    print(f"\n📄 Rapport detaille: {report_path}")
    print("\nPour voir tous les details, consultez le rapport.")
    
    if missing_modules and not args.check_only:
        print("\n💡 CONSEIL: Executez maintenant generate_apex_addin.py pour generer l'add-in avec les stubs")
    
    return 0 if (args.check_only or len(generated_stubs) == len(missing_modules)) else 1

if __name__ == "__main__":
    sys.exit(main()) 