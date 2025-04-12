#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Test complet PowerShell avec nouveaux paramètres terminal.
Référence: chat_045 (2024-04-11 16:20)
Source: chat_044 (Configuration terminal)
"""

import subprocess
import sys
import os
from datetime import datetime

def run_ps_command(command, description):
    """Exécute une commande PowerShell et affiche le résultat."""
    print(f"\n{description} :")
    print("-" * 50)
    result = subprocess.run(
        ["powershell", "-Command", command],
        capture_output=True,
        text=True
    )
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print("ERREUR:", result.stderr, file=sys.stderr)
    return result.returncode == 0

def main():
    """Tests principaux."""
    print("Test PowerShell avec paramètres VSCode")
    print("=" * 50)
    print(f"Démarré à : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    tests = [
        ("$PSVersionTable.PSVersion", "1. Version PowerShell"),
        ("Get-Location", "2. Répertoire courant"),
        ("$env:PATH", "3. Variable PATH"),
        ("Get-ExecutionPolicy", "4. Politique d'exécution"),
        ("""
        $testFile = 'test_cursor.txt'
        $content = 'Test depuis Cursor avec config VSCode'
        Set-Content -Path $testFile -Value $content -Force
        if (Test-Path $testFile) {
            Get-Content $testFile
            Write-Host "✅ Fichier créé avec succès"
        } else {
            Write-Host "❌ Échec de création du fichier"
        }
        """, "5. Test création fichier"),
        ("Get-Process | Select-Object -First 3", "6. Test commande complexe")
    ]
    
    success_count = 0
    for command, description in tests:
        if run_ps_command(command, description):
            success_count += 1
    
    print("\nRésumé des tests")
    print("=" * 50)
    print(f"Tests réussis : {success_count}/{len(tests)}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Erreur globale: {e}", file=sys.stderr) 