#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Wrapper Python pour le pont PowerShell.
Référence: chat_048 (2024-04-11 16:35)
Source: chat_047 (Correction encodage)
"""

import subprocess
import sys
import os
from pathlib import Path

class PowerShellBridge:
    def __init__(self):
        """Initialise le pont PowerShell."""
        self.bridge_script = Path(__file__).parent / "powershell_bridge.ps1"
        if not self.bridge_script.exists():
            raise FileNotFoundError(f"Script pont non trouvé : {self.bridge_script}")
            
    def execute(self, command=None):
        """
        Exécute une commande via le pont PowerShell.
        
        Args:
            command: Commande PowerShell à exécuter (optionnel)
        """
        # Configuration pour UTF-8 sans BOM
        os.environ['PYTHONIOENCODING'] = 'utf-8'
        
        ps_command = [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-NonInteractive",  # Évite les problèmes de console
            "-NoLogo",         # Supprime le logo PowerShell
            "-File", str(self.bridge_script)
        ]
        
        if command:
            ps_command.extend(["-Command", command])
            
        try:
            # Exécution avec UTF-8
            result = subprocess.run(
                ps_command,
                capture_output=True,
                text=True,
                encoding='utf-8',
                env=dict(os.environ, PYTHONIOENCODING='utf-8')
            )
            
            # Affichage direct (déjà en UTF-8)
            if result.stdout:
                print(result.stdout, end='')
            if result.stderr:
                print("ERREUR:", result.stderr, file=sys.stderr, end='')
                
            return result.returncode == 0
            
        except Exception as e:
            print(f"Erreur lors de l'exécution : {e}", file=sys.stderr)
            return False

if __name__ == "__main__":
    # Test du pont PowerShell
    bridge = PowerShellBridge()
    
    # Test 1: Initialisation simple
    print("Test 1: Initialisation du pont")
    bridge.execute()
    
    # Test 2: Commande simple
    print("\nTest 2: Commande Get-Process")
    bridge.execute("Get-Process | Select-Object -First 3") 