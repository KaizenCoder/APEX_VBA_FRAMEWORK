#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Test simple de PowerShell.
Référence: chat_043 (2024-04-11 16:05)
Source: chat_042 (Test simplifié)
"""

import subprocess
import sys

def test_powershell():
    """Test basique de PowerShell."""
    try:
        # Test 1: Version PowerShell
        print("\nTest 1: Version PowerShell")
        subprocess.run(["powershell", "-Command", "$PSVersionTable.PSVersion"], check=True)
        
        # Test 2: Emplacement actuel
        print("\nTest 2: Emplacement actuel")
        subprocess.run(["powershell", "-Command", "Get-Location"], check=True)
        
        # Test 3: Création fichier
        print("\nTest 3: Création fichier")
        subprocess.run([
            "powershell",
            "-Command",
            "New-Item -Path 'test.txt' -ItemType File -Force; Get-Content 'test.txt'"
        ], check=True)
        
    except subprocess.CalledProcessError as e:
        print(f"Erreur PowerShell: {e}", file=sys.stderr)
    except Exception as e:
        print(f"Erreur: {e}", file=sys.stderr)

if __name__ == "__main__":
    test_powershell() 