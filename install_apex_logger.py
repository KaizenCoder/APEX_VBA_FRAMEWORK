#!/usr/bin/env python3
import os
import sys
from pathlib import Path
from datetime import datetime
import shutil
import json
import traceback

# --- CONFIGURATION MODIFIABLE ---
CONFIG = {
    "root_dir": "apex_cursor_logger",
    "tools_subdir": "tools/cursor",
    "logs_subdir": "logs",
    "prompts_subdir": "prompts",
    "default_agent": "GPT-4",
    "default_note": "+",
    "max_prompt_length": 160,
    "max_response_length": 300,
    "encoding": "utf-8",
    "backup_existing": True
}

class ApexLoggerInstaller:
    def __init__(self, config=None):
        self.config = config or CONFIG
        self.root = Path(self.config["root_dir"])
        self.tools_dir = self.root / self.config["tools_subdir"]
        self.logs_dir = self.root / self.config["logs_subdir"]
        self.prompts_dir = self.root / self.config["prompts_subdir"]
        
        self.journal_file = self.logs_dir / "cursor-journal.md"
        self.autolog_script = self.tools_dir / "cursor-autolog.py"
        self.bash_wrapper = self.tools_dir / "log_cursor.sh"
        self.powershell_wrapper = self.tools_dir / "log_cursor.ps1"
        self.prompts_file = self.prompts_dir / "prompter_template.md"

    def validate_paths(self):
        """Valide et crée les chemins nécessaires"""
        try:
            if self.root.exists() and self.config["backup_existing"]:
                backup_dir = self.root.parent / f"{self.root.name}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                if self.root.exists():
                    shutil.copytree(self.root, backup_dir, dirs_exist_ok=True)
                    print(f"[✓] Backup créé : {backup_dir}")

            for directory in [self.tools_dir, self.logs_dir, self.prompts_dir]:
                directory.mkdir(parents=True, exist_ok=True)
                print(f"[✓] Dossier créé : {directory}")
        except Exception as e:
            print(f"[✗] Erreur lors de la validation des chemins : {str(e)}")
            traceback.print_exc()
            sys.exit(1)

    def generate_autolog_script(self):
        """Génère le script Python principal"""
        try:
            content = f'''#!/usr/bin/env python3
import os
from datetime import datetime
import sys
import json

CONFIG = {json.dumps(self.config, indent=2)}
LOG_PATH = os.path.join(os.path.dirname(__file__), f"../../{self.config['logs_subdir']}/cursor-journal.md")

def log_cursor_interaction(prompt, response, agent=CONFIG["default_agent"], note=CONFIG["default_note"]):
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    entry = f"""---

### {{now}}

**Prompt :** {{prompt.strip()[:CONFIG["max_prompt_length"]]}}{{("..." if len(prompt) > CONFIG["max_prompt_length"] else "")}}
**Agent IA :** {{agent}}
**Avis :** {{note}}
**Réponse (extrait) :** {{response.strip()[:CONFIG["max_response_length"]]}}{{("..." if len(response) > CONFIG["max_response_length"] else "")}}
"""
    try:
        with open(LOG_PATH, "a", encoding=CONFIG["encoding"]) as f:
            f.write(entry)
        print(f"[✓] Interaction ajoutée au journal : {{LOG_PATH}}")
    except Exception as e:
        print(f"[✗] Erreur lors de l'écriture : {{str(e)}}")
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: cursor-autolog.py '<prompt>' '<response>' [agent] [note]")
        sys.exit(1)
    prompt = sys.argv[1]
    response = sys.argv[2]
    agent = sys.argv[3] if len(sys.argv) > 3 else CONFIG["default_agent"]
    note = sys.argv[4] if len(sys.argv) > 4 else CONFIG["default_note"]
    log_cursor_interaction(prompt, response, agent, note)
'''
            self.autolog_script.write_text(content, encoding=self.config["encoding"])
            self.autolog_script.chmod(0o755)
            print(f"[✓] Script Python créé : {self.autolog_script}")
        except Exception as e:
            print(f"[✗] Erreur lors de la génération du script Python : {str(e)}")
            traceback.print_exc()
            sys.exit(1)

    def generate_bash_wrapper(self):
        """Génère le wrapper Bash"""
        try:
            content = '''#!/bin/bash
if [ "$#" -lt 2 ]; then
  echo "Usage: ./log_cursor.sh \\"Prompt ici\\" \\"Réponse ici\\" [Agent] [Note]"
  exit 1
fi

PROMPT="$1"
RESPONSE="$2"
AGENT=${3:-GPT-4}
NOTE=${4:-+}

SCRIPT_DIR="$(dirname "$0")"
python3 "$SCRIPT_DIR/cursor-autolog.py" "$PROMPT" "$RESPONSE" "$AGENT" "$NOTE"
'''
            self.bash_wrapper.write_text(content, encoding=self.config["encoding"])
            self.bash_wrapper.chmod(0o755)
            print(f"[✓] Wrapper Bash créé : {self.bash_wrapper}")
        except Exception as e:
            print(f"[✗] Erreur lors de la génération du wrapper Bash : {str(e)}")
            traceback.print_exc()
            sys.exit(1)

    def generate_powershell_wrapper(self):
        """Génère le wrapper PowerShell"""
        try:
            content = '''param (
    [Parameter(Mandatory=$true)][string]$Prompt,
    [Parameter(Mandatory=$true)][string]$Response,
    [string]$Agent = "GPT-4",
    [string]$Note = "+"
)

$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
python "$ScriptPath\\cursor-autolog.py" "$Prompt" "$Response" "$Agent" "$Note"
'''
            self.powershell_wrapper.write_text(content, encoding=self.config["encoding"])
            print(f"[✓] Wrapper PowerShell créé : {self.powershell_wrapper}")
        except Exception as e:
            print(f"[✗] Erreur lors de la génération du wrapper PowerShell : {str(e)}")
            traceback.print_exc()
            sys.exit(1)

    def generate_prompt_template(self):
        """Génère le template de prompt"""
        try:
            content = '''## Objectif
[Description claire de l'objectif]

## Contraintes
- [Contrainte 1]
- [Contrainte 2]
- [Contrainte 3]

## Livrables attendus
1. [Livrable 1]
2. [Livrable 2]
3. [Livrable 3]

## Critères de validation
- [Critère 1]
- [Critère 2]
- [Critère 3]
'''
            self.prompts_file.write_text(content, encoding=self.config["encoding"])
            print(f"[✓] Template de prompt créé : {self.prompts_file}")
        except Exception as e:
            print(f"[✗] Erreur lors de la génération du template : {str(e)}")
            traceback.print_exc()
            sys.exit(1)

    def initialize_journal(self):
        """Initialise le journal des interactions"""
        try:
            content = f"### {datetime.now().strftime('%Y-%m-%d')} – Projet APEX\n\n*Journal initialisé.*\n"
            self.journal_file.write_text(content, encoding=self.config["encoding"])
            print(f"[✓] Journal initialisé : {self.journal_file}")
        except Exception as e:
            print(f"[✗] Erreur lors de l'initialisation du journal : {str(e)}")
            traceback.print_exc()
            sys.exit(1)

    def generate_config_file(self):
        """Sauvegarde la configuration"""
        try:
            config_file = self.root / "logger_config.json"
            with open(config_file, "w", encoding=self.config["encoding"]) as f:
                json.dump(self.config, f, indent=2)
            print(f"[✓] Configuration sauvegardée : {config_file}")
        except Exception as e:
            print(f"[✗] Erreur lors de la sauvegarde de la configuration : {str(e)}")
            traceback.print_exc()
            sys.exit(1)

    def install(self):
        """Procède à l'installation complète"""
        print("\n=== Installation du Logger APEX ===\n")
        try:
            self.validate_paths()
            self.generate_autolog_script()
            self.generate_bash_wrapper()
            self.generate_powershell_wrapper()
            self.generate_prompt_template()
            self.initialize_journal()
            self.generate_config_file()
            
            print("\n=== Installation terminée avec succès ===")
            print(f"\nDossier d'installation : {self.root.resolve()}")
            print("\nPour utiliser le logger :")
            print("1. Bash : ./log_cursor.sh 'prompt' 'réponse' [agent] [note]")
            print("2. PowerShell : .\\log_cursor.ps1 -Prompt 'prompt' -Response 'réponse' [-Agent agent] [-Note note]")
            print("\nPour personnaliser la configuration, modifiez logger_config.json")
            
        except Exception as e:
            print(f"\n[✗] Erreur lors de l'installation : {str(e)}")
            traceback.print_exc()
            sys.exit(1)

if __name__ == "__main__":
    installer = ApexLoggerInstaller()
    installer.install() 