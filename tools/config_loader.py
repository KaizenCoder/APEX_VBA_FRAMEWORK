import json
from pathlib import Path
from typing import Dict, Any

class ConfigLoader:
    def __init__(self, config_path: str = "config/orchestrator_config.json"):
        self.config_path = Path(config_path)
        self.config: Dict[str, Any] = {}
        self.load_config()

    def load_config(self) -> None:
        """Charge la configuration depuis le fichier JSON."""
        try:
            self.config = json.loads(self.config_path.read_text(encoding='utf-8'))
        except Exception as e:
            print(f"Erreur lors du chargement de la configuration : {e}")
            self.config = self._get_default_config()

    def _get_default_config(self) -> Dict[str, Any]:
        """Retourne une configuration par défaut en cas d'erreur."""
        return {
            "general": {
                "workspace_path": ".",
                "environment": "development",
                "debug_mode": True
            },
            "logging": {
                "logs_path": "logs/",
                "activity_log": "vscode_activity.log",
                "monitor_keywords": ["ERROR", "EXCEPTION", "CRASH"]
            },
            "dashboard": {
                "update_interval_sec": 1800,
                "path": "docs/implementation/VSCODE_TRACKING_DASHBOARD.md"
            }
        }

    def get_config(self) -> Dict[str, Any]:
        """Retourne la configuration complète."""
        return self.config

    def get_value(self, *keys: str, default: Any = None) -> Any:
        """Récupère une valeur spécifique dans la configuration."""
        current = self.config
        try:
            for key in keys:
                current = current[key]
            return current
        except (KeyError, TypeError):
            return default

    def save_config(self) -> None:
        """Sauvegarde la configuration dans le fichier."""
        try:
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            self.config_path.write_text(
                json.dumps(self.config, indent=4),
                encoding='utf-8'
            )
        except Exception as e:
            print(f"Erreur lors de la sauvegarde de la configuration : {e}")

# Exemple d'utilisation
if __name__ == "__main__":
    config = ConfigLoader()
    
    # Exemple de lecture de configuration
    log_path = config.get_value("logging", "logs_path")
    update_interval = config.get_value("dashboard", "update_interval_sec")
    
    print(f"Chemin des logs : {log_path}")
    print(f"Intervalle de mise à jour : {update_interval} secondes") 