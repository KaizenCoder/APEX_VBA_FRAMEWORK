import sys
from pathlib import Path
from config_loader import ConfigLoader
from splash_screen import SplashScreen

class ApexOrchestrator:
    def __init__(self):
        self.config = ConfigLoader()
        self.init_paths()

    def init_paths(self):
        """Initialise les chemins nécessaires."""
        paths = [
            Path("logs"),
            Path("config"),
            Path("assets")
        ]
        for path in paths:
            path.mkdir(parents=True, exist_ok=True)

    def start(self):
        """Démarre l'orchestrateur avec splash screen."""
        def on_splash_complete():
            # Ici on lance l'interface principale
            self.run()

        splash = SplashScreen(duration=2)
        splash.show(on_splash_complete)

    def run(self):
        """Lance l'interface principale."""
        # Ici on lancera l'interface choisie (CLI, Web ou GUI)
        interface_type = self.config.get_value("ui", "default_interface", default="gui")
        
        if interface_type == "cli":
            from orchestrator_tui import menu
            menu()
        elif interface_type == "web":
            from orchestrator_web_flask_live import app
            app.run(
                host=self.config.get_value("ui", "web", "host"),
                port=self.config.get_value("ui", "web", "port")
            )
        else:  # gui par défaut
            from orchestrator_gui_tkinter_live import start_gui
            start_gui()

if __name__ == "__main__":
    orchestrator = ApexOrchestrator()
    orchestrator.start() 