# Apex Orchestrator – Documentation interne

## Objectif
Fournir un orchestrateur autonome pour piloter :
- les tests de performance Excel (framework Apex VBA)
- la mise à jour automatique du dashboard
- la journalisation continue et la surveillance des erreurs

## Structure du projet

- `src/` : code source Python (agents, interface CLI/Tkinter/Flask)
- `tools/` : scripts PowerShell utilisés par l’orchestrateur
- `config/` : paramètres de configuration JSON
- `assets/` : icône et splash screen
- `docs/` : documentation markdown ou historique

## Fonctionnalités

- Interface CLI (TUI), GUI (Tkinter), Web (Flask)
- Splash screen au démarrage (version `.exe`)
- Exécutable unique `.exe` + installateur `.msi` (Inno Setup)
- Logs auto dans `logs/`, alertes dans `anomaly_alerts.log`
- Configuration modulaire via `config/config.json`

## Prochaines étapes

- Génération `.exe` via PyInstaller
- Génération `.msi` via Inno Setup
- Version USB sans splash

> Mainteneur : Vous
