# 🎯 Gestion des Règles Cursor

## Vue d'ensemble
Ce module gère l'automatisation des règles Cursor dans le framework APEX VBA. Il assure la prise en compte systématique des règles `.cursor-rules` indépendamment des sessions.

## 📋 Fonctionnalités
- Chargement automatique des règles
- Validation de l'environnement
- Traçabilité des sessions
- Installation/Désinstallation simplifiée
- Intégration VS Code

## 🚀 Installation

### Prérequis
- PowerShell 5.1 ou supérieur
- VS Code (recommandé)
- Git pour Windows

### Installation Automatique
```powershell
# Dans PowerShell, exécuter :
.\tools\workflow\cursor\Install-CursorRules.ps1
```

### Installation Manuelle
1. Copier les fichiers dans `tools/workflow/cursor/`
2. Exécuter `Register-CursorHooks.ps1`
3. Configurer VS Code avec `settings.json`

## 🔄 Désinstallation

### Désinstallation Automatique
```powershell
# Désinstallation complète
.\tools\workflow\cursor\Uninstall-CursorRules.ps1

# Options disponibles
-RemoveSessionFiles  # Supprime les fichiers de session
-Force              # Sans confirmation
```

### Désinstallation Manuelle
1. Supprimer les hooks du profil PowerShell
2. Nettoyer les variables d'environnement
3. Restaurer la configuration VS Code

## 📁 Structure des Fichiers
```
tools/workflow/cursor/
├── Install-CursorRules.ps1     # Script d'installation
├── Uninstall-CursorRules.ps1   # Script de désinstallation
├── Register-CursorHooks.ps1    # Enregistrement des hooks
├── Unregister-CursorHooks.ps1  # Suppression des hooks
├── Test-CursorRules.ps1        # Validation des règles
└── README.md                   # Documentation
```

## 🔍 Fonctionnement

### Hooks PowerShell
- Surveillance du changement de répertoire
- Détection automatique de `.cursor-rules`
- Chargement et validation des règles

### Variables d'Environnement
- `CURSOR_WORKSPACE` : Chemin du workspace
- `CURSOR_RULES_LOADED` : État des règles

### Fichiers de Session
- Format : `.cursor-session-{timestamp}.json`
- Contient : workspace, timestamp, version
- Utilisé pour : traçabilité, audit, debug

## ⚙️ Configuration

### PowerShell
```powershell
# Configuration manuelle
$env:CURSOR_RULES_LOADED = $true
```

### VS Code
```json
{
    "workspaceInit.tasks": [
        {
            "name": "Initialize Cursor Rules",
            "runOn": ["workspaceOpen"]
        }
    ]
}
```

## 🔒 Sécurité
- Validation avant modification
- Sauvegarde automatique
- Confirmation utilisateur
- Logs des changements

## 🐛 Dépannage

### Problèmes Courants
1. **Règles non chargées**
   ```powershell
   # Vérifier l'état
   $env:CURSOR_RULES_LOADED
   ```

2. **Hooks non actifs**
   ```powershell
   # Réinstaller les hooks
   .\Register-CursorHooks.ps1
   ```

3. **Erreurs VS Code**
   - Vérifier `.vscode/settings.json`
   - Recharger la fenêtre VS Code

## 📚 Documentation Associée
- [Guide d'Encodage](../../docs/requirements/powershell_encoding.md)
- [Architecture Core](../../docs/Components/CoreArchitecture.md)
- [Conventions Git](../../docs/GIT_COMMIT_CONVENTION.md)

## 🤝 Contribution
1. Fork le projet
2. Créer une branche (`git checkout -b feature/AmazingFeature`)
3. Commit les changements (`git commit -m 'Add AmazingFeature'`)
4. Push la branche (`git push origin feature/AmazingFeature`)
5. Ouvrir une Pull Request

## 📝 Licence
Distribué sous la licence MIT. Voir `LICENSE` pour plus d'informations.

## ✨ Auteurs
- APEX Framework Team 

# 🛠️ Scripts d'Intégration VS Code/Cursor

## Vue d'ensemble
Ce dossier contient l'ensemble des scripts PowerShell nécessaires pour gérer l'intégration entre VS Code et Cursor.

## 📋 Structure des Scripts

### Configuration et Installation
- `Configure-IDEIntegration.ps1` : Configuration initiale de l'intégration
- `Uninstall-CursorVSCode.ps1` : Désinstallation propre de l'intégration
- `Restore-CursorVSCode.ps1` : Restauration depuis une sauvegarde

### Tests et Validation
- `Test-IDEIntegration.ps1` : Tests de l'intégration
- `Test-UninstallCursorVSCode.ps1` : Tests de désinstallation
- `Analyze-IntegrationResults.ps1` : Analyse des résultats de test

## 🚀 Utilisation

### Configuration Initiale
```powershell
# Configuration standard
.\Configure-IDEIntegration.ps1

# Configuration avec options
.\Configure-IDEIntegration.ps1 -Force -EnableSharing
```

### Désinstallation
```powershell
# Désinstallation standard
.\Uninstall-CursorVSCode.ps1

# Désinstallation avec options
.\Uninstall-CursorVSCode.ps1 -Force -KeepSettings -NoBackup
```

### Tests
```powershell
# Tests d'intégration
.\Test-IDEIntegration.ps1

# Tests de désinstallation
.\Test-UninstallCursorVSCode.ps1 -Detailed

# Analyse des résultats
.\Analyze-IntegrationResults.ps1 -GenerateReport
```

## 📊 Rapports et Résultats

### Structure des Dossiers
```
tools/workflow/cursor/
├── tests/
│   ├── results/         # Résultats des tests
│   └── reports/         # Rapports d'analyse
├── backup/              # Sauvegardes de configuration
└── logs/               # Journaux d'exécution
```

### Types de Rapports
1. **Rapports de Test**
   - Format JSON
   - Résultats détaillés
   - Métriques de performance

2. **Rapports d'Analyse**
   - Format Markdown
   - Synthèse graphique
   - Recommandations

## 🔧 Maintenance

### Sauvegarde
```powershell
# Création manuelle d'une sauvegarde
.\Backup-VSCodeConfiguration.ps1

# Restauration depuis une sauvegarde
.\Restore-CursorVSCode.ps1 -BackupPath "path/to/backup"
```

### Validation
```powershell
# Validation de l'installation
.\Test-Installation.ps1

# Validation de la désinstallation
.\Test-Uninstallation.ps1
```

## 🐛 Dépannage

### Problèmes Courants
1. **Échec de Configuration**
   ```powershell
   .\Repair-Configuration.ps1
   ```

2. **Synchronisation Perdue**
   ```powershell
   .\Reset-Synchronization.ps1
   ```

3. **Conflits d'Extensions**
   ```powershell
   .\Resolve-ExtensionConflicts.ps1
   ```

## 🔒 Sécurité

### Bonnes Pratiques
1. Toujours exécuter avec les droits appropriés
2. Vérifier les sauvegardes régulièrement
3. Monitorer les logs d'exécution
4. Valider après chaque modification

### Permissions Requises
- Accès en lecture/écriture au dossier `.vscode`
- Accès aux variables d'environnement
- Droits d'installation d'extensions

## 📚 Documentation Associée
- [Guide d'Intégration](../../docs/vscode/IDE_INTEGRATION.md)
- [Guide de Désinstallation](../../docs/vscode/UNINSTALL.md)
- [Guide de Test](../../docs/vscode/TESTING.md)

## 🤝 Support
- Issues : [GitHub Issues](https://github.com/org/repo/issues)
- Wiki : [Documentation](https://github.com/org/repo/wiki)
- Email : support@organization.com

## ✨ Notes
- Les scripts sont maintenus par l'équipe APEX Framework
- Version minimale de PowerShell requise : 5.1
- Compatible Windows 10/11 