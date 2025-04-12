# 📝 Workflow des Sessions de Développement

## 🎯 Objectif
Ce workflow fournit une structure organisée pour documenter et suivre les sessions de développement dans le framework APEX VBA. Il assure la cohérence, la traçabilité et la qualité de la documentation du développement.

## 🗂️ Structure
```
workflow/
├── sessions/          # Sessions de développement
│   ├── YYYY/         # Année
│   │   └── MM/       # Mois
│   │       ├── active/    # Sessions en cours
│   │       └── completed/ # Sessions terminées
├── scripts/          # Scripts de gestion
├── docs/            # Documentation détaillée
└── git-hooks/       # Hooks de validation Git
```

## 🛠️ Outils Principaux

### Scripts PowerShell
- **`Manage-SessionMigration.ps1`**
  - Migration des sessions
  - Organisation automatique
  - Validation de structure
  ```powershell
  .\Manage-SessionMigration.ps1 -Action [Migrate|Validate|Simulate]
  ```

- **`Test-SessionMarkdownFormat.ps1`**
  - Validation du format Markdown
  - Vérification de l'encodage
  - Correction automatique
  ```powershell
  .\Test-SessionMarkdownFormat.ps1 -Path "sessions" [-FixEncoding]
  ```

### Hooks Git
- **`pre-commit`** : Validation avant commit
- **`commit-msg`** : Validation du message de commit

## 📋 Format des Sessions

### Nom de Fichier
- Format : `YYYY_MM_DD_description.md`
- Exemple : `2024_04_11_pipeline_validation.md`

### Structure Requise
```markdown
# 🧭 Session de travail – YYYY-MM-DD

## 🎯 Objectif(s)
- Point 1
- Point 2

## 📌 Suivi des tâches
| Tâche | Module | Statut | Commentaire |
|-------|--------|--------|-------------|
| ... | ... | ✅/⏳ | ... |

## ✅ Clôture de session
- Résumé
- Prochaines étapes
```

## 🔄 Workflow Typique

1. **Création de Session**
   ```powershell
   # Créer un nouveau fichier dans sessions/
   Copy-Item templates/session_template.md "sessions/YYYY_MM_DD_description.md"
   ```

2. **Développement**
   - Remplir les sections au fur et à mesure
   - Committer régulièrement (validation automatique)

3. **Clôture**
   - Compléter la section "Clôture"
   - Valider le format
   ```powershell
   .\Test-SessionMarkdownFormat.ps1 -Path "sessions/YYYY_MM_DD_description.md"
   ```

4. **Migration**
   ```powershell
   .\Manage-SessionMigration.ps1 -Action Migrate
   ```

## 📚 Documentation
- Guide détaillé : `docs/GUIDE_MIGRATION.md`
- Résumé du workflow : `docs/SESSION_WORKFLOW_SUMMARY.md`
- Documentation des outils : `docs/WORKFLOW_TOOLS.md`

## 🔍 Validation et Maintenance
- Exécution régulière des scripts de validation
- Vérification de l'organisation des sessions
- Maintenance des hooks Git
- Mise à jour de la documentation

## 🤝 Contribution
1. Créer une session pour les modifications
2. Suivre le format standard
3. Tester les modifications
4. Soumettre avec documentation mise à jour

## 📊 Rapports
- `migration_report.md` : Détails des migrations
- `validation_report.md` : Résultats des validations
- `session_migration.log` : Logs techniques
- `migration_statistics.csv` : Statistiques d'exécution
``` 