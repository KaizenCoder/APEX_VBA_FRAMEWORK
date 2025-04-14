# Journal des Interactions Cursor - APEX Framework
> Format standardisé pour le suivi des interactions avec l'IA

## Format d'Entrée
```markdown
### YYYY-MM-DD HH:MM - [TYPE_INTERACTION]

**Prompt :** [Description courte de la demande]
**Contexte :** [Module/Fonction concerné]
**IA :** [Version IA utilisée]

#### 📋 Demande Détaillée
[Description complète de la demande]

#### 💡 Réponse IA
- Action : [Action entreprise]
- Fichiers modifiés : [Liste des fichiers]
- Temps : [Durée approximative]

#### 📊 Évaluation
- ✅ Points positifs : [Liste]
- ⚠️ Points d'attention : [Liste]
- 🔄 Modifications apportées : [Si applicable]

#### 📝 Décision
- Status : [Accepté/Modifié/Rejeté]
- Raison : [Justification]
- Suite : [Prochaines étapes]
```

## Entrées

### 2024-04-14 16:10 - SETUP_INITIAL

**Prompt :** Création de la structure documentaire APEX
**Contexte :** Configuration initiale
**IA :** Claude-3-Sonnet

#### 📋 Demande Détaillée
Mise en place de la structure documentaire pour le framework APEX avec playbook, règles Cursor et templates.

#### 💡 Réponse IA
- Action : Création des dossiers et fichiers de base
- Fichiers modifiés : 
  - docs/guidelines/playbook.md
  - docs/guidelines/cursor-rules.json
  - docs/_templates/playbook.template.md
  - tools/workflow/cursor/cursor-journal.md
- Temps : ~10 minutes

#### 📊 Évaluation
- ✅ Points positifs :
  - Structure claire et organisée
  - Documentation complète
  - Templates réutilisables
- ⚠️ Points d'attention :
  - Maintenir la cohérence des versions
  - Assurer les mises à jour régulières
  - Vérifier l'intégration CI/CD

#### 📝 Décision
- Status : Accepté
- Raison : Structure conforme aux besoins du projet
- Suite : Intégration continue et automatisation du suivi

<!-- 
Instructions pour l'utilisation du journal :
1. Créer une nouvelle entrée pour chaque interaction significative
2. Remplir tous les champs du template
3. Maintenir la chronologie inverse (plus récent en haut)
4. Ajouter des tags si nécessaire pour la recherche
5. Faire un commit après chaque nouvelle entrée
--> 