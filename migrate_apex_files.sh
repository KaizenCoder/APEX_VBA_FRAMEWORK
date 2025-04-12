#!/bin/bash
# ==========================================================================
# Script : migrate_apex_files.sh
# Version : 1.0
# Purpose : Migration des fichiers vers la nouvelle architecture Apex Framework v1.1
# Date : 10/04/2025
# ==========================================================================

echo "===== MIGRATION DES FICHIERS APEX FRAMEWORK v1.1 ====="
echo ""

TODAY=$(date +%Y-%m-%d)
MIGRATION_COMMENT="' Migrated to"
MIGRATION_COMMENT2="' Part of the APEX Framework v1.1 architecture refactoring"

# Fonction pour ajouter un commentaire de migration en début de fichier
add_migration_comment() {
    local source_file=$1
    local target_file=$2
    local target_dir=$(dirname "$target_file")
    
    # Créer le répertoire cible s'il n'existe pas
    mkdir -p "$target_dir"
    
    # Vérifier si le fichier source existe
    if [ ! -f "$source_file" ]; then
        echo "[ERREUR] Fichier source introuvable: $source_file"
        return 1
    fi
    
    # Copier le fichier
    cp "$source_file" "$target_file"
    
    # Ajouter le commentaire de migration en début de fichier
    comment_line="$MIGRATION_COMMENT $target_dir - $TODAY"
    sed -i "1i$comment_line" "$target_file"
    sed -i "2i$MIGRATION_COMMENT2" "$target_file"
    
    echo "[OK] Migré: $source_file -> $target_file"
}

echo "[INFO] Migration des fichiers Core..."
# Core - Logger et configuration
add_migration_comment "src/core/clsLogger.cls" "apex-core/clsLogger.cls"
add_migration_comment "src/core/modConfigManager.bas" "apex-core/modConfigManager.bas"
add_migration_comment "src/core/modSecurityDPAPI.bas" "apex-core/modSecurityDPAPI.bas"
add_migration_comment "src/core/modVersionInfo.bas" "apex-core/modVersionInfo.bas"
add_migration_comment "src/core/modReleaseValidator.bas" "apex-core/modReleaseValidator.bas"
add_migration_comment "src/core/modEnvVars.bas" "apex-core/modEnvVars.bas"

# Core - Utilitaires
add_migration_comment "src/utils/modTextUtils.bas" "apex-core/utils/modTextUtils.bas"
add_migration_comment "src/utils/modDateUtils.bas" "apex-core/utils/modDateUtils.bas"
add_migration_comment "src/utils/modFileUtils.bas" "apex-core/utils/modFileUtils.bas"

# Core - Testing
add_migration_comment "src/architecture/modTestAssertions.bas" "apex-core/testing/modTestAssertions.bas"
add_migration_comment "src/architecture/modTestRegistry.bas" "apex-core/testing/modTestRegistry.bas"
add_migration_comment "src/architecture/modTestRunner.bas" "apex-core/testing/modTestRunner.bas"
add_migration_comment "src/architecture/clsTestSuite.cls" "apex-core/testing/clsTestSuite.cls"

# Core - Interfaces
add_migration_comment "src/Interfaces/ILoggerBase.cls" "apex-core/interfaces/ILoggerBase.cls"
add_migration_comment "src/Interfaces/IPlugin.cls" "apex-core/interfaces/IPlugin.cls"
add_migration_comment "src/Classes/clsPluginManager.cls" "apex-core/clsPluginManager.cls"

echo "[INFO] Migration des fichiers Métier..."

# Métier - Recette
add_migration_comment "src/recette/clsReportGenerator.cls" "apex-metier/recette/clsReportGenerator.cls"
add_migration_comment "src/recette/clsTableComparer.cls" "apex-metier/recette/clsTableComparer.cls"
add_migration_comment "src/recette/modRecipeComparer.bas" "apex-metier/recette/modRecipeComparer.bas"

# Métier - XML
add_migration_comment "src/xml/clsXmlNode.cls" "apex-metier/xml/clsXmlNode.cls"
add_migration_comment "src/xml/clsXmlParser.cls" "apex-metier/xml/clsXmlParser.cls"
add_migration_comment "src/xml/clsXmlConfigManager.cls" "apex-metier/xml/clsXmlConfigManager.cls"
add_migration_comment "src/xml/clsXmlFlattener.cls" "apex-metier/xml/clsXmlFlattener.cls"
add_migration_comment "src/xml/clsXmlDiffer.cls" "apex-metier/xml/clsXmlDiffer.cls"
add_migration_comment "src/xml/clsXmlValidator.cls" "apex-metier/xml/clsXmlValidator.cls"

# Métier - Outlook
add_migration_comment "src/outlook/clsAttachmentProcessor.cls" "apex-metier/outlook/clsAttachmentProcessor.cls"
add_migration_comment "src/outlook/clsMailFetcher.cls" "apex-metier/outlook/clsMailFetcher.cls"
add_migration_comment "src/outlook/clsMailBuilder.cls" "apex-metier/outlook/clsMailBuilder.cls"
add_migration_comment "src/outlook/clsOutlookClient.cls" "apex-metier/outlook/clsOutlookClient.cls"

# Métier - Database
add_migration_comment "src/Classes/clsDbAccessor.cls" "apex-metier/database/clsDbAccessor.cls"
add_migration_comment "src/Classes/clsQueryBuilder.cls" "apex-metier/database/clsQueryBuilder.cls"
add_migration_comment "src/Classes/clsAccessDriver.cls" "apex-metier/database/clsAccessDriver.cls"
add_migration_comment "src/Interfaces/IDbDriver.cls" "apex-metier/database/interfaces/IDbDriver.cls"
add_migration_comment "src/Interfaces/IDbAccessorBase.cls" "apex-metier/database/interfaces/IDbAccessorBase.cls"
add_migration_comment "src/Interfaces/IQueryBuilder.cls" "apex-metier/database/interfaces/IQueryBuilder.cls"

# Métier - ORM
add_migration_comment "src/Classes/clsOrmBase.cls" "apex-metier/orm/clsOrmBase.cls"
add_migration_comment "src/Classes/clsRelationMetadata.cls" "apex-metier/orm/clsRelationMetadata.cls"
add_migration_comment "src/Interfaces/IRelationMetadata.cls" "apex-metier/orm/interfaces/IRelationMetadata.cls"
add_migration_comment "src/Interfaces/IRelationalObject.cls" "apex-metier/orm/interfaces/IRelationalObject.cls"

echo "[INFO] Vérification des éléments UI..."
# UI - Vérification des fichiers créés précédemment
if [ -f "apex-ui/ribbon/customUI.xml" ]; then
    echo "[OK] Fichier trouvé: apex-ui/ribbon/customUI.xml"
else
    echo "[ATTENTION] Fichier manquant: apex-ui/ribbon/customUI.xml"
fi

if [ -f "apex-ui/handlers/modRibbonCallbacks.bas" ]; then
    echo "[OK] Fichier trouvé: apex-ui/handlers/modRibbonCallbacks.bas"
else
    echo "[ATTENTION] Fichier manquant: apex-ui/handlers/modRibbonCallbacks.bas"
fi

echo "[INFO] Comptage des fichiers migrés..."
echo "Fichiers Core: $(find apex-core -type f | wc -l)"
echo "Fichiers Métier: $(find apex-metier -type f | wc -l)"
echo "Fichiers UI: $(find apex-ui -type f | wc -l)"

echo ""
echo "===== MIGRATION DES FICHIERS TERMINÉE ====="
echo "Vérifiez que tous les fichiers ont été correctement migrés."
echo "Pour finaliser, mettez à jour la documentation et testez la build avec tools/BuildRelease.bat" 