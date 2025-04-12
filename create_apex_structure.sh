#!/bin/bash
# ==========================================================================
# Script : create_apex_structure.sh
# Version : 1.0
# Purpose : Création de la structure de répertoires pour Apex Framework v1.1
# Date : 10/04/2025
# ==========================================================================

echo "===== CRÉATION DE LA STRUCTURE APEX FRAMEWORK v1.1 ====="
echo ""

# Créer la structure de répertoires Core
echo "[INFO] Création de la structure Apex.Core..."
mkdir -p apex-core/interfaces
mkdir -p apex-core/utils
mkdir -p apex-core/testing

# Créer la structure de répertoires Métier
echo "[INFO] Création de la structure Apex.Métier..."
mkdir -p apex-metier/recette
mkdir -p apex-metier/xml
mkdir -p apex-metier/outlook
mkdir -p apex-metier/database/interfaces
mkdir -p apex-metier/orm/interfaces
mkdir -p apex-metier/restapi

# Créer la structure de répertoires UI
echo "[INFO] Création de la structure Apex.UI..."
mkdir -p apex-ui/ribbon
mkdir -p apex-ui/forms
mkdir -p apex-ui/handlers

# Créer le répertoire pour le wiki local
echo "[INFO] Création des répertoires supplémentaires..."
mkdir -p wiki_local
mkdir -p interop
mkdir -p roadmap
mkdir -p handover

# Vérification de la structure créée
echo "[INFO] Vérification de la structure..."
find apex-core apex-metier apex-ui -type d | sort

echo ""
echo "===== STRUCTURE DE RÉPERTOIRES CRÉÉE AVEC SUCCÈS ====="
echo "Pour continuer avec la migration des fichiers, suivez les instructions dans MIGRATION_APEX_FRAMEWORK.md" 