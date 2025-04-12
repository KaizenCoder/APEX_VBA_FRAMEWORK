#!/bin/bash
# Script de vérification de la configuration WSL pour APEX VBA Framework

# Fonction pour afficher les sections
print_section() {
    echo -e "\n\033[1;36m=== $1 ===\033[0m"
}

# Vérifier le système
print_section "Informations Système"
echo "Distribution: $(lsb_release -ds)"
echo "Version du noyau: $(uname -r)"
echo "Architecture: $(uname -m)"

# Vérifier /etc/wsl.conf
print_section "Configuration WSL"
if [ -f /etc/wsl.conf ]; then
    echo "Fichier /etc/wsl.conf trouvé:"
    echo "-------------------------"
    cat /etc/wsl.conf
    echo "-------------------------"
else
    echo -e "\033[1;31mFichier /etc/wsl.conf non trouvé!\033[0m"
    echo "Exécutez le script configure_wsl.ps1 pour créer ce fichier."
fi

# Vérifier les points de montage
print_section "Points de Montage Windows"
echo "Points de montage Windows:"
mount | grep "/mnt" | sort

# Vérifier Git
print_section "Configuration Git"
if command -v git &> /dev/null; then
    echo "Git est installé:"
    git --version
    echo -e "\nConfiguration Git:"
    git config --list
else
    echo -e "\033[1;31mGit n'est pas installé!\033[0m"
    echo "Installez Git avec: sudo apt install git"
fi

# Test d'accès aux fichiers
print_section "Test d'Accès aux Fichiers"
PROJECT_DIR="/mnt/d/Dev/Apex_VBA_FRAMEWORK"

if [ -d "$PROJECT_DIR" ]; then
    echo "Répertoire du projet trouvé: $PROJECT_DIR"
    
    # Test d'écriture
    TEST_FILE="$PROJECT_DIR/wsl_test_file"
    if touch "$TEST_FILE" 2>/dev/null; then
        echo -e "\033[1;32mTest d'écriture réussi ✅\033[0m"
        rm "$TEST_FILE"
    else
        echo -e "\033[1;31mTest d'écriture échoué ❌\033[0m"
        echo "Erreur: Impossible d'écrire dans $PROJECT_DIR"
    fi
    
    # Test de lecture
    echo -e "\nFichiers dans le répertoire racine du projet:"
    ls -la "$PROJECT_DIR" | head -n 10
    echo "... (plus de fichiers)"
else
    echo -e "\033[1;31mRépertoire du projet non trouvé: $PROJECT_DIR\033[0m"
    echo "Vérifiez le chemin d'accès et la structure des points de montage."
fi

# Vérifier les permissions et utilisateurs
print_section "Utilisateurs et Permissions"
echo "Utilisateur actuel: $(whoami)"
echo "ID utilisateur: $(id -u)"
echo "ID groupe: $(id -g)"
echo "Tous les groupes: $(id -Gn)"

# Vérifier la connexion réseau
print_section "Connectivité Réseau"
echo "Test de connectivité Internet:"
if ping -c 1 github.com &> /dev/null; then
    echo -e "\033[1;32mConnexion à github.com réussie ✅\033[0m"
else
    echo -e "\033[1;31mImpossible de se connecter à github.com ❌\033[0m"
fi

# Résumé
print_section "Résumé"
echo "Vérifiez les résultats ci-dessus pour identifier les problèmes potentiels."
echo "Pour résoudre les problèmes, consultez le guide docs/WSL_SETUP_GUIDE.md"
echo "ou exécutez le script tools/workflow/scripts/configure_wsl.ps1" 