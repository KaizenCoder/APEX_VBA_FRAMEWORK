#!/usr/bin/env python3
"""
Tests unitaires pour le gestionnaire d'historique des chats.

Ce module contient l'ensemble des tests unitaires pour valider le fonctionnement
du gestionnaire d'historique des chats (ChatHistoryManager).

Les tests couvrent:
- Initialisation et configuration
- Validation des numéros de chat et des références
- Création et gestion des fichiers
- Parsing et analyse des références
- Export des statistiques
- Rotation automatique des fichiers

Chaque test est documenté et inclut des assertions pour vérifier
le comportement attendu du système.

Référence: chat_039 (2024-04-11 15:26) - Documentation des tests
Source: chat_038 (2024-04-11 15:25) - Documentation du code
"""

import unittest
import tempfile
import shutil
import json
from pathlib import Path
from datetime import datetime
from manage_chat_history import ChatHistoryManager

class TestChatHistoryManager(unittest.TestCase):
    """
    Suite de tests pour la classe ChatHistoryManager.
    
    Cette classe contient tous les tests unitaires nécessaires pour valider
    le bon fonctionnement du gestionnaire d'historique des chats.
    
    Chaque méthode de test se concentre sur un aspect spécifique du système
    et utilise un environnement temporaire pour éviter toute interférence.
    """

    def setUp(self):
        """
        Initialisation avant chaque test.
        
        Crée un répertoire temporaire et initialise une nouvelle instance
        du gestionnaire pour chaque test.
        """
        self.test_dir = tempfile.mkdtemp()
        self.manager = ChatHistoryManager(self.test_dir)

    def tearDown(self):
        """
        Nettoyage après chaque test.
        
        Supprime le répertoire temporaire et tous ses contenus.
        """
        shutil.rmtree(self.test_dir)

    def test_init(self):
        """
        Test de l'initialisation du gestionnaire.
        
        Vérifie:
        - Création du répertoire de base
        - Initialisation correcte des attributs
        - État initial des statistiques
        """
        self.assertTrue(Path(self.test_dir).exists())
        self.assertIsNone(self.manager.current_file)
        self.assertEqual(self.manager.current_stats, {})

    def test_validate_chat_number(self):
        """
        Test de la validation des numéros de chat.
        
        Vérifie:
        - Acceptation des numéros valides (1-999)
        - Rejet des numéros invalides
        - Gestion des types incorrects
        """
        # Tests valides
        for num in [1, 500, 999]:
            try:
                self.manager.validate_chat_number(num)
            except Exception as e:
                self.fail(f"validate_chat_number({num}) a levé une exception inattendue: {e}")

        # Tests invalides
        invalid_numbers = [
            (0, ValueError),
            (1000, ValueError),
            ("123", TypeError),
            (3.14, TypeError)
        ]
        for num, expected_error in invalid_numbers:
            with self.assertRaises(expected_error):
                self.manager.validate_chat_number(num)

    def test_create_new_file(self):
        """
        Test de création d'un nouveau fichier d'historique.
        
        Vérifie:
        - Création physique du fichier
        - Format correct du contenu
        - Présence des sections requises
        """
        file = self.manager._create_new_file(1, 30)
        self.assertTrue(file.exists())
        content = file.read_text()
        self.assertIn("Journal des Références de Chat (001-030)", content)

    def test_get_current_file(self):
        """
        Test de récupération du fichier actuel.
        
        Vérifie:
        - Création du premier fichier si nécessaire
        - Cohérence des appels successifs
        """
        file1 = self.manager.get_current_file()
        self.assertTrue(file1.exists())
        file2 = self.manager.get_current_file()
        self.assertEqual(file1, file2)

    def test_update_chat_reference(self):
        """
        Test de mise à jour des références.
        
        Vérifie:
        - Ajout correct d'une nouvelle entrée
        - Format des champs (action, impact, source, références)
        - Mise à jour des statistiques
        """
        self.manager.update_chat_reference(
            chat_num=1,
            action="Test action",
            impact="Impact test",
            source="chat_001",
            references=["chat_002", "chat_003"]
        )
        
        current_file = self.manager.get_current_file()
        content = current_file.read_text()
        self.assertIn("Test action", content)
        self.assertIn("Impact test", content)
        self.assertIn("chat_001", content)
        self.assertIn("chat_002", content)
        self.assertIn("chat_003", content)

    def test_invalid_chat_number(self):
        """
        Test des numéros de chat invalides.
        
        Vérifie:
        - Rejet des numéros hors limites
        - Messages d'erreur appropriés
        """
        with self.assertRaises(ValueError):
            self.manager.update_chat_reference(1000, "Test", "Impact")

    def test_check_rotation(self):
        """
        Test de la rotation des fichiers.
        
        Vérifie:
        - Création de nouveaux fichiers aux bons intervalles
        - Conservation des anciens fichiers
        """
        self.manager.check_rotation(1)
        self.manager.check_rotation(30)
        self.manager.check_rotation(31)
        files = list(Path(self.test_dir).glob("*.md"))
        self.assertGreaterEqual(len(files), 2)

    def test_parse_chat_reference(self):
        """
        Test du parsing des références.
        
        Vérifie:
        - Extraction correcte des statistiques
        - Comptage des impacts majeurs
        - Suivi des références et sources
        """
        content = """### chat_001
- Action: Test
- Impact: Impact majeur
- Source: chat_002
- Références: chat_003, chat_004

### chat_002
- Action: Test 2
- Impact: Impact mineur
"""
        stats = self.manager.parse_chat_reference(content)
        self.assertEqual(stats['total'], 2)
        self.assertEqual(stats['major_impact'], 1)
        self.assertEqual(stats['last_chat'], 2)
        self.assertEqual(len(stats['sources']), 1)
        self.assertEqual(len(stats['references']), 2)
        self.assertIn("002", stats['sources'])
        self.assertIn("003", stats['references'])
        self.assertIn("004", stats['references'])

    def test_impact_detection(self):
        """
        Test de la détection des impacts.
        
        Vérifie:
        - Détection des différents niveaux d'impact
        - Comptage correct des impacts majeurs
        """
        content = """### chat_001
- Impact: Impact majeur
### chat_002
- Impact: Impact critique
### chat_003
- Impact: Impact mineur
"""
        stats = self.manager.parse_chat_reference(content)
        self.assertEqual(stats['major_impact'], 2)

    def test_export_stats(self):
        """
        Test de l'export des statistiques.
        
        Vérifie:
        - Création du fichier JSON
        - Structure correcte des données
        - Présence de tous les champs requis
        """
        self.manager.update_chat_reference(
            chat_num=1,
            action="Test export",
            impact="Impact majeur",
            source="chat_001",
            references=["chat_002"]
        )

        self.manager.export_stats()
        stats_file = self.manager.base_path / "chat_history_stats.json"
        self.assertTrue(stats_file.exists())

        with open(stats_file, 'r', encoding='utf-8') as f:
            stats = json.load(f)
            self.assertIn('total', stats)
            self.assertIn('major_impact', stats)
            self.assertIn('last_chat', stats)
            self.assertIn('references', stats)
            self.assertIn('sources', stats)
            self.assertIn('export_date', stats)
            self.assertIn('file_name', stats)

    def test_invalid_reference_format(self):
        """
        Test de la validation du format des références.
        
        Vérifie:
        - Rejet des formats invalides
        - Validation des sources et références
        - Messages d'erreur appropriés
        """
        invalid_tests = [
            ("invalid_chat", "source"),
            ("chat_abc", "source"),
            ("chat_0", "source"),
            ("chat_1000", "source")
        ]

        for ref, param_type in invalid_tests:
            with self.assertRaises(ValueError):
                if param_type == "source":
                    self.manager.update_chat_reference(1, "Test", "Impact", source=ref)
                else:
                    self.manager.update_chat_reference(1, "Test", "Impact", references=[ref])

    def test_auto_export_on_multiple_ten(self):
        """
        Test de l'export automatique tous les 10 chats.
        
        Vérifie:
        - Déclenchement de l'export aux bons intervalles
        - Mise à jour du fichier de statistiques
        - Gestion des timestamps
        """
        stats_file = self.manager.base_path / "chat_history_stats.json"
        
        # Test avec un numéro non multiple de 10
        self.manager.update_chat_reference(1, "Test", "Impact")
        self.assertFalse(stats_file.exists())
        
        # Test avec un multiple de 10
        self.manager.update_chat_reference(10, "Test", "Impact")
        self.assertTrue(stats_file.exists())
        
        # Vérification de la mise à jour
        mtime_1 = stats_file.stat().st_mtime
        self.manager.update_chat_reference(20, "Test", "Impact")
        mtime_2 = stats_file.stat().st_mtime
        self.assertGreater(mtime_2, mtime_1)

    def test_references_and_sources_tracking(self):
        """
        Test du suivi des références et sources.
        
        Vérifie:
        - Enregistrement correct des sources
        - Suivi des références multiples
        - Mise à jour des statistiques
        """
        self.manager.update_chat_reference(
            chat_num=1,
            action="Test tracking",
            impact="Impact test",
            source="chat_002",
            references=["chat_003", "chat_004"]
        )
        
        current_file = self.manager.get_current_file()
        content = current_file.read_text(encoding='utf-8')
        stats = self.manager.parse_chat_reference(content)
        
        self.assertEqual(len(stats['sources']), 1)
        self.assertEqual(len(stats['references']), 2)
        self.assertIn("002", stats['sources'])
        self.assertIn("003", stats['references'])
        self.assertIn("004", stats['references'])

def main():
    """
    Point d'entrée pour l'exécution des tests.
    
    Permet d'exécuter les tests directement depuis ce fichier.
    """
    unittest.main()

if __name__ == '__main__':
    main() 