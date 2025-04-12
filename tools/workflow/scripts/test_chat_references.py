#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Tests unitaires pour le gestionnaire de références de chat.
Référence: chat_039 (2024-04-11 15:35)
Source: chat_038 (Gestion erreurs silencieuses)
"""

import unittest
import os
import shutil
import tempfile
import logging
from datetime import datetime
from manage_chat_references import ChatReferenceManager, ChatReferenceError

class TestChatReferenceManager(unittest.TestCase):
    def setUp(self):
        """Initialisation avant chaque test."""
        self.test_dir = tempfile.mkdtemp()
        # Désactivation du logging pendant les tests
        logging.getLogger().setLevel(logging.CRITICAL)
        self.manager = ChatReferenceManager(base_dir=self.test_dir)
        
    def tearDown(self):
        """Nettoyage après chaque test."""
        shutil.rmtree(self.test_dir)
        
    def test_init_directories(self):
        """Test de la création des répertoires."""
        self.assertTrue(os.path.exists(self.test_dir))
        self.assertTrue(os.path.exists('tools/workflow/logs'))
        
    def test_validate_chat_number(self):
        """Test de la validation des numéros de chat."""
        self.assertTrue(self.manager.validate_chat_number(1))
        self.assertTrue(self.manager.validate_chat_number(999))
        self.assertFalse(self.manager.validate_chat_number(0))
        self.assertFalse(self.manager.validate_chat_number(1000))
        
    def test_get_current_file(self):
        """Test de la génération du nom de fichier."""
        self.manager.last_chat = 35
        current_file = self.manager.get_current_file()
        self.assertRegex(current_file, r"\d{4}-\d{2}-\d{2}_\d{4}_chat_references_031-060\.md$")
        
    def test_format_chat_entry(self):
        """Test du formatage des entrées de chat."""
        entry = self.manager.format_chat_entry(
            chat_num=38,
            action="Test action",
            impact="majeur",
            source="chat_037",
            refs=["chat_036"]
        )
        self.assertIn("### chat_038", entry)
        self.assertIn("- Action: Test action", entry)
        self.assertIn("- Impact: Impact majeur", entry)
        self.assertIn("- Source: chat_037", entry)
        self.assertIn("- Références: chat_036", entry)
        
    def test_update_chat_reference(self):
        """Test de la mise à jour des références."""
        # Test création initiale
        self.manager.update_chat_reference(
            chat_num=38,
            action="Test action",
            impact="majeur",
            source="chat_037",
            refs=["chat_036"]
        )
        
        current_file = self.manager.get_current_file()
        self.assertTrue(os.path.exists(current_file))
        
        with open(current_file, 'r', encoding='utf-8') as f:
            content = f.read()
            self.assertIn("### chat_038", content)
            self.assertIn("Test action", content)
            self.assertIn("Impact majeur", content)
            
    def test_calculate_stats(self):
        """Test du calcul des statistiques."""
        content = """
        ### chat_001
        - Impact: Impact majeur
        ### chat_002
        - Impact: Impact mineur
        ### chat_003
        - Impact: Impact critique
        """
        stats = self.manager._calculate_stats(content)
        self.assertEqual(stats['total'], 3)
        self.assertEqual(stats['major_impact'], 2)
        
    def test_invalid_chat_number(self):
        """Test de la gestion des numéros de chat invalides."""
        with self.assertRaises(ValueError):
            self.manager.update_chat_reference(
                chat_num=0,
                action="Test invalide",
                impact="majeur"
            )
            
    def test_file_creation_error(self):
        """Test de la gestion des erreurs de création de fichier."""
        # Rendre le répertoire en lecture seule
        os.chmod(self.test_dir, 0o444)
        with self.assertRaises(ChatReferenceError):
            self.manager.update_chat_reference(
                chat_num=38,
                action="Test erreur",
                impact="majeur"
            )
            
    def test_invalid_file_structure(self):
        """Test de la gestion des erreurs de structure de fichier."""
        # Créer un fichier avec une structure invalide
        current_file = self.manager.get_current_file()
        os.makedirs(os.path.dirname(current_file), exist_ok=True)
        with open(current_file, 'w', encoding='utf-8') as f:
            f.write("# Fichier invalide")
            
        with self.assertRaises(ChatReferenceError):
            self.manager.update_chat_reference(
                chat_num=38,
                action="Test erreur",
                impact="majeur"
            )
            
if __name__ == '__main__':
    unittest.main() 