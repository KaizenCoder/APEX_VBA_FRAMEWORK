#!/usr/bin/env python3
"""
Module de gestion de l'historique des chats pour le framework APEX VBA.

Ce module fournit une interface pour gérer et suivre l'historique des conversations
avec les assistants IA dans le cadre du développement du framework APEX VBA.

Fonctionnalités principales:
- Création et gestion des fichiers d'historique
- Validation des références de chat
- Suivi des impacts et des relations entre les chats
- Export des statistiques au format JSON
- Rotation automatique des fichiers d'historique

Utilisation typique:
    manager = ChatHistoryManager()
    manager.update_chat_reference(
        chat_num=1,
        action="Description de l'action",
        impact="Impact majeur - Description",
        source="chat_001",
        references=["chat_002", "chat_003"]
    )

Référence: chat_038 (2024-04-11 15:25) - Documentation du code
Source: chat_037 (2024-04-11 15:18) - Correction du test de parsing
"""

import os
import re
import json
import time
import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union

class ChatHistoryManager:
    """
    Gestionnaire de l'historique des chats.
    
    Cette classe gère l'enregistrement, la validation et le suivi des références
    de chat dans le cadre du développement du framework APEX VBA.
    
    Attributes:
        base_path (Path): Chemin vers le répertoire de stockage des fichiers d'historique
        current_file (Optional[Path]): Fichier d'historique actuellement utilisé
        current_stats (Dict): Statistiques de la session courante
        CHAT_MIN (int): Numéro minimum autorisé pour un chat (1)
        CHAT_MAX (int): Numéro maximum autorisé pour un chat (999)
        CHAT_REF_PATTERN (str): Pattern regex pour valider les références de chat
    """

    def __init__(self, base_path: str = "tools/workflow/chat_history"):
        """
        Initialise le gestionnaire d'historique.
        
        Args:
            base_path (str): Chemin vers le répertoire de stockage des fichiers
                           d'historique. Par défaut: "tools/workflow/chat_history"
        """
        self.base_path = Path(base_path)
        self.base_path.mkdir(parents=True, exist_ok=True)
        self.current_file: Optional[Path] = None
        self.current_stats: Dict = {}
        self.CHAT_MIN = 1
        self.CHAT_MAX = 999
        self.CHAT_REF_PATTERN = r"^chat_\d{3}$"

    def validate_chat_number(self, chat_num: int) -> None:
        """
        Valide un numéro de chat.
        
        Args:
            chat_num (int): Numéro de chat à valider
            
        Raises:
            TypeError: Si le numéro n'est pas un entier
            ValueError: Si le numéro est hors limites (1-999)
        """
        if not isinstance(chat_num, int):
            raise TypeError("Le numéro de chat doit être un entier")
        if not self.CHAT_MIN <= chat_num <= self.CHAT_MAX:
            raise ValueError(f"Le numéro de chat doit être entre {self.CHAT_MIN} et {self.CHAT_MAX}")

    def validate_chat_reference(self, ref: str) -> bool:
        """
        Valide le format d'une référence de chat.
        
        Args:
            ref (str): Référence à valider (format: chat_XXX)
            
        Returns:
            bool: True si la référence est valide, False sinon
        """
        if not isinstance(ref, str):
            return False
        if not re.match(self.CHAT_REF_PATTERN, ref):
            return False
        chat_num = int(ref.split('_')[1])
        return self.CHAT_MIN <= chat_num <= self.CHAT_MAX

    def get_current_file(self) -> Path:
        """
        Trouve ou crée le fichier d'historique actuel.
        
        Returns:
            Path: Chemin vers le fichier d'historique actuel
        """
        files = sorted(self.base_path.glob("*.md"), reverse=True)
        if not files:
            return self._create_new_file(1, 30)
        return files[0]

    def _create_new_file(self, start: int, end: int) -> Path:
        """
        Crée un nouveau fichier d'historique.
        
        Args:
            start (int): Premier numéro de chat du fichier
            end (int): Dernier numéro de chat du fichier
            
        Returns:
            Path: Chemin vers le nouveau fichier créé
            
        Raises:
            ValueError: Si end <= start
        """
        self.validate_chat_number(start)
        self.validate_chat_number(end)
        if end <= start:
            raise ValueError("La fin doit être supérieure au début")

        now = datetime.datetime.now()
        filename = f"{now.strftime('%Y-%m-%d_%H%M')}_chat_references_{start:03d}-{end:03d}.md"
        new_file = self.base_path / filename
        
        template = f"""# Journal des Références de Chat ({start:03d}-{end:03d})
*Période: {now.strftime('%Y-%m-%d %H:%M')} - En cours*
*Dernière mise à jour: chat_{start:03d}*

## 📋 Structure du Journal
- Plage de chats: {start:03d}-{end:03d}
- Mise à jour automatique tous les 10 chats
- Format standardisé des références

## 🔄 Plage Actuelle (chat_{start:03d} à chat_{end:03d})

## 📊 Statistiques
- Total des chats référencés: 0
- Chats avec impact majeur: 0
- Dernier chat: {start:03d}
- Prochaine mise à jour automatique: chat_{(start+9):03d}
"""
        new_file.write_text(template, encoding='utf-8')
        return new_file

    def parse_chat_reference(self, content: str) -> Dict:
        """
        Parse les références de chat dans le contenu.
        
        Args:
            content (str): Contenu du fichier d'historique à analyser
            
        Returns:
            Dict: Statistiques extraites du contenu
                - total: nombre d'entrées principales
                - major_impact: nombre d'impacts majeurs
                - last_chat: dernier numéro de chat
                - references: ensemble des références
                - sources: ensemble des sources
        """
        stats = {
            'total': 0,
            'major_impact': 0,
            'last_chat': 0,
            'references': set(),
            'sources': set()
        }
        
        # Extraction des entrées principales de chat
        main_entries = re.finditer(r"### chat_(\d{3})", content)
        chat_numbers = [int(match.group(1)) for match in main_entries]
        if chat_numbers:
            stats['last_chat'] = max(chat_numbers)
            stats['total'] = len(chat_numbers)
        
        # Extraction des impacts
        impact_pattern = r"Impact:\s*([^\n]+)"
        impacts = re.findall(impact_pattern, content)
        stats['major_impact'] = len([i for i in impacts if 
            any(kw in i.lower() for kw in ['majeur', 'critique', 'fondamental', 'structurel'])])
        
        # Extraction des sources et références
        source_pattern = r"Source:\s*chat_(\d{3})"
        ref_pattern = r"Références:\s*([^\n]+)"
        
        sources = re.findall(source_pattern, content)
        stats['sources'].update(sources)
        
        for ref_list in re.findall(ref_pattern, content):
            stats['references'].update(re.findall(r"chat_(\d{3})", ref_list))
        
        return stats

    def export_stats(self, output_file: Optional[Path] = None) -> None:
        """
        Exporte les statistiques en JSON.
        
        Args:
            output_file (Optional[Path]): Chemin du fichier de sortie
                Si None, utilise le nom par défaut dans base_path
        """
        if output_file is None:
            output_file = self.base_path / "chat_history_stats.json"
        
        current_file = self.get_current_file()
        content = current_file.read_text(encoding='utf-8')
        stats = self.parse_chat_reference(content)
        
        # Conversion des sets en listes pour JSON
        stats['references'] = sorted(list(stats['references']))
        stats['sources'] = sorted(list(stats['sources']))
        
        # Ajout métadonnées
        stats['export_date'] = datetime.datetime.now().isoformat()
        stats['file_name'] = current_file.name
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(stats, f, indent=2, ensure_ascii=False)

    def update_chat_reference(self, chat_num: int, action: str, impact: str, 
                            source: Optional[str] = None, references: Optional[List[str]] = None) -> None:
        """
        Met à jour les références avec une nouvelle entrée.
        
        Args:
            chat_num (int): Numéro du chat
            action (str): Description de l'action réalisée
            impact (str): Description de l'impact
            source (Optional[str]): Référence source (format: chat_XXX)
            references (Optional[List[str]]): Liste des références (format: chat_XXX)
            
        Raises:
            ValueError: Si le format des références est invalide
        """
        self.validate_chat_number(chat_num)
        
        if source and not self.validate_chat_reference(source):
            raise ValueError(f"Format de source invalide: {source}")
            
        if references:
            invalid_refs = [ref for ref in references if not self.validate_chat_reference(ref)]
            if invalid_refs:
                raise ValueError(f"Format de référence(s) invalide(s): {', '.join(invalid_refs)}")

        current_file = self.get_current_file()
        content = current_file.read_text(encoding='utf-8')
        
        # Création de la nouvelle entrée
        now = datetime.datetime.now()
        new_entry = f"\n### chat_{chat_num:03d} ({now.strftime('%Y-%m-%d %H:%M')})\n"
        new_entry += f"- Action: {action}\n"
        new_entry += f"- Impact: {impact}\n"
        
        if source:
            new_entry += f"- Source: {source}\n"
            
        if references:
            new_entry += f"- Références: {', '.join(references)}\n"

        # Insertion dans la section appropriée
        section_pattern = r"## 🔄 Plage Actuelle \(chat_\d{3} à chat_\d{3}\)"
        content = re.sub(section_pattern, f"\\g<0>\n{new_entry}", content)

        # Mise à jour des statistiques
        stats = self.parse_chat_reference(content)
        stats_section = f"""## 📊 Statistiques
- Total des chats référencés: {stats['total']}
- Chats avec impact majeur: {stats['major_impact']}
- Dernier chat: {stats['last_chat']:03d}
- Prochaine mise à jour automatique: chat_{((stats['last_chat']//10 + 1)*10):03d}
"""
        content = re.sub(r"## 📊 Statistiques.*?(?=\n\n|$)", stats_section, content, flags=re.DOTALL)
        
        # Sauvegarde
        current_file.write_text(content, encoding='utf-8')
        
        # Export des stats si c'est un multiple de 10
        if chat_num % 10 == 0:
            time.sleep(0.1)  # Petit délai pour assurer la différence de timestamp
            self.export_stats()

    def check_rotation(self, chat_num: int) -> None:
        """
        Vérifie si un nouveau fichier doit être créé.
        
        Args:
            chat_num (int): Numéro de chat actuel
            
        Note:
            Crée un nouveau fichier tous les 30 chats
        """
        self.validate_chat_number(chat_num)
        if chat_num % 30 == 1:  # Nouveau fichier tous les 30 chats
            self._create_new_file(chat_num, chat_num + 29)

def main():
    """Fonction principale pour les tests."""
    manager = ChatHistoryManager()
    
    # Exemple d'utilisation
    try:
        manager.update_chat_reference(
            chat_num=38,
            action="Documentation du code Python",
            impact="Impact majeur: Documentation complète du gestionnaire d'historique",
            source="chat_037",
            references=["chat_036"]
        )
        print("✅ Référence ajoutée avec succès")
    except Exception as e:
        print(f"❌ Erreur: {str(e)}")

if __name__ == "__main__":
    main() 