#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script de gestion des références de chat pour le framework APEX.
Référence: chat_039 (2024-04-11 15:35)
Source: chat_038 (Gestion erreurs silencieuses)
"""

import os
import re
import json
import logging
import sys
from datetime import datetime
from typing import Dict, List, Optional, Tuple
from pathlib import Path

# Configuration du logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join('tools', 'workflow', 'logs', 'chat_references.log')),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

class ChatReferenceError(Exception):
    """Erreur personnalisée pour la gestion des références de chat."""
    pass

class ChatReferenceManager:
    def __init__(self, base_dir: str = "tools/workflow/chat_history"):
        """
        Initialise le gestionnaire de références.
        Crée les répertoires nécessaires s'ils n'existent pas.
        """
        self.base_dir = base_dir
        self.current_file = None
        self.current_range = (0, 0)
        self.last_chat = 0
        
        try:
            # Création des répertoires nécessaires
            Path(base_dir).mkdir(parents=True, exist_ok=True)
            Path('tools/workflow/logs').mkdir(parents=True, exist_ok=True)
            logger.info(f"Répertoires initialisés: {base_dir}")
        except Exception as e:
            logger.error(f"Erreur lors de l'initialisation des répertoires: {str(e)}")
            raise ChatReferenceError(f"Erreur d'initialisation: {str(e)}")
        
    def validate_chat_number(self, chat_num: int) -> bool:
        """Valide un numéro de chat."""
        try:
            is_valid = 1 <= chat_num <= 999
            if not is_valid:
                logger.warning(f"Numéro de chat invalide: {chat_num}")
            return is_valid
        except Exception as e:
            logger.error(f"Erreur lors de la validation du numéro de chat: {str(e)}")
            return False
        
    def get_current_file(self) -> str:
        """Détermine le fichier actuel basé sur le dernier chat."""
        try:
            current_time = datetime.now().strftime("%Y-%m-%d_%H%M")
            range_start = ((self.last_chat - 1) // 30) * 30 + 1
            range_end = range_start + 29
            filename = f"{current_time}_chat_references_{range_start:03d}-{range_end:03d}.md"
            full_path = os.path.join(self.base_dir, filename)
            logger.debug(f"Fichier courant: {full_path}")
            return full_path
        except Exception as e:
            logger.error(f"Erreur lors de la génération du nom de fichier: {str(e)}")
            raise ChatReferenceError(f"Erreur de génération de fichier: {str(e)}")
        
    def format_chat_entry(self, chat_num: int, action: str, impact: str,
                         source: Optional[str] = None, refs: Optional[List[str]] = None) -> str:
        """Formate une entrée de chat."""
        try:
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M")
            entry = [
                f"### chat_{chat_num:03d} ({current_time})",
                f"- Action: {action}",
                f"- Impact: Impact {impact}"
            ]
            if source:
                entry.append(f"- Source: {source}")
            if refs:
                entry.append(f"- Références: {', '.join(refs)}")
            return "\n".join(entry)
        except Exception as e:
            logger.error(f"Erreur lors du formatage de l'entrée: {str(e)}")
            raise ChatReferenceError(f"Erreur de formatage: {str(e)}")
        
    def update_chat_reference(self, chat_num: int, action: str, impact: str,
                            source: Optional[str] = None, refs: Optional[List[str]] = None) -> bool:
        """Met à jour les références de chat."""
        try:
            if not self.validate_chat_number(chat_num):
                raise ValueError(f"Numéro de chat invalide: {chat_num}")
                
            self.last_chat = max(self.last_chat, chat_num)
            current_file = self.get_current_file()
            
            # Création du fichier si nécessaire
            if not os.path.exists(current_file):
                logger.info(f"Création du nouveau fichier: {current_file}")
                range_start = ((chat_num - 1) // 30) * 30 + 1
                range_end = range_start + 29
                template = self._create_file_template(range_start, range_end)
                os.makedirs(os.path.dirname(current_file), exist_ok=True)
                with open(current_file, 'w', encoding='utf-8') as f:
                    f.write(template)
            
            # Mise à jour du fichier
            with open(current_file, 'r', encoding='utf-8') as f:
                content = f.read()
                
            # Ajout de la nouvelle entrée
            entry = self.format_chat_entry(chat_num, action, impact, source, refs)
            if "## 🔄 Plage Actuelle" in content:
                content = content.replace(
                    "## 🔄 Plage Actuelle",
                    f"## 🔄 Plage Actuelle\n\n{entry}\n"
                )
                logger.info(f"Entrée ajoutée pour chat_{chat_num:03d}")
            else:
                logger.error("Section 'Plage Actuelle' non trouvée dans le fichier")
                raise ChatReferenceError("Structure de fichier invalide")
            
            # Mise à jour des statistiques
            stats = self._calculate_stats(content)
            stats_section = self._format_stats(stats)
            content = re.sub(r"## 📊 Statistiques.*$", stats_section, content, flags=re.DOTALL)
            
            # Sauvegarde du fichier
            try:
                with open(current_file, 'w', encoding='utf-8') as f:
                    f.write(content)
                logger.info(f"Fichier mis à jour avec succès: {current_file}")
            except Exception as e:
                logger.error(f"Erreur lors de la sauvegarde du fichier: {str(e)}")
                raise ChatReferenceError(f"Erreur de sauvegarde: {str(e)}")
                
            return True
            
        except Exception as e:
            logger.error(f"Erreur lors de la mise à jour des références: {str(e)}")
            raise ChatReferenceError(f"Erreur de mise à jour: {str(e)}")
        
    def _create_file_template(self, range_start: int, range_end: int) -> str:
        """Crée le template pour un nouveau fichier."""
        try:
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M")
            return f"""# Journal des Références de Chat ({range_start:03d}-{range_end:03d})
*Période: {current_time} - En cours*
*Dernière mise à jour: chat_{self.last_chat:03d}*

## 📋 Structure du Journal
- Plage de chats: {range_start:03d}-{range_end:03d}
- Mise à jour automatique tous les 10 chats
- Format standardisé des références

## 🔄 Plage Actuelle

## 📊 Statistiques
- Total des chats référencés: 0
- Chats avec impact majeur: 0
- Dernier chat: {self.last_chat:03d}
- Prochaine mise à jour automatique: chat_{((self.last_chat // 10) + 1) * 10:03d}
"""
        except Exception as e:
            logger.error(f"Erreur lors de la création du template: {str(e)}")
            raise ChatReferenceError(f"Erreur de template: {str(e)}")
        
    def _calculate_stats(self, content: str) -> Dict:
        """Calcule les statistiques du fichier."""
        try:
            stats = {
                'total': len(re.findall(r"### chat_\d+", content)),
                'major_impact': len(re.findall(r"Impact majeur|Impact critique", content)),
                'last_chat': self.last_chat,
                'next_update': ((self.last_chat // 10) + 1) * 10
            }
            logger.debug(f"Statistiques calculées: {stats}")
            return stats
        except Exception as e:
            logger.error(f"Erreur lors du calcul des statistiques: {str(e)}")
            raise ChatReferenceError(f"Erreur de statistiques: {str(e)}")
        
    def _format_stats(self, stats: Dict) -> str:
        """Formate la section des statistiques."""
        return f"""## 📊 Statistiques
- Total des chats référencés: {stats['total']}
- Chats avec impact majeur: {stats['major_impact']}
- Dernier chat: {stats['last_chat']:03d}
- Prochaine mise à jour automatique: chat_{stats['next_update']:03d}"""

if __name__ == "__main__":
    try:
        manager = ChatReferenceManager()
        # Example usage:
        manager.update_chat_reference(
            chat_num=39,
            action="Amélioration de la gestion des erreurs",
            impact="majeur",
            source="chat_038",
            refs=["chat_037"]
        )
        logger.info("Mise à jour des références effectuée avec succès")
    except ChatReferenceError as e:
        logger.error(f"Erreur lors de l'exécution: {str(e)}")
        sys.exit(1) 