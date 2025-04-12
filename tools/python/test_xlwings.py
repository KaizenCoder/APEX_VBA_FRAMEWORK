import xlwings as xw
import pandas as pd
from datetime import datetime
import os
import time

def test_xlwings_integration():
    print("Démarrage du test xlwings...")
    
    # Définir le chemin complet du fichier
    current_dir = os.path.abspath(os.path.dirname(__file__))
    excel_path = os.path.join(current_dir, "Test_APEX_Framework.xlsx")
    print(f"Création du classeur à l'emplacement : {excel_path}")
    
    # Fermer toutes les instances d'Excel
    try:
        xw.apps.active.quit()
    except:
        pass
    
    # Créer une nouvelle instance d'Excel avec des options explicites
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    app.screen_updating = True
    print("Instance Excel créée - Visible =", app.visible)
    
    # Créer un nouveau classeur
    wb = app.books.add()
    time.sleep(1)  # Attendre que Excel soit prêt
    
    # Forcer la fenêtre Excel au premier plan
    app.api.WindowState = xw.constants.WindowState.xlMaximized
    app.activate(steal_focus=True)
    
    wb.save(excel_path)
    sheet = wb.sheets[0]
    sheet.name = "Test APEX"  # Renommer la feuille
    
    # Écrire des données
    data = {
        'Test': ['xlwings', 'Python', 'Excel', 'APEX Framework'],
        'Status': ['OK', 'OK', 'OK', 'En cours'],
        'Date': [datetime.now()] * 4
    }
    df = pd.DataFrame(data)
    
    # Écrire le DataFrame
    sheet.range('A1').value = df
    
    # Formater
    sheet.range('A1:C1').color = (217, 217, 217)  # Gris clair
    sheet.range('A1:C5').api.Borders.LineStyle = 1
    
    # Ajuster les colonnes
    sheet.autofit()
    
    # Sauvegarder
    wb.save()
    print("Test terminé avec succès!")
    print(f"Fichier sauvegardé: {excel_path}")
    print("Excel devrait être visible à l'écran")
    print("État de la fenêtre Excel:", "Visible" if app.visible else "Non visible")
    
    # Garder Excel ouvert pour vérification
    input("Appuyez sur Entrée pour fermer Excel...")
    wb.close()
    app.quit()

if __name__ == '__main__':
    test_xlwings_integration() 