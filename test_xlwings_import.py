import sys
print(f"Python version: {sys.version}")
print(f"Python executable: {sys.executable}")

try:
    import xlwings as xw
    print(f"xlwings version: {xw.__version__}")
    print("xlwings importé avec succès")
except ImportError as e:
    print(f"Erreur lors de l'importation de xlwings: {e}")

input("Appuyez sur une touche pour continuer...") 