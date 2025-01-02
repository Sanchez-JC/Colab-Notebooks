import pandas as pd
import openpyxl
import os

def archivos():
  archivos = os.listdir()
  archivos.remove("Plantilla_0.xlsx")
  archivos.remove(".config")
  archivos.remove("sample_data")
  print(archivos)
  try:
    archivos.remove(".ipynb_checkpoints")
    archivos.remove("ageso_submain.py")
    archivos.remove("__pycache__")
  except:
    pass
  return archivos

def importar_abrir(nombre_archivo): #IMPORTA Y ABRE ARCHIVO
  ruta_datos = archivos[nombre_archivo]
  datos = pd.read_excel(ruta_datos)
  ruta_plantilla = "/content/Plantilla_0.xlsx"
  plantilla = openpyxl.load_workbook(ruta_plantilla)
  total_trabajadores = datos.shape[nombre_archivo]
  return

