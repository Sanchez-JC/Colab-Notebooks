import pandas as pd
import openpyxl
import os

def archivos():
  archivos = os.listdir()
  elementos_a_eliminar = ["Plantilla_0.xlsx", ".config", "sample_data", ".ipynb_checkpoints", "ageso_submain.py", "__pycache__"]
  for elemento_a_eliminar in elementos_a_eliminar:
    try:
      archivos.remove(elemento_a_eliminar)
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

