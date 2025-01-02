import pandas as pd
import openpyxl
import os

def submain(numero_de_archivos):
  archivos = os.listdir()
  archivos.remove("Plantilla_0.xlsx")
  archivos.remove(".config")
  archivos.remove("sample_data")
  try:
    archivos.remove(".ipynb_checkpoints")
  except:
    pass

  #IMPORTA Y ABRE ARCHIVO
  ruta_datos = archivos[numero_de_archivos]
  datos = pd.read_excel(ruta_datos)
  ruta_plantilla = "/content/Plantilla_0.xlsx"
  plantilla = openpyxl.load_workbook(ruta_plantilla)
  total_trabajadores = datos.shape[numero_de_archivos]
  print(archivos)

