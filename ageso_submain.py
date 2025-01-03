import pandas as pd
import openpyxl
import os

def listar_directorio():
  archivos = os.listdir()
  elementos_a_eliminar = ["Plantilla_0.xlsx", ".config", "sample_data", ".ipynb_checkpoints", "ageso_submain.py", "__pycache__"]
  for elemento_a_eliminar in elementos_a_eliminar:
    try:
      archivos.remove(elemento_a_eliminar)
    except:
      pass
  with open("text.txt","w") as file:
    file.write("I am learning Python!\n")
    file.write("I am really enjoying it!\n")
    file.write("And I want to add more lines to say how much I like it")
  return archivos

### FUNCIONES AUXILIARES ###

def ordenar(lista_porcentajes, lista_etiquetas): #El número de elementos es la cantidad de elementos diferentes de la lista. P. ej: para los sexos, el numero es 2, puesto que hay solo dos elementos diferentes "F" y "M". Lista_de_porcentajes es una lista con los porcentajes de la clase
  diccionario = dict(zip(lista_etiquetas, lista_porcentajes))
  diccionario_ordenado = dict(sorted(diccionario.items(), key=lambda x: x[1], reverse=True))
  return diccionario_ordenado

def clave_valor(diccionario, posicion):
  clave, valor = list(diccionario.items())[posicion]
  return clave, valor

def editor_valores(hoja, columnas, filas, valores):
  for i in columnas:
    for j, k in zip(filas, valores):
      celda = f"{i}{j}"
      hoja[celda] = k

def editor_porcentajes(hoja, columnas, filas, valores):
  for i in columnas:
    for j, k in zip(filas, valores):
      celda = f"{i}{j}"
      hoja[celda] = str(k) + "%"

def editor_conclusion(hoja, columnas, filas, valores):
  celda = f"{columnas}{filas}"
  hoja[celda] = valores

def ceros(lista):
  ceros = lista.count(0)
  return ceros

def get_patology_name(code):
  CIE = CIECodes() #Llama la librería e importa el diccionario con los códigos y nombres
  if CIE.info(code = code) is None: #Algunos códigos no están en la librería, aquí se verifica si están o no
    name = "None" #Si el código de patología no se encuentra a la librería, se le asigna nombre "None" a la patología con ese código
    print("Código no encontrado: ", code) #Imprime el código que no está en la librería
  else: #Si el código está en la librería
    name = CIE.info(code = code)["description"] #Obtiene el nombre de la patología asignada a dicho código de patología
  return name


    
archivos = listar_directorio()
def main(indice_archivos):

 #IMPORTA Y ABRE ARCHIVO
  ruta_datos = archivos[indice_archivos]
  datos = pd.read_excel(ruta_datos)
  ruta_plantilla = "/content/Plantilla_0.xlsx"
  plantilla = openpyxl.load_workbook(ruta_plantilla)
  total_trabajadores = datos.shape[indice_archivos]
  
##### TABLA INFORMACIÓN TRABAJADORES #####

  fechas = list(datos["fecha"])
  nombres = list(datos["nombre"])
  documentos = list(datos["documento"])
  ocupacion = list(datos["ocupacion"])
  
  hoja_informacion = plantilla["TRABAJ"]
  filas_informacion = [i for i in range(3, len(fechas)+3)]
  
  editor_valores(hoja_informacion, "B", filas_informacion, fechas)
  editor_valores(hoja_informacion, "C", filas_informacion, nombres)
  editor_valores(hoja_informacion, "D", filas_informacion, documentos)
  editor_valores(hoja_informacion, "E", filas_informacion, ocupacion)
  return

### SEXOS ###

  sexos = list(datos["genero"])

  #Clasificador y contador
  elementos_sexos = ["F", "M"]
  etiquetas_sexos = ["Femenino", "Masculino"]
  numero_femenino = sexos.count(elementos_sexos[0])
  numero_masculino = sexos.count(elementos_sexos[1])
  
  #Validación
  if numero_femenino + numero_masculino != total_trabajadores:
    popup("Error: la suma de los sexos es diferente al número total de trabajadores")
  
  #Porcentajes
  porcentaje_femenino = round(numero_femenino / total_trabajadores * 100)
  porcentaje_masculino = round(numero_masculino / total_trabajadores * 100)
  porcentajes_sexos = [porcentaje_femenino, porcentaje_masculino]
  orden_sexos = ordenar(porcentajes_sexos, etiquetas_sexos) #diccionario que tiene por clave las etiquetas y por valor los porcentajes ordenados
  S_may = clave_valor(orden_sexos, 0)[0] #Sexo mayor
  S_men = clave_valor(orden_sexos, 1)[0] #Sexo menor
  P_may = clave_valor(orden_sexos, 0)[1] #Porcentaje mayor
  P_men = clave_valor(orden_sexos, 1)[1] #Porcentaje menor
  
  b1 = f"De acuerdo con el sexo, se observó que el {P_may}% de la población es de Sexo {S_may}, mientras que un {P_men}% es de Sexo {S_men}."
  b2 = f"Se observó que el {P_may}% de la población es de Sexo {S_may}, mientras que un {P_men}% es de Sexo {S_men}."
  b3 = f"En relación con el sexo, se observó que el {P_may}% de la población es de Sexo {S_may} y un {P_men}% es de Sexo {S_men}."
  B = [b1, b2, b3]
  conclusion_sexos = B[random.randint(0, len(B)-1)]
  
  #Editor
  hoja_sexos = plantilla["SEX"]
  
  v_columnas_sexos = "B"
  v_filas_sexos = [6, 7, 8]
  v_valores_sexos = [numero_femenino, numero_masculino, numero_femenino + numero_masculino]
  p_columnas_sexos = "C"
  p_filas_sexos = [6, 7, 8]
  p_valores_sexos = [porcentaje_femenino, porcentaje_masculino, 100]
  
  editor_valores(hoja_sexos, v_columnas_sexos, v_filas_sexos, v_valores_sexos)
  editor_porcentajes(hoja_sexos, p_columnas_sexos, p_filas_sexos,p_valores_sexos)
  editor_conclusion(hoja_sexos, "A", 10, conclusion_sexos)
  plantilla.save(ruta_datos[3:])
  print("Filename:", ruta_datos[3:])
  return 
