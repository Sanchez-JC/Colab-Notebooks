import pandas as pd
import openpyxl
import random
from cie.cie10 import CIECodes #Importar librería CIE10
import re
import itertools
import os
from IPython.display import display, HTML

def listar_directorio():
  archivos = os.listdir()
  elementos_a_eliminar = ["Plantilla_0.xlsx", ".config", "sample_data", ".ipynb_checkpoints", "ageso_submain.py", "__pycache__"]
  for elemento_a_eliminar in elementos_a_eliminar:
    try:
      archivos.remove(elemento_a_eliminar)
    except:
      pass
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

def popup(mensaje):
    popup_html = f"""
    <script>
    alert("{mensaje}");
    </script>
    """
    display(HTML(popup_html))

### FUNCIÓN PRINCIPAL ###   
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

  ########## EDADES ##########
  edades = list(datos["edad"])
  
  #Clasificador y contador
  edades_18_23 = [i for i in edades if 0 <= i <= 23]
  edades_24_29 = [i for i in edades if 24 <= i <= 29]
  edades_30_35 = [i for i in edades if 30 <= i <= 35]
  edades_36_41 = [i for i in edades if 36 <= i <= 41]
  edades_41 = [i for i in edades if i > 41]
  etiquetas_edades = ["18 a 23 años", "24 a 29 años", "30 a 35 años", "36 a 41 años", "Mayores de 41 años"]
  elementos_edades = [len(edades_18_23), len(edades_24_29), len(edades_30_35), len(edades_36_41), len(edades_41)]
  
  #Validación
  if sum(elementos_edades) != total_trabajadores:
    popup("Error: la suma de las edades es diferente al número total de trabajadores")
  
  #porcentajes
  porcentaje_18_23 = round(len(edades_18_23) / total_trabajadores * 100)
  porcentaje_24_29 = round(len(edades_24_29) / total_trabajadores * 100)
  porcentaje_30_35 = round(len(edades_30_35) / total_trabajadores * 100)
  porcentaje_36_41 = round(len(edades_36_41) / total_trabajadores * 100)
  porcentaje_41 = round(len(edades_41) / total_trabajadores * 100)
  porcentajes_edades = [porcentaje_18_23, porcentaje_24_29, porcentaje_30_35, porcentaje_36_41, porcentaje_41]
  orden_edades = ordenar(porcentajes_edades, etiquetas_edades)
  
  #Conclusión
  E1 = clave_valor(orden_edades, 0)[0] #label
  E2 = clave_valor(orden_edades, 1)[0] #label
  E3 = clave_valor(orden_edades, 2)[0] #label
  E4 = clave_valor(orden_edades, 3)[0] #label
  E5 = clave_valor(orden_edades, 4)[0] #label
  P1 = clave_valor(orden_edades, 0)[1] #porcentaje
  P2 = clave_valor(orden_edades, 1)[1] #porcentaje
  P3 = clave_valor(orden_edades, 2)[1] #porcentaje
  P4 = clave_valor(orden_edades, 3)[1] #porcentaje
  P5 = clave_valor(orden_edades, 4)[1] #porcentaje
  
  etiquetas_edades_ordenadas = [E1, E2, E3, E4, E5]
  porcentajes_edades_ordenados = [P1, P2, P3, P4, P5]
  valores_edades_ordenadas = elementos_edades
  
  ceros_edades = ceros(elementos_edades)
  
  S_0 = "De acuerdo con las edades de los trabajadores, se realizó una clasificación de ellos en diferentes grupos etarios, donde se encontró que el "
  S_1 = "Teniendo en cuenta las edades de los trabajadores, se realizó una clasificación de ellos en diferentes grupos etarios, donde se encontró que el "
  S_2 = "Con base en las edades de los trabajadores, se realizó una clasificación de ellos en diferentes grupos etarios, donde se encontró que el "
  S_3 = "Según las edades de los trabajadores, se realizó una clasificación de ellos en diferentes grupos etarios, donde se encontró que el "
  S = [S_0, S_1, S_2, S_3]
  
  e0 = f"el {P1}% de la población se encuentra entre {E1}, un {P2}% están entre {E2}, un {P3}% tienen entre {E3}, un {P4}% están entre {E4} y un {P5}% se encuentran entre {E5}"
  e1 = f"el {P1}% de la población se encuentra entre {E1}, un {P2}% están entre {E2}, un {P3}% tienen entre {E3} y un {P4}% están entre {E4}"
  e2 = f"el {P1}% de la población se encuentra entre {E1}, un {P2}% están entre {E2} y un {P3}% tienen entre {E3}"
  e3 = f"el {P1}% de la población se encuentra entre {E1} y un {P2}% están entre {E2}"
  e4 = f"el {P1}% de la población se encuentra entre {E1}"
  
  e = [e0, e1, e2, e3, e4]
  
  conclusion_edades = S[random.randint(0, len(e)-2)] + e[ceros_edades]
  
  #Editor
  hoja_edades = plantilla["GRUPOS ETARIOS"]
  
  elementos_edades.append(total_trabajadores)
  v_columnas_edades = "C"
  v_filas_edaes = [5, 6, 7, 8, 9, 10]
  v_valores_edades = elementos_edades
  porcentajes_edades.append(100)
  p_columnas_edades = "D"
  p_filas_edades = [5, 6, 7, 8, 9, 10]
  p_valores_edades = porcentajes_edades
  
  editor_valores(hoja_edades, v_columnas_edades, v_filas_edaes, v_valores_edades)
  editor_porcentajes(hoja_edades, p_columnas_edades, p_filas_edades, p_valores_edades)
  editor_conclusion(hoja_edades, "B", 12, conclusion_edades)

  ############## ESTADO CIVIL #################
  estados_civiles = list(datos["estado_civil"])
  
  #clasificador y contador
  etiquetas_estados = ["Solteros", "Casados", "en Unión Libre", "Separados",  "Viudos"]
  elementos_estados = [estados_civiles.count("Soltero"), estados_civiles.count("Casado"), estados_civiles.count("Union_libre"), estados_civiles.count("Separado") + estados_civiles.count("Divorciado"), estados_civiles.count("Viudo")]
  
  #Validación
  if sum(elementos_estados) != total_trabajadores:
    popup("Error: la suma de los estados civiles es diferente al número total de trabajadores")
  
  #Porcentajes
  porcentaje_solteros = round(elementos_estados[0] / total_trabajadores * 100)
  porcentaje_casados = round(elementos_estados[1] / total_trabajadores * 100)
  porcentaje_union_libre = round(elementos_estados[2] / total_trabajadores * 100)
  porcentaje_separados = round(elementos_estados[3] / total_trabajadores * 100)
  porcentaje_viudos = round(elementos_estados[4] / total_trabajadores * 100)
  porcentajes_estados = [porcentaje_solteros, porcentaje_casados, porcentaje_union_libre, porcentaje_separados, porcentaje_viudos]
  orden_estados = ordenar(porcentajes_estados, etiquetas_estados)
  
  #Conclusión
  EC1 = clave_valor(orden_estados, 0)[0] #label
  EC2 = clave_valor(orden_estados, 1)[0] #label
  EC3 = clave_valor(orden_estados, 2)[0] #label
  EC4 = clave_valor(orden_estados, 3)[0] #label
  EC5 = clave_valor(orden_estados, 4)[0] #label
  PEC1 = clave_valor(orden_estados, 0)[1] #porcentaje
  PEC2 = clave_valor(orden_estados, 1)[1] #porcentaje
  PEC3 = clave_valor(orden_estados, 2)[1] #porcentaje
  PEC4 = clave_valor(orden_estados, 3)[1] #porcentaje
  PEC5 = clave_valor(orden_estados, 4)[1] #porcentaje
  
  etiquetas_estados_ordenados = [EC1, EC2, EC3, EC4, EC5]
  porcentajes_estados_ordenados = [PEC1, PEC2, PEC3, PEC4, PEC5]
  valores_estados_ordenados = elementos_estados
  
  intro_0_estados_civiles = "En relación con el estado civil, el "
  intro_1_estados_civiles = "En cuanto al estado civil, se encontró que el "
  intro_2_estados_civiles = "Respecto con el estado civil, se encontró que el "
  intro_3_estados_civiles = "En cuanto al estado civil, se observó que el "
  intros_estados_civiles = [intro_0_estados_civiles, intro_1_estados_civiles, intro_2_estados_civiles, intro_3_estados_civiles]
  
  conclu_estado_civil_0 = f"el {PEC1}% de la población están {EC1}, un {PEC2}% son {EC2}, {PEC3}% son {EC3}, un {PEC4}% se encuentran {EC4} y un {PEC5}% son {EC5}"
  conclu_estado_civil_1 = f"el {PEC1}% de la población están {EC1}, un {PEC2}% son {EC2}, {PEC3}% son {EC3} y un {PEC4}% se encuentran {EC4}"
  conclu_estado_civil_2 = f"el {PEC1}% de la población están {EC1}, un {PEC2}% son {EC2} y {PEC3}% son {EC3}"
  conclu_estado_civil_3 = f"el {PEC1}% de la población están {EC1} y un {PEC2}% son {EC2}"
  conclu_estado_civil_4 = f"el {PEC1}% de la población están {EC1}"
  conclu_estados_civiles = [conclu_estado_civil_0, conclu_estado_civil_1, conclu_estado_civil_2, conclu_estado_civil_3, conclu_estado_civil_4]
  
  ceros_estados_civiles = ceros(elementos_estados)
  
  conclusion_estados_civiles = intros_estados_civiles[random.randint(0, len(intros_estados_civiles)-1)] + conclu_estados_civiles[ceros_estados_civiles]
  
  #Editor
  elementos_estados.append(total_trabajadores)
  hoja_estados_civiles = plantilla["ESTADO CIVIL"]
  v_columnas_estados_civiles = "C"
  V_filas_estados_civiles = [5, 6, 7, 8, 9, 10]
  v_valores_estados_civiles = elementos_estados
  porcentajes_estados.append(100)
  p_columnas_estados_civiles = "D"
  p_filas_estados_civiles = [5, 6, 7, 8, 9, 10]
  p_valores_estados_civiles = porcentajes_estados
  
  editor_valores(hoja_estados_civiles, v_columnas_estados_civiles, V_filas_estados_civiles, v_valores_estados_civiles)
  editor_porcentajes(hoja_estados_civiles, p_columnas_estados_civiles, p_filas_estados_civiles, p_valores_estados_civiles)
  editor_conclusion(hoja_estados_civiles, "B", 12, conclusion_estados_civiles)

  ########## ESCOLARIDADES #############
  escolaridades = list(datos["escolaridad"])
  
  #Clasificador y contador
  etiquetas_escolaridades = ["Analfabeta", "Primaria Incompleta", "Primaria Completa", "Secundaria Incompleta",
                             "Secundaria Completa", "Técnica Incompleta", "Técnica Completa", "Tecnológico Incompleto",
                             "Tecnológico Completo", "Universitario Incompleto", "Universitario Completo", "Postgrado Completo"]
  
  elementos_escolaridades = [escolaridades.count("Analfabeta"), escolaridades.count("Primaria_incompleta"), escolaridades.count("Primaria_completa"),
                             escolaridades.count("Secundaria_incompleta"), escolaridades.count("Secundaria_completa"), escolaridades.count("Tecnico_incompleto"),
                             escolaridades.count("Tecnico_completo"), escolaridades.count("Tecnologico_incompleto"), escolaridades.count("Tecnologico_completo"),
                             escolaridades.count("Universitario_incompleto"), escolaridades.count("Universitario_completo"), escolaridades.count("Estudios_posgrado")]
  
  #Validación
  if sum(elementos_escolaridades) != total_trabajadores:
    popup("Error: la suma de las escolaridades es diferente al número total de trabajadores")
  
  #Porcentajes
  porcentaje_analfabeta = round(elementos_escolaridades[0] / total_trabajadores * 100)
  porcentaje_primaria_incompleta = round(elementos_escolaridades[1] / total_trabajadores * 100)
  porcentaje_primaria_completa = round(elementos_escolaridades[2] / total_trabajadores * 100)
  porcentaje_secundaria_incompleta = round(elementos_escolaridades[3] / total_trabajadores * 100)
  porcentaje_secundaria_completa = round(elementos_escolaridades[4] / total_trabajadores * 100)
  porcentaje_tecnico_incompleto = round(elementos_escolaridades[5] / total_trabajadores * 100)
  porcentaje_tecnico_completo = round(elementos_escolaridades[6] / total_trabajadores * 100)
  porcentaje_tecnologico_incompleto = round(elementos_escolaridades[7] / total_trabajadores * 100)
  porcentaje_tecnologico_completo = round(elementos_escolaridades[8] / total_trabajadores * 100)
  porcentaje_universitario_incompleto = round(elementos_escolaridades[9] / total_trabajadores * 100)
  porcentaje_universitario_completo = round(elementos_escolaridades[10] / total_trabajadores * 100)
  porcentaje_postgrado_completo = round(elementos_escolaridades[11] / total_trabajadores * 100)
  porcentajes_escolaridades = [porcentaje_analfabeta, porcentaje_primaria_incompleta, porcentaje_primaria_completa, porcentaje_secundaria_incompleta,
                               porcentaje_secundaria_completa, porcentaje_tecnico_incompleto, porcentaje_tecnico_completo, porcentaje_tecnologico_incompleto,
                               porcentaje_tecnologico_completo, porcentaje_universitario_incompleto, porcentaje_universitario_completo, porcentaje_postgrado_completo]
  orden_escolaridades = ordenar(porcentajes_escolaridades, etiquetas_escolaridades)
  
  #Conclusiones
  ES1 = clave_valor(orden_escolaridades, 0)[0] #label
  ES2 = clave_valor(orden_escolaridades, 1)[0] #label
  ES3 = clave_valor(orden_escolaridades, 2)[0] #label
  ES4 = clave_valor(orden_escolaridades, 3)[0] #label
  ES5 = clave_valor(orden_escolaridades, 4)[0] #label
  ES6 = clave_valor(orden_escolaridades, 5)[0] #label
  ES7 = clave_valor(orden_escolaridades, 6)[0] #label
  ES8 = clave_valor(orden_escolaridades, 7)[0] #label
  ES9 = clave_valor(orden_escolaridades, 8)[0] #label
  ES10 = clave_valor(orden_escolaridades, 9)[0] #label
  ES11 = clave_valor(orden_escolaridades, 10)[0] #label
  ES12 = clave_valor(orden_escolaridades, 11)[0] #label
  PES1 = clave_valor(orden_escolaridades, 0)[1] #porcentaje
  PES2 = clave_valor(orden_escolaridades, 1)[1] #porcentaje
  PES3 = clave_valor(orden_escolaridades, 2)[1] #porcentaje
  PES4 = clave_valor(orden_escolaridades, 3)[1] #porcentaje
  PES5 = clave_valor(orden_escolaridades, 4)[1] #porcentaje
  PES6 = clave_valor(orden_escolaridades, 5)[1] #porcentaje
  PES7 = clave_valor(orden_escolaridades, 6)[1] #porcentaje
  PES8 = clave_valor(orden_escolaridades, 7)[1] #porcentaje
  PES9 = clave_valor(orden_escolaridades, 8)[1] #porcentaje
  PES10 = clave_valor(orden_escolaridades, 9)[1] #porcentaje
  PES11 = clave_valor(orden_escolaridades, 10)[1] #porcentaje
  PES12 = clave_valor(orden_escolaridades, 11)[1] #porcentaje
  
  etiquetas_escolaridades_ordenadas = [ES1, ES2, ES3, ES4, ES5, ES6, ES7, ES8, ES9, ES10, ES11, ES12]
  porcentajes_escolaridades_ordenados = [PES1, PES2, PES3, PES4, PES5, PES6, PES7, PES8, PES9, PES10, PES11, PES12]
  valores_escolaridades_ordenadas = elementos_escolaridades
  
  
  intro_0_escolaridades = "En relación con la escolaridad, el "
  intro_1_escolaridades = "En cuanto a la escolaridad, se encontró que el "
  intro_2_escolaridades = "Respecto con la escolaridad, se encontró que el "
  intro_3_escolaridades = "En cuanto a la escolaridad, se observó que el "
  intros_escolaridades = [intro_0_escolaridades, intro_1_escolaridades, intro_2_escolaridades, intro_3_escolaridades]
  
  conclu_escolaridad_0 = f"{PES1}% de la población cuentan con estudios de {ES1}, un {PES2}% con estudios de {ES2}, {PES3}% con estudios de {ES3}, un {PES4}% con estudios de {ES4}, {PES5}% cuentan con estudios de {ES5}, {PES6}% presentan estudios de {ES6}, {PES7}% tienen estudios de {ES7}, {PES8}% con estudios de {ES8}, {PES9}% presentan estudios de {ES9}, {PES10}% tienen estudios de {ES10}, {PES11}% cuentan con estudios de {ES11} y un {PES12}% con estudios de {ES12}"
  conclu_escolaridad_1 = f"{PES1}% de la población cuentan con estudios de {ES1}, un {PES2}% con estudios de {ES2}, {PES3}% con estudios de {ES3}, un {PES4}% con estudios de {ES4}, {PES5}% cuentan con estudios de {ES5}, {PES6}% presentan estudios de {ES6}, {PES7}% tienen estudios de {ES7}, {PES8}% con estudios de {ES8}, {PES9}% presentan estudios de {ES9}, {PES10}% tienen estudios de {ES10} y un {PES11}% con estudios de {ES11}"
  conclu_escolaridad_2 = f"{PES1}% de la población cuentan con estudios de {ES1}, un {PES2}% con estudios de {ES2}, {PES3}% con estudios de {ES3}, un {PES4}% con estudios de {ES4}, {PES5}% cuentan con estudios de {ES5}, {PES6}% presentan estudios de {ES6}, {PES7}% tienen estudios de {ES7}, {PES8}% con estudios de {ES8}, {PES9}% presentan estudios de {ES9} y un {PES10}% tienen estudios de {ES10}"
  conclu_escolaridad_3 = f"{PES1}% de la población cuentan con estudios de {ES1}, un {PES2}% con estudios de {ES2}, {PES3}% con estudios de {ES3}, un {PES4}% con estudios de {ES4}, {PES5}% cuentan con estudios de {ES5}, {PES6}% presentan estudios de {ES6}, {PES7}% tienen estudios de {ES7}, {PES8}% con estudios de {ES8} y un {PES9}% presentan estudios de {ES9}"
  conclu_escolaridad_4 = f"{PES1}% de la población cuentan con estudios de {ES1}, un {PES2}% con estudios de {ES2}, {PES3}% con estudios de {ES3}, un {PES4}% con estudios de {ES4}, {PES5}% cuentan con estudios de {ES5}, {PES6}% presentan estudios de {ES6}, {PES7}% tienen estudios de {ES7} y un {PES8}% con estudios de {ES8}"
  conclu_escolaridad_5 = f"{PES1}% de la población cuentan con estudios de {ES1}, un {PES2}% con estudios de {ES2}, {PES3}% con estudios de {ES3}, un {PES4}% con estudios de {ES4}, {PES5}% cuentan con estudios de {ES5} y un {PES6}% presentan estudios de {ES6}"
  conclu_escolaridad_6 = f"{PES1}% de la población cuentan con estudios de {ES1}, un {PES2}% con estudios de {ES2}, {PES3}% con estudios de {ES3}, un {PES4}% con estudios de {ES4} y un {PES5}% cuentan con estudios de {ES5}"
  conclu_escolaridad_7 = f"{PES1}% de la población cuentan con estudios de {ES1}, un {PES2}% con estudios de {ES2}, {PES3}% con estudios de {ES3} y un {PES4}% con estudios de {ES4}"
  conclu_escolaridad_8 = f"{PES1}% de la población cuentan con estudios de {ES1}, un {PES2}% con estudios de {ES2} y un {PES3}% con estudios de {ES3}"
  conclu_escolaridad_9 = f"{PES1}% de la población cuentan con estudios de {ES1} y un {PES2}% con estudios de {ES2}"
  conclu_escolaridad_10 = f"{PES1}% de la población cuentan con estudios de {ES1}"
  conclu_escolaridades = [conclu_escolaridad_0, conclu_escolaridad_1, conclu_escolaridad_2, conclu_escolaridad_3, conclu_escolaridad_4, conclu_escolaridad_5, conclu_escolaridad_6, conclu_escolaridad_7, conclu_escolaridad_8, conclu_escolaridad_9, conclu_escolaridad_10]
  
  ceros_escolaridades = ceros(elementos_escolaridades)
  
  conclusion_escolaridades = intros_escolaridades[random.randint(0, len(intros_escolaridades)-1)] + conclu_escolaridades[ceros_escolaridades]
  
  #editor
  hoja_escolaridades = plantilla["ESCOLARIDAD"]
  elementos_escolaridades.append(total_trabajadores)
  v_columnas_escolaridades = "C"
  V_filas_escolaridades = [13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26]
  v_valores_escolaridades = elementos_escolaridades
  porcentajes_escolaridades.append(100)
  p_columnas_escolaridades = "D"
  p_filas_escolaridades = [13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26]
  p_valores_escolaridades = porcentajes_escolaridades
  porcentajes_escolaridades
  
  editor_valores(hoja_escolaridades, v_columnas_escolaridades, V_filas_escolaridades, v_valores_escolaridades)
  editor_porcentajes(hoja_escolaridades, p_columnas_escolaridades, p_filas_escolaridades, p_valores_escolaridades)
  editor_conclusion(hoja_escolaridades, "B", 27, conclusion_escolaridades)
  
  ############## FUMADORES ##############
  fumadores = list(datos["habitos_tabaquismo1"])
  
  #Clasificador y contador
  etiquetas_fumadores = ["No Fumadores", "Fumadores", "Ex-fumadores"]
  elementos_fumadores = [fumadores.count("No Fuma"), fumadores.count("Fumador") + fumadores.count("Ocasional") , fumadores.count("Ex-fumador")]
  
  #Validación
  if sum(elementos_fumadores) != total_trabajadores:
    popup("Error: la suma de los fumadores es diferente al número total de trabajadores")
  
  #Porcentajes
  porcentaje_no_fuma = round(elementos_fumadores[0] / total_trabajadores * 100)
  porcentaje_fuma = round(elementos_fumadores[1] / total_trabajadores * 100)
  porcentaje_ex_fumador = round(elementos_fumadores[2] / total_trabajadores * 100)
  porcentajes_fumadores = [porcentaje_no_fuma, porcentaje_fuma, porcentaje_ex_fumador]
  orden_fumadores = ordenar(porcentajes_fumadores, etiquetas_fumadores)
  
  #Conclusión
  FF1 = clave_valor(orden_fumadores, 0)[0] #label
  FF2 = clave_valor(orden_fumadores, 1)[0] #label
  FF3 = clave_valor(orden_fumadores, 2)[0] #label
  FP1 = clave_valor(orden_fumadores, 0)[1] #porcentaje
  FP2 = clave_valor(orden_fumadores, 1)[1] #porcentaje
  FP3 = clave_valor(orden_fumadores, 2)[1] #porcentaje
  
  etiquetas_fumadores_ordenadas = [FF1, FF2, FF3]
  porcentajes_fumadores_ordenados = [FP1, FP2, FP3]
  valores_fumadores_ordenados = elementos_fumadores
  
  intro_0_fumadores = "De acuerdo con los hábitos toxicológicos, en cuanto al hábito de fumar, el "
  intro_2_fumadores = "En relación con el hábito de fumar, el "
  intro_3_fumadores = "Según los hábitos toxicológicos, en cuanto al hábito de fumar, el "
  intro_4_fumadores = "Con respecto al hábito de fumar, se encontró que el "
  intros_fumadores = [intro_0_fumadores, intro_2_fumadores, intro_3_fumadores, intro_4_fumadores]
  
  conclu_fumadores_0 = f"{FP1}% de la población son {FF1}, un {FP2}% son {FF2} y un {FP3}% son {FF3}"
  conclu_fumadores_1 = f"{FP1}% de la población son {FF1} y un {FP2}% son {FF2}"
  conclu_fumadores_2 = f"{FP1}% de la población son {FF1}"
  conclu_fumadores = [conclu_fumadores_0, conclu_fumadores_1, conclu_fumadores_2]
  
  ceros_fumadores = ceros(elementos_fumadores)
  
  conclusion_fumadores = intros_fumadores[random.randint(0, len(intros_fumadores)-1)] + conclu_fumadores[ceros_fumadores]
  
  #Editor
  hoja_fumadores = plantilla["FUMADOR"]
  elementos_fumadores.append(total_trabajadores)
  v_columnas_fumadores = "C"
  V_filas_fumadores = [4, 5, 6, 7]
  v_valores_fumadores = elementos_fumadores
  porcentajes_fumadores.append(100)
  p_columnas_fumadores = "D"
  p_filas_fumadores = [4, 5, 6, 7]
  p_valores_fumadores = porcentajes_fumadores
  
  editor_valores(hoja_fumadores, v_columnas_fumadores, V_filas_fumadores, v_valores_fumadores)
  editor_porcentajes(hoja_fumadores, p_columnas_fumadores, p_filas_fumadores, p_valores_fumadores)
  editor_conclusion(hoja_fumadores, "B", 10, conclusion_fumadores)
  
  ################# TOMADORES ##########
  tomadores = list(datos["habitos_licor2"])
  
  #Clasificador y contador
  etiquetas_tomadores = ["No Consumidor", "Consumidor"]
  elementos_tomadores = [tomadores.count("Ninguno"), sum(1 for i in tomadores if i != "Ninguno")]
  
  #Validación
  if sum(elementos_tomadores) != total_trabajadores:
    popup("Error: la suma de los tomadores es diferente al número total de trabajadores")
  
  #Porcentajes
  porcentaje_no_toma = round(elementos_tomadores[0] / total_trabajadores * 100)
  porcentaje_toma = round(elementos_tomadores[1] / total_trabajadores * 100)
  porcentajes_tomadores = [porcentaje_no_toma, porcentaje_toma]
  orden_tomadores = ordenar(porcentajes_tomadores, etiquetas_tomadores)
  
  
  #Conclusión
  TT1 = clave_valor(orden_tomadores, 0)[0] #label
  TT2 = clave_valor(orden_tomadores, 1)[0] #label
  TP1 = clave_valor(orden_tomadores, 0)[1] #porcentaje
  TP2 = clave_valor(orden_tomadores, 1)[1] #porcentaje
  
  etiquetas_tomadores_ordenadas = [TT1, TT2]
  porcentajes_tomadores_ordenados = [TP1, TP2]
  valores_tomadores_ordenados = elementos_tomadores
  
  intro_0_tomadores = "De acuerdo con los hábitos toxicológicos, en cuanto al consumo de bebidas alcohólicas, el "
  intro_1_tomadores = "En relación con el consumo de bebidas alcohólicas, se encontró que el "
  intro_2_tomadores = "En cuanto al consumo de bebidas alcohólicas, se encontró que el "
  intro_3_tomadores = "Según el consumo de bebidas alcohólicas, se encontró que el "
  intros_tomadores = [intro_0_tomadores, intro_1_tomadores, intro_2_tomadores, intro_3_tomadores]
  
  conclu_tomadores_0 = f"{TP1}% de la población manifestó ser {TT1}, mientras un {TP2}% son {TT2}."
  conclu_tomadores_1 = f"{TP1}% de la población manifesto ser {TT1}."
  conclusiones_tomadores = [conclu_tomadores_0, conclu_tomadores_1]
  
  ceros_tomadores = ceros(elementos_tomadores)
  
  conclusion_tomadores = intros_tomadores[random.randint(0, len(intros_tomadores)-1)] + conclusiones_tomadores[ceros_tomadores]
  
  #Editor
  hoja_tomadores = plantilla["ALCOHOL"]
  elementos_tomadores.append(total_trabajadores)
  v_columnas_tomadores = "C"
  V_filas_tomadores = [5, 6, 7]
  v_valores_tomadores = elementos_tomadores
  porcentajes_tomadores.append(100)
  p_columnas_tomadores = "D"
  p_filas_tomadores = [5, 6, 7]
  p_valores_tomadores = porcentajes_tomadores
  
  editor_valores(hoja_tomadores, v_columnas_tomadores, V_filas_tomadores, v_valores_tomadores)
  editor_porcentajes(hoja_tomadores, p_columnas_tomadores, p_filas_tomadores, p_valores_tomadores)
  editor_conclusion(hoja_tomadores, "B", 9, conclusion_tomadores)
  
  ############# ACTIVIDAD FÍSICA ##################
  ejercicio = list(datos["habitos_deportes2"])
  
  #Clasificador y contador
  etiquetas_ejercicio = ["Realizar", "No Realizar"]
  elementos_ejercicio = [sum(1 for i in ejercicio if i != "Ninguno"), ejercicio.count("Ninguno")]
  
  #Validación
  if sum(elementos_ejercicio) != total_trabajadores:
    popup("Error: la suma de los ejercicios es diferente al número total de trabajadores")
  
  #Porcentajes
  porcentaje_realiza = round(elementos_ejercicio[0] / total_trabajadores * 100)
  porcentaje_no_realiza = round(elementos_ejercicio[1] / total_trabajadores * 100)
  porcentajes_ejercicio = [porcentaje_realiza , porcentaje_no_realiza]
  orden_ejercicio = ordenar(porcentajes_ejercicio, etiquetas_ejercicio)
  
  #Conclusión
  EE1 = clave_valor(orden_ejercicio, 0)[0] #label
  EE2 = clave_valor(orden_ejercicio, 1)[0] #label
  EP1 = clave_valor(orden_ejercicio, 0)[1] #porcentaje
  EP2 = clave_valor(orden_ejercicio, 1)[1] #porcentaje
  
  etiquetas_ejercicio_ordenadas = [EE1, EE2]
  porcentajes_ejercicio_ordenados = [EP1, EP2]
  valores_ejercicio_ordenados = elementos_ejercicio
  
  intro_0_ejercicio = "En relación con la práctica de actividad física, el "
  intro_1_ejercicio = "En cuanto a la práctica de actividad física, el "
  intro_2_ejercicio = "Con respecto a la práctica de actividad física, el "
  intro_3_ejercicio = "A cerca de la práctica de actividad física, el "
  intros_ejercicio = [intro_0_ejercicio, intro_1_ejercicio, intro_2_ejercicio, intro_3_ejercicio]
  
  conclu_ejercicio_0 = f"{EP1}% de la población manifestó {EE1} actividad física, mientras un {EP2}% manifestó {EE2} ejercicio."
  conclu_ejercicio_1 = f"{EP1}% de la población manifestó {EE1} actividad física."
  conclusiones_ejercicio = [conclu_ejercicio_0, conclu_ejercicio_1]
  
  ceros_ejercicio = ceros(elementos_ejercicio)
  
  conclusion_ejercicio = intros_ejercicio[random.randint(0, len(intros_ejercicio)-1)] + conclusiones_ejercicio[ceros_ejercicio]
  
  #Editor
  hoja_ejercicio = plantilla["EJERCICIO"]
  elementos_ejercicio.append(total_trabajadores)
  v_columnas_ejercicio = "C"
  V_filas_ejercicio = [5, 6, 7]
  v_valores_ejercicio = elementos_ejercicio
  porcentajes_ejercicio.append(100)
  p_columnas_ejercicio = "D"
  p_filas_ejercicio = [5, 6, 7]
  p_valores_ejercicio = porcentajes_ejercicio
  
  editor_valores(hoja_ejercicio, v_columnas_ejercicio, V_filas_ejercicio, v_valores_ejercicio)
  editor_porcentajes(hoja_ejercicio, p_columnas_ejercicio, p_filas_ejercicio, p_valores_ejercicio)
  editor_conclusion(hoja_ejercicio, "B", 9, conclusion_ejercicio)
  
  ########### IMC ################
  imc = list(datos["imc"])
  
  #Clasificador y contador
  etiquetas_imc = ["Bajo Peso", "Normal", "Sobrepeso", "Obesidad I o Leve", "Obesidad II o Moderada", "Obesidad III o Severa"]
  elementos_imc = [sum(1 for i in imc if i <= 18), sum(1 for i in imc if 18 <= i < 25),
                   sum(1 for i in imc if 25 <= i < 30), sum(1 for i in imc if 30 <= i < 35),
                   sum(1 for i in imc if 35 <= i < 40), sum(1 for i in imc if i >= 40)]
  
  #Validación
  if sum(elementos_imc) != total_trabajadores:
    popup("Error: la suma de los imcs es diferente al número total de trabajadores")
  
  #Porcentajes
  porcentaje_bajo = round(elementos_imc[0] / total_trabajadores * 100)
  porcentaje_normal = round(elementos_imc[1] / total_trabajadores * 100)
  porcentaje_sobrepeso = round(elementos_imc[2] / total_trabajadores * 100)
  porcentaje_obesidad_1 = round(elementos_imc[3] / total_trabajadores * 100)
  porcentaje_obesidad_2 = round(elementos_imc[4] / total_trabajadores * 100)
  porcentaje_obesidad_3 = round(elementos_imc[5] / total_trabajadores * 100)
  porcentajes_imc = [porcentaje_bajo, porcentaje_normal, porcentaje_sobrepeso, porcentaje_obesidad_1, porcentaje_obesidad_2, porcentaje_obesidad_3]
  orden_imc = ordenar(porcentajes_imc, etiquetas_imc)
  
  #Conclusión
  IM1 = clave_valor(orden_imc, 0)[0] #label
  IM2 = clave_valor(orden_imc, 1)[0] #label
  IM3 = clave_valor(orden_imc, 2)[0] #label
  IM4 = clave_valor(orden_imc, 3)[0] #label
  IM5 = clave_valor(orden_imc, 4)[0] #label
  IM6 = clave_valor(orden_imc, 5)[0] #label
  IP1 = clave_valor(orden_imc, 0)[1] #porcentaje
  IP2 = clave_valor(orden_imc, 1)[1] #porcentaje
  IP3 = clave_valor(orden_imc, 2)[1] #porcentaje
  IP4 = clave_valor(orden_imc, 3)[1] #porcentaje
  IP5 = clave_valor(orden_imc, 4)[1] #porcentaje
  IP6 = clave_valor(orden_imc, 5)[1] #porcentaje
  
  etiquetas_imc_ordenadas = [IM1, IM2, IM3, IM4, IM5, IM6]
  porcentajes_imc_ordenados = [IP1, IP2, IP3, IP4, IP5, IP6]
  valores_imc_ordenados = elementos_imc
  
  intro_0_imc = "De acuerdo con el peso y estatura de los trabajadores se calculó el Índice de Masa Corporal (IMC), donde se encontró que el "
  intro_1_imc = "Teniendo en cuenta el peso y estatura de los trabajadores se calculó el Índice de Masa Corporal (IMC), donde se encontró que el "
  intro_2_imc = "Con base en el peso y estatura de los trabajadores, se realizó el cálculo del Índice de Masa Corporal, donde se encontró que el "
  intro_3_imc = "Según el peso y estatura de los trabajadores, se realizó el cálculo del Índice de Masa Corporal (IMC), donde se encontró que el "
  intros_imc = [intro_0_imc, intro_1_imc, intro_2_imc, intro_3_imc]
  
  conclu_imc_0 = f"{IP1}% de la población se encuentra en {IM1}, un {IP2}% están en {IM2}, {IP3}% se encuentran en {IM3}, {IP4}% en {IM4}, {IP5}% en {IM5} y un {IP6}% se encuentran en {IM6}"
  conclu_imc_1 = f"{IP1}% de la población se encuentra en {IM1}, un {IP2}% están en {IM2}, {IP3}% se encuentran en {IM3}, {IP4}% en {IM4} y un {IP5}% se encuentran en {IM5}"
  conclu_imc_2 = f"{IP1}% de la población se encuentra en {IM1}, un {IP2}% están en {IM2}, {IP3}% se encuentran en {IM3} y un {IP4}% se encuentran en {IM4}"
  conclu_imc_3 = f"{IP1}% de la población se encuentra en {IM1}, un {IP2}% están en {IM2} y un {IP3}% se encuentran en {IM3}"
  conclu_imc_4 = f"{IP1}% de la población se encuentra en {IM1} y un {IP2}% están en {IM2}"
  conclu_imc_5 = f"{IP1}% de la población se encuentra en {IM1}"
  conclu_imc = [conclu_imc_0, conclu_imc_1, conclu_imc_2, conclu_imc_3, conclu_imc_4, conclu_imc_5]
  
  ceros_imc = ceros(elementos_imc)
  conclusion_imc = intros_imc[random.randint(0, len(intros_imc)-1)] + conclu_imc[ceros_imc]
  
  #Editor
  hoja_imc = plantilla["IMC"]
  elementos_imc.append(total_trabajadores)
  v_columnas_imc = "C"
  V_filas_imc = [5, 6, 7, 8, 9, 10, 11]
  v_valores_imc = elementos_imc
  porcentajes_imc.append(100)
  p_columnas_imc = "D"
  p_filas_imc = [5, 6, 7, 8, 9, 10, 11]
  p_valores_imc = porcentajes_imc
  
  editor_valores(hoja_imc, v_columnas_imc, V_filas_imc, v_valores_imc)
  editor_porcentajes(hoja_imc, p_columnas_imc, p_filas_imc, p_valores_imc)
  editor_conclusion(hoja_imc, "B", 13, conclusion_imc)
  
  #### ACCIDENTES LABORALES ####
  accidentes = list(datos["obs_antecedpatocupacional"])
  
  #Clasificador y contador
  etiquetas_accidentes = ["No haber", "Si haber"]
  elementos_accidentes = [sum(1 for i in accidentes if i[0:2] == "No"),
                          sum(1 for i in accidentes if i[0:2] != "No")]
  
  #Validación
  if sum(elementos_accidentes) != total_trabajadores:
    popup("Error: la suma de los accidentes laborales es diferente al número total de trabajadores")
  
  #Porcentajes
  porcentaje_no_accidentes = round(elementos_accidentes[0] / total_trabajadores * 100)
  porcentaje_accidentes = round(elementos_accidentes[1] / total_trabajadores * 100)
  porcentajes_accidentes = [porcentaje_no_accidentes, porcentaje_accidentes]
  orden_accidentes = ordenar(porcentajes_accidentes, etiquetas_accidentes)
  
  #Conclusión
  AA1 = clave_valor(orden_accidentes, 0)[0] #label
  AA2 = clave_valor(orden_accidentes, 1)[0] #label
  AP1 = clave_valor(orden_accidentes, 0)[1] #porcentaje
  AP2 = clave_valor(orden_accidentes, 1)[1] #porcentaje
  
  etiquetas_accidentes_ordenadas = [AA1, AA2]
  porcentajes_accidentes_ordenados = [AP1, AP2]
  valores_accidentes_ordenados = elementos_accidentes
  
  intro_acc_1 = "Con respecto al historial de accidentes laborales, se encontró que el "
  intro_acc_2 = "En relación al historial de accidentes laborales, se observó que el "
  intro_acc_3 = "En cuanto al historial de accidentes laborales, se constató que el "
  intro_acc_4 = "Con referencia al historial de accidentes laborales, se evidenció que el "
  intros_accidentes = [intro_acc_1, intro_acc_2, intro_acc_3, intro_acc_4]
  
  conclu_acc_0 = f"{AP1}% de la población manifestó {AA1} tenido accidentes laborales, mientras un {AP2}% manifesto {AA2} tenido accidentes en el trabajo."
  conclu_acc_1 = f"{AP1}% de la población manifestó {AA1} tenido accidentes laborales."
  conclusiones_acc = [conclu_acc_0, conclu_acc_1]
  
  ceros_accidentes = ceros(elementos_accidentes)
  
  conclusion_accidentes = intros_accidentes[random.randint(0, len(intros_accidentes)-1)] + conclusiones_acc[ceros_accidentes]
  
  #Editor
  hoja_accidentes = plantilla["ACCID TRAB"]
  elementos_accidentes.append(total_trabajadores)
  v_columnas_accidentes = "B"
  V_filas_accidentes = [6, 7, 8]
  v_valores_accidentes = elementos_accidentes
  porcentajes_accidentes.append(100)
  p_columnas_accidentes = "C"
  p_filas_accidentes = [6, 7, 8]
  p_valores_accidentes = porcentajes_accidentes
  
  editor_valores(hoja_accidentes, v_columnas_accidentes, V_filas_accidentes, v_valores_accidentes)
  editor_porcentajes(hoja_accidentes, p_columnas_accidentes, p_filas_accidentes, p_valores_accidentes)
  editor_conclusion(hoja_accidentes, "A", 10, conclusion_accidentes)
  
  ######## RIESGOS ACTUALES ########
  riesgos = list(datos["obs_antecedocupacional"])
  
  #clasificador y contador
  etiquetas_riesgos = ["Psicosociales", "Mecánicos", "Biológico", "Químicos", "Públicos",
      "Seguridad Industrial", "Físicos", "Ergonómicos", "Eléctricos"]
  
  elementos_riesgos = [sum(list(str(i).count("Psicosociales") for i in riesgos)),
                       sum(list(str(i).count("Mecánicos") for i in riesgos)),
                       sum(list(str(i).count("Biológicos") for i in riesgos)),
                       sum(list(str(i).count("Químicos") for i in riesgos)),
                       sum(list(str(i).count("Públicos") for i in riesgos)),
                       sum(list(str(i).count("Seguridad Industrial") for i in riesgos)),
                       sum(list(str(i).count("Físicos") for i in riesgos)),
                       sum(list(str(i).count("Ergonómicos") for i in riesgos))
                       + sum(list(str(i).count("Posturas y movimientos") for i in riesgos)),
                       sum(list(str(i).count("Eléctricos") for i in riesgos))]
  
  total_riesgos = sum(elementos_riesgos)
  total_riesgos_teoricos = sum(list(str(i).count("FR:") for i in riesgos))
  
  #Validación
  if sum(elementos_riesgos) != total_riesgos_teoricos:
    popup("Advertencia: la suma de los riesgos laborales es diferente al número total de Riesgos")
  
  #Porcentajes
  porcentaje_psicosociales = round(elementos_riesgos[0] / total_riesgos * 100)
  porcentaje_mecanicos = round(elementos_riesgos[1] / total_riesgos * 100)
  porcentaje_biologicos = round(elementos_riesgos[2] / total_riesgos * 100)
  porcentaje_quimicos = round(elementos_riesgos[3] / total_riesgos * 100)
  porcentaje_publicos = round(elementos_riesgos[4] / total_riesgos * 100)
  porcentaje_seguridad = round(elementos_riesgos[5] / total_riesgos * 100)
  porcentaje_fisicos = round(elementos_riesgos[6] / total_riesgos * 100)
  porcentaje_ergonomicos = round(elementos_riesgos[7] / total_riesgos * 100)
  porcentaje_electricos = round(elementos_riesgos[8] / total_riesgos * 100)
  
  porcentajes_riesgos = [porcentaje_psicosociales, porcentaje_mecanicos, porcentaje_biologicos,
                         porcentaje_quimicos, porcentaje_publicos, porcentaje_seguridad,
                         porcentaje_fisicos, porcentaje_ergonomicos, porcentaje_electricos]
  orden_riesgos = ordenar(porcentajes_riesgos, etiquetas_riesgos)
  
  #Conclusión
  RR1 = clave_valor(orden_riesgos, 0)[0] #label
  RR2 = clave_valor(orden_riesgos, 1)[0] #label
  RR3 = clave_valor(orden_riesgos, 2)[0] #label
  RR4 = clave_valor(orden_riesgos, 3)[0] #label
  RR5 = clave_valor(orden_riesgos, 4)[0] #label
  RR6 = clave_valor(orden_riesgos, 5)[0] #label
  RR7 = clave_valor(orden_riesgos, 6)[0] #label
  RR8 = clave_valor(orden_riesgos, 7)[0] #label
  RR9 = clave_valor(orden_riesgos, 8)[0] #label
  RP1 = clave_valor(orden_riesgos, 0)[1] #porcentaje
  RP2 = clave_valor(orden_riesgos, 1)[1] #porcentaje
  RP3 = clave_valor(orden_riesgos, 2)[1] #porcentaje
  RP4 = clave_valor(orden_riesgos, 3)[1] #porcentaje
  RP5 = clave_valor(orden_riesgos, 4)[1] #porcentaje
  RP6 = clave_valor(orden_riesgos, 5)[1] #porcentaje
  RP7 = clave_valor(orden_riesgos, 6)[1] #porcentaje
  RP8 = clave_valor(orden_riesgos, 7)[1] #porcentaje
  RP9 = clave_valor(orden_riesgos, 8)[1] #porcentaje
  
  etiquetas_riesgos_ordenadas = [RR1, RR2, RR3, RR4, RR5, RR6, RR7, RR8, RR9]
  porcentajes_riesgos_ordenados = [RP1, RP2, RP3, RP4, RP5, RP6, RP7, RP8, RP9]
  valores_riesgos_ordenados = elementos_riesgos
  
  intro_0_riesgos = "En relación con la exposición laboral al factor de riesgo actual, se encontró que el "
  intro_1_riesgos = "En cuanto a la exposición laboral al factor de riesgo actual, se observó que el "
  intro_2_riesgos = "Con respecto a la exposición laboral al factor de riesgo actual, se constató que el "
  intro_3_riesgos = "En lo que respecta a la exposición laboral al factor de riesgo actual, se evidenció  el "
  intros_riesgos = [intro_0_riesgos, intro_1_riesgos, intro_2_riesgos, intro_3_riesgos]
  
  conclu_riesgos_0 = f"{RP1}% de los factores de riesgo encontrados corresponde a Factores {RR1}, un {RP2}% corresponde a Factores {RR2}, {RP3}% a los Factores {RR3}, un {RP4}% a Factores {RR4}, {RP5}% corresponde a Factores {RR5}, {RP6}% a los Factores {RR6}, de igual manera, un {RP7}% a los Factores {RR7} y finalmente un {RP8}% y {RP9} corresponden a Factores {RR8} y {RR9} respectivamente."
  conclu_riesgos_1 = f"{RP1}% de los factores de riesgo encontrados corresponde a Factores {RR1}, un {RP2}% corresponde a Factores {RR2}, {RP3}% a los Factores {RR3}, un {RP4}% corresponde a Factores {RR4}, {RP5}% corresponde a Factores {RR5}, {RP6}% a los Factores {RR6} y finalmente, un {RP7}% y un {RP8}% corresponden a los Factores {RR7} y {RR8} respectivamente."
  conclu_riesgos_2 = f"{RP1}% de los factores de riesgo encontrados corresponde a Factores {RR1}, un {RP2}% corresponde a Factores {RR2}, {RP3}% a los Factores {RR3}, un {RP4}% corresponde a Factores {RR4}, {RP5}% corresponde a Factores {RR5} y finalemente un {RP6}% y un {RP7}% corresponden a los Factores {RR6} y {RR7} respectivamente."
  conclu_riesgos_3 = f"{RP1}% de los factores de riesgo encontrados corresponde a Factores {RR1}, un {RP2}% corresponde a Factores {RR2}, {RP3}% a los Factores {RR3}, un {RP4}% corresponde a Factores {RR4} y finalmente un {RP5}% y un {RP6} corresponden a Factores {RR5} y {RR6} respectivamente."
  conclu_riesgos_4 = f"{RP1}% de los factores de riesgo encontrados corresponde a Factores {RR1}, un {RP2}% corresponde a Factores {RR2}, {RP3}% a los Factores {RR3} y finalmente un {RP4}% y un {RP5}% corresponden a Factores {RR4} y {RR5} respectivamente."
  conclu_riesgos_5 = f"{RP1}% de los factores de riesgo encontrados corresponde a Factores {RR1}, un {RP2}% corresponde a Factores {RR2} y finalmente un {RP3}% y un {RP4}% corresponden a factores {RR3} y {RR4} respectivamente."
  conclu_riesgos_6 = f"{RP1}% de los factores de riesgo encontrados corresponde a Factores {RR1} y un {RP2}% y {RP3} corresponden a Factores {RR2} y {RR3} respectivamente."
  conclu_riesgos_7 = f"{RP1}% de los factores de riesgo encontrados corresponde a Factores {RR1} y un {RP2}% corresponde a Factores {RR2}."
  conclu_riesgos_8 = f"{RP1}% de los factores de riesgo encontrados corresponde a Factores {RR1}."
  conclusiones_riesgos = [conclu_riesgos_0, conclu_riesgos_1, conclu_riesgos_2, conclu_riesgos_3, conclu_riesgos_4, conclu_riesgos_5, conclu_riesgos_6, conclu_riesgos_7, conclu_riesgos_8]
  
  ceros_riesgos = ceros(elementos_riesgos)
  
  conclusion_riesgos = intros_riesgos[random.randint(0, len(intros_riesgos)-1)] + conclusiones_riesgos[ceros_riesgos]
  
  #Editor
  hoja_riesgos = plantilla["FACTOR DE RIESGO ACTUAL"]
  elementos_riesgos.append(total_riesgos)
  v_columnas_riesgos = "B"
  V_filas_riesgos = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
  v_valores_riesgos = elementos_riesgos
  porcentajes_riesgos.append(100)
  p_columnas_riesgos = "C"
  p_filas_riesgos = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
  p_valores_riesgos = porcentajes_riesgos
  
  editor_valores(hoja_riesgos, v_columnas_riesgos, V_filas_riesgos, v_valores_riesgos)
  editor_porcentajes(hoja_riesgos, p_columnas_riesgos, p_filas_riesgos, p_valores_riesgos)
  editor_conclusion(hoja_riesgos, "A", 15, conclusion_riesgos)
  
  #### PATOLOGÍAS ####
  patologias = list(datos["obs_diagnostico"])
  
  #Clasificador y contador
  codigos = [re.findall(r'CIE10\|([A-Z0-9]+):', l) for l in patologias] #Obtiene solo los códigos
  codigos = list(itertools.chain(*codigos))
  
  solo_codigos = list(set(codigos))
  valores_solo_patologias = [codigos.count(i) for i in solo_codigos]
  nombres_patologias = [get_patology_name(k) for k in solo_codigos]
  
  #Totales
  total_patologias = sum(valores_solo_patologias)
  
  #Editor
  hoja_patologias = plantilla["PATOLOGIAS"]
  filas_patologias = [i for i in range(19, len(nombres_patologias) + 19)]
  
  editor_valores(hoja_patologias, "B", filas_patologias, nombres_patologias)
  editor_valores(hoja_patologias, "C", filas_patologias, valores_solo_patologias)
  editor_conclusion(hoja_patologias, "B", len(nombres_patologias) + 20, "Total")
  editor_conclusion(hoja_patologias, "C", len(nombres_patologias) + 20, total_patologias)

  plantilla.save(ruta_datos[3:])
  print("Filename:", ruta_datos[3:])
  return 
   
