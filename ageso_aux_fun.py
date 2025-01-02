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
