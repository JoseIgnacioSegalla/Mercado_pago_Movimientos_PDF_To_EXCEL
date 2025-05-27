import fitz  # PyMuPDF
import pandas as pd
import re


ruta = "C:/Users/ignac/Desktop/pdftoexcel/"
nombre_pdf = "11-24"
# Abre el archivo PDF
pdf_documento = fitz.open(ruta + nombre_pdf + ".pdf")

titulos = ['Fecha', 'Descripcion', 'ID de operacion', 'Valor', 'Saldo']

datos = []
for num_pagina in range(len(pdf_documento)):
    pagina = pdf_documento[num_pagina]
    contenido = pagina.get_text("text")
    datos.extend(contenido.split())



datos_filtrados = []
total_paginas = len(pdf_documento)
f = 0
enter = False
estado = True
pos = 2
while f < len(datos):

    if "generación:" == datos[f]:
        print("entro")
        f += 7
       

    if f"{pos}/{total_paginas}" == datos[f]:
        enter = False
        pos += 1
    
    elif "Fecha" in datos[f]:
        enter = False

    elif len(datos[f].split('-')) == 3:
        enter = True
      
        
    

    if enter:
        datos_filtrados.append(datos[f])
        
        

    f += 1


i = 0
pi = 0
datos_concatenados = []
while i < len(datos_filtrados):

    match = re.match(r"([a-zA-ZñÑ,.]+)([0-9]+)", datos_filtrados[i])
    

    if len(datos_filtrados[i].split('-')) == 3:
        pi = i
        datos_concatenados.append(datos_filtrados[i])
    elif match and len(match.group(2)) >= 8:


        datos_concatenados.append(" ".join(datos_filtrados[pi + 1 : i]) + " " + match.group(1))
        datos_concatenados.append(match.group(2))
            
    elif len(datos_filtrados[i]) >= 8 and datos_filtrados[i].isdigit():

        datos_concatenados.append(" ".join(datos_filtrados[pi + 1 : i]))
        datos_concatenados.append(datos_filtrados[i])

    elif datos_filtrados[i] == '$':
        datos_concatenados.append(datos_filtrados[i] + datos_filtrados[i + 1])

    i+= 1
    

t = 0
fila = []
datos_final = []
while t < len(datos_concatenados):
    if len(fila) < 5:
        fila.append(datos_concatenados[t])
    if len(fila) == 5:
        datos_final.append(fila)
        fila = []
    t += 1

# Crear un DataFrame de pandas
df = pd.DataFrame(datos_final, columns=titulos) 

# Guardar el DataFrame en un archivo Excel
df.to_excel(nombre_pdf +'.xlsx', index=False)
