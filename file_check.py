""" 
Sistema de Alertas Tempranas Monex México
Programa 1: Validación de Archivos

Este código tiene la finalidad de realizar una revisión de los archivos propuestos en el
libro de excel '4_Rutas_Archivos.xlsx' para la correcta implementación del modelo MONEX. 
Se examinan las rutas de alojamiento, existencia y nombre de los mismos, además del formato y
las variables que hacen parte del cálculo, obteniendo como output la descripción de si el
archivo está útil para siguientes pasos ó encaso los errores en contrados de la inspección. 

El código se encuentra estructurado en 4 partes:
    1) Librerías necesarias para la ejecución.
    2) Carga de catálogos.
    3) Revisión de archivos.
    4) Output.

IMPORTANTE: En caso de obtener algún archivo con e mínimo de errores, no será posible continuar
con el resto de scripts para el cálculo del modelo.
"""



#################################################
#    1 Librerías necesarias para la ejecución   #
#################################################

from pathlib import Path       # Librería para interactuar con archivos de sistema.
from datetime import datetime  # Librería para manejo de Fechas.
import os                      # Librería para interacturar con el sistema operativo.
from os.path import abspath    # " "
from os import scandir, getcwd # " "
import win32com.client         # Librería que permite interactuar con el SO Windows.
import numpy as np             # Librería para manipulación de datos numéricos.
import pandas as pd            # Librería para manipulación de datos en una estructura DataFrame.
import pyxlsb                  # Librería que auxilia la lectura de archivos '.xlsb'.


#################################################
#             2 Carga de catálogos              #
#################################################

"""
A continuación se debe proporcionar la ruta del catálogo de archivos que 
serán evaluados para los siguientes procesos. Importante colocar la ruta
específica donde se encuentra el archivo almacenado dentro de la pc.

Adicional, se debe ingresar la fecha para el cual se de desea hacer el cálculo.
Es importante que se tenga la siguiente estructura: (yyyy, m, 1), ejemplo: (2021, 3 , 1).
"""

## Solo se debe de modificar las rutas y la fecha

# Ingresar ruta de catálogo
Ruta_Catalogo_Rutas = "C:/Users/52551/Documents/4_Rutas_Archivos.xlsx"
#Ruta_Catalogo_Campos = "C:/Users/52551/Desktop/Modelo_MONEX/2_Catalogo_Layouts.xlsx"
#Ruta_Output = "C:/Users/52551/Desktop/"

# Colocar fecha del mes de cálculo, (yyyy, m, 1) .
fecha = datetime(2021, 4, 1) 

## No se modifica el código
# Se leerán las rutas del archivo 4_Rutas_Archivos.xlsx, haciendo modificación en ellas.

Rutas_aux = pd.read_excel(Ruta_Catalogo_Rutas, sheet_name="RUTAS_FIJAS", skiprows=0)

# Se reemplazan "\" por "/"
Ruta_Catalogo_Campos = Rutas_aux.iloc [0, 2]
Ruta_Catalogo_Campos = Ruta_Catalogo_Campos.replace("\\","/")

Ruta_Output = Rutas_aux.iloc [4, 2]
Ruta_Output = Ruta_Output.replace("\\","/") + "/"

# Se definen las listas a utilizar, estas guardaran valores de acuerdo a la ejecución del siguiente código
Resultado = []

# Se valida la existencia del catálogo de rutas
if Path(Ruta_Catalogo_Rutas).is_file():
    try:
        # Se lee el archivo '4_Rutas_Archivos.xlsx' usando la librería pandas y se almacena en el objeto Rutas.
        Rutas = pd.read_excel(Ruta_Catalogo_Rutas, sheet_name = "RUTAS", skiprows = 0)
        Rutas_2 = Rutas[Rutas["Fecha"] == fecha] # Una vez teniendo el archivo, se realiza el filtro por el mes que se desea analizar.
        Rutas_2.reset_index(inplace = True, drop = True) # Se retiran índices de los renglones del DataFrame.

        # Se extrae en forma de listas las variables necesarias del catálogo de rutas por cada objeto listado a continuación:
        Rutas_Archivos = Rutas_2[Rutas_2["Fecha"] == fecha]["Ruta"] #Rutas_2["Ruta"].tolist()
        IDs_Archivos = Rutas_2[Rutas_2["Fecha"] == fecha]["ID"]
        Hojas_Archivos = Rutas_2[Rutas_2["Fecha"] == fecha]["Hoja"]
        Pwd_Archivos = Rutas_2[Rutas_2["Fecha"] == fecha]["Contraseña"]
        Salto_Filas = Rutas_2[Rutas_2["Fecha"] == fecha]["SkipRows"]
        Ext_Archivos = [os.path.splitext(archivo)[1] for archivo in Rutas_Archivos] # Se obtienen extensiones de los archivos.
        
        #if len(Rutas_Archivos) != 12: # si no se tienen los 12 archivos necesarios el programa imprimirá el siguiente error:
        #    print(f"Error en el número de archivos con la fecha: {fecha}" )
    except:
        Resultado.append("Error al abrir el archivo de rutas") #### Aquí hay observación
else: # En caso de que no se encuentre el catálogo, se imprimirá lo siguiente:
    pass
    #print("Error al leer el catalogo de Rutas")  


# Se valida la existencia del catálogo de campos
Hoja_Catalogo_Campos = ["INFLINDISCREDITO", "CARTAS_CREDITO","BASE_CLIENTES", "REP_VENCIDOS", "GRUPOS_RIESGO", "BASE_INSUMOS", "MODELO_CALIFICACION", "RFC","GARANTIAS","CALIFICA","WATCH","MODELOS20"]
if Path(Ruta_Catalogo_Campos).is_file():
    print("")
#else:
    #print("Error al leer el Catalogo de Campos")



#################################################
#             3 Revisión de archivos            #
#################################################

"""
Se valida que los archivos en el catálogo de rutas cumplan:
    1. Ruta: Se valida que la ruta exista.
    2. Nombre: Se valida que el archivo exista.
    3. Formato: Se valida que el formato se pueda leer.
    4. Pestaña: Se valida que la pestaña (en caso de ser Excel) exista.
    5. Layout: Se valida que las variables contenidas en los archivos coincidan con el Layout del catálogo de campos.
"""

## No se modifica el código

# Se definen las listas a utilizar, estas guardaran valores de acuerdo a la ejecución del siguiente código
Resultado = []
Ruta_Califica_AUX, Ruta_Califica = [], []
Ruta_Modelos20_AUX, Ruta_Modelos20 = [], []

# Se inicializá un ciclo for desde 0 hasta 11, lo que se hará es ir recorriendo cada una de las rutas y
# archivos que se tienen en el catálogo.

for i in range(12):
    # Se carga el layout del archivo a analizar.
    Campos = pd.read_excel(Ruta_Catalogo_Campos, sheet_name = Hoja_Catalogo_Campos[i], skiprows = 0)
    Data_Frame_Aux = pd.DataFrame() # Se crea un DataFrame vacío.
    
    # El ID = 6 corresponde a la Base_Insumos, que por estar protegido con contraseña necesita otro tratamiento
    # tal cual se muestra a continuación:
    if IDs_Archivos[i] == 6: # Primero se valida si el archivo existe y si se encuentra en alguno de los formatos '.xlsx' o '.xlsm'
        if Path(Rutas_Archivos[i]).is_file() and ((Ext_Archivos[i] == '.xlsx') or (Ext_Archivos[i]=='.xlsm')):
            try:
                xlApp = win32com.client.Dispatch("Excel.Application") # Se interactua con el sistema para informar que se desea acceder a una aplicación de excel.
                xlwb = xlApp.Workbooks.Open(Rutas_Archivos[i],False, True, None, Password=Pwd_Archivos[i]) # Se realiza la apertura del archivo usando la contraseña proporcionada.
                Base_Insumos = xlwb.Sheets(Hojas_Archivos[i]) # Del archivo abierto, se toma la hoja 'Insumos'.
                Base_Insumos_Columnas = Base_Insumos.Range(Base_Insumos.Cells(Salto_Filas[i], 1), Base_Insumos.Cells(Salto_Filas[i], 200)).value # Se toman los valores de la hoja necesarios para el cálculo.
                Base_Insumos_Aux = np.transpose(np.matrix(Base_Insumos_Columnas)) 
                Nombre_Columnas_BI = [Base_Insumos_Aux[i,0] for i in range(len(Base_Insumos_Aux[:,0]))]
                # A continuación se recorren los nombres de las variables para identificar alguna faltante.
                Variable_Error = [variable for variable in Campos if variable not in Nombre_Columnas_BI]
                if len(Variable_Error) >= 1: # En caso de faltar alguna variable se obtendrá el siguiente mensaje y será almacenado en la lista Resultado.
                    Resultado.append("Se encontró el archivo y la pestaña indicada, pero no se encontraron las variables: " + str(Variable_Error))
                else: # En caso de estar todo correcto se obtiene el siguiente resultado
                    Resultado.append("Se encontró el archivo, la pestaña y las variables necesarias")
            except:
                Resultado.append("Error al abrir el archivo") # Se obtiene error por no poder abrir archivo.
        else: 
            Resultado.append("Error en la ruta y formato proporcionado") # Este mensaje se obtiene al no tener la ruta y/o formato correctamente.

    # El ID = 10 corresponde a la Ruta_Califica, que se espera sea una ruta y el código leerá cada uno de los archivos
    # en formato '.cvs'.
    elif IDs_Archivos[i] == 10:
        if os.path.isdir(Rutas_Archivos[i]): # Se valida que la ruta exista, en caso contrario se obtendrá un mensaje de error.
            Ruta_Califica_AUX = [abspath(arch.path) for arch in scandir(Rutas_Archivos[i]) if arch.is_file()] # Se leen cada una de las rutas.
            Ruta_Califica = [arch for arch in Ruta_Califica_AUX if os.path.splitext(arch)[1] == '.csv'] # Se lee cada archivo en '.csv'
            Resultado.append(f"Se leyó la ruta con: {len(Ruta_Califica)} archivos con formato csv")
        else: 
            Resultado.append("Error en la ruta proporcionada")
    
    # El ID = 12 corresponde a la Ruta_Modelos20 y se espera sea una ruta, el código contabiliza 
    # el número de archivos Excel (.xlsm o .xlsx)
    elif IDs_Archivos[i] == 12:
        if os.path.isdir(Rutas_Archivos[i]): # Se valida que la ruta exista, en caso contrario se obtendrá un mensaje de error.
            Ruta_Modelos20_AUX = [abspath(arch.path) for arch in scandir(Rutas_Archivos[i]) if arch.is_file()] # Se leen cada una de las rutas.
            Ruta_Modelos20 = [arch for arch in Ruta_Modelos20_AUX if os.path.splitext(arch)[1]=='.xlsx' or os.path.splitext(arch)[1]=='.xlsm'] # Se lee cada archivo en formato '.xlsx' o '.xlsm'
            Resultado.append(f"Se leyó la ruta con: {len(Ruta_Modelos20)} archivos con formato xlsx o xlsm")
        else: 
            Resultado.append("Error en la ruta proporcionada")
    
    # El resto de los archivos se podrán leer en cualquier formato Excel ('.xlsx' o '.xlsm'), 
    # formato binario ('.xlsb') y en '.csv'. En caso de que los archivos existan se validará el layout.
    else:
        if Path(Rutas_Archivos[i]).is_file(): # Se valida que la ruta exista, en caso contrario se obtendrá un mensaje de error.
            if Ext_Archivos[i] == '.xlsm' or Ext_Archivos[i] == '.xlsx': # Si la ruta existe, entonces se identifica si el formato es '.xlsx' o '.xlsm'.
                try:
                    Data_Frame_Aux = pd.read_excel(Rutas_Archivos[i], sheet_name=Hojas_Archivos[i], skiprows=Salto_Filas[i]) # El archivo se lee y se almacena como DataFrame.
                    Variable_Error = [variable for variable in Campos if variable not in Data_Frame_Aux.columns] # Se identifican las variables en el archivo.
                    if len(Variable_Error) >= 1: # En caso de faltar alguna variable se obtendrá el siguiente mensaje y será almacenado en la lista Resultado.
                        Resultado.append(f"Se encontró el archivo y la pestaña indicada, pero no se encontraron las variables:  {Variable_Error}")
                    else: 
                        Resultado.append("Se encontró el archivo, la pestaña y las variables necesarias")
                except:
                    Resultado.append("Se encontró el archivo, pero no la pestaña indicada") # Se obtiene error por no poder abrir archivo.
            
            elif Ext_Archivos[i] == '.xlsb': # Si la ruta existe, entonces se identifica si el formato es '.xlsb'.
                try:
                    Data_Frame_Aux = pd.read_excel(Rutas_Archivos[i], sheet_name=Hojas_Archivos[i], engine='pyxlsb', skiprows=Salto_Filas[i]) # El archivo se lee y se almacena como DataFrame.
                    Variable_Error = [variable for variable in Campos if variable not in Data_Frame_Aux.columns] # Se identifican las variables en el archivo.
                    if len(Variable_Error) >= 1: # En caso de faltar alguna variable se obtendrá el siguiente mensaje y será almacenado en la lista Resultado.
                        Resultado.append(f"Se encontró el archivo y la pestaña indicada, pero no se encontraron las variables:  {Variable_Error}")
                    else: 
                        Resultado.append("Se encontró el archivo, la pestaña y las variables necesarias")
                except:
                    Resultado.append("Se encontró el archivo, pero no la pestaña indicada") # Se obtiene error por no poder abrir archivo.

            elif Ext_Archivos[i] == '.csv': # Si la ruta existe, entonces se identifica si el formato es '.csv'.
                try:
                    Data_Frame_Aux = pd.read_csv(Rutas_Archivos[i], encoding='latin-1', delimiter=",",skiprows=Salto_Filas[i]) # El archivo se lee y se almacena como DataFrame.
                    Variable_Error = [variable for variable in Campos if variable not in Data_Frame_Aux.columns] # Se identifican las variables en el archivo.
                    if len(Variable_Error) >= 1: # En caso de faltar alguna variable se obtendrá el siguiente mensaje y será almacenado en la lista Resultado.
                        Resultado.append(f"Se encontró el archivo y la pestaña indicada, pero no se encontraron las variables:  {Variable_Error}")
                    else: 
                        Resultado.append("Se encontró el archivo, la pestaña y las variables necesarias")
                except:
                    Resultado.append("Se encontró el archivo, pero no la pestaña indicada") # Se obtiene error por no poder abrir archivo.
            else:
                Resultado.append("Error en el formato, corroborar que sea: xlsx, xlsm, xlsb o csv") # Error obtenido al no tener el formato deseado.
                
        else:
            Resultado.append("Error en la ruta proporcionada")



#################################################
#             4 Guardar Output                  #
#################################################

"""
Se guarda los comentarios respecto de los archivos, en caso de encontrar 
algún error se deberá de corregirlo antes de proceder al siguiente código.
"""

Archivo_Out = Rutas_2[Rutas_2["Fecha"]==fecha]
# Se realiza un Join entre el DataFrame Archivo_Out y los resultados de la revisión de los archivos.
Archivo_Out_1 = pd.merge(Archivo_Out, pd.DataFrame(Resultado), left_index=True, right_index=True)
# Se guarda en un '.xlsx' de acuerdo a la ruta especificada.
Archivo_Out_1.to_excel(Ruta_Output + '1_Rutas_Archivos_Output_'+str(fecha.year)+'_'+str(fecha.month)+'.xlsx', sheet_name='RUtAS')