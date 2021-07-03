"""
Sistema de Alertas Tempranas Monex México
Programa 2: Extracción de Modelos 20s

Este código tiene la finalidad de leer los modelos tanto financieros como corporativos
para extraer la información útil para próximos cálculos, se realiza una limpieza de datos
para homologar campos (se convierte a mayúsculas, se quitan espacios, guiones, 
convertir formatos, etc). Una vez terminado el proceso, los nuevos registros se adjuntan a 
archivos históricos para llevar un control de lo sucedido mes a mes.

El código se encuentra estructurado en 6 partes:
    1) Librerías necesarias para la ejecución.
    2) Carga de catálogos y archivos históricos.
    3) Conteo de los Modelos 20.
    4) Extracción de datos.
    5) Limpieza de datos.
    6) Guardado del nuevo histórico.

IMPORTANTE: Este programa solo correrá si previamente ya fueron validados los archivos
usando el programa 1.
"""

#################################################
#    1 Librerías necesarias para la ejecución   #
#################################################

import time                     # Librería para trabajar con el reloj de sistema. 
import tkinter as tk            # Librería para realizar un GUI.
from functools import partial   # Librería que permite interactuar con funciones.
import pathlib                  # Librería para interactuar con archivos de sistema.
from pathlib import Path        # Librería para interactuar con archivos de sistema.
from datetime import datetime   # Librería para manejo de Fechas.
import os                       # Librería para interacturar con el sistema operativo.
from os.path import abspath     # " "
from os import scandir, getcwd  # " "
import win32com.client          # Librería que permite interactuar con el SO Windows.
import numpy as np              # Librería para manipulación de datos numéricos.
import pandas as pd             # Librería para manipulación de datos en una estructura DataFrame.
import pyxlsb                   # Librería que auxilia la lectura de archivos '.xlsb'.
import sys                      # Librería para obtener información de sistema.
from xlrd import *              # Librería para lectura de archivos '.xls'
import csv                      # Librería para lectura de archivos '.csv'
import openpyxl                 # Librería para interactuar con libros de Excel.

def extraccion_modelos(label_result, ruta, anio, mes):

#################################################
#  2 Carga de catálogos y archivos históricos   #
#################################################

    
    Ruta_Catalogo_Rutas = ruta.get().replace("\\","/") # Se reemplaza '\' por '/'
    anio_ = int(anio.get()) # Se toma el valor de año.
    mes_ = int(mes.get())   # Se toma el valor de mes.
    fecha = datetime(anio_, mes_, 1) # Se genera una fecha con los valores de año y mes.

    # Se realiza la prueba para leer el archivo dados los inputs.
    try:
        Rutas_aux = pd.read_excel(Ruta_Catalogo_Rutas, sheet_name="RUTAS_FIJAS", skiprows=0)
    except:
        time.sleep(2)
        label_result.config(text="Ruta o archivo inexistente") # En caso de no poder leer el archivo
        # en la interfaz se mostrará el mensaje.

    # Si el archivo se leyó, se toma el valor de la posición [2,2], correspondiente a la ruta de modelos.
    Ruta_Catalogo_Modelos = Rutas_aux.iloc[2, 2]
    Ruta_Catalogo_Modelos = Ruta_Catalogo_Modelos.replace("\\","/") # Se reemplaza '\' por '/'

    # Si el archivo se leyó, se toma el valor de la posición [4,2], correspondiente a la ruta de tratamiento.
    Ruta_Output = Rutas_aux.iloc [4, 2]
    Ruta_Output = Ruta_Output.replace("\\","/") + "/" # Se reemplaza '\' por '/' y se agrega '/'

    #fecha = datetime(2021, 4, 1)

    if fecha.month == 1:
        mes = 12
        anio = fecha.year-1
    else:
        mes = fecha.month-1
        anio = fecha.year

    Ruta_Historico_Modelos20 = Ruta_Output+'2_Hist_Modelos_20_'+str(anio)+'_'+str(mes)+'.xlsx'

    # Se valida la existencia del catálogo de rutas y se lee la ruta de los modelos 20
    #if Path(Ruta_Catalogo_Rutas).is_file():
    try:
        Rutas = pd.read_excel(Ruta_Catalogo_Rutas, sheet_name = "RUTAS", skiprows = 0) # Se lee el archivo '4_Rutas_Archivos.xlsx' usando la librería pandas y se almacena en el objeto Rutas.
        Ruta_M20 = Rutas[(Rutas["Fecha"] == fecha) & (Rutas["Archivo"] == "Ruta_Modelos20")].iloc[0]["Ruta"] # Una vez teniendo el archivo, se realiza el filtro por el mes que
        # se desea analizar y se toma la ruta del archivo Ruta_Modelos20.
    except:
        label_result.config(text="Mes o año incorrecto", bg="red")    
    # Se lee y valida la existencia del catálogo de modelos 20, que es el archivo donde se direccionan las celdas de cada modelo.
    if Path(Ruta_Catalogo_Modelos).is_file():
        CAT_Corp = pd.read_excel(Ruta_Catalogo_Modelos, sheet_name = "CORPO", skiprows = 0) # Se lee la hoja "CORPO", hace referencia a los modelos corporativos.
        CAT_Fin = pd.read_excel(Ruta_Catalogo_Modelos, sheet_name = "FIN", skiprows = 0) # Se lee la hoja "FIN", hace referencia a los modelos financieros.
    
    # Se lee y valida la existencia del catálogo de modelos 20, que es el archivo donde se direccionan las celdas de cada modelo
    #if Path(Ruta_Historico_Modelos20).is_file():
    try:
        HIST_Corp = pd.read_excel(Ruta_Historico_Modelos20, sheet_name = "Hist_Corp", skiprows = 0) # Se lee la hoja "Hist_Corp", hace referencia a los modelos corporativos históricos.
        HIST_Fin = pd.read_excel(Ruta_Historico_Modelos20, sheet_name = "Hist_Fin", skiprows = 0) # Se lee la hoja "Hist_Fin", hace referencia a los modelos financieros históricos.
    except:
        pass

    #################################################
    #          3 Conteo de los Modelos 20           #
    #################################################

    """
    En esta sección se analizan los modelos de años anteriores para obervar si se cumple con los formatos
    necesarios para ejecutar pasos siguientes, de cada modelo se tomará su ruta, extensión ('.xlsx' y '.xlsm') y nombre.
    Los modelos se toman de la ruta especificada en el archivo '4_Rutas_Archivos.xlsx', que se almacenó en el objeto
    Ruta_M20.
    """

    # No se modifica el código

    #Función que extrae la ruta y nombre de todos los archivos en la ruta proporcionada.
    def Nombre_Archivos(ruta = getcwd()): # Por default se obtiene la ruta al que el sistema apunta.
        if os.path.isdir(ruta):
            return [abspath(archivo.path) for archivo in scandir(ruta) if archivo.is_file()]
        else:
            return []

    # No se modifica el código

    #Con la ruta completa de cada archivo se extrae el nombre, formato, ruta.
    Modelos_20 = Nombre_Archivos(Ruta_M20) # Modelos_20 es una lista.
    lista_rutas, lista_ext, lista_nombres = [], [], []

    for Modelo_20 in Modelos_20: # Se recorre cada modelo para dividirlos en diferentes listas
        ruta, ext = os.path.splitext(Modelo_20)
        nombre = Path(Modelo_20).stem
        lista_rutas.append(Modelo_20) # Se almacenan las rutas de cada archivo.
        lista_ext.append(ext)         # Se almacenan los formatos de cada archivo.
        lista_nombres.append(nombre)  # Se almacenan los nombres de cada archivo.

    # No se modifica el código

    #Se crea un DataFrame con las Rutas, Extensiones y Nombres de los modelos 20 como control.
    Control_df = pd.DataFrame(list(zip(lista_rutas,lista_ext,lista_nombres)), columns=['Ruta_Completa', 'Extensión','Nombre'])

    # Se valida que los formatos de los archivos sean aceptados
    Formatos_Validos = ['.xlsm', '.xlsx'] # Formatos válidos para leer, siempre será solo '.xlsx' y '.xlsm'
    Control_df["Estatus_Num"] = Control_df.Extensión.isin(Formatos_Validos)# + 0
    Control_df.loc[Control_df['Estatus_Num'] == 1 , 'Estatus'] = 'Formato aceptado'
    Control_df.loc[Control_df['Estatus_Num'] == 0 , 'Estatus'] = 'Formato No aceptado'

    Archivos_Aceptado = list((Control_df[Control_df['Estatus_Num'] != 0])['Ruta_Completa']) # Se toman los archivos que tengan el formato correcto.
    Archivos = list(Control_df['Ruta_Completa']) # Se toman todos los archivos de modelos.


    #################################################
    #             4 Extracción de datos             #
    #################################################

    """
    Se extrae toda la información de cada uno de los modelos 20. No se podrán cargar modelos 20 que:
        - No tengan las pestañas necesarias.
        - El formato sea distinto de '.xlsx' o '.xslm'
        - Estén protegidos con contraseña.
        - Tengan algún error de origen.
        - Este paso puede tardar en función del número de modelos 20.
    """

    # No se modifica el código

    #Se define la función que extrae los datos de cada uno de los Modelos 20
    def lectura_datos(Modelo_20_Excel, Catalogo):
        # La función busca el valor, dado un archivo de Excel, la hoja y celda
        Hoja, Celdas, Valores = Catalogo["Fuente"], Catalogo["Celda"], []
        i = 0
        for celda in Celdas:
            Valores.append(Modelo_20_Excel[Hoja[i]][celda].value)
            i = i + 1

        return Valores

    # Se extrae la información para cada uno de los archivos

    # Se crea un DataFrame vacío para corpo y financieros con las mismas columnas que el catálogo.
    df_Modelos_20_Corp = pd.DataFrame(columns = CAT_Corp["Variable"])
    df_Modelos_20_Fin = pd.DataFrame(columns = CAT_Fin["Variable"])

    Archivo_Corp, Archivo_Fin, Archivo_NA, Archivo_Incompleto = [], [], [], []

    # Se inicia un ciclo hasta el número de archivos en la ruta
    i = 0
    for Modelo_20 in Archivos:
        print(str(round(i/len(Archivos),3) * 100) + "%")
        
        # Se valida que el archivo tenga el formato permitido y se procede a abrir el archivo.
        if Modelo_20 in Archivos_Aceptado:
            Modelo_20_Excel = openpyxl.load_workbook(filename = Modelo_20, data_only = True)
            hojas = Modelo_20_Excel.sheetnames
            
            #En caso de tener el formato, se valida que tenga las pestañas necesarias según el modelo.
            if ("MODELO FINANCIERO" in hojas) and ("CARÁTULA" in hojas): # Modelo Corporativo.
                df_Modelos_20_Corp.loc[i] = lectura_datos(Modelo_20_Excel, CAT_Corp)
                Archivo_Corp.append(Modelo_20)
                
            elif ("ARRENDADORA" in hojas) and ("CARÁTULA" in hojas): # Modelo Financiero.
                df_Modelos_20_Fin.loc[i] = lectura_datos(Modelo_20_Excel, CAT_Fin)
                Archivo_Fin.append(Modelo_20)
                
            else:
                print("Documento incompleto:" + Modelo_20)
                Archivo_Incompleto.append(Modelo_20)
            
        else:
            print("Docuento no leido:" + Modelo_20)
            Archivo_NA.append(Modelo_20)

        i = i+1
        
    df_Modelos_20_Corp["Archivo"] = Archivo_Corp
    df_Modelos_20_Fin["Archivo"] = Archivo_Fin

    #################################################
    #              5 Limpieza de datos              #
    #################################################


    """
    Una vez leeidos los Modelos 20 del mes, se relizán las siguientes actividades:
        * Unión con la información histórica.
        * Quitar registros duplicados.
        * Limpiar campos importantes. Ejemplo: Quitar espacios en el RFC.
    """

    # No se modifica el código

    # Se agrega el registro de fecha a cada DataFrame.
    df_Modelos_20_Corp["Fecha_Ejecucion"] = fecha
    df_Modelos_20_Fin["Fecha_Ejecucion"] = fecha

    # Se unifica con la base historíca
    full_data_corpo = pd.concat([HIST_Corp, df_Modelos_20_Corp], ignore_index = True,sort = False) # Se unen los nuevos registros con los registros históricos de los modelos corporativos.
    full_data_fin = pd.concat([HIST_Fin, df_Modelos_20_Fin], ignore_index = True, sort = False) # Se unen los nuevos registros con los registros históricos de los modelos financieros.

    # Se eliminan registros duplicados (en caso de existir) y se reestablecen índices.
    full_data_corpo = full_data_corpo.drop_duplicates()
    full_data_corpo.reset_index(inplace = True, drop = True)
    full_data_fin = full_data_fin.drop_duplicates()
    full_data_fin.reset_index(inplace = True, drop = True)


    # No se modifica el código

    ## Modelo Corporativo
    # Se limpian las variables más mportantes de la base de corporativos: 
    # Mayusculas, quitar espacios, quitar guiones, convertir formatos, etc.

    # A la columna 'RFC' se le hacen los siguientes ajustes:
    full_data_corpo['RFC'] = full_data_corpo['RFC'].astype(str)          # Se convierte a un string.
    full_data_corpo['RFC'] = full_data_corpo['RFC'].str.replace(' ', '') # Se le quitan espacios.
    full_data_corpo['RFC'] = full_data_corpo['RFC'].str.replace('-', '') # Se reemplazan guiones por sin espacio.
    full_data_corpo['RFC'] = full_data_corpo['RFC'].str.replace('_', '') # Se reemplazan guiones bajos por sin espacio.
    full_data_corpo['RFC'] = full_data_corpo['RFC'].str.upper()          # Se convierte todo el texto a mayúsculas.


    # A la columna 'ACREDITADO' se le hacen los siguientes ajustes:
    full_data_corpo['ACREDITADO'] = full_data_corpo['ACREDITADO'].astype(str) # Se convierte a un string.
    full_data_corpo['ACREDITADO'] = full_data_corpo['ACREDITADO'].str.upper() # Se convierte todo el texto a mayúsculas.

    try:
        full_data_corpo['FECHA'] = pd.to_datetime(full_data_corpo['FECHA'], errors = "coerce") # Se intenta convertir 'FECHA' a un formato de fecha estandarizado.
    except:
        #print()
        pass

    full_data_corpo['EMPLEADOS'] = pd.to_numeric(full_data_corpo['EMPLEADOS'], errors = "coerce") # 'EMPLEADOS' se convierte a un valor numérico.
    full_data_corpo["ELABORO"] = full_data_corpo["ELABORO"].str.upper()       # Se convierte todo el texto a mayúsculas.
    full_data_corpo["PROMOTOR"] = full_data_corpo["PROMOTOR"].str.upper()     # Se convierte todo el texto a mayúsculas.
    full_data_corpo["REGIONAL"] = full_data_corpo["REGIONAL"].str.upper()     # Se convierte todo el texto a mayúsculas.
    #full_data_corpo["F_CONTITUCION"] = full_data_corpo["F_CONTITUCION"].str.upper()
    full_data_corpo["ACTIVIDAD"] = full_data_corpo["ACTIVIDAD"].str.upper()   # Se convierte todo el texto a mayúsculas.
    full_data_corpo["SECTOR"] = full_data_corpo["SECTOR"].str.upper()         # Se convierte todo el texto a mayúsculas.
    full_data_corpo["GRUPO"] = full_data_corpo["GRUPO"].str.upper()           # Se convierte todo el texto a mayúsculas.
    full_data_corpo["ART_73"] = full_data_corpo["ART_73"].str.upper()         # Se convierte todo el texto a mayúsculas.
    #full_data_corpo["CLIENTE DESDE"] = full_data_corpo["CLIENTE DESDE"].str.upper()
    full_data_corpo["MF_ACREDITADO"] = full_data_corpo["MF_ACREDITADO"].str.upper() # Se convierte todo el texto a mayúsculas.
    full_data_corpo["MF_GRUPO"] = full_data_corpo["MF_GRUPO"].str.upper()     # Se convierte todo el texto a mayúsculas.
    #full_data_corpo["MF_FECHA_ELABORACION"] = full_data_corpo["MF_FECHA_ELABORACION"].str.upper()
    full_data_corpo["MF_CIFRAS"] = full_data_corpo["MF_CIFRAS"].str.upper()   # Se convierte todo el texto a mayúsculas.


    # Las siguientes columnas se convierten a un formato de fecha estandarizado.
    full_data_corpo['FECHA_1'] = pd.to_datetime(full_data_corpo['FECHA_1'], errors="coerce")
    full_data_corpo['FECHA_2'] = pd.to_datetime(full_data_corpo['FECHA_2'], errors="coerce")
    full_data_corpo['FECHA_3'] = pd.to_datetime(full_data_corpo['FECHA_3'], errors="coerce")
    full_data_corpo['FECHA_4'] = pd.to_datetime(full_data_corpo['FECHA_4'], errors="coerce")
    full_data_corpo['FECHA_5'] = pd.to_datetime(full_data_corpo['FECHA_5'], errors="coerce")

    # En la columna 29 y posteriores todos los campos son númericos.
    j = 1
    n_columns = len(full_data_corpo.columns)
    for columna in full_data_corpo.columns:
        if j >= 29 and j <= n_columns - 2:
            full_data_corpo[columna] = pd.to_numeric(full_data_corpo[columna], errors="coerce") # Se convierte a un valor numérico.
        j += 1


    ## Modelo financiero

    # No se modifica el código

    # Se limpia las variables más mportantes de la base de corporativos: 
    # Mayusculas, quitar espacios, quitar guiones, convertir formatos, etc.


    # A la columna 'RFC' se le hacen los siguientes ajustes: Convertir a string, reemplazar '_' y '-' 
    # por espacios no vacíos y todo a mayúsculas.
    full_data_fin['RFC'] = full_data_fin['RFC'].astype(str)
    full_data_fin['RFC'] = full_data_fin['RFC'].str.replace(' ', '')
    full_data_fin['RFC'] = full_data_fin['RFC'].str.replace('-', '')
    full_data_fin['RFC'] = full_data_fin['RFC'].str.replace('_', '')
    full_data_fin['RFC'] = full_data_fin['RFC'].str.upper()

    full_data_fin['ACREDITADO'] = full_data_fin['ACREDITADO'].astype(str)
    full_data_fin['ACREDITADO'] = full_data_fin['ACREDITADO'].str.upper()


    try:
        full_data_fin['FECHA'] = pd.to_datetime(full_data_fin['FECHA'], errors="coerce") # Se intenta convertir 'FECHA' a un formato de fecha estandarizado.
    except:
        #print()
        pass

    full_data_fin['EMPLEADOS'] = pd.to_numeric(full_data_fin['EMPLEADOS'], errors="coerce")

    # Las siguientes columnas se convierten todo a mayúsculas.
    full_data_fin["ELABORO"] = full_data_fin["ELABORO"].str.upper()
    full_data_fin["PROMOTOR"] = full_data_fin["PROMOTOR"].str.upper()
    full_data_fin["REGIONAL"] = full_data_fin["REGIONAL"].str.upper()
    full_data_fin["F_CONTITUCION"] = full_data_fin["F_CONTITUCION"].str.upper()
    full_data_fin["ACTIVIDAD"] = full_data_fin["ACTIVIDAD"].str.upper()
    full_data_fin["SECTOR"] = full_data_fin["SECTOR"].str.upper()
    full_data_fin["GRUPO"] = full_data_fin["GRUPO"].str.upper()
    full_data_fin["ART_73"] = full_data_fin["ART_73"].str.upper()
    full_data_fin["CLIENTE DESDE"] = full_data_fin["CLIENTE DESDE"].str.upper()
    full_data_fin["MF_ACREDITADO"] = full_data_fin["MF_ACREDITADO"].str.upper()
    full_data_fin["MF_GRUPO"] = full_data_fin["MF_GRUPO"].str.upper()
    #full_data_fin["MF_FECHA_ELABORACION"] = full_data_fin["MF_FECHA_ELABORACION"].str.upper()
    full_data_fin["MF_CIFRAS"] = full_data_fin["MF_CIFRAS"].str.upper()

    # Las siguientes columnas se convierten a un formato de fecha estandarizado.
    full_data_fin['FECHA_1'] = pd.to_datetime(full_data_fin['FECHA_1'], errors="coerce")
    full_data_fin['FECHA_2'] = pd.to_datetime(full_data_fin['FECHA_2'], errors="coerce")
    full_data_fin['FECHA_3'] = pd.to_datetime(full_data_fin['FECHA_3'], errors="coerce")
    full_data_fin['FECHA_4'] = pd.to_datetime(full_data_fin['FECHA_4'], errors="coerce")
    full_data_fin['FECHA_5'] = pd.to_datetime(full_data_fin['FECHA_5'], errors="coerce")

    # En la columna 29 y posteriores todos los campos son númericos
    l = 1
    n_columns = len(full_data_fin.columns)
    for columna in full_data_fin.columns:
        if l >= 29 and l <= n_columns - 2:
            full_data_fin[columna] = pd.to_numeric(full_data_fin[columna], errors="coerce")
        l = l + 1


    #################################################
    #       6 Guardado del nuevo histórico          #
    #################################################

    #Ruta_Output = "C:/Users/52551/Desktop/Monex_Tratamiento/"
    # Se guardan los nuevos archivos históricos en un archivo con 2 hojas.
    with pd.ExcelWriter(Ruta_Output + '2_Hist_Modelos_20_'+str(fecha.year)+'_'+str(fecha.month)+'.xlsx') as writer:  
        full_data_corpo.to_excel(writer, sheet_name='Hist_Corp', index=False)
        full_data_fin.to_excel(writer, sheet_name='Hist_Fin', index=False)

    # Se guarda los listados de archivos que no se pudieron leer por hojas faltantes y/o formatos no validos
    (pd.DataFrame(Archivo_Incompleto)).to_excel(Ruta_Output + '2_1_Archivos_Incompletos_'+str(fecha.year)+'_'+str(fecha.month)+'.xlsx', sheet_name='Archivos_Incompletos', index=False)
    (pd.DataFrame(Archivo_NA)).to_excel(Ruta_Output + '2_2_Archivos_No_Leidos_'+str(fecha.year)+'_'+str(fecha.month)+'.xlsx', sheet_name='Archivos_No_Leidos', index=False)


    #print(full_data_fin.head() )
    time.sleep(1)
    label_result.config(text="Terminado")