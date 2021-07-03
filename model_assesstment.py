""" 
Sistema de Alertas Tempranas Monex México
Programa 3: Extraccion de Insumos y Generación de Modelo

Este código tiene la finalidad de realizar una revisión de los archivos propuestos en el
libro de excel '4_Rutas_Archivos.xlsx' para la correcta implementación del modelo MONEX. 
Se examinan las rutas de alojamiento, existencia y nombre de los mismos, además del formato y
las variables que hacen parte del cálculo, obteniendo como output la descripción de si el
archivo está útil para siguientes pasos ó encaso los errores en contrados de la inspección. 

El código se encuentra estructurado en 4 partes:
    1) Librerías necesarias para la ejecución.
    2) Carga de catálogos.
    3) Extracción de archivos.
    4) Limpieza de archivos.
    5) Cruce de Tablas.
    6) Evaluación del modelo SMART.
        6.1) Variables del Modelo.
        6.2) Ejecución del Modelo.
    7) Resumen Reporte.

IMPORTANTE: En caso de obtener algún archivo con e mínimo de errores, no será posible continuar
con el resto de scripts para el cálculo del modelo.
"""


#################################################
#    1 Librerías necesarias para la ejecución   #
#################################################


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



#################################################
#               2 Carga de catálogos            #
#################################################


# Sólo se debe de modificar las rutas y la fecha

Ruta_Catalogo_Rutas = "C:/Users/52551/Documents/4_Rutas_Archivos.xlsx"
#Ruta_Catalogo_Campos = "C:/Users/52551/Desktop/Modelo_MONEX/2_Catalogo_Layouts.xlsx"
#Ruta_Catalogo_Campos_Rename = "C:/Users/52551/Desktop/Modelo_MONEX/2_Catalogo_Layouts_Rename.xlsx"
#Ruta_Output= "C:/Users/52551/Desktop/"

fecha = datetime(2021, 4, 1)

Rutas_aux = pd.read_excel(Ruta_Catalogo_Rutas, sheet_name="RUTAS_FIJAS", skiprows=0)

# Se reemplazan "\" por "/"
Ruta_Catalogo_Campos = Rutas_aux.iloc [0, 2]
Ruta_Catalogo_Campos = Ruta_Catalogo_Campos.replace("\\","/")

Ruta_Catalogo_Campos_Rename = Rutas_aux.iloc [1, 2]
Ruta_Catalogo_Campos_Rename = Ruta_Catalogo_Campos_Rename.replace("\\","/")

Ruta_Output = Rutas_aux.iloc [4, 2]
Ruta_Output = Ruta_Output.replace("\\","/") + "/"


Rutas_aux_tc = pd.read_excel(Ruta_Catalogo_Rutas, sheet_name="TC", skiprows=0)
TC = Rutas_aux_tc[(Rutas_aux_tc['Anio']==fecha.year) & (Rutas_aux_tc['Mes']==fecha.month) & (Rutas_aux_tc['Dia']==1)].iloc[0,3]


# No se modifica el código

# Se valida la existencia del catálogo de rutas
if Path(Ruta_Catalogo_Rutas).is_file():

    # Se lee el archivo '4_Rutas_Archivos.xlsx' usando la librería pandas y se almacena en el objeto Rutas.
    Rutas_0 = pd.read_excel(Ruta_Catalogo_Rutas, sheet_name = "RUTAS", skiprows = 0)
    Rutas = Rutas_0[Rutas_0["Fecha"] == fecha] # Una vez teniendo el archivo, se realiza el filtro por el mes que se desea analizar.
    Rutas.reset_index(inplace = True, drop = True) # Se retiran índices de los renglones del DataFrame.    

    
    # Se extrae en forma de listas las variables necesarias del catálogo de rutas por cada objeto listado a continuación:
    Rutas_Archivos = Rutas[Rutas["Fecha"] == fecha]["Ruta"]
    IDs_Archivos = Rutas[Rutas["Fecha"] == fecha]["ID"]
    Hojas_Archivos = Rutas[Rutas["Fecha"] == fecha]["Hoja"]
    Pwd_Archivos = Rutas[Rutas["Fecha"] == fecha]["Contraseña"]
    Salto_Filas = Rutas[Rutas["Fecha"] == fecha]["SkipRows"]
    Ext_Archivos = [os.path.splitext(archivo)[1] for archivo in Rutas_Archivos]

    if len(Rutas_Archivos) != 12: # si no se tienen los 12 archivos necesarios el programa imprimirá el siguiente error:
        print(f"Error en el número de archivos con la fecha: {fecha}" )

else:
    print("Error al leer el catalogo de Rutas")

# Se valida la existencia del catálogo de campos
Hoja_Catalogo_Campos=["INFLINDISCREDITO", "CARTAS_CREDITO","BASE_CLIENTES", "REP_VENCIDOS", "GRUPOS_RIESGO", "BASE_INSUMOS", "MODELO_CALIFICACION", "RFC","GARANTIAS","CALIFICA","WATCH","MODELOS20"]
#if Path(Ruta_Catalogo_Campos).is_file():
#    print("")
#else:
#    print("Error al leer el Catalogo de Campos")



#################################################
#           3 Extracción de archivos            #
#################################################


# No se modifica el código

# Función que permite leer cualquier archivo Excel ('.xlsm' o '.xlsx'), Binario ('.xlsb') o '.csv'. 
# Además valida nuevamente que todos los campos coincidan con el Layout esperado, tiene como parámetro
# el índice de la hoja que se desea leer del archivo '2_Catalogo_Layouts.xlsx'.
def extraccion_archivos(i):
    # Se carga el catálogo (Layout) de variables esperadas
    Campos = pd.read_excel(Ruta_Catalogo_Campos, sheet_name = Hoja_Catalogo_Campos[i], skiprows=0)
    Data_Frame_Aux = pd.DataFrame() # Se crea un DataFrame vacío.
    
    # Se valida que el archivo exista el archivo
    if Path(Rutas_Archivos[i]).is_file(): # Se valida que la ruta exista, en caso contrario se obtendrá un mensaje de error.
        # Se valida el formato
        if Ext_Archivos[i] == '.xlsm' or Ext_Archivos[i] == '.xlsx': # Si la ruta existe, entonces se identifica si el formato es '.xlsx' o '.xlsm'.
            # En caso de ser Excel (.xlsm o .xlsx) se utilizá el siguiente código 
            try:
                Data_Frame_Aux = pd.read_excel(Rutas_Archivos[i], sheet_name = Hojas_Archivos[i], skiprows=Salto_Filas[i]) # El archivo se lee y se almacena como DataFrame.
                Variable_Error = [variable for variable in Campos if variable not in Data_Frame_Aux.columns] # Se identifican las variables en el archivo.
                if len(Variable_Error) == 0: # En caso de no faltar ninguna variable será almacenado el auxiliar en un DataFrame final.
                    Data_Frame_final = Data_Frame_Aux
                    # print("Se leyó correctamente el archivo: " + Rutas_Archivos[i])
                else:
                    #print("Se encontró el archivo " + Rutas_Archivos[i] + " y la pestaña indicada, pero no se encontraron las variables: " + str(Variable_Error))
                    Data_Frame_final = pd.DataFrame() # En caso de alguna falla, se devuelve un DF vacío
            except:
                #Resultado.append("Se encontró el archivo, pero no la pestaña indicada")
                Data_Frame_final = pd.DataFrame()
                
        elif Ext_Archivos[i] == '.xlsb':
            # En caso de ser Binario (.xlsb) se utilizá el siguiente código 
            try:
                Data_Frame_Aux = pd.read_excel(Rutas_Archivos[i], sheet_name=Hojas_Archivos[i], engine='pyxlsb', skiprows=Salto_Filas[i]) # El archivo se lee y se almacena como DataFrame.
                Variable_Error = [variable for variable in Campos if variable not in Data_Frame_Aux.columns] # Se identifican las variables en el archivo.
                if len(Variable_Error) == 0: # En caso de no faltar ninguna variable será almacenado el auxiliar en un DataFrame final.
                    Data_Frame_final = Data_Frame_Aux
                    #print("Se leyó correctamente el archivo: " + Rutas_Archivos[i])
                else:
                    #print("Se encontró el archivo " + Rutas_Archivos[i] + " y la pestaña indicada, pero no se encontraron las variables: " + str(Variable_Error))
                    Data_Frame_final = pd.DataFrame() # En caso de alguna falla, se devuelve un DF vacío
            except:
                #Resultado.append("Se encontró el archivo, pero no la pestaña indicada")
                Data_Frame_final = pd.DataFrame() # En caso de alguna falla, se devuelve un DF vacío
                
        elif Ext_Archivos[i] == '.csv':
            # En caso de ser ,csv se utilizá el siguiente código 
            try:
                Data_Frame_Aux = pd.read_csv(Rutas_Archivos[i], encoding='latin-1', delimiter=",",skiprows=Salto_Filas[i]) # El archivo se lee y se almacena como DataFrame.
                Variable_Error = [variable for variable in Campos if variable not in Data_Frame_Aux.columns] # Se identifican las variables en el archivo.
                if len(Variable_Error) == 0: # En caso de no faltar ninguna variable será almacenado el auxiliar en un DataFrame final.
                    Data_Frame_final = Data_Frame_Aux
                    # print("Se leyó correctamente el archivo: " + Rutas_Archivos[i])
                else:
                    #print("Se encontró el archivo " + Rutas_Archivos[i] + ", pero no se encontraron las variables: " + str(Variable_Error))
                    Data_Frame_final = pd.DataFrame() # En caso de alguna falla, se devuelve un DF vacío
            except:
                #Resultado.append("Se encontró el archivo, pero no la pestaña indicada")
                Data_Frame_final = pd.DataFrame() # En caso de alguna falla, se devuelve un DF vacío
        else:
            #Resultado.append("Error en el formato, corroborar que sea: xlsx, xlsm, xlsb o csv")
            Data_Frame_final = pd.DataFrame()

    else:
        #Resultado.append("Error en la ruta proporcionada")
        Data_Frame_final = pd.DataFrame() # En caso de alguna falla, se devuelve un DF vacío
    
    Campos_rename = pd.read_excel(Ruta_Catalogo_Campos_Rename, sheet_name=Hoja_Catalogo_Campos[i], skiprows=0) # Se lee el archivo '2_Catalogo_Layouts_Rename.xlsx'.
    Data_Frame_final_2 = Data_Frame_final[Campos.columns] # Se realiza una compia del DF Final con las columnas de la hoja leída
    Data_Frame_final_2.columns = Campos_rename.columns # Se renombran las columnas.
    return Data_Frame_final_2 # Se retorna el DataFrame

# Se evaluá la función de extracción para todos los archivos necesarios 
# excepto: Base_Insumos, Califica y Modelos 20, ya que necesitan un tratamiento especial.
# Todos los objetos de abajo son DataFrames que llaman a la función extraccion_archivos.
InfLinDisCredito = extraccion_archivos(0)
Cartas_Credito = extraccion_archivos(1)
Sucursales = extraccion_archivos(2)
Reporte_Vencidos = extraccion_archivos(3)
Grupo_Riesgo = extraccion_archivos(4)
#Base_Insumos=extraccion_archivos(5)
Modelo_Calif = extraccion_archivos(6)
RFC = extraccion_archivos(7)
Garantias = extraccion_archivos(8)
#Califica=extraccion_archivos(9)
Watch = extraccion_archivos(10)
#Hist_Corpo=extraccion_archivos(11)
#Hist_Fin=extraccion_archivos(11)


# El ID = 6 (i=5) corresponde a la Base_Insumos, 
# que por estar protegido con contraseña necesita otro tratamiento
i = 5
Campos = pd.read_excel(Ruta_Catalogo_Campos, sheet_name = Hoja_Catalogo_Campos[i], skiprows=0) 
# Primero se valida si el archivo existe y si se encuentra en alguno de los formatos '.xlsx' o '.xlsm'
if Path(Rutas_Archivos[i]).is_file() and ((Ext_Archivos[i] == '.xlsx') or (Ext_Archivos[i]=='.xlsm')):
    try:       
        xlApp = win32com.client.Dispatch("Excel.Application") # Se interactua con el sistema para informar que se desea acceder a una aplicación de excel.
        xlwb = xlApp.Workbooks.Open(Rutas_Archivos[i],False, True, None, Password=Pwd_Archivos[i]) # Se realiza la apertura del archivo usando la contraseña proporcionada.
        Base_Insumos = xlwb.Sheets(Hojas_Archivos[i]) # Del archivo abierto, se toma la hoja 'Insumos'.
        Base_Insumos_Columnas = Base_Insumos.Range(Base_Insumos.Cells(Salto_Filas[i], 1), Base_Insumos.Cells(Salto_Filas[i], 200)).value # Se toman los valores de la hoja necesarios para el cálculo.
        Base_Insumos_Aux = np.transpose(np.matrix(Base_Insumos_Columnas))
        
        Nombre_Columnas_BI = [Base_Insumos_Aux[j,0] for j in range(len(Base_Insumos_Aux[:,0]))]
        Base_Insumos_Extraccion = Base_Insumos.Range(Base_Insumos.Cells(4, 1), Base_Insumos.Cells(6000, 200)).value # Se toman algunos valores de la Base de Insumos.
        Variable_Error=[variable for variable in Campos if variable not in Nombre_Columnas_BI] # A continuación se recorren los nombres de las variables para identificar alguna faltante.

        if len(Variable_Error) == 0: # En caso de no faltar ninguna variable será almacenado en un nuevo DataFrame llamada Base_Insumos_Limpia.
            Base_Insumos_Limpia = pd.DataFrame(Base_Insumos_Extraccion, columns = Nombre_Columnas_BI)
            #print("Se leyó correctamente el archivo: " + Rutas_Archivos[i])
        else:
            #print("Se encontró el archivo y la pestaña indicada, pero no se encontraron las variables: " + str(Variable_Error))
            Base_Insumos_Limpia = pd.DataFrame() # En caso de algún fallo se tendrá un DF vacío
    except:
        pass
        #print("Error al abrir el archivo")
#else: 
#    print("Error en la ruta y formato proporcionado")

Campos_rename = pd.read_excel(Ruta_Catalogo_Campos_Rename, sheet_name=Hoja_Catalogo_Campos[i], skiprows=0)
Base_Insumos_Limpia_2 = Base_Insumos_Limpia[Campos.columns] # Se realiza una compia del DF Base_Insumos_Limpia con las columnas de la hoja leída
Base_Insumos_Limpia_2.columns = Campos_rename.columns # Se renombran las columnas.



# El ID = 10  (i=9) corresponde a la Ruta_Califica, que se espera sea una ruta y 
# el código leerá cada uno de los archivos '.csv'.
i  =9
full_data_CALIFICA = pd.DataFrame() # Se crea un DF vacío

if os.path.isdir(Rutas_Archivos[i]): # Primero se valida si el archivo existe
    Ruta_Califica_AUX = [abspath(arch.path) for arch in scandir(Rutas_Archivos[i]) if arch.is_file()] # Se leen cada una de las rutas.
    Ruta_Califica = [arch for arch in Ruta_Califica_AUX if os.path.splitext(arch)[1]=='.csv'] # Se lee cada archivo en '.csv' y se guarda la ruta en una lista.
    #print("Se leyó la ruta con: " + str(len(Ruta_Califica))+" archivos con formato csv")
    for Ruta in Ruta_Califica: 
        if Path(Ruta).is_file():
            CALIFICA = pd.read_csv(Ruta, encoding='latin-1', delimiter="|") # Se lee cada archivo para generar un único llamado full_data_CALIFICA
            full_data_CALIFICA = pd.concat([full_data_CALIFICA, CALIFICA], ignore_index=True,sort=False)
        #else:
        #    print("Error al leer el Califica")
#else: 
#    print("Error en la ruta proporcionada")

# Se lee el arcchivo 2_Hist_Modelos_20 ya generado previamente, se leen los modelos 
# financieros y corporativos.
ruta_M20 = Ruta_Output + '2_Hist_Modelos_20_'+str(fecha.year)+'_'+str(fecha.month)+'.xlsx'
if Path(ruta_M20).is_file():
    Hist_Corpo = pd.read_excel(ruta_M20, sheet_name = "Hist_Corp", skiprows=0)
    Hist_Fin = pd.read_excel(ruta_M20, sheet_name = "Hist_Fin", skiprows=0)

#print(full_data_CALIFICA.head() )
#print(Hist_Corpo.info())
#print(Hist_Fin.head())


#################################################
#              4 Limpieza de datos              #
#################################################

"""
Una vez cargados todos los archivos, se procede a realizar la limpieza, que en general 
consiste en lo siguiente:

- Quitar registros duplicados.
- Limpieza de los campos prioritarios, Ejemplo: Quitar espacios en el RFC.
- Creación de IDs.
- Reemplazo de valores incorrectos.
- Revisión de formato.
"""

# Cartas_Credito
Cartas_Credito = Cartas_Credito.dropna(how='all') # Se quitan registros vacíos del DataFrame Cartas_Credito.
Cartas_Credito['CC_STATUS'] = Cartas_Credito['CC_STATUS'].str.upper() # Se convierte texto a mayúscula.
Cartas_Credito.loc[Cartas_Credito['CC_STATUS'] == "VIGENTE",'CC_STATUS'] = "Vigente"
Cartas_Credito.loc[Cartas_Credito['CC_STATUS'] == "VENCIDA",'CC_STATUS'] = "Vencida"

# RFC
RFC['RFC'] = RFC['RFC'].astype(str) # Se convierte a string.
RFC['RFC'] = RFC['RFC'].str.replace(' ', '') # Se quitan espacios.
RFC['RFC'] = RFC['RFC'].str.replace('-', '') # Se cambian guiones por sin espacio.
RFC['RFC'] = RFC['RFC'].str.replace('_', '') # Se cambian guiones bajos por sin espacio.
RFC['RFC'] = RFC['RFC'].str.upper() # Se convierte texto a mayúscula.


#Base_Insumos_Limpia_3 = pd.DataFrame(Base_Insumos_Limpia_2)

Base_Insumos_Limpia_3 = Base_Insumos_Limpia_2.copy() # Se tiene una copia de la base de insumos.
try:
    #Base_Insumos_Limpia_3['BI_EEFF'] = pd.to_datetime(Base_Insumos_Limpia_3['BI_EEFF'], errors="coerce", utc=True)
    Base_Insumos_Limpia_3["BI_EEFF"] = Base_Insumos_Limpia_3["BI_EEFF"].dt.tz_convert(None) # Se transforma fecha a tipo timezone.
except:
    pass
Base_Insumos_Limpia_3 = Base_Insumos_Limpia_3.dropna(how='all') # Se quitan registros vacíos del DataFrame.
Base_Insumos_Limpia_3 = Base_Insumos_Limpia_3.dropna(subset=["BI_ID"]) # Se quitan registros vacíos de la columna BI_ID.


#  Hist_Corpo e Hist_Fin
Hist_Corpo = Hist_Corpo.sort_values(by=['RFC','FECHA','Fecha_Ejecucion'], ascending=False, na_position='last') # Se ordena la tabla.
Hist_Corpo_1 = Hist_Corpo.drop_duplicates(subset=['RFC']) # Se eliminan registros duplicados de acuerdo a la columna RFC.
Hist_Fin = Hist_Fin.sort_values(by=['RFC','FECHA','Fecha_Ejecucion'], ascending=False, na_position='last') # Se ordena la tabla.
Hist_Fin_1 = Hist_Fin.drop_duplicates(subset=['RFC']) # Se eliminan registros duplicados de acuerdo a la columna RFC.


# Al Data Frame se le realizan las siguientes modificaiones en algunas columnas:

# Se reemplaza "'"" por espacio vacío.
full_data_CALIFICA["FOLIO"] = full_data_CALIFICA["FOLIO"].str.replace("'", '')
full_data_CALIFICA["TIPORESPUESTA"] = full_data_CALIFICA["TIPORESPUESTA"].str.replace("'", '')
full_data_CALIFICA["FECHA CONSULTA"] = full_data_CALIFICA["FECHA CONSULTA"].str.replace("'", '')
full_data_CALIFICA["IDCARACTERISTICA"] = full_data_CALIFICA["IDCARACTERISTICA"].str.replace("'", '')
full_data_CALIFICA["VALORCARACTERISTICA"] = full_data_CALIFICA["VALORCARACTERISTICA"].str.replace("'", '')

# Espacio por espacio vacío.
full_data_CALIFICA["TIPORESPUESTA"] = full_data_CALIFICA["TIPORESPUESTA"].str.replace(" ", '')
full_data_CALIFICA["FOLIO"] = full_data_CALIFICA["FOLIO"].str.replace(" ", '')
full_data_CALIFICA["FECHA CONSULTA"] = full_data_CALIFICA["FECHA CONSULTA"].str.replace(" ", '')
full_data_CALIFICA["RFC"] = full_data_CALIFICA["RFC"].str.replace(" ", '')
full_data_CALIFICA["PRIMERNOMBRE"] = full_data_CALIFICA["PRIMERNOMBRE"].str.replace(" ", '')
full_data_CALIFICA["SEGUNDONOMBRE"] = full_data_CALIFICA["SEGUNDONOMBRE"].str.replace(" ", '')
full_data_CALIFICA["APELLIDOPATERNO"] = full_data_CALIFICA["APELLIDOPATERNO"].str.replace(" ", '')
full_data_CALIFICA["IDCARACTERISTICA"] = full_data_CALIFICA["IDCARACTERISTICA"].str.replace(" ", '')
full_data_CALIFICA["NOMBRECARACTERISTICA"] = full_data_CALIFICA["NOMBRECARACTERISTICA"].str.replace(" ", '')
full_data_CALIFICA["ERRORCARACTERISTICA"] = full_data_CALIFICA["ERRORCARACTERISTICA"].str.replace(" ", '')
full_data_CALIFICA["VALORCARACTERISTICA"] = full_data_CALIFICA["VALORCARACTERISTICA"].str.replace(" ", '')

# "--" y "-" por espacio vacío.
full_data_CALIFICA["VALORCARACTERISTICA"] = full_data_CALIFICA["VALORCARACTERISTICA"].str.replace("--", '')
full_data_CALIFICA["VALORCARACTERISTICA"] = full_data_CALIFICA["VALORCARACTERISTICA"].str.replace("-", '')

full_data_CALIFICA_1 = full_data_CALIFICA.copy() # Se hace una copia de full_data_CALIFICA

# Se convierte a tipo entero.
full_data_CALIFICA_1['IDCONSULTA'] = full_data_CALIFICA_1['IDCONSULTA'].astype(int)
full_data_CALIFICA_1['TIPORESPUESTA'] = full_data_CALIFICA_1['TIPORESPUESTA'].astype(int)
full_data_CALIFICA_1['FOLIO'] = full_data_CALIFICA_1['FOLIO'].astype(int)
full_data_CALIFICA_1['FECHA CONSULTA'] = full_data_CALIFICA_1['FECHA CONSULTA'].astype(int)
full_data_CALIFICA_1['IDCARACTERISTICA'] = full_data_CALIFICA_1['IDCARACTERISTICA'].astype(int)
full_data_CALIFICA['VALORCARACTERISTICA'] = full_data_CALIFICA['VALORCARACTERISTICA']*1

full_data_CALIFICA_1['FECHA_CONS_NUEVA'] = pd.to_datetime(full_data_CALIFICA_1['FECHA CONSULTA'], format='%d%m%Y') # Se convierte a fecha en el formato dmy.
full_data_CALIFICA_1 = full_data_CALIFICA_1.sort_values(by='FECHA_CONS_NUEVA', ascending=True, na_position='last') # Se ordena la tabla.
full_data_CALIFICA_2 = full_data_CALIFICA_1.drop_duplicates(subset=['IDCONSULTA',"IDCARACTERISTICA"]) # Se eliminan duplicados de acuerdo a las columnas IDCONSULTA y IDCARACTERISTICA


CARACTERISTICAS = list(full_data_CALIFICA["NOMBRECARACTERISTICA"].unique()) #Se genera una lista con valores únicos de la columna NOMBRECARACTERISTICA de DF full_data_CALIFICA.
CALIFICA_FULL_H = full_data_CALIFICA_2.drop_duplicates(subset=['FOLIO'])["IDCONSULTA"] # Se eliminan duplicados de acuerdo a las columnas FOLIO y IDCONSULTA.

Nombres = ["IDCONSULTA"]
for caracteristica in CARACTERISTICAS: # Se recorre cada característica de la lista CARACTERISTICAS
    # para obtener solo los registros de los campos "IDCONSULTA" y "VALORCARACTERISTICA"
    # que cumplen según la columna NOMBRECARACTERISTICA, se almacenan en Aux_Califica.
    Aux_Califica = (full_data_CALIFICA_2[full_data_CALIFICA_2['NOMBRECARACTERISTICA'] == caracteristica])[["IDCONSULTA","VALORCARACTERISTICA"]]
    
    # Se realiza un Left Join entre CALIFICA_FULL_H y Aux_Califica.
    CALIFICA_FULL_H = pd.merge(CALIFICA_FULL_H, Aux_Califica, left_on='IDCONSULTA', right_on='IDCONSULTA', how="left")
    
    Nombres.append(caracteristica)

#for Caracteristica in CARACTERISTICAS:
#    Nombres.append(Caracteristica)
    
CALIFICA_FULL_H.columns = Nombres # Se renombran las columnas según a la lista Nombres.

try:
    CALIFICA_FULL_H = CALIFICA_FULL_H.drop([''], axis=1) # Se eliminan columnas vacías.
except:
    pass
    #print("No hay líneas extras")

# Cartas_Credito_1
Cartas_Credito_1 = pd.DataFrame(columns=InfLinDisCredito.columns) # Se crea un DF vacío con el nombre de las columnas InfLinDisCredito.
# Se toaman los valores de ciertas columnas del DF Cartas_Credito.
Cartas_Credito_1["Inf_No_Cliente"] = Cartas_Credito["CC_NUMERO_CTE_OVATION"]
Cartas_Credito_1["Inf_Contrato_Monex"] = Cartas_Credito["CC_NUMERO_CTE_OVATION"]
Cartas_Credito_1["Inf_Cliente"] = Cartas_Credito["CC_CLIENTE"]
Cartas_Credito_1["Inf_Divisa"] = Cartas_Credito["CC_CURRENCY"]
Cartas_Credito_1["Inf_Producto"] = "Cartas de Crédito"
Cartas_Credito_1["Inf_Sublin"] = 1
Cartas_Credito_1["Inf_Sub_Credito"] = Cartas_Credito["CC_NUMBER"]
Cartas_Credito_1["Inf_Cartera"] = Cartas_Credito["CC_STATUS"]
Cartas_Credito_1["Inf_Saldo"] = Cartas_Credito["CC_AMOUNT"]
Cartas_Credito_1.loc[Cartas_Credito['CC_CURRENCY'] == "USD",'TC'] = TC
Cartas_Credito_1.loc[Cartas_Credito['CC_CURRENCY'] == "MXN",'TC'] = 1
Cartas_Credito_1["Inf_Saldo_Valorizado"] = Cartas_Credito["CC_AMOUNT"] * Cartas_Credito_1["TC"]
Cartas_Credito_1["Inf_Total_Valorizada"] = Cartas_Credito_1["Inf_Saldo_Valorizado"]
Cartas_Credito_1["Inf_Codigo"] = ""

# Se genera el DF auxiliar InfLinDisCredito_Aux, con algunas columnas de InfLinDisCredito
InfLinDisCredito_Aux = InfLinDisCredito[['Inf_No_Cliente','Inf_Contrato_Monex','Inf_Cliente','Inf_Regional','Inf_Calificacion_CNBV','Inf_Porcentaje_Reserva','Inf_Actividad','Inf_Ventas_Totales_Anuales']]
InfLinDisCredito_Aux = InfLinDisCredito_Aux.drop_duplicates(subset=['Inf_No_Cliente']) # Se eliminan duplicados según la columna Inf_No_Cliente.
# Se realiza un Left Join entre Cartas_Credito_1 y InfLinDisCredito_Aux
Cartas_Credito_2 = pd.merge(Cartas_Credito_1, InfLinDisCredito_Aux, left_on='Inf_No_Cliente', right_on='Inf_No_Cliente', how="left", suffixes=('', '_Inf'))

# Se generan nuevas columnas con registros de otras columnas cuyos valores no son nulos.
Cartas_Credito_2.loc[Cartas_Credito_2["Inf_Contrato_Monex_Inf"].notnull(),'Inf_Contrato_Monex'] = Cartas_Credito_2["Inf_Contrato_Monex_Inf"]
Cartas_Credito_2.loc[Cartas_Credito_2["Inf_Cliente_Inf"].notnull(),'Inf_Cliente'] = Cartas_Credito_2["Inf_Cliente_Inf"]
Cartas_Credito_2.loc[Cartas_Credito_2["Inf_Regional_Inf"].notnull(),'Inf_Regional'] = Cartas_Credito_2["Inf_Regional_Inf"]
Cartas_Credito_2.loc[Cartas_Credito_2["Inf_Calificacion_CNBV_Inf"].notnull(),'Inf_Calificacion_CNBV'] = Cartas_Credito_2["Inf_Calificacion_CNBV_Inf"]
Cartas_Credito_2.loc[Cartas_Credito_2["Inf_Porcentaje_Reserva_Inf"].notnull(),'Inf_Porcentaje_Reserva'] = Cartas_Credito_2["Inf_Porcentaje_Reserva_Inf"]
Cartas_Credito_2.loc[Cartas_Credito_2["Inf_Actividad_Inf"].notnull(),'Inf_Actividad'] = Cartas_Credito_2["Inf_Actividad_Inf"]
Cartas_Credito_2.loc[Cartas_Credito_2["Inf_Ventas_Totales_Anuales_Inf"].notnull(),'Inf_Ventas_Totales_Anuales'] = Cartas_Credito_2["Inf_Ventas_Totales_Anuales_Inf"]

# Se eliminan columnas.
Cartas_Credito_3 = Cartas_Credito_2.drop(['Inf_Contrato_Monex_Inf','Inf_Cliente_Inf','Inf_Regional_Inf','Inf_Calificacion_CNBV_Inf','Inf_Porcentaje_Reserva_Inf','Inf_Actividad_Inf','Inf_Ventas_Totales_Anuales_Inf'], axis=1)

Aux = Cartas_Credito_3[['Inf_No_Cliente','Inf_Cliente']] # Se toman columnas
Aux2 = Aux.drop_duplicates(subset=['Inf_No_Cliente']) # Se eliminan duplicados.

# Se realiza un Left Join entre Cartas_Credito_3 y Aux2
Cartas_Credito_3 = Cartas_Credito_3.drop(['Inf_Cliente'], axis=1)
Cartas_Credito_4 = pd.merge(Cartas_Credito_3, Aux2, left_on='Inf_No_Cliente', right_on='Inf_No_Cliente', how="left", suffixes=('', ''))

# Full_Data_0
Full_Data_0 = pd.concat([InfLinDisCredito, Cartas_Credito_4]) # Se hace una unión entre InfLinDisCredito y Cartas_Credito_4.
# Se convierten columnas a tipo numérico.
Full_Data_0['Inf_No_Cliente'] = pd.to_numeric(Full_Data_0['Inf_No_Cliente'], errors="raise")
Full_Data_0['Inf_Sublin'] = pd.to_numeric(Full_Data_0['Inf_Sublin'], errors="coerce")
Full_Data_0['Inf_Contrato_Monex'] = pd.to_numeric(Full_Data_0['Inf_Contrato_Monex'], errors="coerce")

Full_Data_0 = Full_Data_0.dropna(subset=["Inf_Sub_Credito"]) # Se eliminan registros no válidos según Inf_Sub_Credito.

# Se convierten columnas a tipo entero.
Full_Data_0['Inf_No_Cliente'] = Full_Data_0['Inf_No_Cliente'].astype(int)
Full_Data_0['Inf_Sublin'] = Full_Data_0['Inf_Sublin'].astype(int)
Full_Data_0['Inf_Contrato_Monex'] = Full_Data_0['Inf_Contrato_Monex'].astype(int)

# Se realizan concatenaciones de varias columnas.
Full_Data_0["ID_Linea"] = Full_Data_0["Inf_No_Cliente"].astype(str)+Full_Data_0["Inf_Divisa"].astype(str)+Full_Data_0["Inf_Producto"].astype(str)+Full_Data_0["Inf_Cartera"].astype(str)+Full_Data_0["Inf_Sublin"].astype(str)
Full_Data_0["ID_Linea_Rep_Ven"] = Full_Data_0["Inf_Contrato_Monex"].astype(str)+Full_Data_0["Inf_Producto"].astype(str)+Full_Data_0["Inf_Sub_Credito"].astype(str)+Full_Data_0["Inf_Sublin"].astype(str)+Full_Data_0["Inf_Codigo"].astype(str)
Full_Data_0.loc[Full_Data_0['Inf_Cartera']=="Vigente",'Cartera_Estatus'] = "activo"
Full_Data_0.loc[Full_Data_0['Inf_Cartera']=="Vencida",'Cartera_Estatus'] = "car_venc"
Full_Data_0.loc[Full_Data_0['Inf_Cartera']=="Pendiente",'Cartera_Estatus'] = "activo"
Full_Data_0["ID_Linea_Modelo"] = Full_Data_0["Inf_No_Cliente"].astype(str)+Full_Data_0["Inf_Producto"]+Full_Data_0["Cartera_Estatus"]+Full_Data_0["Inf_Divisa"].astype(str)

# Se realizan concatenaciones de varias columnas.
Reporte_Vencidos["ID_Linea_Rep_Ven"] = Reporte_Vencidos["RV_Contrato"].astype(str)+Reporte_Vencidos["RV_Producto"]+Reporte_Vencidos["RV_Sub"].astype(str)+Reporte_Vencidos["RV_Sub_Linea"].astype(str)+Reporte_Vencidos["RV_Codigo"]
Reporte_Vencidos_1 = Reporte_Vencidos.drop_duplicates(subset=['ID_Linea_Rep_Ven']) # Se eliminan registros no válidos según ID_Linea_Rep_Ven.

#print('Base creada con '+str(len(Full_Data_0.columns))+' columnas y '+str(len(Full_Data_0))+' filas')


#################################################
#               5 Cruce de Tablas               #
#################################################


# Se realiza un Left Join entre Full_Data_0 y Reporte_Vencidos_1 para obtener una base llamada Full_Data_1.
Full_Data_1 = pd.merge(Full_Data_0, Reporte_Vencidos_1, left_on='ID_Linea_Rep_Ven', right_on='ID_Linea_Rep_Ven', how="left", suffixes=('', '_RV'))

# Se convierten columnas a tipo numérico.
Full_Data_1['Inf_Sub_Credito'] = pd.to_numeric(Full_Data_1['Inf_Sub_Credito'], errors="coerce")
Full_Data_1['Inf_Saldo'] = pd.to_numeric(Full_Data_1['Inf_Saldo'], errors="coerce")
Full_Data_1['Inf_Interes'] = pd.to_numeric(Full_Data_1['Inf_Interes'], errors="coerce")
Full_Data_1['Inf_Mora'] = pd.to_numeric(Full_Data_1['Inf_Mora'], errors="coerce")
Full_Data_1['Inf_Saldo_Valorizado'] = pd.to_numeric(Full_Data_1['Inf_Saldo_Valorizado'], errors="coerce")
Full_Data_1['Inf_Total_Valorizada'] = pd.to_numeric(Full_Data_1['Inf_Total_Valorizada'], errors="coerce")
Full_Data_1['Inf_Total_USD'] = pd.to_numeric(Full_Data_1['Inf_Total_USD'], errors="coerce")
Full_Data_1['Inf_Total_MXN'] = pd.to_numeric(Full_Data_1['Inf_Total_MXN'], errors="coerce")
Full_Data_1['Inf_Porcentaje_Reservaf'] = pd.to_numeric(Full_Data_1['Inf_Porcentaje_Reserva'], errors="coerce")
Full_Data_1['Inf_Monto_CNBV'] = pd.to_numeric(Full_Data_1['Inf_Monto_CNBV'], errors="coerce")
Full_Data_1['Inf_Mora_Orden'] = pd.to_numeric(Full_Data_1['Inf_Mora_Orden'], errors="coerce")
Full_Data_1['Inf_Ventas_Totales_Anuales'] = pd.to_numeric(Full_Data_1['Inf_Ventas_Totales_Anuales'], errors="coerce")
Full_Data_1['Inf_Autorizado'] = pd.to_numeric(Full_Data_1['Inf_Autorizado'], errors="coerce")
Full_Data_1['Inf_Dispuesto'] = pd.to_numeric(Full_Data_1['Inf_Dispuesto'], errors="coerce")
Full_Data_1['Inf_Disponible'] = pd.to_numeric(Full_Data_1['Inf_Disponible'], errors="coerce")
Full_Data_1['RV_Dias_Irregular'] = pd.to_numeric(Full_Data_1['RV_Dias_Irregular'], errors="coerce")
Full_Data_1['RV_Capital_Vencido'] = pd.to_numeric(Full_Data_1['RV_Capital_Vencido'], errors="coerce")
Full_Data_1['RV_Total'] = pd.to_numeric(Full_Data_1['RV_Total'], errors="coerce")

# Se convierten columnas a tipo string y todo a mayúsculas.
Full_Data_1['ID_Linea'] = Full_Data_1['ID_Linea'].str.upper()

#print('Base creada con '+str(len(Full_Data_1.columns))+' columnas y '+str(len(Full_Data_1))+' filas')

# Listas del tipo de variables
Var_Agrupadas = ['Inf_No_Cliente',"Inf_Contrato_Monex","Inf_Divisa","Inf_Producto","Inf_Cartera","Inf_Cliente","Inf_Regional","Inf_TC","Inf_Calificacion_CNBV","Inf_Porcentaje_Reserva","Inf_Codigo","Inf_Sublin","Inf_Proposito","Inf_Destino","Inf_Revolvente","Inf_Fec_Inicio_Contrato","Inf_Fec_Venc_Contrato","Inf_Garantia","Inf_Actividad","Inf_Descripcion","Inf_Ventas_Totales_Anuales","TC","ID_Linea","Cartera_Estatus","ID_Linea_Modelo"]
Var_No_Agrupadas = ['ID_Linea','Inf_Sub_Credito',"Inf_Saldo","Inf_Interes","Inf_Mora","Inf_Saldo_Valorizado","Inf_Total_Valorizada","Inf_Total_USD","Inf_Total_MXN","Inf_Monto_CNBV","Inf_Mora_Orden","Inf_Autorizado","Inf_Dispuesto","Inf_Disponible","RV_Dias_Irregular","RV_Capital_Vencido","RV_Total"]

Full_Data_1_Agrup = Full_Data_1[Var_Agrupadas] # Se filtra Full_Data_1 solo con las columnas de la lista Var_Agrupadas.
Full_Data_1_Agrup = Full_Data_1_Agrup.drop_duplicates(subset=Var_Agrupadas) # Se eliman duplicados.
Full_Data_1_Agrup = Full_Data_1_Agrup.drop_duplicates(subset='ID_Linea') # Se eliminan duplicados según la columna ID_Linea.

Full_Data_1_No_Agrup = Full_Data_1[Var_No_Agrupadas] # Se filtra Full_Data_1 solo con las columnas de la lista Var_No_Agrupadas.

# Se realizan los siguientes cálculos agrupados por ID_Linea: 
Full_Data_1_No_Agrup = Full_Data_1_No_Agrup.groupby("ID_Linea").agg(
    Inf_Sub_Credito_C = pd.NamedAgg(column="Inf_Sub_Credito", aggfunc="count"),   # Conteo de Inf_Sub_Credito
    Inf_Saldo_S = pd.NamedAgg(column="Inf_Saldo", aggfunc="sum"),                 # Suma de Inf_Saldo.
    Inf_Interes_S = pd.NamedAgg(column="Inf_Interes", aggfunc="sum"),             # Suma de Inf_Interes.
    Inf_Mora_S = pd.NamedAgg(column="Inf_Mora", aggfunc="sum"),                   # Suma de Inf_Mora.
    Inf_Saldo_Valorizado_S = pd.NamedAgg(column="Inf_Saldo_Valorizado", aggfunc="sum"), # Suma de Inf_Saldo_Valorizado.
    Inf_Total_Valorizada_S = pd.NamedAgg(column="Inf_Total_Valorizada", aggfunc="sum"), # Suma de Inf_Total_Valorizado.
    Inf_Total_USD_S = pd.NamedAgg(column="Inf_Total_USD", aggfunc="sum"),         # Suma de Inf_Total_USD. 
    Inf_Total_MXN_S = pd.NamedAgg(column="Inf_Total_MXN", aggfunc="sum"),         # Suma de Inf_Total_MXN.
    Inf_Monto_CNBV_S = pd.NamedAgg(column="Inf_Monto_CNBV", aggfunc="sum"),       # Suma de Inf_Monto_CNBV.
    Inf_Mora_Orden_S = pd.NamedAgg(column="Inf_Mora_Orden", aggfunc="sum"),       # Suma de Inf_Mora_Orden.
    Inf_Autorizado_Max = pd.NamedAgg(column="Inf_Autorizado", aggfunc="max"),     # Máximo Inf_Autorizado.
    Inf_Dispuesto_Max = pd.NamedAgg(column="Inf_Dispuesto", aggfunc="max"),       # Máximo Inf_Dispuesto.
    Inf_Disponible_Max = pd.NamedAgg(column="Inf_Disponible", aggfunc="max"),     # Máximo Inf_Disponible.
    RV_Dias_Irregular_Max = pd.NamedAgg(column="RV_Dias_Irregular", aggfunc="max"), # Máximo RV_Dias_Irregular.
    RV_Dias_Irregular_Min = pd.NamedAgg(column="RV_Dias_Irregular", aggfunc="min"), # Mínimo RV_Dias_Irregular.
    RV_Capital_Vencido_S = pd.NamedAgg(column="RV_Capital_Vencido", aggfunc="sum"), # Suma de RV_Capital_Vencido.
    RV_Total_S = pd.NamedAgg(column="RV_Total", aggfunc="sum")                      # Suma de RV_Total.
)   
Full_Data_1_No_Agrup["ID"] = list(Full_Data_1_No_Agrup.index) # Se asocia los índices de Full_Data_1_No_Agrup a una columna
# que sirve como llave.

# Se realiza un Left Join entre Full_Data_1_Agrup y Full_Data_1_No_Agrup para generar Full_Data_2.
Full_Data_2 = pd.merge(Full_Data_1_Agrup, Full_Data_1_No_Agrup, left_on='ID_Linea', right_on='ID', how="left")

Sucursales_sin_dup = Sucursales.drop_duplicates(subset=['S_No_Contrato']) # Se eliminan duplicados.

# Se realiza un Left Join entre Full_Data_2 y Sucursales_sin_dup para generar Full_Data_3.
Full_Data_3 = pd.merge(Full_Data_2, Sucursales_sin_dup, left_on='Inf_Contrato_Monex', right_on='S_No_Contrato', how="left", suffixes=('', '_Suc'))
Full_Data_3 = Full_Data_3.drop(['S_No_Contrato'], axis=1) # Se elimina la columna S_No_Contrato.


RFC_1 = RFC.drop_duplicates(subset=['Cliente']) # Se eliminan duplicados.
RFC_1 = RFC_1[["Cliente","RFC"]] # Se toman las columnas Cliente y RFC.
# Se realiza un Left Join entre Full_Data_3 y RFC_1 para generar Full_Data_4.
Full_Data_4 = pd.merge(Full_Data_3, RFC_1, left_on='Inf_No_Cliente', right_on='Cliente', how="left", suffixes=('', '_RFC'))
Full_Data_4 = Full_Data_4.drop(['Cliente'], axis=1) # Se elimina la columna Cliente.


Grupo_Riesgo_1 = Grupo_Riesgo.drop_duplicates(subset=['GR_Ovation']) # Se eliminan duplicados.
# Se realiza un Left Join entre Full_Data_4 y Grupo_Riesgo_1 para generar Full_Data_5.
Full_Data_5 = pd.merge(Full_Data_4, Grupo_Riesgo_1, left_on='Inf_No_Cliente', right_on='GR_Ovation', how="left", suffixes=('', '_GR'))
Full_Data_5 = Full_Data_5.drop(['GR_Ovation'], axis=1) # Se elimina la columna GR_Ovation.
Full_Data_5.loc[Full_Data_5['GR_Grupo_Riesgo'].isna(),'GR_Grupo_Riesgo'] = Full_Data_5["Inf_No_Cliente"] # Se llenan elementos vacíos con Inf_No_Cliente.


Base_Insumos_Limpia_3["BI_ID"] = pd.to_numeric(Base_Insumos_Limpia_3["BI_ID"], errors="coerce") # Se convierte a numérico la columna BI_ID.
Base_Insumos_Limpia_4=Base_Insumos_Limpia_3.drop_duplicates(subset=["BI_ID"]) # Se eliminan duplicados.
# Se realiza un Left Join entre Full_Data_5 y Base_Insumos_Limpia_4 para generar Full_Data_6.
Full_Data_6 = pd.merge(Full_Data_5, Base_Insumos_Limpia_4, left_on='Inf_No_Cliente', right_on='BI_ID', how="left", suffixes=('', '_BI'))
Full_Data_6 = Full_Data_6.drop(['BI_ID'], axis=1) # Se elimina la columna

Modelo_Calif_1=Modelo_Calif.drop_duplicates(subset=["MNC_ID"]) # Se eliminan duplicados.
# Se realiza un Left Join entre Full_Data_6 y Modelo_Calif_1 para generar Full_Data_7.
Full_Data_7 = pd.merge(Full_Data_6, Modelo_Calif_1, left_on='ID_Linea_Modelo', right_on='MNC_ID', how="left", suffixes=('', '_M'))
Full_Data_7 = Full_Data_7.drop(['MNC_ID'], axis=1) # Se elimina la columna

CALIFICA_FULL_H_1 = CALIFICA_FULL_H.drop_duplicates(subset=["IDCONSULTA"]) # Se eliminan duplicados.
# Se realiza un Left Join entre Full_Data_7 y CALIFICA_FULL_H_1 para generar Full_Data_8.
Full_Data_8 = pd.merge(Full_Data_7, CALIFICA_FULL_H_1, left_on='Inf_No_Cliente', right_on='IDCONSULTA', how="left", suffixes=('', '_Cal'))
Full_Data_8 = Full_Data_8.drop(['IDCONSULTA'], axis=1) # Se elimina la columna

Watch_1 = Watch.sort_values(by='W_FECHA_HIT', ascending=True, na_position='last') # Se ordenan registros de manera ascendente de acuerdo a columna W_FECHA_HIT.
Watch_1 = Watch_1.drop_duplicates(subset=["W_RFC"]) # Se eliminan duplicados.
# Se realiza un Left Join entre Full_Data_8 y Watch_1 para generar Full_Data_9.
Full_Data_9 = pd.merge(Full_Data_8, Watch_1, left_on='RFC', right_on='W_RFC', how="left", suffixes=('', '_W'))
Full_Data_9 = Full_Data_9.drop(['W_RFC'], axis=1) # Se elimina la columna

# De Hist_Corpo_1 se realiza una copia y solo se toman las columnas necesarias.
Hist_Corpo_2 = pd.DataFrame(Hist_Corpo_1)
Hist_Corpo_2[["FECHA_F","CRECIMIENTO_VENTAS_F","RAZON_LIQUIDEZ_F","APALANCAMIENTO_CAPITAL_F","ROE_F","COBERTURA_DEUDA_FLUJO_F","MARGEN_EBITDA_F","EBITDA_ANUAL_F","Disponible_F","PASIVO_FIN_F"]]=Hist_Corpo_2[["FECHA_5","CRECIMIENTO_VENTAS_5","RAZON_LIQUIDEZ_5","APALANCAMIENTO_CAPITAL_5","ROE_5","COBERTURA_DEUDA_FLUJO_5","MARGEN_EBITDA_5","EBITDA_ANUAL_5","Disponible_5","PASIVO_FIN_5"]]

# Se toman valores de columnas de acuerdo a las condiciones.
Hist_Corpo_2.loc[Hist_Corpo_2["FECHA_5"].dt.month!=12,"FECHA_F"] = Hist_Corpo_2["FECHA_3"]
Hist_Corpo_2.loc[Hist_Corpo_2["FECHA_5"].dt.month==12,"CRECIMIENTO_VENTAS_F"] = Hist_Corpo_2["CRECIMIENTO_VENTAS_3"]
Hist_Corpo_2.loc[Hist_Corpo_2["FECHA_5"].dt.month==12,"RAZON_LIQUIDEZ_F"] = Hist_Corpo_2["RAZON_LIQUIDEZ_3"]
Hist_Corpo_2.loc[Hist_Corpo_2["FECHA_5"].dt.month==12,"APALANCAMIENTO_CAPITAL_F"] = Hist_Corpo_2["APALANCAMIENTO_CAPITAL_3"]
Hist_Corpo_2.loc[Hist_Corpo_2["FECHA_5"].dt.month==12,"ROE_F"] = Hist_Corpo_2["ROE_3"]
Hist_Corpo_2.loc[Hist_Corpo_2["FECHA_5"].dt.month==12,"COBERTURA_DEUDA_FLUJO_F"] = Hist_Corpo_2["COBERTURA_DEUDA_FLUJO_3"]
Hist_Corpo_2.loc[Hist_Corpo_2["FECHA_5"].dt.month==12,"MARGEN_EBITDA_F"] = Hist_Corpo_2["MARGEN_EBITDA_3"]
Hist_Corpo_2.loc[Hist_Corpo_2["FECHA_5"].dt.month==12,"EBITDA_ANUAL_F"] = Hist_Corpo_2["EBITDA_ANUAL_3"]
Hist_Corpo_2.loc[Hist_Corpo_2["FECHA_5"].dt.month==12,"Disponible_F"] = Hist_Corpo_2["Disponible_3"]
Hist_Corpo_2.loc[Hist_Corpo_2["FECHA_5"].dt.month==12,"PASIVO_FIN_F"] = Hist_Corpo_2["PASIVO_FIN_3"]

Hist_Corpo_3 = Hist_Corpo_2[["RFC","FECHA","EMPLEADOS","Archivo","FECHA_F","CRECIMIENTO_VENTAS_F","RAZON_LIQUIDEZ_F","APALANCAMIENTO_CAPITAL_F","ROE_F","COBERTURA_DEUDA_FLUJO_F","MARGEN_EBITDA_F","EBITDA_ANUAL_F","Disponible_F","PASIVO_FIN_F"]]
Hist_Corpo_4 = Hist_Corpo_3.drop_duplicates(subset=["RFC"]) # Se eliminan duplicados.


# De Hist_Fin_1 se realiza una copia y solo se toman las columnas necesarias.
Hist_Fin_2 = pd.DataFrame(Hist_Fin_1)
Hist_Fin_2[["FECHA_F","IMOR_F","PROVISIONES_A_CARTERA_F","CAPITAL_A_ACTIVO_TOTAL_F","PRESTAMO_BANCARIOS_A_CARTERA_F","UTILIDAD_A_CAPITAL_SIN_RESERVAS_F","GASTOS_OP_A_INGRESOS_F"]] = Hist_Fin_2[["FECHA_5","IMOR_5","PROVISIONES_A_CARTERA_5","CAPITAL_A_ACTIVO_TOTAL_5","PRESTAMO_BANCARIOS_A_CARTERA_5","UTILIDAD_A_CAPITAL_SIN_RESERVAS_5","GASTOS_OP_A_INGRESOS_5"]]

# Se toman valores de columnas de acuerdo a las condiciones.
Hist_Fin_2.loc[Hist_Fin_2["FECHA_5"].dt.month!=12,"FECHA_F"] = Hist_Fin_2["FECHA_3"]
Hist_Fin_2.loc[Hist_Fin_2["FECHA_5"].dt.month==12,"IMOR_F"] = Hist_Fin_2["IMOR_3"]
Hist_Fin_2.loc[Hist_Fin_2["FECHA_5"].dt.month==12,"CAPITAL_A_ACTIVO_TOTAL_F"] = Hist_Fin_2["CAPITAL_A_ACTIVO_TOTAL_3"]
Hist_Fin_2.loc[Hist_Fin_2["FECHA_5"].dt.month==12,"PRESTAMO_BANCARIOS_A_CARTERA_F"] = Hist_Fin_2["PRESTAMO_BANCARIOS_A_CARTERA_3"]
Hist_Fin_2.loc[Hist_Fin_2["FECHA_5"].dt.month==12,"UTILIDAD_A_CAPITAL_SIN_RESERVAS_F"] = Hist_Fin_2["UTILIDAD_A_CAPITAL_SIN_RESERVAS_3"]
Hist_Fin_2.loc[Hist_Fin_2["FECHA_5"].dt.month==12,"GASTOS_OP_A_INGRESOS_F"] = Hist_Fin_2["GASTOS_OP_A_INGRESOS_3"]

Hist_Fin_3 = Hist_Fin_2[["RFC","FECHA","EMPLEADOS","Archivo","FECHA_F","IMOR_F","PROVISIONES_A_CARTERA_F","CAPITAL_A_ACTIVO_TOTAL_F","PRESTAMO_BANCARIOS_A_CARTERA_F","UTILIDAD_A_CAPITAL_SIN_RESERVAS_F","GASTOS_OP_A_INGRESOS_F"]]
Hist_Fin_4 = Hist_Fin_3.drop_duplicates(subset=["RFC"]) # Se eliminan duplicados.

# Se realiza una unión de las tablas Hist_Corpo_4 y Hist_Fin_4
EEFF = pd.concat([Hist_Corpo_4, Hist_Fin_4], sort=False) 
EEFF = EEFF.drop_duplicates(subset=["RFC"]) # Se eliminan duplicados.
EEFF_1 = EEFF[EEFF["RFC"].notna()] # Se filtra solo por registros no vacíos en la columna RFC.

# Se realiza un Left Join entre Full_Data_9 y EEFF_1 para generar Full_Data_10.
Full_Data_10 = pd.merge(Full_Data_9, EEFF_1, left_on='RFC', right_on='RFC', how="left", suffixes=('', '_EEFF'))

#Full_Data_10.loc['Dias_Mora_Watch'] = 0
# Se colocan valores en la columna Dias_Mora_Watch de acuerdo a las condiciones cumplidas.
Full_Data_10.loc[Full_Data_10['W_Imp_29_DIAS'] > 0 , 'Dias_Mora_Watch'] = 29
Full_Data_10.loc[Full_Data_10['W_Imp_59_DIAS'] > 0 , 'Dias_Mora_Watch'] = 59
Full_Data_10.loc[Full_Data_10['W_Imp_89_DIAS'] > 0 , 'Dias_Mora_Watch'] = 89
Full_Data_10.loc[Full_Data_10['W_Imp_119_DIAS'] > 0 , 'Dias_Mora_Watch'] = 119
Full_Data_10.loc[Full_Data_10['W_Imp_179_DIAS'] > 0 , 'Dias_Mora_Watch'] = 179
Full_Data_10.loc[Full_Data_10['W_Imp_MAS_179_DIAS'] > 0 , 'Dias_Mora_Watch'] = 180

Full_Data_10.loc[Full_Data_10['W_Imp_29_DIAS_B'] > 0 , 'Dias_Mora_Watch_B'] = 29
Full_Data_10.loc[Full_Data_10['W_Imp_59_DIAS_B'] > 0 , 'Dias_Mora_Watch_B'] = 59
Full_Data_10.loc[Full_Data_10['W_Imp_89_DIAS_B'] > 0 , 'Dias_Mora_Watch_B'] = 89
Full_Data_10.loc[Full_Data_10['W_Imp_119_DIAS_B'] > 0 , 'Dias_Mora_Watch_B'] = 119
Full_Data_10.loc[Full_Data_10['W_Imp_179_DIAS_B'] > 0 , 'Dias_Mora_Watch_B'] = 179
Full_Data_10.loc[Full_Data_10['W_Imp_MAS_179_DIAS_B'] > 0 , 'Dias_Mora_Watch_B'] = 180

Full_Data_10['FECHA_SMART'] =  fecha # Se asigna fecha.

# Se genera la columna Deuda_Neta_EBITDA_F = (Disponible_F - PASIVO_FIN_F) / (EBITDA_ANUAL_F)
Full_Data_10["Deuda_Neta_EBITDA_F"] = (Full_Data_10["Disponible_F"]-Full_Data_10["PASIVO_FIN_F"]) / (Full_Data_10["EBITDA_ANUAL_F"])

# Se almacena la información.
#Ruta_Output_1 = "C:/Users/52551/Desktop/Monex_Tratamiento/"
Full_Data_10.to_excel(Ruta_Output + '3_Full_Data_Lineas_'+str(fecha.year)+'_'+str(fecha.month)+'.xlsx', sheet_name='LINEAS', index=False)

print('Base creada con '+str(len(Full_Data_10.columns))+' columnas y '+str(len(Full_Data_10))+' filas')


#################################################
#          6 Evaluación del modelo SMART        #
#################################################

"""
Se cargan los archivos necesarios para la ejecución del código.

- 3_Modelo: Contiene las variables y rangos del modelo.
- Full Data de Líneas: Base de datos con el total de las líneas del mes.
"""


#Ruta_Tratamiento= "C:/Users/52551/Desktop/Modelo_MONEX/2_Tratamiento"
#Ruta_Outputs = "C:/Users/52551/Desktop/Modelo_MONEX/3_Outputs/"

# ya se tiene fecha=datetime(2021, 4, 1) 

#Ruta_Catalogo_Modelo= "C:/Users/52551/Desktop/Modelo_MONEX/3_Modelo.xlsx"
#Ruta_Hist_Cli= "C:/Users/jesus/Documents/Monex_Python/1_Inputs/Hist_Cli_2020_0.xlsx"

#Ruta_Lineas = Ruta_Tratamiento + '3_Full_Data_Lineas_'+str(fecha.year)+'_'+str(fecha.month)+'.xlsx' # Archivo Full_Data_10

Ruta_Catalogo_Modelo = Rutas_aux.iloc [3, 2]
Ruta_Catalogo_Modelo = Ruta_Catalogo_Modelo.replace("\\","/")

Ruta_Tratamiento = Rutas_aux.iloc [4, 2]
Ruta_Tratamiento = Ruta_Tratamiento.replace("\\","/") + "/"

Ruta_Outputs = Rutas_aux.iloc [5, 2]
Ruta_Outputs = Ruta_Outputs.replace("\\","/") + "/"

Ruta_Lineas = Ruta_Tratamiento + '3_Full_Data_Lineas_'+str(fecha.year)+'_'+str(fecha.month)+'.xlsx'

# Se configura el mes y año de análisis.
if fecha.month == 1: 
    Mes = 12
    Anio = fecha.year-1
else:
    Mes = fecha.month-1
    Anio = fecha.year

# Se leen las rutas de archivos necesarios.
Ruta_Lineas_Ant = Ruta_Outputs + '4_Full_Data_Final_'+str(Anio)+'_'+str(Mes)+'.xlsx'
Ruta_Lineas_Ant_Acum = Ruta_Outputs + '4_Full_Data_Final_Acum_'+str(Anio)+'_'+str(Mes)+'.xlsx'

# Se valida la existencia del catálogo modelo y se cargan los datos.
if Path(Ruta_Catalogo_Modelo).is_file():
    Modelo_Corp = pd.read_excel(Ruta_Catalogo_Modelo, sheet_name="CORPO", skiprows=0)
    Modelo_Fin = pd.read_excel(Ruta_Catalogo_Modelo, sheet_name="FIN", skiprows=0)
#else:
#    print("Error al leer el Catalogo de Campos")


# Se valida las líneas del mes y se cargan los datos.
if Path(Ruta_Lineas).is_file():
    Full_Data_Lineas = pd.read_excel(Ruta_Lineas, sheet_name="LINEAS", skiprows=0)
#else:
#    print("Error al leer el el archivo de Líneas")

#Full_Data_Lineas = Full_Data_10.copy()


#################################################
#            6.1 Variables del modelo           #
#################################################


try:
    Full_Data_Lineas_1 = Full_Data_Lineas.drop(columns=['Unnamed: 0']) # Se intenta eliminar una columna.
except:
    Full_Data_Lineas_1 = pd.DataFrame(Full_Data_Lineas)

# Listas con tipos de variables.
Var_Agrupadas_Modelo = ['Inf_No_Cliente','Inf_Contrato_Monex','Inf_Cliente','Inf_Regional','Inf_Calificacion_CNBV','Inf_Porcentaje_Reserva','Inf_Actividad','Inf_Ventas_Totales_Anuales','S_Empleado','S_Sucural','S_Region','RFC','GR_Grupo_Riesgo','BI_ANTIGUEDAD_EEFF','BI_ANEXO','MNC_VENTAS_ANUALES','MNC_TIPO_CALIFICACIÓN','MNC_EEFF_OCUPADOS','MNC_AUDITADOS','MNC_ANTIGUEDAD_EEFF','MNC_PUNTAJE_CUANT','MNC_PUNTAJE_CUAL','MNC_PUNTAJE_CREDITICIO','MNC_PI', 'BK12_CLEAN','BK12_NUM_CRED','BK12_NUM_TC_ACT','NBK12_NUM_CRED','BK12_NUM_EXP_PAIDONTIME','BK12_PCT_PROMT','NBK12_PCT_PROMT','BK12_PCT_SAT','NBK12_PCT_SAT','BK24_PCT_60PLUS','NBK24_PCT_60PLUS','NBK12_COMM_PCT_PLUS','BK12_PCT_90PLUS','BK12_DPD_PROM','BK12_IND_QCRA','BK12_MAX_CREDIT_AMT','MONTHS_ON_FILE_BANKING','MONTHS_SINCE_LAST_OPEN_BANKING','BK_IND_PMOR','BK24_IND_EXP','12_INST','BK_DEUDA_TOT','BK_DEUDA_CP','NBK_DEUDA_TOT', 'NBK_DEUDA_CP','DEUDA_TOT','DEUDA_TOT_CP','W_FECHA_HIT','W_OTORGANTE','W_TIPO_CREDITO','W_MONEDA','W_FECHA_APERTURA','W_FECHA_CIERRE','W_PLAZO','W_MONTO_INICIAL','W_SALDO_VIGENTE','W_SALDO_VENCIDO','W_MAX_DIAS_VENCIMIENTO','W_Imp_29_DIAS','W_Imp_59_DIAS','W_Imp_89_DIAS','W_Imp_119_DIAS','W_Imp_179_DIAS','W_Imp_MAS_179_DIAS','W_QUITA','W_QUEBRANTO','W_DACION','W_PAGO','FECHA','EMPLEADOS','Archivo','FECHA_F','CRECIMIENTO_VENTAS_F', 'RAZON_LIQUIDEZ_F','APALANCAMIENTO_CAPITAL_F','ROE_F','COBERTURA_DEUDA_FLUJO_F','MARGEN_EBITDA_F','EBITDA_ANUAL_F','Disponible_F','PASIVO_FIN_F','IMOR_F','PROVISIONES_A_CARTERA_F','CAPITAL_A_ACTIVO_TOTAL_F','PRESTAMO_BANCARIOS_A_CARTERA_F','UTILIDAD_A_CAPITAL_SIN_RESERVAS_F','GASTOS_OP_A_INGRESOS_F','Dias_Mora_Watch','Dias_Mora_Watch_B','Deuda_Neta_EBITDA_F']
Var_No_Agrupadas_Modelo = ['Inf_No_Cliente','Inf_Cartera','Inf_Fec_Venc_Contrato','ID_Linea','Inf_Saldo_S','Inf_Interes_S','Inf_Mora_S','Inf_Saldo_Valorizado_S','Inf_Total_Valorizada_S','Inf_Total_USD_S','Inf_Total_MXN_S','Inf_Monto_CNBV_S','Inf_Mora_Orden_S','Inf_Autorizado_Max','Inf_Dispuesto_Max','Inf_Disponible_Max','RV_Dias_Irregular_Max','RV_Dias_Irregular_Min','RV_Capital_Vencido_S','RV_Total_S','MNC_EI_VALORIZADA','MNC_SP','MNC_SALDO_TOTAL','MNC_RESERVAS_TOTAL','MNC_EI','MNC_PCT_PE','MNC_GRADO_RIESGO']

Full_Data_Cli = Full_Data_Lineas_1[Var_Agrupadas_Modelo] # Del DF Full_Data_Lineas_1 se toman variables agrupadas.
Full_Data_Cli_1 = Full_Data_Cli.drop_duplicates(subset = Var_Agrupadas_Modelo) # Se eliminan registros duplicados.
Full_Data_Cli_1 = Full_Data_Cli_1.drop_duplicates(subset='Inf_No_Cliente') # Se elimina columna.


Full_Data_Lineas_2 = Full_Data_Lineas_1[Var_No_Agrupadas_Modelo] # Del DF Full_Data_Lineas_1 se toman variables no agrupadas.
Full_Data_Lineas_2.loc[:,'Inf_Cartera_Num'] = 0 # Se coloca 0s en la columna Inf_Cartera_Num.
Full_Data_Lineas_2.loc[Full_Data_Lineas_2['Inf_Cartera'] =='Vencida' , 'Inf_Cartera_Num'] = 1 # Se coloca 1s en la columna Inf_Cartera_Num donde
#Inf_Cartera sea 'Vencida'.

Full_Data_Lineas_2.loc[:,'MNC_GRADO_RIESGO_Num'] = 0 # Se coloca 0s en la columna MNC_GRADO_RIESGO_Num.
#Full_Data_Lineas_2['MNC_GRADO_RIESGO_Num']=0
# Se coloca un valor entre 1 y 9 en la columna MNC_GRADO_RIESGO_Num de acuerdo a la condición de la columna MNC_GRADO_RIESGO.
Full_Data_Lineas_2.loc[Full_Data_Lineas_2['MNC_GRADO_RIESGO'] =='A1' , 'MNC_GRADO_RIESGO_Num'] = 1
Full_Data_Lineas_2.loc[Full_Data_Lineas_2['MNC_GRADO_RIESGO'] =='A2' , 'MNC_GRADO_RIESGO_Num'] = 2
Full_Data_Lineas_2.loc[Full_Data_Lineas_2['MNC_GRADO_RIESGO'] =='B1' , 'MNC_GRADO_RIESGO_Num'] = 3
Full_Data_Lineas_2.loc[Full_Data_Lineas_2['MNC_GRADO_RIESGO'] =='B2' , 'MNC_GRADO_RIESGO_Num'] = 4
Full_Data_Lineas_2.loc[Full_Data_Lineas_2['MNC_GRADO_RIESGO'] =='B3' , 'MNC_GRADO_RIESGO_Num'] = 5
Full_Data_Lineas_2.loc[Full_Data_Lineas_2['MNC_GRADO_RIESGO'] =='C1' , 'MNC_GRADO_RIESGO_Num'] = 6
Full_Data_Lineas_2.loc[Full_Data_Lineas_2['MNC_GRADO_RIESGO'] =='C2' , 'MNC_GRADO_RIESGO_Num'] = 7
Full_Data_Lineas_2.loc[Full_Data_Lineas_2['MNC_GRADO_RIESGO'] =='D' , 'MNC_GRADO_RIESGO_Num'] = 8
Full_Data_Lineas_2.loc[Full_Data_Lineas_2['MNC_GRADO_RIESGO'] =='E' , 'MNC_GRADO_RIESGO_Num'] = 9

Full_Data_Lineas_3 = Full_Data_Lineas_2.drop(['MNC_GRADO_RIESGO','Inf_Cartera'], axis=1) # se quitan columnas.

# Se realizan los cálculos siguientes agrupados por la columna Inf_No_Cliente.
Full_Data_Lineas_4 = Full_Data_Lineas_3.groupby("Inf_No_Cliente").agg(
    Inf_Fec_Venc_Contrato_Min = pd.NamedAgg(column="Inf_Fec_Venc_Contrato", aggfunc="min"), # Se calcula el mínimo.
    Num_lineas = pd.NamedAgg(column="ID_Linea", aggfunc="count"), # Se realiza el conteo.
    # Se realiza suma de cada una de las variables.
    Inf_Saldo_S = pd.NamedAgg(column="Inf_Saldo_S", aggfunc="sum"),
    Inf_Interes_S = pd.NamedAgg(column="Inf_Interes_S", aggfunc="sum"),
    Inf_Mora_S = pd.NamedAgg(column="Inf_Mora_S", aggfunc="sum"),
    Inf_Saldo_Valorizado_S = pd.NamedAgg(column="Inf_Saldo_Valorizado_S", aggfunc="sum"),
    Inf_Total_Valorizada_S = pd.NamedAgg(column="Inf_Total_Valorizada_S", aggfunc="sum"),
    Inf_Total_USD_S = pd.NamedAgg(column="Inf_Total_USD_S", aggfunc="sum"),
    Inf_Total_MXN_S = pd.NamedAgg(column="Inf_Total_MXN_S", aggfunc="sum"),
    Inf_Monto_CNBV_S = pd.NamedAgg(column="Inf_Monto_CNBV_S", aggfunc="sum"),
    Inf_Mora_Orden_S = pd.NamedAgg(column="Inf_Mora_Orden_S", aggfunc="sum"),
    Inf_Autorizado_Max_S = pd.NamedAgg(column="Inf_Autorizado_Max", aggfunc="sum"),
    Inf_Dispuesto_Max_S = pd.NamedAgg(column="Inf_Dispuesto_Max", aggfunc="sum"),
    Inf_Disponible_Max_S = pd.NamedAgg(column="Inf_Disponible_Max", aggfunc="sum"),
    RV_Dias_Irregular_Max = pd.NamedAgg(column="RV_Dias_Irregular_Max", aggfunc="max"), # Se calcula el máximo.
    RV_Dias_Irregular_Min = pd.NamedAgg(column="RV_Dias_Irregular_Min", aggfunc="min"), # Se calcula el mínimo.
    RV_Capital_Vencido_S = pd.NamedAgg(column="RV_Capital_Vencido_S", aggfunc="sum"),
    RV_Total_S = pd.NamedAgg(column="RV_Total_S", aggfunc="sum"),
    MNC_EI_VALORIZADA_S = pd.NamedAgg(column="MNC_EI_VALORIZADA", aggfunc="sum"),
    MNC_SP_Avg = pd.NamedAgg(column="MNC_SP", aggfunc="mean"), # Se realiza el promedio.
    MNC_SALDO_TOTAL_S = pd.NamedAgg(column="MNC_SALDO_TOTAL", aggfunc="sum"),
    MNC_RESERVAS_S = pd.NamedAgg(column="MNC_RESERVAS_TOTAL", aggfunc="sum"),
    MNC_EI_S = pd.NamedAgg(column="MNC_EI", aggfunc="sum"),
    MNC_PCT_PE_Avg = pd.NamedAgg(column="MNC_PCT_PE", aggfunc="mean"), # Se realiza el promedio.
    Lineas_Vencidas = pd.NamedAgg(column="Inf_Cartera_Num", aggfunc="sum"),
    MNC_GRADO_RIESGO_Min = pd.NamedAgg(column="MNC_GRADO_RIESGO_Num", aggfunc="min"), # Se calcula el mínimo.
    MNC_GRADO_RIESGO_Max = pd.NamedAgg(column="MNC_GRADO_RIESGO_Num", aggfunc="max"), # Se calcula el máximo.
    MNC_GRADO_RIESGO_Avg = pd.NamedAgg(column="MNC_GRADO_RIESGO_Num", aggfunc="mean") # Se realiza el promedio.
)
Full_Data_Lineas_4["ID"] = list(Full_Data_Lineas_4.index)
#Full_Data_Cli_1=Full_Data_Cli_1.drop(['Inf_No_Cliente'], axis=1)

# Se realiza un Left Join entre Full_Data_Cli_1 y Full_Data_Lineas_4
Full_Data_Cli_2 = pd.merge(Full_Data_Cli_1, Full_Data_Lineas_4, left_on='Inf_No_Cliente', right_on='ID', how="left", suffixes=('', '_2'))
Full_Data_Cli_2 = Full_Data_Cli_2.drop(['ID'], axis=1) # Se elimina columna ID.

Full_Data_Cli_2['Calif_Cli'] = Full_Data_Cli_2['MNC_RESERVAS_S']/Full_Data_Cli_2['MNC_EI_S'] # Se obtiene la columna
# Calif_Cli que es la división entre MNC_RESERVAS_S y MNC_EI_S

# Se realiza un ponderado de acuerdo al resultado anterior.
Full_Data_Cli_2.loc[:,'GRADO_RIESGO_CLI'] = 0
Full_Data_Cli_2.loc[Full_Data_Cli_2['Calif_Cli'] <=0.009 , 'GRADO_RIESGO_CLI'] = 1
Full_Data_Cli_2.loc[Full_Data_Cli_2['Calif_Cli'] <=0.015 , 'GRADO_RIESGO_CLI'] = 2
Full_Data_Cli_2.loc[Full_Data_Cli_2['Calif_Cli'] <=0.02 , 'GRADO_RIESGO_CLI'] = 3
Full_Data_Cli_2.loc[Full_Data_Cli_2['Calif_Cli'] <=0.025 , 'GRADO_RIESGO_CLI'] = 4
Full_Data_Cli_2.loc[Full_Data_Cli_2['Calif_Cli'] <=0.05 , 'GRADO_RIESGO_CLI'] = 5
Full_Data_Cli_2.loc[Full_Data_Cli_2['Calif_Cli'] <=0.1 , 'GRADO_RIESGO_CLI'] = 6
Full_Data_Cli_2.loc[Full_Data_Cli_2['Calif_Cli'] <=0.155 , 'GRADO_RIESGO_CLI'] = 7
Full_Data_Cli_2.loc[Full_Data_Cli_2['Calif_Cli'] <=0.45, 'GRADO_RIESGO_CLI'] = 8
Full_Data_Cli_2.loc[Full_Data_Cli_2['Calif_Cli'] >0.45, 'GRADO_RIESGO_CLI'] = 9

if Anio > 2019: # En caso de ser año mayor a 2019, se realiza un Left Join entre Full_Data_Cli_2 y Grado_Riesgo_Ant_CLI.
    Full_Data_Final_Ant = pd.read_excel(Ruta_Lineas_Ant, sheet_name="Full_Data", skiprows=0)

    Grado_Riesgo_Ant = Full_Data_Final_Ant[['Inf_No_Cliente','GRADO_RIESGO_CLI']]
    Grado_Riesgo_Ant.columns = ['Inf_No_Cliente','GRADO_RIESGO_CLI_ANT']
    Grado_Riesgo_Ant_CLI = Grado_Riesgo_Ant.drop_duplicates(subset=['Inf_No_Cliente'])
    Full_Data_Cli_2 = pd.merge(Full_Data_Cli_2, Grado_Riesgo_Ant_CLI, left_on='Inf_No_Cliente', right_on='Inf_No_Cliente', how="left")
else: # En caso contrario la columna GRADO_RIESGO_CLI_ANT toma valores de GRADO_RIESGO_CLI
    Full_Data_Cli_2[['GRADO_RIESGO_CLI_ANT']] = Full_Data_Cli_2[['GRADO_RIESGO_CLI']]


# Ahora se harán los siguientes cálculos
#Sub división de la cartera
Full_Data_Cli_3 = pd.DataFrame(Full_Data_Cli_2) # Se realiza copia de Full_Data_Cli_2.
Full_Data_Cli_3['ANEXO_Modelo'] = Full_Data_Cli_3["MNC_TIPO_CALIFICACIÓN"] # Se llena ANEXO_Modelo con MNC_TIPO_CALIFICACIÓN.
# Se llena ANEXO_Modelo si se encuentra vacío con BI_ANEXO.
Full_Data_Cli_3.loc[Full_Data_Cli_3["ANEXO_Modelo"].isna(),"ANEXO_Modelo"] = Full_Data_Cli_3["BI_ANEXO"]

# Se genera columna CARTERA con registros según ANEXO_Modelo.
Full_Data_Cli_3.loc[:,'CARTERA'] = ""
Full_Data_Cli_3.loc[Full_Data_Cli_3["ANEXO_Modelo"]=="ANEXO 19","CARTERA"]="PROJECT FINANCE"
Full_Data_Cli_3.loc[Full_Data_Cli_3["ANEXO_Modelo"]=="ANEXO 20","CARTERA"]="FINANCIERO"
Full_Data_Cli_3.loc[Full_Data_Cli_3["ANEXO_Modelo"]=="ANEXO 21","CARTERA"]="PYMES"
Full_Data_Cli_3.loc[Full_Data_Cli_3["ANEXO_Modelo"]=="ANEXO 22","CARTERA"]="CORPORATIVO"
# Se eliminan columnas MNC_TIPO_CALIFICACIÓN,BI_ANEXO
Full_Data_Cli_3 = Full_Data_Cli_3.drop(['MNC_TIPO_CALIFICACIÓN','BI_ANEXO'], axis=1)

#Sub división de la cartera
# Se genera columna S_Region_2 con registros según S_Region.
Full_Data_Cli_3.loc[:,'S_Region_2']=""
Full_Data_Cli_3.loc[Full_Data_Cli_3["S_Region"]=="Norte","S_Region_2"] = "Region Norte (NOR)"
Full_Data_Cli_3.loc[Full_Data_Cli_3["S_Region"]=="Metropolitana","S_Region_2"] = "Region Metropolitana (MET)"
Full_Data_Cli_3.loc[Full_Data_Cli_3["S_Region"]=="Sur","S_Region_2"]= "Region Sur (SUR)"
Full_Data_Cli_3.loc[Full_Data_Cli_3["S_Region"]=="Centro","S_Region_2"]= "Region Centro (CEN)"
Full_Data_Cli_3.loc[Full_Data_Cli_3["S_Region"]=="Noreste","S_Region_2"]= "Region Norte (NOR)"
Full_Data_Cli_3.loc[Full_Data_Cli_3["S_Region"]=="Occidente","S_Region_2"]= "Region Occidente (OCC)"
# Se genera columna Regional_Modelo similar a Inf_Regional.
Full_Data_Cli_3['Regional_Modelo'] = Full_Data_Cli_3["Inf_Regional"]
# Se llena Regional_Modelo si se encuentra vacío con S_Region_2.
Full_Data_Cli_3.loc[Full_Data_Cli_3["Regional_Modelo"].isna(),"Regional_Modelo"] = Full_Data_Cli_3["S_Region_2"]
Full_Data_Cli_3 = Full_Data_Cli_3.drop(['S_Region','S_Region_2','Inf_Regional'], axis=1) # Se eliminan columnas S_Region,S_Region_2,Inf_Regional

#Antiguedad de EEDD
Full_Data_Cli_3['Antiguedad_Modelo'] = Full_Data_Cli_3["MNC_ANTIGUEDAD_EEFF"] # Se genera columna
Full_Data_Cli_3.loc[Full_Data_Cli_3["Antiguedad_Modelo"].isna(),"Antiguedad_Modelo"] = Full_Data_Cli_3["BI_ANTIGUEDAD_EEFF"]
Full_Data_Cli_3=Full_Data_Cli_3.drop(['MNC_ANTIGUEDAD_EEFF','BI_ANTIGUEDAD_EEFF'], axis=1)
Full_Data_Cli_3.loc[Full_Data_Cli_3["CARTERA"]=='PYMES',"Antiguedad_Modelo"] = 0

#Ventas Totales Netas
Full_Data_Cli_3['Ventas_Total_Netas'] = Full_Data_Cli_3["MNC_VENTAS_ANUALES"] # Se genera columna
Full_Data_Cli_3.loc[Full_Data_Cli_3["Ventas_Total_Netas"].isna(),"Ventas_Total_Netas"] = Full_Data_Cli_3["Inf_Ventas_Totales_Anuales"]
Full_Data_Cli_3=Full_Data_Cli_3.drop(['MNC_VENTAS_ANUALES','Inf_Ventas_Totales_Anuales'], axis=1) # Se eliminan columnas

#Perdida Esperada
# Se genera columna Perdida_Esperada = MNC_RESERVAS_S / MNC_EI_S
Full_Data_Cli_3['Perdida_Esperada'] = Full_Data_Cli_3["MNC_RESERVAS_S"] / Full_Data_Cli_3["MNC_EI_S"]
Full_Data_Cli_3.loc[Full_Data_Cli_3["Perdida_Esperada"] > 1,"Perdida_Esperada"] = None
Full_Data_Cli_3.loc[Full_Data_Cli_3["Perdida_Esperada"] < 0,"Perdida_Esperada"] = None

#Califica
Full_Data_Cli_3['Califica_Atrasos'] = Full_Data_Cli_3["BK12_CLEAN"] # Se genera columna
Full_Data_Cli_3['Porcentaje_Pago_Bancos'] = Full_Data_Cli_3["BK12_PCT_PROMT"] # Se genera columna
Full_Data_Cli_3['Porcentaje_Pago_NO_Bancos'] = Full_Data_Cli_3["NBK12_PCT_PROMT"] # Se genera columna
Full_Data_Cli_3['Dias_Prom_Mora'] = Full_Data_Cli_3["BK12_DPD_PROM"] # Se genera columna
Full_Data_Cli_3['QCR_ultimos_12M'] = Full_Data_Cli_3["BK12_IND_QCRA"]  # Se genera columna
Full_Data_Cli_4=Full_Data_Cli_3.drop(['BK12_CLEAN','BK12_PCT_PROMT','BK12_PCT_PROMT','NBK12_PCT_PROMT','BK12_IND_QCRA'], axis=1) # Se eliminan columnas.

#Incremento Notches
Full_Data_Cli_4.loc[Full_Data_Cli_4['GRADO_RIESGO_CLI_ANT'].isna() , 'GRADO_RIESGO_CLI_ANT'] = 9999
Full_Data_Cli_4.loc[Full_Data_Cli_4['GRADO_RIESGO_CLI_ANT'] != 9999 , 'Cambio_Calif'] = Full_Data_Cli_4['GRADO_RIESGO_CLI_ANT']-Full_Data_Cli_4['GRADO_RIESGO_CLI']
Full_Data_Cli_4.loc[Full_Data_Cli_4['GRADO_RIESGO_CLI_ANT'] == 9999 , 'Cambio_Calif'] = 0

mil = """

Se llegó hasta aquí

 """


print('Base creada con '+str(len(Full_Data_Cli_4.columns))+' columnas y '+str(len(Full_Data_Cli_4))+' filas')

print(mil)
Ruta_Output_1 = "C:/Users/52551/Desktop/Monex_Tratamiento/"
#Full_Data_Cli_4.to_excel(Ruta_Output_1 + 'Full_Data_Cli_4'+str(fecha.year)+'_'+str(fecha.month)+'.xlsx', sheet_name='Full_Data', index=False)

#################################################
#            6.2 Ejecución del Modelo           #
#################################################

"""
Se evalúa el modelo en función de las variables pre-cargadas desde el Catálogo Modelo.
"""
#Full_Data_Cli_4 = pd.read_excel(Ruta_Output + 'Full_Data_Cli_4'+str(fecha.year)+'_'+str(fecha.month)+'.xlsx')

Variables_Corpo, Variables_Fin, Variables_T = [], [], []

# Se hará un recorrido por cada variable de Modelo_Corp.
for i in range(len(Modelo_Corp['Variable'])): # se mejora con enumerate.
    
    Variable = Modelo_Corp['Variable'][i] # Se toma el valor de variable.
    new_v = "Calif_Corp_"+str(Variable)
    
    if Modelo_Corp['Relacion'][i] == ">=": # En caso de que exista una relación ">=", se realiza:
        Full_Data_Cli_4.loc[:,new_v] = 0 # Se genera nueva columna en 0.

        # Se asigna un valor ponderado a la columna recien creada si se cumplen las condiciones.
        #print( f"{Modelo_Corp['Relacion'][i]} \n {Full_Data_Cli_4[Variable]>= Modelo_Corp['Rango_Leve'][i]}" )
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable] >= Modelo_Corp['Rango_Leve'][i],new_v] = 1
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable] >= Modelo_Corp['Rango_Medio'][i],new_v] = 5
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable] >= Modelo_Corp['Rango_Grave'][i],new_v] = 25
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable] >= Modelo_Corp['Rango_Moroso'][i],new_v] = 125
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable].isna(),new_v] = 0 # Si el valor de la nueva columna llega a estar vacío, se asigna 0.
        # Se ingresa el nombre de la variable en las listas.
        Variables_Corpo.append(new_v)
        Variables_T.append(new_v)
    else: # En caso de que no exista una relación ">=", se realiza:
        Full_Data_Cli_4.loc[:,new_v] = 0 # Se genera nueva columna en 0.
        # Se asigna un valor ponderado a la columna recien creada si se cumplen las condiciones.
        #print(f"{type(Full_Data_Cli_4[Variable] )} \n {type(Modelo_Corp['Rango_Leve'][i])}" )
        #print( f"{Modelo_Corp['Relacion'][i]} \n {Full_Data_Cli_4[Variable]< Modelo_Corp['Rango_Leve'][i]} " )
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable] < Modelo_Corp['Rango_Leve'][i],new_v] = 1
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable] < Modelo_Corp['Rango_Medio'][i],new_v] = 5
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable] < Modelo_Corp['Rango_Grave'][i],new_v] = 25
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable] < Modelo_Corp['Rango_Moroso'][i],new_v] = 125
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable].isna(),new_v] = 0 # Si el valor de la nueva columna llega a estar vacío, se asigna 0.
        # Se ingresa el nombre de la variable en las listas.
        Variables_Corpo.append(new_v)
        Variables_T.append(new_v)

for i in range(len(Modelo_Fin['Variable'])): # Se hará un recorrido por cada variable de Modelo_Fin.
    #Variable = Modelo_Fin['Variable'][i]

    Variable = Modelo_Fin['Variable'][i] # Se toma el valor de variable.
    new_v = "Calif_Fin_"+str(Variable)
    
    if Modelo_Fin['Relacion'][i] == ">=":
        Full_Data_Cli_4.loc[:,new_v] = 0
        # Se asigna un valor ponderado a la columna recien creada si se cumplen las condiciones.
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable] >= Modelo_Fin['Rango_Leve'][i],new_v] = 1
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable] >= Modelo_Fin['Rango_Medio'][i],new_v] = 5
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable] >= Modelo_Fin['Rango_Grave'][i],new_v] = 25
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable] >= Modelo_Fin['Rango_Moroso'][i],new_v] = 125
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable].isna(),new_v] = 0 # Si el valor de la nueva columna llega a estar vacío, se asigna 0.
        # Se ingresa el nombre de la variable en las listas.
        Variables_Fin.append(new_v)
        Variables_T.append(new_v)
    else:
        Full_Data_Cli_4.loc[:,new_v] = 0
        # Se asigna un valor ponderado a la columna recien creada si se cumplen las condiciones.
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable] < Modelo_Fin['Rango_Leve'][i],new_v] = 1
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable] < Modelo_Fin['Rango_Medio'][i],new_v] = 5
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable] < Modelo_Fin['Rango_Grave'][i],new_v] = 25
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable] < Modelo_Fin['Rango_Moroso'][i],new_v] = 125
        Full_Data_Cli_4.loc[Full_Data_Cli_4[Variable].isna(),new_v] = 0 # Si el valor de la nueva columna llega a estar vacío, se asigna 0.
        # Se ingresa el nombre de la variable en las listas.
        Variables_Fin.append(new_v)
        Variables_T.append(new_v)

# Se realuza la suma de de los valores de Variables_Corpo y Variables_Fin
Full_Data_Cli_4['SMART_Corp'] = Full_Data_Cli_4[Variables_Corpo].sum(axis=1)
Full_Data_Cli_4['SMART_Fin'] = Full_Data_Cli_4[Variables_Fin].sum(axis=1)

# Se asignan valores a la variable "SMART".
Full_Data_Cli_4.loc[Full_Data_Cli_4["CARTERA"]=="FINANCIERO","SMART"] = Full_Data_Cli_4['SMART_Fin']
Full_Data_Cli_4.loc[Full_Data_Cli_4["CARTERA"]=="CORPORATIVO","SMART"] = Full_Data_Cli_4['SMART_Corp']
Full_Data_Cli_4.loc[Full_Data_Cli_4["CARTERA"]=="PYMES","SMART"] = Full_Data_Cli_4['SMART_Corp']
Full_Data_Cli_4.loc[Full_Data_Cli_4["CARTERA"]=="PROJECT FINANCE","SMART"] = 0

# Se realiza la unión de las variables de Modelo_Corp y Modelo_Fin
Variables_M = pd.concat([Modelo_Corp['Variable'], Modelo_Fin['Variable']], sort=False)
Variables_M = list(Variables_M.drop_duplicates()) # se eliminan duplicados
Variables_M.extend(Variables_T) # se agregan valores de la lista Variables_T
Variables_M.extend(['Inf_No_Cliente','ANEXO_Modelo', 'CARTERA', 'Regional_Modelo','SMART','SMART_Color']) # Se agregan nuevas variables.

# Se agrega una variable semáforo de acuerdo a los valores de la variable "SMART".
Full_Data_Cli_4.loc[:,'SMART_Color'] = 'Verde'
Full_Data_Cli_4.loc[Full_Data_Cli_4['SMART'] >= 1,'SMART_Color']='Verde'
Full_Data_Cli_4.loc[Full_Data_Cli_4['SMART'] >= 5,'SMART_Color']='Naranja'
Full_Data_Cli_4.loc[Full_Data_Cli_4['SMART'] >= 25,'SMART_Color']='Rojo'
Full_Data_Cli_4.loc[Full_Data_Cli_4['SMART'] >= 125,'SMART_Color']='Negro'

# Full_Data_Cli_5 contiene solo las variables contenidas en la lista Variables_M
Full_Data_Cli_5 = Full_Data_Cli_4[Variables_M]

Full_Data_Cli_6 = Full_Data_Cli_5.copy() # Se realiza una copia.
# Se renombran columnas.
Full_Data_Cli_6.rename(columns={'RV_Dias_Irregular_Max': 'D_Irregular_M', 'Dias_Mora_Watch': 'D_Mora_Watch_M', 'Dias_Mora_Watch_B': 'D_Mora_Watch_B_M', 'Califica_Atrasos': 'Califica_Atrasos_M', 'Antiguedad_Modelo': 'Antiguedad_M', 'Deuda_Neta_EBITDA_F': 'DN_EBITDA_M', 'Perdida_Esperada': 'PE_M', 'Cambio_Calif': 'Cambio_Calif_M', 'APALANCAMIENTO_CAPITAL_F': 'APALANCAMIENTO_M', 'QCR_ultimos_12M': 'QCR_12M_M', 'Porcentaje_Pago_Bancos': 'P_Pago_Bancos_M', 'RAZON_LIQUIDEZ_F': 'RAZON_LIQUIDEZ_M', 'ROE_F': 'ROE_M', 'IMOR_F': 'IMOR_M', 'CAPITAL_A_ACTIVO_TOTAL_F': 'PCT_CAPITAL_M', 'UTILIDAD_A_CAPITAL_SIN_RESERVAS_F': 'ROA_M'}, inplace=True)

# Se realiza un Left Join entre Full_Data_Lineas_1 y Full_Data_Cli_6.
Full_Data_Final = pd.merge(Full_Data_Lineas_1, Full_Data_Cli_6, left_on='Inf_No_Cliente', right_on='Inf_No_Cliente', how="left", suffixes=('', '_Model'))
# Se realiza un Left Join entre Full_Data_Final y algunas columnas de Full_Data_Cli_4.
Full_Data_Final_2 = pd.merge(Full_Data_Final, Full_Data_Cli_4[['Inf_No_Cliente','GRADO_RIESGO_CLI','GRADO_RIESGO_CLI_ANT']], left_on='Inf_No_Cliente', right_on='Inf_No_Cliente', how="left", suffixes=('', '_Model'))

# Se guarda archivo en .xlsx
Full_Data_Final_2.to_excel(Ruta_Outputs + '4_Full_Data_Final_'+str(fecha.year)+'_'+str(fecha.month)+'.xlsx', sheet_name='Full_Data', index=False)
#Full_Data_Final_2['Inf_Total_Valorizada_S'].sum()

if Anio > 2019: # Si el análisis es de año > 2019.
    Full_Data_Final_Ant_Acum = pd.read_excel(Ruta_Lineas_Ant_Acum, sheet_name="Full_Data", skiprows=0)
    try:
        Full_Data_Final_Ant_Acum = Full_Data_Final_Ant_Acum.drop(columns=['Unnamed: 0'])
    except:
        pass
    # Se realiza la unión de Full_Data_Final_2 y Full_Data_Final_Ant_Acum
    Full_Data_Final_Acum = pd.concat([Full_Data_Final_2, Full_Data_Final_Ant_Acum], ignore_index=True,sort=False)
    # Se guarda archivo en .xlsx
    Full_Data_Final_Acum.to_excel(Ruta_Outputs + '4_Full_Data_Final_Acum_'+str(fecha.year)+'_'+str(fecha.month)+'.xlsx', sheet_name='Full_Data', index=False)
else:
    Full_Data_Final_2.to_excel(Ruta_Outputs + '4_Full_Data_Final_Acum_'+str(fecha.year)+'_'+str(fecha.month)+'.xlsx', sheet_name='Full_Data', index=False)


print(Full_Data_Final_Acum.info() )


#################################################
#               7 Resumen Reporte               #
#################################################


Ruta_Lineas_Acum = Ruta_Outputs + '4_Full_Data_Final_Acum_'+str(fecha.year)+'_'+str(fecha.month)+'.xlsx'

# Se valida la existencia del catálogo modelo
if Path(Ruta_Lineas_Acum).is_file():
    Full_Data_Acum = pd.read_excel(Ruta_Lineas_Acum, sheet_name="Full_Data", skiprows=0)
    Full_Data_Acum["ID"] = Full_Data_Acum["FECHA_SMART"].astype(str)+Full_Data_Acum["CARTERA"].astype(str)+Full_Data_Acum["Inf_Cartera"].astype(str)+Full_Data_Acum["Inf_Producto"].astype(str)+Full_Data_Acum["Regional_Modelo"].astype(str)+Full_Data_Acum["SMART_Color"].astype(str)+Full_Data_Acum["S_Sucural"].astype(str)
#else:
#    print("Error al leer el archivo")


#Full_Data_Acum[['FECHA_SMART', 'CARTERA', 'Inf_Saldo_Valorizado_S', 'Inf_Cartera', 'Inf_Producto', 'Regional_Modelo', 'SMART_Color', 'S_Sucural', 'ID_Linea']]
#Full_Data_Acum["ID"]=Full_Data_Acum["FECHA_SMART"].astype(str)+Full_Data_Acum["CARTERA"].astype(str)+Full_Data_Acum["Inf_Cartera"].astype(str)+Full_Data_Acum["Inf_Producto"].astype(str)+Full_Data_Acum["Regional_Modelo"].astype(str)

# Se toman las columnas necesarias de Full_Data_Acum y se elimanan columnas
Full_Data_Acum_2 = Full_Data_Acum[['ID', 'FECHA_SMART', 'CARTERA', 'Inf_Cartera', 'Inf_Producto', 'Regional_Modelo', 'SMART_Color', 'S_Sucural']]
Full_Data_Acum_3 = Full_Data_Acum_2.drop_duplicates(subset=['ID', 'FECHA_SMART', 'CARTERA', 'Inf_Cartera', 'Inf_Producto', 'Regional_Modelo', 'SMART_Color', 'S_Sucural'])

# Se realiza un agrupado por ID y se realiza una suma de Inf_Total_Valorizada_S y conteo de ID_Linea
Full_Data_Acum_No_Agrup = Full_Data_Acum.groupby('ID').agg(
    Inf_Total_Valorizada_S=pd.NamedAgg(column="Inf_Total_Valorizada_S", aggfunc="sum"),
    Conteo=pd.NamedAgg(column="ID_Linea", aggfunc="count")
)
# Se toma como índice ID
Full_Data_Acum_No_Agrup["ID"]=list(Full_Data_Acum_No_Agrup.index)
Full_Data_Acum_No_Agrup.reset_index(inplace=True, drop=True)

# Se hace un left Jpin entre Full_Data_Acum_3 y Full_Data_Acum_No_Agrup
Full_Data_Acum_4 = pd.merge(Full_Data_Acum_3, Full_Data_Acum_No_Agrup, left_on='ID', right_on='ID', how="left")
Full_Data_Acum_5 = Full_Data_Acum_4.drop(columns=['ID'])
# Se renombran columnas
Full_Data_Acum_5.rename(columns={'FECHA_SMART': 'Fecha', 'CARTERA': 'Tipo de Cartera', 'Inf_Cartera': 'Estatus', 'Regional_Modelo': 'Regional', 'SMART_Color': 'SMART', 'S_Sucural': 'Sucursal', 'Inf_Total_Valorizada_S':'Saldo Total MXN', 'Conteo':'Lineas'}, inplace=True)

Full_Data_Acum_5.to_excel(Ruta_Outputs + '5_Reporte_'+str(fecha.year)+'_'+str(fecha.month)+'.xlsx', sheet_name='Full_Data', index=False)
