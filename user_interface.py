""" 
Sistema de Alertas Tempranas Monex México
Programa Inicial: Pantalla de Inputs

Este código tiene el objetivo de generar la interfaz de usuario mediante la librería tkinter
donde ingresará los valores necesarios para la ejecución de los programas que evaluan, cargar y producen modelos.
Se estructura en 2 partes:
    1) Librerías.
    2) Generación de Pantalla de Usuario.
"""

#################################################
#    1 Librerías necesarias para la ejecución   #
#################################################

import tkinter as tk            # Librería para realizar un GUI.
from PIL import Image, ImageTk  # Librería para interactuar con imágenes en la GUI.
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
from functools import partial   # Librería que permite interactuar con funciones.
from file_check import revision_archivos # Se carga la función del archivo revision_archivos_0
from model_extraction import extraccion_modelos # Se carga la función del archivo extraccion_modelos_20_1
from model_assesstment import ejecucion_modelo # Se carga la función del archivo ejecucion_modelo_2
   


#################################################
#      2 Generación de Pantalla de Usuario      #
#################################################

"""
Al objeto raiz se le asignan todas las propiedades de la clase tk y es la base para
la generación de la interfaz de usuario.
"""

raiz = tk.Tk() # Se asignan propiedades.
raiz.title("Panel de control SMART") # Se asigna título a la pantalla
raiz.resizable(0,0) # Se estable que el usuario no pueda expandir ni reducir el tamaño de pantalla.
raiz.geometry("700x500") # Se establece tamaño para pantalla widthxheight .

ruta = tk.StringVar() # Al objeto ruta se le asigna que será de tipo string.
anio = tk.IntVar()    # Al objeto anio se le asigna que será de tipo int.
mes = tk.IntVar()     # Al objeto me se le asigna que será de tipo int.

x_ = 30 # Se esablece una posición inicial x para los inputs y botones
y_ = 70 # Se esablece una posición inicial y para los inputs y botones

# Se generan etiquetas, sus propiedades de texto y fuente, además de su posición en pantalla.
ruta_label = tk.Label(raiz, text = "Ingresa Ruta", font=('Arial bold',14) ).place(x = x_, y = y_)       
anio_label = tk.Label(raiz, text = "Ingresa Año", font=('Arial bold',14) ).place(x = x_, y = y_+70) 
mes_label = tk.Label(raiz, text = "Ingresa Mes", font=('Arial bold',14) ).place(x = x_, y = y_+140)  

# Se generan cuadros de input, sus valores y tipo de imput, además de su posición en pantalla.
ruta_entry = tk.Entry(raiz, textvariable = ruta).place(x = x_ + 130, y = y_, width = 400, height = 25)
anio_entry = tk.Entry(raiz, textvariable = anio).place(x = x_ + 130, y = y_+70, width = 400, height = 25) 
mes_entry = tk.Entry(raiz, textvariable = mes).place(x = x_ + 130, y = y_+140, width = 400, height = 25)

# Se genera etiqueta con sus propiedades de fuente, además de su posición en pantalla. Servirá para mostrar
# un texto de resultado de acuerdo a lo que el usuario ingrese.
labelResult = tk.Label(raiz, font=('Arial bold',12))
labelResult.place(x = x_ + 240, y = y_ + 340 )

# Ahora se asignan las funciones que se ejecutarán en cada uno de los botones,
# estas funciones están en un archivo diferente para el proyecto.
# el orden de las funciones y botones es la forma correcta de ejecutar el proceso.
revision_archivo = partial(revision_archivos,labelResult, ruta, anio, mes)
extraccion_modelo = partial(extraccion_modelos,labelResult, ruta, anio, mes)
ejecucion_modelos = partial(ejecucion_modelo,labelResult, ruta, anio, mes)

# Se generan botones con sus propiedades y posición, cada botón ejecuta la tarea
# definida en cada uno de los archivos.

# Ejecuta la revisión de archivos usando el programa revision_archivos_0
archivos_sbmitbtn = tk.Button(raiz, text = "Revisión de Archivos",activebackground = "green",\
    activeforeground = "blue", font=('Arial bold',12), command= revision_archivo ).place(x = x_ + 50, y = y_ + 220)

# Ejecuta la carga de modelos 20 usando el programa extraccion_modelos_20_1
modelos_sbmitbtn1 = tk.Button(raiz, text = "Leer Modelos 20",activebackground = "green",\
    activeforeground = "blue", font=('Arial bold',12), command= extraccion_modelo ).place(x = x_ + 250, y = y_ + 220)

# Ejecuta la evaluación de modelo usando el programa ejecucion_modelo_2
ejecucion_sbmitbtn2 = tk.Button(raiz, text = "Ejecución de Modelo",activebackground = "green",\
    activeforeground = "blue", font=('Arial bold',12), command= ejecucion_modelos ).place(x = x_ + 450, y = y_ + 220)

# Código para cargar imágen de MONEX
load = Image.open("monex.jpg")
load = load.resize((130, 200))
render = ImageTk.PhotoImage(load)
img = tk.Label(raiz, image = render)
img.image = render
img.place(x=0, y=0, width = 150, height = 40)



raiz.mainloop() # Inicia la pantalla

