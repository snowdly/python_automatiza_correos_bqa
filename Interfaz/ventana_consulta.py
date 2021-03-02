import logging
import tkinter as tk
import tkinter.font as tkFont
from tkinter import filedialog
import os

# variables de flujo
Continuar_Proceso = False


def seleccion_fichero(directorio_inicial):
    Continuar_Proceso = False
    Fichero_Auditoria = ""
    root = tk.Tk()
    root.title("Automatización Correos Facturas BQA")
    width = 770
    height = 240
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    root.geometry(alignstr)
    root.resizable(width=False, height=False)
    # LabelFrame general
    ftl = tkFont.Font(family='Times', size=13)
    labelframe = tk.LabelFrame(root, text='Seleccionar Carpeta para resultados ...', font=ftl)
    labelframe.pack(expand='yes', fill='both')
    ft = tkFont.Font(family='Times', size=12)
    GLabel_Fichero = tk.Label(labelframe, font=ftl, text="Resultados en:", justify="center")
    GLabel_Fichero.place(x=25, y=35, width=120, height=30)
    TxtFichero = tk.Entry(labelframe, text="", font=ft)
    TxtFichero.place(x=25 + 130, y=35, width=430, height=25)

    def btnSeleccionar_command():
        nonlocal Continuar_Proceso
        nonlocal Fichero_Auditoria
        nonlocal directorio_inicial
        # Verifica si existe directorio
        # directorio_inicial = "C:/EMR_Auditorias_Python/"
        if not os.path.exists(directorio_inicial):
            directorio_inicial = "/"
        root.filename = filedialog.askdirectory(initialdir=directorio_inicial, title="Seleccionar Auditoría")
        # Change label contents
        if root.filename != '':
            TxtFichero.delete(0, tk.END)
        TxtFichero.insert(0, root.filename)
        Fichero_Auditoria = root.filename
        # GLabel_Fichero.configure(text= root.filename)

    TxtFichero.delete(0, tk.END)
    TxtFichero.insert(0, directorio_inicial)
    Fichero_Auditoria = directorio_inicial

    btnSeleccionar = tk.Button(labelframe, text="Seleccionar", font=ftl, command=btnSeleccionar_command)
    btnSeleccionar.place(x=25 + 120 + 450, y=35)

    # BOTONES
    def btnAceptar_command():
        nonlocal Continuar_Proceso
        Continuar_Proceso = True
        print("Fichero de Auditoría: " + Fichero_Auditoria)
        logging.debug("Fichero de Auditoría: " + Fichero_Auditoria)
        root.destroy()

    def btnCancelar_command():
        nonlocal Continuar_Proceso
        Continuar_Proceso = False
        root.destroy()

    btnAceptar = tk.Button(root, text="Aceptar", font=ftl, command=btnAceptar_command)
    btnAceptar.place(relx="0.3", rely="0.70")
    btnCancelar = tk.Button(root, text="Cancelar", font=ftl, command=btnCancelar_command)
    btnCancelar.place(relx="0.5", rely="0.70")

    # PROCESO PRINCIPAL
    root.mainloop()

    return Continuar_Proceso, Fichero_Auditoria
