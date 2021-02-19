import base64
import os
import shutil
import datetime
import re
from pdf2image import convert_from_path
import pytesseract
from PIL import Image


# Proceso de decodificacion
def sender_decode(sender):
    parsed_string = sender.split("?")
    decoded = base64.b64decode(parsed_string[3]).decode(parsed_string[1], "ignore")
    return decoded


# Crear carpeta resultado
def crea_carpeta_resultado(vroot):
    if not os.path.exists(os.path.join(vroot)):
        os.makedirs(os.path.join(vroot))
    if not os.path.exists(os.path.join(vroot, 'Aceptacion')):
        os.makedirs(os.path.join(vroot, 'Aceptacion'))
    if not os.path.exists(os.path.join(vroot, 'Aceptacion', 'TempAceptacion')):
        os.makedirs(os.path.join(vroot, 'Aceptacion', 'TempAceptacion'))
    if not os.path.exists(os.path.join(vroot, 'Pedido')):
        os.makedirs(os.path.join(vroot, 'Pedido'))
    if not os.path.exists(os.path.join(vroot, 'Pedido', 'TempPedido')):
        os.makedirs(os.path.join(vroot, 'Pedido', 'TempPedido'))


# Eliminar carpetas temporales
def elimina_carpetas_temporales(vroot):
    if os.path.exists(os.path.join(vroot, 'Aceptacion', 'TempAceptacion')):
        shutil.rmtree(os.path.join(vroot, 'Aceptacion', 'TempAceptacion'))
    if os.path.exists(os.path.join(vroot, 'Pedido', 'TempPedido')):
        shutil.rmtree(os.path.join(vroot, 'Pedido', 'TempPedido'))


# Crear carpeta temporal
def crea_carpeta(vroot, carpeta_file):
    # carpeta_file = 'TempCompras'
    if not os.path.exists(os.path.join(vroot, carpeta_file)):
        os.makedirs(os.path.join(vroot, carpeta_file))


# Eliminar carpeta temporal
def elimina_carpeta(vroot, carpeta_file):
    if os.path.exists(os.path.join(vroot, carpeta_file)):
        shutil.rmtree(os.path.join(vroot, carpeta_file))


# Convierte una fecha a un nombre de carpeta
def convertir_fecha_carpeta():
    mydate = datetime.datetime.now()
    nombre = 'Actas&Pedidos_' + mydate.strftime('%Y%m%d')
    return nombre


# TRANSFORMA UN FICHERO PDF EN TEXTO POR MEDIO DE OCR
def fichero_pdf_imagen_texto(PDF_file, ficheros_respaldo, carpeta_trabajo, nameTemp):
    d = dict()
    fichero_nombre, fichero_extension = os.path.splitext(os.path.basename(PDF_file))
    tempDir = os.path.join(carpeta_trabajo, nameTemp, fichero_nombre)
    if not os.path.exists(tempDir):
        os.makedirs(tempDir)
    try:
        '''
        Part #1 : Converting PDF to images
        '''
        # Store all the pages of the PDF in a variable
        pages = convert_from_path(PDF_file, poppler_path=os.path.join(ficheros_respaldo, 'poppler-0.68.0/bin'),
                                  strict=False)
        # Counter to store images of each page of PDF to image
        image_counter = 1
        # Iterate through all the pages stored above
        for page in pages:
            filename = "page_" + str(image_counter) + ".jpg"
            # Save the image of the page in system
            page.save(os.path.join(tempDir, fichero_nombre + '_' + filename), 'JPEG')
            # Increment the counter to update filename
            image_counter = image_counter + 1

        ''' 
        Part #2 - Recognizing text from the images using OCR 
        '''
        # Variable to get count of total number of pages
        filelimit = image_counter - 1
        # Creating a text file to write the output
        outfile = os.path.join(tempDir, fichero_nombre + "_text.txt")
        # Open the file in append mode so that
        # All contents of all images are added to the same file
        f = open(outfile, "a")
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
        # https://tesseract-ocr.github.io/tessdoc/4.0-with-LSTM.html#400-alpha-for-windows
        # https://github.com/UB-Mannheim/tesseract/wiki

        # Iterate from 1 to total number of pages
        for i in range(1, filelimit + 1):
            # Set filename to recognize text from
            # Again, these files will be:
            # page_1.jpg
            # page_2.jpg
            # ....
            # page_n.jpg
            filename = "page_" + str(i) + ".jpg"

            # Recognize the text as string in image using pytesserct
            text = str(((pytesseract.image_to_string(
                Image.open(os.path.join(tempDir, fichero_nombre + '_' + filename)))))).upper()

            # The recognized text is stored in variable text
            # Any string processing may be applied on text
            # Here, basic formatting has been done:
            # In many PDFs, at line ending, if a word can't
            # be written fully, a 'hyphen' is added.
            # The rest of the word is written in the next line
            # Eg: This is a sample text this word here GeeksF-
            # orGeeks is half on first line, remaining on next.
            # To remove this, we replace every '-\n' to ''.
            text = text.replace('-\n', '')

            # Finally, write the processed text to the file.
            f.write(text)

        # Close the file after writing all the text.
        f.close()
        d['Fichero_Texto'] = outfile
        d['Error'] = ''

    except Exception as e:
        d['Fichero_Texto'] = ''
        d['Error'] = 'Error al convertir en texto el fichero ' + PDF_file + '   ' + e

    return d


def pdf_renombra_mueve(PDF_file, nombre_nuevo, destino):
    fichero_nombre, fichero_extension = os.path.splitext(os.path.basename(PDF_file))

    shutil.copyfile(PDF_file, os.path.join(destino, nombre_nuevo + fichero_extension))


def extrae_los_datos_compras(fichero):
    vcompras = {
        'N_Pedido': '',
        'Importe': '',
        'Observaciones': '',
        'Empresa_Emite': '',
        'Orden_Entrega': '',
        'Documento_Referencia': '',
        'Descripcion': '',
        'Cantidad': '',
        'Precio_Unitario': '',
        'Subtotal': '',
        'Fecha_Grabacion': '',
        'Codigo_A': '',
    }
    # Extrae N_Pedido
    datoreg = 'DE PEDIDO'
    lista_encontrados = []
    try:
        pattern = re.compile(datoreg, re.IGNORECASE)
        with open(fichero, "rt") as myfile:
            for line in myfile:
                if pattern.search(line) != None:
                    lista_encontrados.append(line)
                    break

        r1 = re.findall(r"\d{6,}", lista_encontrados[0])
        vcompras['N_Pedido'] = r1[0]
    except Exception as e:
        print(e)
        vcompras['N_Pedido'] = ''

    # Extrae Importe
    datoreg = r"\|.*"
    lista_encontrados = []
    try:
        pattern = re.compile(datoreg, re.IGNORECASE)
        with open(fichero, "rt") as myfile:
            for line in myfile:
                if pattern.search(line) != None:
                    lista_encontrados.append(line)
                    break

        r1 = re.findall(r"\d{1,}\,\d{1,}", lista_encontrados[0])
        vcompras['Importe'] = r1[1]
    except Exception as e:
        print(e)
        vcompras['Importe'] = ''

    # Extrae Observaciones
    datoreg = r".*BQA.*"
    lista_encontrados = []
    try:
        pattern = re.compile(datoreg, re.IGNORECASE)
        with open(fichero, "rt") as myfile:
            for line in myfile:
                if pattern.search(line) != None:
                    lista_encontrados.append(line)
                    break
        vcompras['Observaciones'] = lista_encontrados[0].rstrip()

    except Exception as e:
        print(e)
        vcompras['Observaciones'] = ''

    # Extrae Empresa Emite
    datoreg = r"^ORANGE \w.*"
    lista_encontrados = []
    try:
        pattern = re.compile(datoreg, re.IGNORECASE)
        with open(fichero, "rt") as myfile:
            for line in myfile:
                if pattern.search(line) != None:
                    lista_encontrados.append(line)
                    break
        vcompras['Empresa_Emite'] = lista_encontrados[0].rstrip()
    except Exception as e:
        print(e)
        vcompras['Empresa_Emite'] = ''

    # Extrae Orden Entrega
    datoreg = r"^DOCUMENTO DE REFERENCIA.*"
    lista_encontrados = []
    try:
        pattern = re.compile(datoreg, re.IGNORECASE)
        with open(fichero, "rt") as myfile:
            for line in myfile:
                if pattern.search(line) != None:
                    lista_encontrados.append(line)
                    break

        r1 = re.findall(r"46\d{8,}", lista_encontrados[0])
        vcompras['Orden_Entrega'] = r1[0]
        vcompras['Documento_Referencia'] = r1[0]
    except Exception as e:
        print(e)
        vcompras['Orden_Entrega'] = ''
        vcompras['Documento_Referencia'] = ''

    # Extrae Descripcion
    datoreg = r"^DOCUMENTO DE REFERENCIA.*"
    lista_encontrados = []
    try:
        pattern = re.compile(datoreg, re.IGNORECASE)
        with open(fichero, "rt") as myfile:
            for line in myfile:
                if pattern.search(line) != None:
                    lista_encontrados.append(line)
                    break

        r1 = re.findall(r"46\d{8,}", lista_encontrados[0])
        vcompras['Descripcion'] = r1[0]

    except Exception as e:
        print(e)
        vcompras['Descripcion'] = ''
        


    return vcompras


def fichero_pdf_imagen_texto_oc(PDF_file, ficheros_respaldo, carpeta_trabajo, carpetas_trabajo):
    d = dict()
    fichero_nombre, fichero_extension = os.path.splitext(os.path.basename(PDF_file))
    try:
        '''
        Part #1 : Converting PDF to images
        '''
        # Store all the pages of the PDF in a variable
        pages = convert_from_path(PDF_file, poppler_path=os.path.join(ficheros_respaldo, 'poppler-0.68.0/bin'),
                                  strict=False)
        # Iterate through all the pages stored above
        filename = "page_1.jpg"
        # Save the image of the page in system
        # pages[0].save(os.path.join(carpeta_trabajo + '_' + fichero_nombre), 'JPEG')
        pages[0].save(os.path.join(carpeta_trabajo, fichero_nombre + '.jpg'), 'JPEG')
        ''' 
        Part #2 - Recognizing text from the images using OCR 
        '''
        # Creating a text file to write the output
        outfile = os.path.join(carpeta_trabajo, fichero_nombre + ".txt")
        # Open the file in append mode so that
        # All contents of all images are added to the same file
        f = open(outfile, "a")
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'

        filename = "page_1.jpg"
        # Recognize the text as string in image using pytesserct
        text = str(((pytesseract.image_to_string(
            Image.open(os.path.join(carpeta_trabajo, fichero_nombre + '.jpg')))))).upper()
        text = text.replace('-\n', '')

        # Finally, write the processed text to the file.
        f.write(text)
        # Close the file after writing all the text.
        f.close()
        d['Fichero_Texto'] = outfile
        d['Error'] = ''

        # EXTRACCION DE DATOS
        vcompras = extrae_los_datos_compras(outfile)

        try:
            pdf_renombra_mueve(PDF_file, vcompras['N_Pedido'], carpetas_trabajo['Pedido'])
        except Exception as e1:
            print('No se ha podido mover el fichero {}, error {}'.format(PDF_file ,e1))


    except Exception as e:
        d['Fichero_Texto'] = ''
        d['Error'] = 'Error al convertir en texto el fichero ' + PDF_file + '   ' + e

    return d


# BUSCA DATOS EN EL TEXTO OBTENIDO POR IA
def busca_datos_pdf_texto(datoreg, filename):
    d = dict()
    lista_encontrados = []
    d['Error'] = ''
    try:
        pattern = re.compile(datoreg, re.IGNORECASE)
        with open(filename, "rt") as myfile:
            for line in myfile:
                if pattern.search(line) != None:
                    # print(line, end='')
                    lista_encontrados.append(line)
    except Exception as e:
        d['ListaEncontrados'] = lista_encontrados
        d['Error'] = 'Error al buscar datos en  ' + filename + '   ' + e
    d['ListaEncontrados'] = lista_encontrados
    return d


# Función para ubicar cada fichero con un extension determinada, retorna una lista
def lista_extension(rootDir, extension):
    # r=>root, d=>directories, f=>files
    files_in_dir = []
    for r, d, f in os.walk(rootDir):
        for item in f:
            if '.' + extension in item:
                files_in_dir.append(os.path.join(r, item))
    return files_in_dir


def extrae_texto_imagen(fichero, rootDir):
    d = dict()
    d['Texto'] = ''
    fichero_nombre, fichero_extension = os.path.splitext(os.path.basename(fichero))
    d['Fichero'] = fichero_nombre + '.pdf'
    d['Error'] = 'OK'
    try:
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
        ruta_imagen = os.path.join(rootDir, 'Temporal_base',
                                   fichero_nombre)  # r'D:/EMR_Auditorias_Python/Auditorias/PO6100/Carpeta_de_Trabajo\\Temporal_base\\GALE6100E_GALR6100E_ER1_M_ARCA_112601_1'
        texto_a_revisar = ''
        for imagen_pdf in lista_extension(ruta_imagen, 'jpeg'):
            texto_a_revisar = texto_a_revisar + '     ' + pytesseract.image_to_string(imagen_pdf)

        texto_a_revisar = texto_a_revisar.upper()
        d['Texto'] = texto_a_revisar
    except Exception as e:
        d['Texto'] = ''
        d['Error'] = 'Error al intentar leer imagen, verifique que se ha instalado en el equipo  el Tesseract-OCR' + e
    return d


def busca_datos_pdf_texto(datoreg, filename):
    d = dict()
    lista_encontrados = []
    d['Error'] = ''
    try:
        pattern = re.compile(datoreg, re.IGNORECASE)
        with open(filename, "rt") as myfile:
            for line in myfile:
                if pattern.search(line) != None:
                    # print(line, end='')
                    lista_encontrados.append(line)
    except Exception as e:
        d['ListaEncontrados'] = lista_encontrados
        d['Error'] = 'Error al buscar datos en  ' + filename + '   ' + e
    d['ListaEncontrados'] = lista_encontrados
    return d

# PDFS = lista_extension(rootDir='C:/CORREOS_AUTOMATIZACION/Actas&Pedidos_20210216/TempCompras', extension='PDF')
# for elemento in PDFS:
#    print(elemento)
#   rr = fichero_pdf_imagen_texto_oc(PDF_file=elemento,ficheros_respaldo='C:/CORREOS_AUTOMATIZACION/Ficheros_Respaldo',
#                                     carpeta_trabajo='C:/CORREOS_AUTOMATIZACION/Actas&Pedidos_20210216/TempCompras')

# intll = busca_datos_pdf_texto('DE PEDIDO', 'C:/CORREOS_AUTOMATIZACION/Actas&Pedidos_20210216/TempCompras/DOCUMENTO1.txt')
# if (len(intll['ListaEncontrados']) >= 1):
#    print(intll)

c = extrae_los_datos_compras('C:/CORREOS_AUTOMATIZACION/Actas&Pedidos_20210219/Pedido/TempPedido/DOCUMENTO1.txt')
print(c)

#pdf_renombra_mueve(r'C:\CORREOS_AUTOMATIZACION\Actas&Pedidos_20210218\Pedido\TempPedido\DOCUMENTO1.PDF', '7777',
#                   r'C:\CORREOS_AUTOMATIZACION\Actas&Pedidos_20210218\Pedido')
