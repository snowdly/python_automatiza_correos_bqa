from Interfaz import ventana_consulta
from Procesos_Correo import procesos_comunes
from Procesos_Correo import correos_operaciones
import os

# Datos geberales
nominal = 'C:/CORREOS_AUTOMATIZACION'
nominal_respaldo = os.path.join(nominal, 'Ficheros_Respaldo')
vroot = os.path.join(nominal, procesos_comunes.convertir_fecha_carpeta())
host = 'imap-mail.outlook.com'
user = 'pedidos.orange@nae.es'
password = 'Inform@tica2019'
carpetas_trabajo ={
        'vroot': vroot,
        'Aceptacion': os.path.join(vroot, 'Aceptacion'),
        'TempAceptacion': os.path.join(vroot, 'Aceptacion', 'TempAceptacion'),
        'Pedido': os.path.join(vroot, 'Pedido'),
        'TempPedido': os.path.join(vroot, 'Pedido', 'TempPedido')
    }

# # Creacion de carpetas
procesos_comunes.elimina_carpetas_temporales(vroot)
procesos_comunes.crea_carpeta_resultado(vroot)

Continuar_Ventana, Directorio = ventana_consulta.seleccion_fichero(vroot)
if Continuar_Ventana:
    print('Se va trabajar en la ruta: ' + Directorio)
    # # Compras
    # # 1. Leer correos y mover
    # r = correos_operaciones.leer_mover_correos_solo(vhost=host, vuser=user, vpassword=password,
    #                                                 vbuzon="INBOX",
    #                                                 vcriterio='(FROM "WF-BATCH@orange.com" SUBJECT "Pedido de compra")',
    #                                                 vdestino='INBOX/Compras')
    # print(r)

    # 2. Descargar los PDF desde el buzon Compras
    p = correos_operaciones.descargar_adjuntos(vhost=host, vuser=user, vpassword=password,
                                               vbuzon="INBOX/Compras", vroot=vroot,
                                               vcriterio='(FROM "WF-BATCH@orange.com" SUBJECT "Pedido de compra")',
                                               vcarpetatemp= carpetas_trabajo['TempPedido'])
    print(p)

    # 3. Convertir PDF en Imagen Texto y Capturar los datos relevantes
    PDFS = procesos_comunes.lista_extension(rootDir= carpetas_trabajo['TempPedido'], extension='PDF')
    for elemento in PDFS:
       print(elemento)
       rr = procesos_comunes.fichero_pdf_imagen_texto_oc(PDF_file=elemento, ficheros_respaldo=nominal_respaldo,
                                                         carpeta_trabajo=carpetas_trabajo['TempPedido'],
                                                         carpetas_trabajo=carpetas_trabajo)

    # # Actas/Pedidos
    # # 1. Leer correos y mover
    # r = correos_operaciones.leer_mover_correos_solo(vhost=host, vuser=user, vpassword=password,
    #                                                 vbuzon="INBOX",
    #                                                 vcriterio='(FROM "sistema_facturacion@seresnet.com" SUBJECT "de acta de aceptac")',
    #                                                 vdestino='INBOX/Actas')
    # print(r)


else:
    print("Rechazado")
