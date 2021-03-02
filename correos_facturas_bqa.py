from Interfaz import ventana_consulta
from Procesos_Correo import procesos_comunes
from Procesos_Correo import correos_operaciones
import os
import pandas as pd

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
# procesos_comunes.elimina_carpetas_temporales(vroot)
procesos_comunes.crea_carpeta_resultado(vroot)

Continuar_Ventana, Directorio = ventana_consulta.seleccion_fichero(vroot)
if Continuar_Ventana:
    print('Se va trabajar en la ruta: ' + Directorio)
    # # COMPRAS
    # # ==================================================================================================================
    # # 1. Leer correos y mover
    # r = correos_operaciones.leer_mover_correos_solo(vhost=host, vuser=user, vpassword=password,
    #                                                 vbuzon="INBOX",
    #                                                 vcriterio='(FROM "WF-BATCH@orange.com" SUBJECT "Pedido de compra")',
    #                                                 vdestino='INBOX/Compras')
    # print(r)
    #
    # # 2. Descargar los PDF desde el buzon Compras
    # p = correos_operaciones.descargar_adjuntos(vhost=host, vuser=user, vpassword=password,
    #                                            vbuzon="INBOX/Compras", vroot=vroot,
    #                                            vcriterio='(FROM "WF-BATCH@orange.com" SUBJECT "Pedido de compra")',
    #                                            vcarpetatemp= carpetas_trabajo['TempPedido'])
    # print(p)

    # 3. Convertir PDF en Imagen Texto y Capturar los datos relevantes
    PDFS = procesos_comunes.lista_extension(rootDir= carpetas_trabajo['TempPedido'], extension='PDF')
    lconsolidado = []
    for elemento in PDFS:
       #print(elemento)
       rr = procesos_comunes.fichero_pdf_imagen_texto_oc(PDF_file=elemento, ficheros_respaldo=nominal_respaldo,
                                                         carpeta_trabajo=carpetas_trabajo['TempPedido'])
       # Extraer datos y mover
       if rr['Resultado']=='OK':
           vcompras = procesos_comunes.extrae_los_datos_compras(rr['Fichero_Texto'])
           lconsolidado.append(vcompras)
           try:
               procesos_comunes.pdf_renombra_mueve(elemento, vcompras['N_Pedido'], carpetas_trabajo['Pedido'])
           except Exception as e1:
               print('No se ha podido mover el fichero {}, error {}'.format(elemento, e1))
       else:
           print(rr['Error'])

    print(lconsolidado)
    # 4. Generar el excel consolidado
    df = pd.DataFrame(lconsolidado, columns=['N_Pedido', 'Importe', 'Observaciones', 'Empresa_Emite',
                                             'Orden_Entrega', 'Documento_Referencia', 'Descripcion',
                                             'Cantidad', 'Precio_Unitario', 'Subtotal', 'Fecha_Grabacion',
                                             'Codigo_A'])
    df['Importe'] = df['Importe'].astype('float32')
    df['Cantidad'] = df['Cantidad'].astype('float32')
    df['Precio_Unitario'] = df['Precio_Unitario'].astype('float32')
    df['Subtotal'] = df['Subtotal'].astype('float32')
    writer = pd.ExcelWriter(os.path.join(carpetas_trabajo['Pedido'],
                                         'Pedidos_' + procesos_comunes.convertir_fecha_nombre() + '.xlsx'),
                            engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Pedidos', index=False)
    writer.save()
    # ==================================================================================================================

    # # ACTAS ACEPTACION
    # # ==================================================================================================================
    # # 1. Leer correos y mover
    # r = correos_operaciones.leer_mover_correos_solo(vhost=host, vuser=user, vpassword=password,
    #                                                 vbuzon="INBOX",
    #                                                 vcriterio='(FROM "sistema_facturacion@seresnet.com" SUBJECT "de acta de aceptac")',
    #                                                 vdestino='INBOX/Actas')
    # print(r)
    # # 2. Descargar los PDF desde el buzon Compras
    # p = correos_operaciones.descargar_adjuntos_no_cambio(vhost=host, vuser=user, vpassword=password,
    #                                            vbuzon="INBOX/Actas", vroot=vroot,
    #                                            vcriterio='(FROM "sistema_facturacion@seresnet.com" SUBJECT "de acta de aceptac")',
    #                                            vcarpetatemp= carpetas_trabajo['TempAceptacion'])
    # print(p)
    # # 3. Convertir PDF en Imagen Texto y Capturar los datos relevantes
    # PDFS = procesos_comunes.lista_extension(rootDir=carpetas_trabajo['TempAceptacion'], extension='PDF')
    # lconsolidado = []
    # for elemento in PDFS:
    #     print(elemento)
    #     rr = procesos_comunes.fichero_pdf_imagen_texto_oc(PDF_file=elemento, ficheros_respaldo=nominal_respaldo,
    #                                                       carpeta_trabajo=carpetas_trabajo['TempAceptacion'])
    #     # Extraer datos y mover
    #     if rr['Resultado'] == 'OK':
    #         vaceptacion = procesos_comunes.extrae_los_datos_aceptacion(rr['Fichero_Texto'])
    #         lconsolidado.append(vaceptacion)
    #     #     try:
    #     #         procesos_comunes.pdf_renombra_mueve(elemento, vcompras['N_Pedido'], carpetas_trabajo['Pedido'])
    #     #     except Exception as e1:
    #     #         print('No se ha podido mover el fichero {}, error {}'.format(elemento, e1))
    #     else:
    #         print(rr['Error'])
    #
    # print(lconsolidado)
    # # 4. Generar el excel consolidado
    # # df = pd.DataFrame(lconsolidado, columns=['N_Pedido', 'Importe', 'Observaciones', 'Empresa_Emite',
    # #                                          'Orden_Entrega', 'Documento_Referencia', 'Descripcion',
    # #                                          'Cantidad', 'Precio_Unitario', 'Subtotal', 'Fecha_Grabacion',
    # #                                          'Codigo_A'])
    # # df['Importe'] = df['Importe'].astype('float32')
    # # df['Cantidad'] = df['Cantidad'].astype('float32')
    # # df['Precio_Unitario'] = df['Precio_Unitario'].astype('float32')
    # # df['Subtotal'] = df['Subtotal'].astype('float32')
    # # writer = pd.ExcelWriter(os.path.join(carpetas_trabajo['Pedido'],
    # #                                      'Pedidos_' + procesos_comunes.convertir_fecha_nombre() + '.xlsx'),
    # #                         engine='xlsxwriter')
    # # df.to_excel(writer, sheet_name='Pedidos', index=False)
    # # writer.save()

else:
    print("Rechazado")
