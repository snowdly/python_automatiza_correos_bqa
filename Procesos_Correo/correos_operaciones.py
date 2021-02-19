import imaplib
import email
import base64
import os
import shutil
from Procesos_Correo import procesos_comunes


# Proceso que lee todos los correos de un buzon
def leer_mover_correos(vhost, vuser, vpassword, vbuzon, vroot, vcriterio, vdestino):
    respuesta = {
        'resultado': 'OK',
        'cantidad_correos': 0,
        'error': ''
    }
    try:
        mail = imaplib.IMAP4_SSL(vhost, 993)
        mail.login(vuser, vpassword)
        typ, data = mail.select(vbuzon)
        if typ != 'OK':
            raise ValueError('No existe la carpeta {}'.format(vbuzon))
        num_msgs = int(data[0])
        print('Existen {} correos en INBOX'.format(num_msgs))

        # Crear la carpeta
        carpeta_file = 'TempCompras'
        procesos_comunes.elimina_carpeta(vroot, carpeta_file)
        procesos_comunes.crea_carpeta(vroot, carpeta_file)
        svdir = os.path.join(vroot, carpeta_file)

        # Filtrado por Subject
        typ, msgs = mail.search(None, vcriterio)
        print('Existen {} correos con el Subject indicado '.format(msgs[0]))

        # Descargado de adjuntos
        contador = 0
        msgs = msgs[0].split()
        for emailid in msgs:
            resp, data = mail.fetch(emailid, "(RFC822)")
            email_body = data[0][1]
            m = email.message_from_bytes(email_body)
            if m.get_content_maintype() != 'multipart': continue

            for part in m.walk():
                if part.get_content_maintype() == 'multipart': continue
                if part.get('Content-Transfer-Encoding') is None: continue

                filename: str = part.get_filename()
                if filename is None: continue
                try:
                    filename = procesos_comunes.sender_decode(filename)
                except IndexError:
                    filename = part.get_filename()
                if filename is not None:
                    contador = contador + 1
                    sv_path = os.path.join(svdir, filename)
                    fp = open(sv_path, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
                    fichero_nombre, fichero_extension = os.path.splitext(sv_path)
                    os.rename(os.path.join(svdir, fichero_nombre + fichero_extension), os.path.join(svdir,
                                                                                                    fichero_nombre + str(
                                                                                                        contador) + fichero_extension.upper()))

                    '''
                    if not os.path.isfile(sv_path):
                        fp = open(sv_path, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
                        #cambiar de nombre al fichero PDF
                        fichero_nombre, fichero_extension = os.path.splitext(sv_path)
                        os.rename(os.path.join(svdir, fichero_nombre + fichero_extension), os.path.join(svdir,
                                fichero_nombre + str(contador) + fichero_extension))
                    '''
        # Mover el mensaje
        for emailid in msgs:
            print('Moviendo:', emailid)
            # result = mail.uid('COPY', emailid, 'Compras')
            typ, response = mail.copy(emailid, vdestino)
            if typ == 'OK':
                # What are the current flags?
                # typ, response = mail.fetch(emailid, '(FLAGS)')
                # print('Flags before:', response)
                typ, response = mail.store(emailid, '+FLAGS', r'(\Deleted)')
                # typ, response = mail.fetch(emailid, '(FLAGS)')
                # print('Flags after:', response)
                # Really delete the message.
                # typ, response = mail.expunge()
                # print('Expunged:', response)

        mail.close()
        mail.logout()

    except Exception as e:
        respuesta = {
            'resultado': 'KO',
            'cantidad_correos': 0,
            'error': e
        }
    # finally:
    # Eliminar carpeta
    # elimina_carpeta(vroot, carpeta_file)
    return respuesta


# Proceso que lee todos los correos de un buzon
def leer_mover_correos_solo(vhost, vuser, vpassword, vbuzon, vcriterio, vdestino):
    respuesta = {
        'resultado': 'OK',
        'cantidad_correos': 0,
        'error': ''
    }
    try:
        mail = imaplib.IMAP4_SSL(vhost, 993)
        mail.login(vuser, vpassword)
        typ, data = mail.select(vbuzon)
        if typ != 'OK':
            raise ValueError('No existe la carpeta {}'.format(vbuzon))
        num_msgs = int(data[0])
        print('Existen {} correos en INBOX'.format(num_msgs))

        # Filtrado por Subject
        typ, msgs = mail.search(None, vcriterio)

        msgs = msgs[0].split()
        print('Existen {} correos con el Subject indicado '.format(len(msgs)))
        # Mover el mensaje
        for emailid in msgs:
            print('Moviendo:', emailid)
            typ, response = mail.copy(emailid, vdestino)
            if typ == 'OK':
                typ, response = mail.store(emailid, '+FLAGS', r'(\Deleted)')

    except Exception as e:
        respuesta = {
            'resultado': 'KO',
            'cantidad_correos': 0,
            'error': e
        }
    finally:
        mail.close()
        mail.logout()

    return respuesta


def open_connection(vhostname, vusername, vpassword):
    # Read the config file
    # config = configparser.ConfigParser()
    # config.read([os.path.expanduser('~/.pymotw')])

    # Connect to the server
    # hostname = config.get('server', 'hostname')
    # vhostname = 'imap-mail.outlook.com'
    # if verbose:
    #    print('Connecting to', hostname)
    connection = imaplib.IMAP4_SSL(vhostname, 993)

    # Login to our account
    # username = config.get('account', 'username')
    # username = 'ocamn@nae.es'
    # password = config.get('account', 'password')
    # password = 'esFebrero2021!'
    # if verbose:
    #    print('Logging in as', vusername)
    try:
        connection.login(vusername, vpassword)
    except Exception as err:
        print('ERROR:', err)
    return connection


def mover_correos_solo(vhost, vuser, vpassword, vbuzon, vroot, vcriterio, vdestino):
    with open_connection(vhost, vuser, vpassword) as mail:
        print(mail)
        typ, data = mail.select(vbuzon)
        if typ != 'OK':
            raise ValueError('No existe la carpeta {}'.format(vbuzon))
        num_msgs = int(data[0])
        print('Existen {} correos en INBOX'.format(num_msgs))
        # Filtrado por Subject
        typ, msgs = mail.search(None, vcriterio)
        msgs = msgs[0].split()
        print('Existen {} correos con el Subject indicado '.format(len(msgs)))

        # Mover el mensaje
        # for emailid in msgs:
        #     print('Moviendo:', emailid)
        #     typ, response = mail.copy(emailid, vdestino)
        #     if typ == 'OK':
        #         typ, response = mail.store(emailid, '+FLAGS', r'(\Deleted)')


# Proceso de descarga de PDF
def descargar_adjuntos(vhost, vuser, vpassword, vbuzon, vroot, vcriterio, vcarpetatemp):
    respuesta = {
        'resultado': 'OK',
        'cantidad_correos': 0,
        'error': ''
    }
    try:
        # Crear la carpeta
        # carpeta_file = vcarpetatemp
        # procesos_comunes.elimina_carpeta(vroot, carpeta_file)
        # procesos_comunes.crea_carpeta(vroot, carpeta_file)
        # svdir = os.path.join(vroot, carpeta_file)
        svdir = vcarpetatemp

        mail = imaplib.IMAP4_SSL(vhost, 993)
        mail.login(vuser, vpassword)
        typ, data = mail.select(vbuzon)
        if typ != 'OK':
            raise ValueError('No existe la carpeta {}'.format(vbuzon))
        num_msgs = int(data[0])
        print('Existen {} correos en {}'.format(num_msgs, vbuzon))
        respuesta['cantidad_correos'] = num_msgs

        # Filtrado por Subject
        typ, msgs = mail.search(None, vcriterio)
        #print('Existen {} correos con el Subject indicado '.format(msgs[0]))

        # Descargado de adjuntos
        contador = 0
        msgs = msgs[0].split()
        for emailid in msgs:
            resp, data = mail.fetch(emailid, "(RFC822)")
            email_body = data[0][1]
            m = email.message_from_bytes(email_body)
            if m.get_content_maintype() != 'multipart': continue

            for part in m.walk():
                if part.get_content_maintype() == 'multipart': continue
                if part.get('Content-Transfer-Encoding') is None: continue

                filename: str = part.get_filename()
                if filename is None: continue
                try:
                    filename = procesos_comunes.sender_decode(filename)
                except IndexError:
                    filename = part.get_filename()
                if filename is not None:
                    contador = contador + 1
                    sv_path = os.path.join(svdir, filename)
                    fp = open(sv_path, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
                    fichero_nombre, fichero_extension = os.path.splitext(sv_path)
                    os.rename(os.path.join(svdir, fichero_nombre + fichero_extension), os.path.join(svdir,
                                                                                                    fichero_nombre + str(
                                                                                                        contador) + fichero_extension.upper()))
                    print("Se ha descargado: {}".format(fichero_nombre+str(contador)+ fichero_extension.upper()))
        print("Total descargado {} documentos".format(str(contador)))

    except Exception as e:
        respuesta = {
            'resultado': 'KO',
            'cantidad_correos': 0,
            'error': e
        }
    finally:
        mail.close()
        mail.logout()
    return respuesta
