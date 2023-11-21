import subprocess
from configuracion_logging import logger_errores, logger_transacciones
import datetime


def import_and_install(package, install_command):
    try:
        __import__(package)
    except ImportError:
        print(f"Falta la biblioteca '{package}'. Instalando...")
        subprocess.check_call(["pip", "install", install_command])
        try:
            __import__(package)
        except ImportError:
            print(
                f"No se pudo instalar '{package}'. Asegúrate de que se haya instalado correctamente."
            )


try:
    import datetime
    import zipfile
    import win32com.client
    import os
    import time
    import subprocess
    import sys
    import re
    import asyncio
    import math
    import threading
    import pythoncom
    import win32api
    import win32con
    import pyautogui

except ImportError as e:
    missing_libraries = str(e).split("No module named ")[1].split(", ")
    logger_transacciones.info(
        f"Las siguientes bibliotecas faltan: {', '.join(missing_libraries)}. Procediendo a instalarlas..."
    )

    libraries_to_install = [
        ("win32com.client", "pywin32"),
        ("pyautogui", "pyautogui"),
        "datetime",
        "pythoncom",
        "win32api",
        "win32con",
        "pywin32",
        "zipfile",
        "time",
        "re",
        "asyncio",
        "math",
        "threading",
    ]

    for package in libraries_to_install:
        if isinstance(package, tuple):
            package_name, install_command = package
        else:
            package_name, install_command = package, package
        import_and_install(package_name, install_command)

try:
    # Obtener los parámetros del archivo .bat
    ruta_outlook = sys.argv[1]
    carpeta_outlook_origen = sys.argv[2]
    carpeta_outlook_dest = sys.argv[3]
    carpeta_destino = sys.argv[4]
    logger_transacciones.info("Parametros del .bat obtenidos con exito")
except Exception as e:
    logger_errores.error(f"Error al obtener los parametros del archivo .bat {e}")

dominio_bios = "grupobios.co"

nSync = 0

# Abrir Outlook
# subprocess.Popen([ruta_outlook])

# time.sleep(2)

# # Verificar si la carpeta de destino existe y crearla si no
if not os.path.exists(carpeta_destino):
    try:
        logger_transacciones.info(
            f"La carpeta no existe , se creara la carpeta de destino {carpeta_destino}"
        )
        os.makedirs(carpeta_destino)
    except Exception as Error:
        logger_errores.error(
            f"Error al crear la carpeta Destino {carpeta_destino} Motivo >> {Error}"
        )


# Verificar si un string contiene otro
def validar_contenido(string, subcadena):
    return subcadena.lower() in string.lower()


# Función para obtener la ruta destino del archivo
def obtener_destino(asunto, dominio):
    # Si es el dominio de grupo bios, debemos validar el asunto
    folder = carpeta_destino
    bank_folder = ""
    if dominio == dominio_bios:
        # Obtenemos el banco del asunto
        if validar_contenido(asunto, "agrario"):
            bank_folder = "/Banco Agrario"
        elif validar_contenido(asunto, "bogota") or validar_contenido(asunto, "bogotá"):
            bank_folder = "/Banco de Bogotá"
        elif validar_contenido(asunto, "bancolombia panama") or validar_contenido(
            asunto, "bancolombia panamá"
        ):
            bank_folder = "/Bancolombia Panamá"
        elif validar_contenido(asunto, "valores bancolombia") or validar_contenido(
            asunto, "mt940"
        ):
            bank_folder = "/Valores Bancolombia"
        elif validar_contenido(asunto, "bancolombia"):
            bank_folder = "/Bancolombia"
        elif validar_contenido(asunto, "corficolombiana"):
            bank_folder = "/Corficolombiana"
        elif validar_contenido(asunto, "corredores davivienda"):
            bank_folder = "/Corredores Davivienda"
        elif validar_contenido(asunto, "credicorp"):
            bank_folder = "/Credicorp"
        elif validar_contenido(asunto, "davivienda"):
            bank_folder = "/Davivienda"
        elif validar_contenido(asunto, "fidualianza"):
            bank_folder = "/Fidualianza"
        elif validar_contenido(asunto, "itau") or validar_contenido(asunto, "itaú"):
            bank_folder = "/Itau"
    else:
        # Obtenemos el banco del dominio
        if dominio == "davivienda.com":
            bank_folder = "/Davivienda"
        elif dominio == "corredores.com":
            bank_folder = "/Corredores Davivienda"
        elif dominio == "alianza.com.co":
            bank_folder = "/Fidualianza"
        elif dominio == "credicorpcapital.com":
            bank_folder = "/Credicorp"
        elif dominio == "solicitudesgrupobancolombia.com.co":
            bank_folder = "/Bancolombia Panamá"
        elif dominio == "bancolombia.com.co":
            bank_folder = "/Valores Bancolombia"

    # Si la carpeta no existe, la creamos:
    field_folder = folder + bank_folder

    if not os.path.exists(field_folder):
        os.makedirs(field_folder)

    return field_folder


class SyncHandler(object):
    try:
        # Save the dispatch interface to identify the SyncObject if needed
        def set(self, disp):
            self._disp = disp

        def _process(self):
            # Decrement sync counter
            global nSync
            nSync -= 1

            # If nothing left to sync, then send WM_QUIT to thread message loop
            if nSync <= 0:
                win32api.PostThreadMessage(
                    win32api.GetCurrentThreadId(), win32con.WM_QUIT, 0, 0
                )

        def OnSyncStart(self):
            print("Starting sync on", self._disp.Name)

        def OnSyncEnd(self):
            print("Sync complete on", self._disp.Name)
            self._process()

        def OnProgress(self, state, description, value, max):
            print(
                "Sync progress: {0:} {1:} {2:}%".format(
                    self._disp.Name, description, 100 * value / max
                )
            )

        def OnError(self, code, description):
            print("Sync Error", description)
            self._process()

    except Exception as e:
        logger_errores.error(f"ERROR SYNC >> {e}")


# Función para buscar y descargar los archivos adjuntos
async def main():
    global nSync

    Hora_ejecucion = datetime.datetime.now()
    logger_transacciones.info(f"BOT Ejecutado a las {Hora_ejecucion}")

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        syncs = namespace.SyncObjects
        nSync = syncs.Count
        logger_transacciones.info(f"Numero de objetos a Sync>> {nSync}")

        for syncOject in syncs:
            print("Carpeta a sincronizar", syncOject.Name)
            handler = win32com.client.WithEvents(syncOject, SyncHandler)
            handler.set(syncOject)
            syncOject.Start()

        pythoncom.PumpMessages()

        root_folder = namespace.Folders.Item(1)

        # Obtenemos la carpeta de la que se leerán los correos
        carpeta = root_folder.Folders[carpeta_outlook_origen]

        # Obtenemos la carpeta a la que debemos mover el correo una vez leído.
        carpeta_destino_outlook = root_folder.Folders[carpeta_outlook_dest]

        correos_por_hilo = 4
        mensajes = carpeta.Items
        logger_transacciones.info(f"NUMERO DE CORREOS >> {len(mensajes)}")
        hilos = dividir_en_lotes(list(mensajes), correos_por_hilo)
        i = 1
        for items in hilos:
            logger_transacciones.info(f"PROCESO =========== {i} =========== ")
            await asyncio.gather(
                *(
                    procesar_correo(correos, carpeta_destino_outlook)
                    for correos in items
                )
            )
            i += 1

        # Cerrar Outlook:
        outlook.Quit()

    except Exception as e:
        logger_errores.error(f"ERROR CORREO 1 >> {e}")


def dividir_en_lotes(array, tamano_lote):
    lotes = []
    for i in range(0, len(array), tamano_lote):
        aux = []
        for x in range(0, tamano_lote):
            index = i + x
            if index < len(array):
                aux.append(array[index])
        lotes.append(aux)

    return lotes


async def procesar_correo(item, carpeta_destino_outlook):
    try:
        # Obtenemos el asunto y el correo del remitente:
        asunto = item.Subject
        remitente = item.SenderEmailAddress
        logger_transacciones.info(f"ASUNTO >>> {asunto}")
        logger_transacciones.info(f"REMITENTE >>> {remitente}")
        # Extraemos el dominio del remitente utilizando expresiones regulares:
        patron_dominio = r"@([\w\.-]+)"
        if re.search(patron_dominio, remitente) is None:
            dominio = dominio_bios
        else:
            dominio = re.search(patron_dominio, remitente).group(1)
        carpeta_descarga = obtener_destino(asunto, dominio.lower())

        # Descargar archivos adjuntos

        logger_transacciones.info(
            f"NUMERO DE ADJUNTOS >> {len(item.Attachments)} A CARPETA {carpeta_descarga}"
        )
        for adjunto in item.Attachments:
            descargar_archivos(adjunto, carpeta_descarga)
            # await asyncio.gather(*(descargar_archivos(adjunto, carpeta_descarga)))

        item.Move(carpeta_destino_outlook)  # Mueve los correos a la nueva carpeta0
        logger_transacciones.info("CORREO MOVIDO CORRECTAMENTE")
        time.sleep(4)
        # if item.Move(carpeta_destino_outlook):
        #     print(f'ARCHIVO MOVIDO CORRECTAMENTE')

    except Exception as e:
        logger_errores.error(f"ERROR CORREO 2 procesar_correo >> {e}")


def descargar_archivos(adjunto, carpeta_descarga):
    try:
        _, ext = os.path.splitext(adjunto.FileName)

        # Obtener la marca de tiempo actual con milisegundos
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")

        # Verificar si el adjunto tiene extensión .txt
        if ext.lower() == ".txt":
            # Obtener el nombre del archivo adjunto y su extensión
            nombre_archivo, ext = os.path.splitext(adjunto.FileName)

            # Modificar el nombre del archivo agregando el timestamp antes de la extensión
            nuevo_nombre_archivo = f"{nombre_archivo}@@{timestamp}{ext}"

            try:
                adjunto.SaveAsFile(os.path.join(carpeta_descarga, nuevo_nombre_archivo))
            except Exception as e:
                logger_errores.error(f"ERROR AL DESCARGAR EL ARCHIVO (.txt)")

        # Verificar si el adjunto tiene extensión .prn
        if ext.lower() == ".prn":
            # Obtener el nombre del archivo adjunto y su extensión
            nombre_archivo, ext = os.path.splitext(adjunto.FileName)

            # Modificar el nombre del archivo agregando el timestamp antes de la extensión
            nuevo_nombre_archivo = f"{nombre_archivo}@@{timestamp}{ext}"

            try:
                bank_folder = "/Valores Bancolombia"
                carpeta_descarga = carpeta_destino + bank_folder
                adjunto.SaveAsFile(os.path.join(carpeta_descarga, nuevo_nombre_archivo))
            except Exception as e:
                logger_errores.info(f"ERROR AL DESCARGAR EL ARCHIVO (.prn)")

        elif ext.lower() == ".zip":
            # Guardar el archivo ZIP localmente
            archivo_zip = os.path.join(carpeta_descarga, adjunto.FileName)
            adjunto.SaveAsFile(archivo_zip)
            # Descomprimir el archivo ZIP
            with zipfile.ZipFile(archivo_zip, "r") as zip_ref:
                for nombre_archivo in zip_ref.namelist():
                    nombre, ext = os.path.splitext(nombre_archivo)

                    # Cambiar el nombre del archivo extraído
                    nuevo_nombre_archivo = f"{nombre}@@{timestamp}{ext}"
                    ruta_descarga = os.path.join(carpeta_descarga, nuevo_nombre_archivo)

                    # Extraer y guardar el archivo con el nuevo nombre
                    zip_ref.extract(nombre_archivo, carpeta_descarga)
                    os.rename(
                        os.path.join(carpeta_descarga, nombre_archivo), ruta_descarga
                    )

            # Eliminar el archivo ZIP después de la extracción y descarga
            os.remove(archivo_zip)

    except Exception as e:
        logger_errores.error(f"ERROR En  descargar_archivos Tipo de error >> {e}")


if __name__ == "__main__":
    asyncio.run(main())
