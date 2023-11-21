import logging
from logging.handlers import SocketHandler, TimedRotatingFileHandler
import os
from datetime import datetime, timedelta

# Rutas logs
ruta_logs = "logs"
ruta_archivo_transacciones = "logs\BOT_MultiCash_Transacciones.log"
ruta_archivo_errores = "logs\BOT_MultiCash_Errores.log"
# Dias a eliminar logs
DiaE_logs = 7

if not os.path.exists(ruta_logs):
    os.makedirs(ruta_logs)

# Configura el logger para transacciones
logger_transacciones = logging.getLogger("transacciones")

socket_handler = SocketHandler("127.0.0.1", 19996)  # default listening address
logger_transacciones.addHandler(socket_handler)
logger_transacciones.setLevel(logging.INFO)
formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

# Utiliza TimedRotatingFileHandler para rotar por día
file_handler_transacciones = TimedRotatingFileHandler(
    ruta_archivo_transacciones, when="midnight", interval=1, backupCount=7
)
file_handler_transacciones.setFormatter(formatter)
logger_transacciones.addHandler(file_handler_transacciones)

# Configura el logger para errores
logger_errores = logging.getLogger("errores")
logger_errores.addHandler(socket_handler)
logger_errores.setLevel(logging.ERROR)

# Utiliza TimedRotatingFileHandler para rotar por día
file_handler_errores = TimedRotatingFileHandler(
    ruta_archivo_errores, when="midnight", interval=1, backupCount=7
)
file_handler_errores.setFormatter(formatter)
logger_errores.addHandler(file_handler_errores)


def eliminar_archivos_antiguos():
    # Obtener la fecha de hace siete días
    seven_days_ago = datetime.now() - timedelta(DiaE_logs)

    log_files = [ruta_archivo_transacciones, ruta_archivo_errores]

    for file_name in log_files:
        if os.path.exists(file_name):
            file_time = datetime.fromtimestamp(os.path.getctime(file_name))
            if file_time < seven_days_ago:
                logger_transacciones.info(
                    f">> Iniciando procedimiento de eliminacion de logs >> Desde: {seven_days_ago}"
                )
                os.remove(file_name)
