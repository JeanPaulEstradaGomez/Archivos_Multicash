@echo off

REM Establecer la ubicación de la carpeta de Outlook:
set ruta_outlook= "C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE"

REM Establecer la ubicación, dentro de Outlook donde estarán los correos que el programa debe leer:
set carpeta_outlook_origen="Nuevos Multicash" 

REM Establecer la ubicación, dentro de Outlook, para mover los correos, una vez se hayan descargado sus archivos
set carpeta_outlook_dest= "Procesados Multicash"

REM Establecer la ubicación de la carpeta de destino de descarga
set carpeta_destino= "\\10.252.0.106\integraciones\ARCHIVOS_MULTICASH"

REM ejecucion del Bot 
python AutomatizacionMulticash.py %ruta_outlook% %carpeta_outlook_origen% %carpeta_outlook_dest% %carpeta_destino%
