"""
==============================================================
Proyecto : Sistema de Trazabilidad de Fertilización
Módulo   : logger.py
Autor    : Proyecto Fertilización
==============================================================

Responsabilidad

Centralizar todo el registro del procesamiento.

- Mostrar mensajes en consola.
- Guardar mensajes en archivo.
- Registrar errores.
- Registrar tiempo de ejecución.

==============================================================
"""

from pathlib import Path
from datetime import datetime
import logging
import time


class LoggerSistema:

    def __init__(self):

        carpeta_logs = Path("logs")
        carpeta_logs.mkdir(exist_ok=True)

        fecha = datetime.now().strftime("%Y%m%d_%H%M%S")

        archivo = carpeta_logs / f"Proceso_{fecha}.log"

        self.logger = logging.getLogger("ProyectoFertilizacion")

        self.logger.setLevel(logging.INFO)

        self.logger.handlers.clear()

        formatter = logging.Formatter(
            "%(asctime)s | %(levelname)-8s | %(message)s",
            "%Y-%m-%d %H:%M:%S"
        )

        archivo_handler = logging.FileHandler(
            archivo,
            encoding="utf-8"
        )

        archivo_handler.setFormatter(formatter)

        consola = logging.StreamHandler()

        consola.setFormatter(formatter)

        self.logger.addHandler(archivo_handler)

        self.logger.addHandler(consola)

        self.inicio = time.time()

        self.info("=============================================")
        self.info("PROYECTO FERTILIZACIÓN")
        self.info("Inicio del procesamiento")
        self.info("=============================================")

    def info(self, mensaje):

        self.logger.info(mensaje)

    def success(self, mensaje):

        self.logger.info("✔ " + mensaje)

    def warning(self, mensaje):

        self.logger.warning(mensaje)

    def error(self, mensaje):

        self.logger.error(mensaje)

    def paso(self, modulo, funcion):

        self.info(f"Ejecutando -> {modulo}.{funcion}")

    def finalizar(self):

        tiempo = time.time() - self.inicio

        self.info("---------------------------------------------")
        self.success(
            f"Proceso terminado correctamente ({tiempo:.2f} segundos)"
        )