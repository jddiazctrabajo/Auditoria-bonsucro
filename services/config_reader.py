"""
==========================================================
CONFIG READER

Lee toda la configuración del sistema.

==========================================================
"""

from pathlib import Path

import pandas as pd

from models.configuracion import Configuracion


class ConfigReader:

    def __init__(self, logger):

        self.logger = logger

    def leer(self, ruta):

        ruta = Path(ruta)

        self.logger.info("")

        self.logger.info("Leyendo configuración...")

        if not ruta.exists():

            raise FileNotFoundError(ruta)

        libro = pd.ExcelFile(ruta)

        hojas = libro.sheet_names

        self.logger.success(

            f"{len(hojas)} hojas encontradas"

        )

        return Configuracion(

            columnas=self._leer(libro, "COLUMNAS"),

            productos=self._leer(libro, "PRODUCTOS"),

            unidades=self._leer(libro, "UNIDADES"),

            operarios=self._leer(libro, "OPERARIOS"),

            haciendas=self._leer(libro, "HACIENDAS"),

            razones_sociales=self._leer(libro, "RAZONES_SOCIALES"),

            tipos_aplicacion=self._leer(

                libro,

                "TIPOS_APLICACION"

            ),

            parametros=self._leer(

                libro,

                "PARAMETROS"

            )

        )

    def _leer(self, libro, hoja):

        if hoja not in libro.sheet_names:

            self.logger.warning(

                f"No existe la hoja {hoja}"

            )

            return pd.DataFrame()

        self.logger.success(

            f"Hoja {hoja}"

        )

        return pd.read_excel(

            libro,

            sheet_name=hoja

        )