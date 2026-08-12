"""
=========================================================
Proyecto Fertilización

Administrador de archivos

=========================================================
"""

from pathlib import Path


class FileManager:

    def __init__(self):

        pass

    # -----------------------------------------------------

    def buscar_excels(self, carpeta):

        carpeta = Path(carpeta)

        archivos = []

        archivos.extend(carpeta.rglob("*.xlsx"))

        archivos.extend(carpeta.rglob("*.xlsm"))

        return sorted(archivos)

    # -----------------------------------------------------

    def crear_carpeta(self, ruta):

        ruta = Path(ruta)

        ruta.mkdir(

            parents=True,

            exist_ok=True

        )

        return ruta

    # -----------------------------------------------------

    def existe(self, ruta):

        return Path(ruta).exists()

    # -----------------------------------------------------

    def nombre_archivo(self, ruta):

        return Path(ruta).stem

    # -----------------------------------------------------

    def extension(self, ruta):

        return Path(ruta).suffix

    # -----------------------------------------------------

    def listar_carpetas(self, carpeta):

        carpeta = Path(carpeta)

        return [

            x

            for x in carpeta.iterdir()

            if x.is_dir()

        ]