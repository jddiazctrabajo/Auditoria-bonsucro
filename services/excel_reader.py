"""
==============================================================
Proyecto Fertilización
services/excel_reader.py
==============================================================

Responsabilidad
---------------
Lectura de archivos Excel.

Este módulo NO conoce reglas de negocio.
NO realiza cálculos.
NO conoce Bonsucro.

Únicamente se encarga de leer hojas de Excel y devolver
DataFrames limpios.

==============================================================
"""

from pathlib import Path
import pandas as pd
import unicodedata


class ExcelReader:

    def __init__(self):

        # Rango de lectura del formato de entrada
        # Todo lo que esté después de la columna O será ignorado.
        self.RANGO_LECTURA = "A:O"

    # ==========================================================
    # Validar archivo
    # ==========================================================

    def validar_archivo(self, ruta):

        ruta = Path(ruta)

        if not ruta.exists():
            raise FileNotFoundError(f"No existe el archivo:\n{ruta}")

        if ruta.suffix.lower() not in [".xlsx", ".xlsm"]:
            raise ValueError("El archivo debe ser .xlsx o .xlsm")

        return True

    # ==========================================================
    # Listar hojas
    # ==========================================================

    def hojas(self, ruta):

        self.validar_archivo(ruta)

        libro = pd.ExcelFile(ruta)

        return libro.sheet_names

    # ==========================================================
    # Leer hoja completa (sin encabezados)
    # ==========================================================

    def leer_hoja(self, ruta, hoja):

        self.validar_archivo(ruta)

        return pd.read_excel(
            ruta,
            sheet_name=hoja,
            header=None,
            usecols=self.RANGO_LECTURA,
            dtype=object
        )

    # ==========================================================
    # Normalizar texto
    # ==========================================================

    @staticmethod
    def normalizar(texto):

        if pd.isna(texto):
            return ""

        texto = str(texto).upper().strip()

        texto = unicodedata.normalize("NFKD", texto)

        texto = "".join(
            c for c in texto
            if not unicodedata.combining(c)
        )

        texto = " ".join(texto.split())

        return texto

    # ==========================================================
    # Buscar encabezado automáticamente
    # ==========================================================

    def detectar_encabezado(self, df, columnas_esperadas):

        mejor_fila = None
        mejor_score = -1

        columnas = {
            self.normalizar(c)
            for c in columnas_esperadas
        }

        for fila, valores in df.iterrows():

            score = 0

            for valor in valores:

                valor = self.normalizar(valor)

                if valor in columnas:
                    score += 1

            if score > mejor_score:

                mejor_score = score
                mejor_fila = fila

        if mejor_fila is None:
            raise Exception("No fue posible encontrar el encabezado.")

        return mejor_fila

    # ==========================================================
    # Leer tabla automáticamente
    # ==========================================================

    def leer_tabla(self, ruta, hoja, columnas_esperadas):

        # Leer hoja completa
        bruto = self.leer_hoja(ruta, hoja)
        print(bruto.iloc[:15, :15])

        # Buscar encabezado
        fila = self.detectar_encabezado(
            bruto,
            columnas_esperadas
        )

        # Leer únicamente la tabla
        df = pd.read_excel(
            ruta,
            sheet_name=hoja,
            header=fila,
            usecols=self.RANGO_LECTURA,
            dtype=object
        )

        # Eliminar filas totalmente vacías
        df = df.dropna(how="all")

        # Normalizar nombres de columnas
        df.columns = [
            self.normalizar(c)
            for c in df.columns
        ]

        # Reiniciar índice
        df = df.reset_index(drop=True)

        return df

        # ==========================================================
    # Leer cualquier tabla detectando automáticamente el encabezado
    # ==========================================================

    def leer_tabla_auto(self, ruta, hoja=0):

        self.validar_archivo(ruta)

        # Leer hoja completa sin encabezado
        bruto = pd.read_excel(
            ruta,
            sheet_name=hoja,
            header=None,
            dtype=object
        )

        mejor_fila = None
        mejor_score = -1

        # Buscar la fila con mayor cantidad de celdas con texto
        for fila, valores in bruto.iterrows():

            score = 0

            for valor in valores:

                if pd.isna(valor):
                    continue

                texto = self.normalizar(valor)

                if texto != "":
                    score += 1

            if score > mejor_score:

                mejor_score = score
                mejor_fila = fila

        if mejor_fila is None:

            raise Exception(
                "No fue posible detectar automáticamente el encabezado."
            )

        # Leer nuevamente usando la fila encontrada como encabezado
        df = pd.read_excel(

            ruta,

            sheet_name=hoja,

            header=mejor_fila,

            dtype=object

        )
        print(df.iloc[:10, :15])
        print(df.columns)

        # Eliminar filas completamente vacías
        df = df.dropna(how="all")

        # Normalizar nombres de columnas
        df.columns = [

            self.normalizar(c)

            for c in df.columns

        ]

        # Eliminar columnas completamente vacías
        df = df.loc[:, ~df.columns.isna()]
        df = df.loc[:, df.columns != ""]

        # Eliminar columnas duplicadas
        df = df.loc[:, ~df.columns.duplicated()]

        # Reiniciar índice
        df = df.reset_index(drop=True)

        return df



    # ==========================================================
    # Información general
    # ==========================================================

    def informacion(self, ruta):

        libro = pd.ExcelFile(ruta)

        return {

            "archivo": Path(ruta).name,

            "ruta": str(ruta),

            "numero_hojas": len(libro.sheet_names),

            "hojas": libro.sheet_names

        }