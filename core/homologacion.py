# -*- coding: utf-8 -*-
"""
==============================================================
Proyecto Fertilización
core/homologacion.py
==============================================================

Responsabilidad
---------------
Homologar textos del proyecto.

NO realiza cálculos.
NO valida reglas.
NO modifica archivos.

Únicamente transforma valores a un formato estándar.

==============================================================
"""

from pathlib import Path
import pandas as pd
import unicodedata
import re


class Homologador:

    # ==========================================================
    # Constructor
    # ==========================================================

    def __init__(self, ruta_maestro=None):

        self.productos = {}
        self.unidades = {}

        if ruta_maestro:
            self.cargar_maestro(ruta_maestro)

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
    # Leer homologaciones desde maestro.xlsx
    # ==========================================================

    def cargar_maestro(self, ruta):

        ruta = Path(ruta)

        if not ruta.exists():
            return

        xls = pd.ExcelFile(ruta)

        # ------------------------------------------------------
        # Productos
        # ------------------------------------------------------

        if "HOMOLOGACION_PRODUCTOS" in xls.sheet_names:

            df = pd.read_excel(
                xls,
                "HOMOLOGACION_PRODUCTOS"
            )

            for _, fila in df.iterrows():

                entrada = self.normalizar(fila["ENTRADA"])
                salida = self.normalizar(fila["PRODUCTO_OFICIAL"])

                self.productos[entrada] = salida

        # ------------------------------------------------------
        # Unidades
        # ------------------------------------------------------

        if "HOMOLOGACION_UNIDADES" in xls.sheet_names:

            df = pd.read_excel(
                xls,
                "HOMOLOGACION_UNIDADES"
            )

            for _, fila in df.iterrows():

                entrada = self.normalizar(fila["ENTRADA"])
                salida = self.normalizar(fila["OFICIAL"])

                self.unidades[entrada] = salida

    # ==========================================================
    # Homologar producto
    # ==========================================================

    def producto(self, valor):

        valor = self.normalizar(valor)

        return self.productos.get(valor, valor)

    # ==========================================================
    # Homologar unidad
    # ==========================================================

    def unidad(self, valor):

        valor = self.normalizar(valor)

        return self.unidades.get(valor, valor)

    # ==========================================================
    # Homologar Hacienda
    # ==========================================================

    def hacienda(self, valor):

        if pd.isna(valor):
            return ""

        valor = str(valor).strip()

        valor = re.sub(r"\D", "", valor)

        return valor.zfill(6)

    # ==========================================================
    # Homologar Suerte
    # ==========================================================

    def suerte(self, valor):

        if pd.isna(valor):
            return ""

        valor = self.normalizar(valor)

        valor = valor.replace(" ", "")

        m = re.match(r"(\d+)([A-Z]?)", valor)

        if m is None:
            return valor

        numero = m.group(1).zfill(3)

        letra = m.group(2)

        return numero + letra

    # ==========================================================
    # Homologar columnas
    # ==========================================================

    def columnas(self, df):

        nuevas = []

        for c in df.columns:

            c = self.normalizar(c)

            c = c.replace("Á", "A")

            nuevas.append(c)

        df.columns = nuevas

        return df

    # ==========================================================
    # Homologar DataFrame completo
    # ==========================================================

    def dataframe(self, df):

        df = self.columnas(df)

        if "PRODUCTO" in df.columns:
            df["PRODUCTO"] = df["PRODUCTO"].apply(self.producto)

        if "UNIDAD" in df.columns:
            df["UNIDAD"] = df["UNIDAD"].apply(self.unidad)

        if "HACIENDA" in df.columns:
            df["HACIENDA"] = df["HACIENDA"].apply(self.hacienda)

        if "SUERTE" in df.columns:
            df["SUERTE"] = df["SUERTE"].apply(self.suerte)

        return df