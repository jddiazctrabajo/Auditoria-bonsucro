# -*- coding: utf-8 -*-
"""
==============================================================
Proyecto Fertilización
core/prontuario.py
==============================================================

Responsabilidad
---------------
Leer el prontuario una sola vez y consultar información
por HACIENDA + SUERTE.

NO realiza cálculos.
NO valida reglas.
NO modifica archivos.

==============================================================
"""

from pathlib import Path
import pandas as pd
import unicodedata


class Prontuario:

    def __init__(self, ruta):

        self.ruta = Path(ruta)

        self.df = self._cargar()

    # =====================================================
    # Normalizar texto
    # =====================================================

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

    # =====================================================
    # Cargar prontuario
    # =====================================================

    def _cargar(self):

        df = pd.read_excel(
            self.ruta,
            header=5,
            dtype=object
        )

        df.columns = [

            self.normalizar(c)

            for c in df.columns

        ]
        # Convertir todos los nombres de columnas a mayúsculas
        df.columns = (
            df.columns
              .astype(str)
              .str.strip()
              .str.upper()
        )

        # Homologar nombres
        renombrar = {

            "AREA NETA": "AREA",
            "ULT.COR/SIEM.": "ULT_CORTE"

        }

        df.rename(columns=renombrar, inplace=True)

        df["HACIENDA"] = (

            df["HACIENDA"]

            .astype(str)

            .str.strip()

            .str.zfill(6)

        )

        df["SUERTE"] = (

            df["SUERTE"]

            .astype(str)

            .str.strip()

            .str.upper()

        )

        return df

    # =====================================================
    # Buscar registro
    # =====================================================

    def buscar(self, hacienda, suerte):

        hacienda = str(hacienda).zfill(6)

        suerte = str(suerte).upper().strip()

        fila = self.df[

            (self.df["HACIENDA"] == hacienda)

            &

            (self.df["SUERTE"] == suerte)

        ]

        if fila.empty:

            return None

        return fila.iloc[0].to_dict()

    # =====================================================
    # Saber si existe
    # =====================================================

    def existe(self, hacienda, suerte):

        return self.buscar(hacienda, suerte) is not None

    # =====================================================
    # Obtener área
    # =====================================================

    def area(self, hacienda, suerte):

        dato = self.buscar(hacienda, suerte)

        if dato is None:

            return None

        return dato.get("AREA")

    # =====================================================
    # Obtener variedad
    # =====================================================

    def variedad(self, hacienda, suerte):

        dato = self.buscar(hacienda, suerte)

        if dato is None:

            return None

        return dato.get("VARIEDAD")

    # =====================================================
    # Obtener fecha de corte
    # =====================================================

    def ultimo_corte(self, hacienda, suerte):

        dato = self.buscar(hacienda, suerte)

        if dato is None:

            return None

        return dato.get("ULT_CORTE")

    # =====================================================
    # Obtener TCH
    # =====================================================

    def tch(self, hacienda, suerte):

        dato = self.buscar(hacienda, suerte)

        if dato is None:

            return None

        return dato.get("TCH ACTUAL")