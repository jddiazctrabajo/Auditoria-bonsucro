"""
==============================================================
Proyecto Fertilización
core/calculos.py
==============================================================

Responsabilidad
---------------
Contiene únicamente reglas matemáticas.

NO lee archivos.
NO consulta el Maestro.
NO consulta el Prontuario.
NO modifica DataFrames.

Cada función recibe valores y devuelve un resultado.

==============================================================
"""

import pandas as pd
# pyrefly: ignore [missing-import]
import numpy as np


class Calculos:

    # ==========================================================
    # Saber si un valor está vacío
    # ==========================================================

    @staticmethod
    def vacio(valor):

        if pd.isna(valor):
            return True

        if str(valor).strip() == "":
            return True

        return False

    # ==========================================================
    # Convertir cualquier valor a número
    # ==========================================================

    @staticmethod
    def numero(valor):

        if Calculos.vacio(valor):
            return np.nan

        try:

            valor = str(valor).replace(",", ".")

            return float(valor)

        except:

            return np.nan

    # ==========================================================
    # Cantidad = Área × Dosis
    # ==========================================================

    @staticmethod
    def calcular_cantidad(area, dosis):

        area = Calculos.numero(area)
        dosis = Calculos.numero(dosis)

        if pd.isna(area) or pd.isna(dosis):
            return np.nan

        return round(area * dosis, 2)

    # ==========================================================
    # Dosis = Cantidad / Área
    # ==========================================================

    @staticmethod
    def calcular_dosis(area, cantidad):

        area = Calculos.numero(area)
        cantidad = Calculos.numero(cantidad)

        if pd.isna(area) or pd.isna(cantidad):
            return np.nan

        if area == 0:
            return np.nan

        return round(cantidad / area, 2)

    # ==========================================================
    # Área = Cantidad / Dosis
    # ==========================================================

    @staticmethod
    def calcular_area(cantidad, dosis):

        cantidad = Calculos.numero(cantidad)
        dosis = Calculos.numero(dosis)

        if pd.isna(cantidad) or pd.isna(dosis):
            return np.nan

        if dosis == 0:
            return np.nan

        return round(cantidad / dosis, 2)

    # ==========================================================
    # Elemento aplicado (N, P, K, S, Menores)
    # ==========================================================

    @staticmethod
    def calcular_elemento(cantidad_producto, porcentaje):

        cantidad_producto = Calculos.numero(cantidad_producto)
        porcentaje = Calculos.numero(porcentaje)

        if pd.isna(cantidad_producto) or pd.isna(porcentaje):
            return np.nan

        return round(cantidad_producto * porcentaje / 100, 2)

    # ==========================================================
    # Cantidad de producto requerida
    # ==========================================================

    @staticmethod
    def calcular_producto(elemento, porcentaje):

        elemento = Calculos.numero(elemento)
        porcentaje = Calculos.numero(porcentaje)

        if pd.isna(elemento) or pd.isna(porcentaje):
            return np.nan

        if porcentaje == 0:
            return np.nan

        return round(elemento * 100 / porcentaje, 2)

    # ==========================================================
    # Validar igualdad matemática
    # ==========================================================

    @staticmethod
    def coincide(valor1, valor2, tolerancia=0.01):

        valor1 = Calculos.numero(valor1)
        valor2 = Calculos.numero(valor2)

        if pd.isna(valor1) or pd.isna(valor2):
            return False

        return abs(valor1 - valor2) <= tolerancia

    # ==========================================================
    # Validar si existen suficientes datos para calcular
    # ==========================================================

    @staticmethod
    def puede_calcular_area(area, cantidad, dosis):

        return (
            Calculos.vacio(area)
            and not Calculos.vacio(cantidad)
            and not Calculos.vacio(dosis)
        )

    @staticmethod
    def puede_calcular_cantidad(area, cantidad, dosis):

        return (
            not Calculos.vacio(area)
            and Calculos.vacio(cantidad)
            and not Calculos.vacio(dosis)
        )

    @staticmethod
    def puede_calcular_dosis(area, cantidad, dosis):

        return (
            not Calculos.vacio(area)
            and not Calculos.vacio(cantidad)
            and Calculos.vacio(dosis)
        )

    # ==========================================================
    # Saber si los tres valores existen
    # ==========================================================

    @staticmethod
    def datos_completos(area, cantidad, dosis):

        return (
            not Calculos.vacio(area)
            and not Calculos.vacio(cantidad)
            and not Calculos.vacio(dosis)
        )

    @staticmethod
    def puede_calcular_cantidad_por_elemento(cantidad, unidades, porcentaje):

        return (
            Calculos.vacio(cantidad)
            and not Calculos.vacio(unidades)
            and not Calculos.vacio(porcentaje)
        and Calculos.numero(porcentaje) > 0
    )