from pathlib import Path
import pandas as pd
import unicodedata


class Maestro:

    def __init__(self, ruta):

        self.ruta = Path(ruta)

        self.df = self._cargar()

    # ======================================================
    # Normalizar texto
    # ======================================================

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

    # ======================================================
    # Cargar maestro
    # ======================================================

    def _cargar(self):

        df = pd.read_excel(

            self.ruta,

            sheet_name="APORTE FERTILIZANTES",

            usecols="A:H"

        )

        df.columns = [

            self.normalizar(c)

            for c in df.columns

        ]

        df["NOMBRE COMERCIAL"] = (

            df["NOMBRE COMERCIAL"]

            .fillna("")

            .apply(self.normalizar)

        )

        return df

    # ======================================================
    # Saber si existe
    # ======================================================

    def existe_producto(self, producto):

        producto = self.normalizar(producto)

        return producto in set(self.df["NOMBRE COMERCIAL"])

    # ======================================================
    # Obtener fila
    # ======================================================

    def obtener_producto(self, producto):

        producto = self.normalizar(producto)

        fila = self.df[

            self.df["NOMBRE COMERCIAL"] == producto

        ]

        if fila.empty:

            return None

        return fila.iloc[0]

    # ======================================================
    # Obtener aportes
    # ======================================================

    def obtener_aportes(self, producto):

        fila = self.obtener_producto(producto)

        if fila is None:

            return None

        return {

            "N": fila["N"],

            "P": fila["P"],

            "K": fila["K"],

            "S": fila["S"],

            "MENORES": fila["ELEMENTOS MENORES"]

        }

    # ======================================================
    # Listar productos
    # ======================================================

    def listar_productos(self):

        return sorted(

            self.df["NOMBRE COMERCIAL"]

            .dropna()

            .unique()

        )