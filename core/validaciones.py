import re
import pandas as pd


class Validador:

    def __init__(self):

        self.errores = []

    # =====================================================
    # Registrar error
    # =====================================================

    def registrar_error(self, archivo, fila, campo, valor, descripcion):

        self.errores.append({

            "ARCHIVO": archivo,
            "FILA": fila,
            "CAMPO": campo,
            "VALOR": valor,
            "ERROR": descripcion

        })

    # =====================================================
    # Validar Hacienda
    # =====================================================

    def validar_hacienda(self, df):

        if "HACIENDA" not in df.columns:
            return df

        for i, valor in df["HACIENDA"].items():

            valor = "" if pd.isna(valor) else str(valor).strip()

            if len(valor) != 6:

                self.registrar_error(

                    archivo=df.loc[i, "ARCHIVO_ORIGEN"],

                    fila=i + 2,

                    campo="HACIENDA",

                    valor=valor,

                    descripcion="La hacienda debe tener exactamente 6 caracteres"

                )

        return df

    # =====================================================
    # Validar Suerte
    # =====================================================

    def validar_suerte(self, df):

        if "SUERTE" not in df.columns:
            return df

        patron = r"^\d{3}[A-Za-z]?$"

        for i, valor in df["SUERTE"].items():

            valor = "" if pd.isna(valor) else str(valor).strip().upper()

            if not re.match(patron, valor):

                self.registrar_error(

                    archivo=df.loc[i, "ARCHIVO_ORIGEN"],

                    fila=i + 2,

                    campo="SUERTE",

                    valor=valor,

                    descripcion="Formato válido: 100 o 100A"

                )

        return df

    # =====================================================
    # Exportar errores
    # =====================================================

    def exportar(self, ruta):

        if len(self.errores) == 0:

            print("\nNo se encontraron errores.")

            return

        pd.DataFrame(self.errores).to_excel(

            ruta,

            index=False

        )

        print(f"\nErrores exportados en:\n{ruta}")


    # =====================================================
# Marcar error en el DataFrame
# =====================================================

    def marcar_error(self, df, indice, descripcion):

        if "ESTADO" not in df.columns:
            df["ESTADO"] = "OK"

        if "OBSERVACIONES" not in df.columns:
            df["OBSERVACIONES"] = ""

        df.at[indice, "ESTADO"] = "ERROR"

        if df.at[indice, "OBSERVACIONES"] == "":

            df.at[indice, "OBSERVACIONES"] = descripcion

        else:

            df.at[indice, "OBSERVACIONES"] += "; " + descripcion

        return df