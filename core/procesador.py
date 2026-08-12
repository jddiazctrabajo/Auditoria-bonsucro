# -*- coding: utf-8 -*-
"""
==============================================================
Proyecto Fertilización
core/procesador.py
==============================================================

Responsabilidad
---------------
Orquestar todo el proceso de consolidación.

NO realiza cálculos directamente.
NO consulta Excel directamente (usa ExcelReader).

Flujo

Leer archivos
    ↓
Homologar
    ↓
Completar desde Prontuario
    ↓
Completar desde Maestro
    ↓
Calcular
    ↓
Validar
    ↓
Devolver DataFrames

==============================================================
"""

from pathlib import Path
import pandas as pd

from services.excel_reader import ExcelReader
from core.homologacion import Homologador
from core.maestro import Maestro
from core.prontuario import Prontuario
from core.calculos import Calculos
from core.validaciones import Validador


class Procesador:

    # ==========================================================
    # Constructor
    # ==========================================================

    def __init__(self, ruta_maestro, ruta_prontuario):

        self.reader = ExcelReader()

        self.homologador = Homologador(ruta_maestro)

        self.maestro = Maestro(ruta_maestro)

        self.prontuario = Prontuario(ruta_prontuario)

        self.calculos = Calculos()

        self.validador = Validador()

        self.columnas_esperadas = [

            "HACIENDA",
            "SUERTE",
            "AREA",
            "PRODUCTO",
            "CANTIDAD",
            "UNIDAD",
            "DOSIS X HA",
            "UNIDADES - N",
            "UNIDADES - P",
            "UNIDADES - K",
            "UNIDADES - S",
            "UNIDADES - MENORES"
        ]

    # ==========================================================
    # Método principal
    # ==========================================================

    def procesar(self, carpeta):

        df = self._leer_archivos(carpeta)

        df = self._homologar(df)

        df = self._completar_prontuario(df)

        df = self._completar_maestro(df)

        df = self._calcular(df)

        errores = self._validar(df)

        return df, errores

    # ==========================================================
    # Leer archivos
    # ==========================================================

    def _leer_archivos(self, carpeta):

        carpeta = Path(carpeta)

        archivos = list(carpeta.rglob("*.xlsx"))
        archivos.extend(carpeta.rglob("*.xlsm"))

        if len(archivos) == 0:

            raise Exception(
                "No se encontraron archivos Excel."
            )

        lista_df = []

        for archivo in archivos:

            try:

                df = self.reader.leer_tabla(

                    ruta=archivo,

                    hoja="FORMATO FERT",

                    columnas_esperadas=self.columnas_esperadas

                )

                print(df.columns.tolist())

                print(df[
                    [
                        "UNIDADES - N",
                        "UNIDADES - P",
                        "UNIDADES - K",
                        "MENORES"
                    ]
                ].head(10))
                # --------------------------
                # Trazabilidad
                # --------------------------

                df["RAZON_SOCIAL"] = archivo.parent.name

                df["HACIENDA_ARCHIVO"] = archivo.stem

                df["ARCHIVO_ORIGEN"] = archivo.name

                lista_df.append(df)

            except Exception as e:

                self.validador.registrar_error(

                    archivo=archivo.name,

                    fila=0,

                    campo="ARCHIVO",

                    valor="",

                    descripcion=str(e)

                )

        if len(lista_df) == 0:

            raise Exception(
                "No fue posible leer ningún archivo."
            )

        consolidado = pd.concat(

            lista_df,

            ignore_index=True

        )

        consolidado = consolidado.reset_index(drop=True)

        # Estado inicial

        consolidado["ESTADO"] = "OK"

        consolidado["OBSERVACIONES"] = ""

        return consolidado

    # ==========================================================
    # Homologar información
    # ==========================================================

    def _homologar(self, df):
        df = self.homologador.dataframe(df)

        # Normalizar columnas igual que prontuario y maestro
        import unicodedata
        
        def normalizar_texto(texto):
            """Mismo método que Prontuario y Maestro"""
            if pd.isna(texto):
                return ""
            texto = str(texto).upper().strip()
            texto = unicodedata.normalize("NFKD", texto)
            texto = "".join(c for c in texto if not unicodedata.combining(c))
            texto = " ".join(texto.split())
            return texto

        # Normalizar campos clave para el merge
        columnas_normalizar = ["HACIENDA", "SUERTE", "PRODUCTO"]
        
        for col in columnas_normalizar:
            if col in df.columns:
                print(f"DEBUG: Normalizando {col}")
                print(f"  Antes: {df[col].head(3).tolist()}")
                df[col] = df[col].apply(normalizar_texto)
                print(f"  Después: {df[col].head(3).tolist()}")

        # Normalizar nombres de columnas: eliminar espacios, unificar guiones y colapsar espacios, manejar caracteres invisibles
        import re
        def _clean_name(col):
            name = col.strip()
            # Eliminar caracteres invisibles (zero‑width spaces, etc.)
            name = re.sub(r'[\u200B-\u200D\uFEFF]', '', name)
            # Reemplazar cualquier tipo de guion largo por '-'
            name = re.sub(r'[\u2010-\u2015]', '-', name)
            # Normalizar espacios alrededor del guion
            name = re.sub(r'\s*-\s*', ' - ', name)
            # Reemplazar cualquier carácter de espacio Unicode (incluido \u00A0) por espacio simple
            name = re.sub(r"[\s\u00A0]+", " ", name)
            # Colapsar múltiples espacios en uno
            name = re.sub(r"\s+", " ", name)
            return name
        df.rename(columns=lambda c: _clean_name(c), inplace=True)
        # Reconocer y renombrar variantes de columnas de unidades (ej. "UNIDADES - N (g)" o sin guion)
        import re
        def _norm(col):
            # Remove non‑alphanumeric characters and uppercase
            return re.sub(r"[^A-Za-z0-9]", "", col).upper()
        for expected in ["UNIDADES - N", "UNIDADES - P", "UNIDADES - K", "UNIDADES - S", "UNIDADES - MENORES"]:
            if expected in df.columns:
                continue
            elem = expected.split(" - ")[1]
            target_norm = _norm(f"UNIDADES{elem}")
            for col in df.columns:
                if _norm(col) == target_norm:
                    df.rename(columns={col: expected}, inplace=True)
                    break
        # Asegurar que existen todas las columnas de unidades esperadas
        unidades_cols = ["UNIDADES - N", "UNIDADES - P", "UNIDADES - K", "UNIDADES - S", "UNIDADES - MENORES"]
        for uc in unidades_cols:
            if uc not in df.columns:
                df[uc] = pd.NA

        return df

    # ==========================================================
    # Completar desde prontuario
    # ==========================================================

    def _completar_prontuario(self, df):

        prontuario = self.prontuario.df.copy()

        # Renombrar columnas para evitar conflictos
        prontuario = prontuario.rename(columns={

            "AREA": "AREA_PRONTUARIO",

            "VARIEDAD": "VARIEDAD_PRONTUARIO",

            "ULT_CORTE": "ULT_CORTE_PRONTUARIO",

            "TCH ACTUAL": "TCH_ACTUAL_PRONTUARIO"

        })

        columnas = [

            "HACIENDA",

            "SUERTE",

            "AREA_PRONTUARIO",

            "VARIEDAD_PRONTUARIO",

            "ULT_CORTE_PRONTUARIO",

            "TCH_ACTUAL_PRONTUARIO"

        ]

        prontuario = prontuario[columnas]

        print("\nDEBUG - Antes del merge con prontuario:")
        print(f"  Registros en df: {len(df)}")
        print(f"  Haciendas únicas en df: {df['HACIENDA'].nunique()}")
        print(f"  Haciendas en prontuario: {prontuario['HACIENDA'].nunique()}")
        print(f"  Primeras haciendas en df: {df['HACIENDA'].head(3).tolist()}")
        print(f"  Primeras haciendas en prontuario: {prontuario['HACIENDA'].head(3).tolist()}")

        df = df.merge(

            prontuario,

            on=["HACIENDA", "SUERTE"],

            how="left"

        )

        print("\nDEBUG - Después del merge con prontuario:")
        print(f"  Registros con AREA_PRONTUARIO: {df['AREA_PRONTUARIO'].notna().sum()}")
        print(f"  Registros sin AREA_PRONTUARIO: {df['AREA_PRONTUARIO'].isna().sum()}")

        # ---------------------------------------------------
        # Completar área solamente cuando venga vacía
        # ---------------------------------------------------

        if "AREA" in df.columns:

            df["AREA"] = df["AREA"].fillna(df["AREA_PRONTUARIO"])

        # ---------------------------------------------------
        # Agregar columnas nuevas
        # ---------------------------------------------------

        if "VARIEDAD" not in df.columns:

            df["VARIEDAD"] = df["VARIEDAD_PRONTUARIO"]

        if "ULT_CORTE" not in df.columns:

            df["ULT_CORTE"] = df["ULT_CORTE_PRONTUARIO"]

        if "TCH_ACTUAL" not in df.columns:

            df["TCH_ACTUAL"] = df["TCH_ACTUAL_PRONTUARIO"]

        # Eliminar columnas auxiliares

        df = df.drop(

            columns=[

                "AREA_PRONTUARIO",

                "VARIEDAD_PRONTUARIO",

                "ULT_CORTE_PRONTUARIO",

                "TCH_ACTUAL_PRONTUARIO"

            ],

            errors="ignore"

        )

        return df

    # ==========================================================
    # Completar desde maestro
    # ==========================================================

    def _completar_maestro(self, df):

        maestro = self.maestro.df.copy()

        maestro = maestro.rename(columns={

            "NOMBRE COMERCIAL": "PRODUCTO",

            "N": "PORC_N",

            "P": "PORC_P",

            "K": "PORC_K",

            "S": "PORC_S",

            "ELEMENTOS MENORES": "PORC_MENORES"

        })

        # Normalizar PRODUCTO en maestro usando el mismo método que prontuario
        import unicodedata
        
        def normalizar_texto(texto):
            """Mismo método que Prontuario"""
            if pd.isna(texto):
                return ""
            texto = str(texto).upper().strip()
            texto = unicodedata.normalize("NFKD", texto)
            texto = "".join(c for c in texto if not unicodedata.combining(c))
            texto = " ".join(texto.split())
            return texto
        
        maestro["PRODUCTO"] = maestro["PRODUCTO"].apply(normalizar_texto)

        print("\nDEBUG - Maestro después de normalizar PRODUCTO:")
        print(maestro[["PRODUCTO", "PORC_N"]].head())

        columnas = [

            "PRODUCTO",

            "PORC_N",

            "PORC_P",

            "PORC_K",

            "PORC_S",

            "PORC_MENORES"

        ]

        maestro = maestro[columnas]

        print("\nDEBUG - Antes del merge con maestro:")
        print(f"  Registros en df: {len(df)}")
        print(f"  Productos únicos en df: {df['PRODUCTO'].nunique()}")
        print(f"  Productos en maestro: {len(maestro)}")
        print(f"  Productos únicos en maestro: {maestro['PRODUCTO'].nunique()}")
        print(f"  Primeros productos en df: {df['PRODUCTO'].head(3).tolist()}")
        print(f"  Primeros productos en maestro: {maestro['PRODUCTO'].head(3).tolist()}")

        df = df.merge(

            maestro,

            on="PRODUCTO",

            how="left"

        )

        print("\nDEBUG - Después del merge con maestro:")
        print(f"  Registros con PORC_N: {df['PORC_N'].notna().sum()}")
        print(f"  Registros sin PORC_N: {df['PORC_N'].isna().sum()}")
        print(f"  Primeros PORC_N: {df['PORC_N'].head(5).tolist()}")

        return df

    # ==========================================================
    # Realizar cálculos
    # ==========================================================

    # ==========================================================
# Completar Área
# ==========================================================

    def _calcular_area(self, df):

        if "AREA" not in df.columns:
            return df

        mascara = (

            df["AREA"].isna()

            &

            df["CANTIDAD"].notna()

            &

            df["DOSIS X HA"].notna()

        )

        df.loc[mascara, "AREA"] = (

            df.loc[mascara, "CANTIDAD"].astype(float)

            /

            df.loc[mascara, "DOSIS X HA"].astype(float)

        )

        return df

    # ==========================================================
# Completar Cantidad
# ==========================================================

    def _calcular_cantidad(self, df):

        mascara = (

            df["CANTIDAD"].isna()

            &

            df["AREA"].notna()

            &

            df["DOSIS X HA"].notna()

        )

        df.loc[mascara, "CANTIDAD"] = (

            df.loc[mascara, "AREA"].astype(float)

            *

            df.loc[mascara, "DOSIS X HA"].astype(float)

        )

        return df

    # ==========================================================
# Calcular dosis
# ==========================================================

    def _calcular_dosis(self, df):

        mascara = (

            df["DOSIS X HA"].isna()

            &

            df["AREA"].notna()

            &

            df["CANTIDAD"].notna()

        )

        df.loc[mascara, "DOSIS X HA"] = (

            df.loc[mascara, "CANTIDAD"].astype(float)

            /

            df.loc[mascara, "AREA"].astype(float)

        )

        return df

    # ==========================================================
# Calcular Cantidad total (kg/ha × ha)
# ==========================================================

    def _calcular_elementos(self, df):

        print("\n================ INICIO _calcular_elementos ================\n")

        print("Antes del cálculo:")
        print(df[[
            "UNIDADES - N",
            "UNIDADES - P",
            "UNIDADES - K",
            "CANTIDAD"
        ]].head())

        elementos = {
            "N": "PORC_N",
            "P": "PORC_P",
            "K": "PORC_K",
            "S": "PORC_S",
            "MENORES": "PORC_MENORES"
        }

        # Asegurar que CANTIDAD sea numérica
        df["CANTIDAD"] = pd.to_numeric(
            df["CANTIDAD"],
            errors="coerce"
        )

        for elemento, columna in elementos.items():

            if columna not in df.columns:
                print(f"⚠️  Columna {columna} no encontrada, saltando {elemento}")
                continue

            # Convertir porcentaje a numérico
            df[columna] = pd.to_numeric(
                df[columna],
                errors="coerce"
            )

            nombre_salida = f"UNIDADES - {elemento}"

            # Convertir la columna existente a numérica
            df[nombre_salida] = pd.to_numeric(
                df[nombre_salida],
                errors="coerce"
            )

            # Calcular únicamente donde el valor esté vacío
            mascara = (
                df[nombre_salida].isna()
                &
                df["CANTIDAD"].notna()
                &
                df[columna].notna()
            )

            print(f"\n{nombre_salida} - Filas calculadas: {mascara.sum()}")

            df.loc[mascara, nombre_salida] = (
                df.loc[mascara, "CANTIDAD"]
                *
                df.loc[mascara, columna]
                / 100
            ).round(4)

        print("\nDespués del cálculo:")
        print(df[[
            "UNIDADES - N",
            "UNIDADES - P",
            "UNIDADES - K",
            "UNIDADES - S",
            "UNIDADES - MENORES",
            "CANTIDAD"
        ]].head())

        print("\n================ FIN _calcular_elementos ================\n")

        return df

# ==========================================================
# Calcular cantidad del producto a partir de unidades
# ==========================================================

    def _calcular_producto(self, df):

        print("\n================ INICIO _calcular_producto ================\n")
        print("Columnas disponibles:")
        print(df.columns.tolist())

        for idx in df.index:

            print("\n===================================================\n")
            print("FILA:", idx)

            # Si ya existe la cantidad no hacer nada
            cantidad = Calculos.numero(df.at[idx, "CANTIDAD"])
            if not pd.isna(cantidad):
                print(">> Ya tiene cantidad. Se omite.")
                continue

            producto = df.at[idx, "PRODUCTO"]
            print("Producto:", producto)
            if pd.isna(producto):
                print(">> Producto vacío.")
                continue

            aportes = self.maestro.obtener_aportes(producto)
            print("Aportes:", aportes)
            if aportes is None:
                print(">> No se encontró el producto en el maestro.")
                continue

            # Buscar el primer elemento con unidades suministradas y porcentaje > 0
            for elemento in ["N", "P", "K", "S", "MENORES"]:
                columna_unidades = f"UNIDADES - {elemento}"
                if columna_unidades not in df.columns:
                    continue

                unidades_raw = df.at[idx, columna_unidades]
                unidades = Calculos.numero(unidades_raw)
                porcentaje = Calculos.numero(aportes.get(elemento, 0))
                print("\nElemento:", elemento)
                print("Columna esperada:", columna_unidades)
                print("Valor crudo:", unidades_raw)
                print("Unidades:", unidades)
                print("Porcentaje:", porcentaje)

                if Calculos.puede_calcular_cantidad_por_elemento(cantidad, unidades, porcentaje):
                    cantidad_calc = Calculos.calcular_producto(unidades, porcentaje)
                    print("Cantidad calculada:", cantidad_calc)
                    df.at[idx, "CANTIDAD"] = cantidad_calc
                    cantidad = cantidad_calc
                    break

            # Después de calcular CANTIDAD, intentar recomputar DOSIS X HA y ÁREA si faltan
            area = Calculos.numero(df.at[idx, "AREA"])
            dosis = Calculos.numero(df.at[idx, "DOSIS X HA"])

            # Recalcular DOSIS X HA si falta y AREA está disponible
            if pd.isna(dosis) and not pd.isna(area) and not pd.isna(cantidad):
                nueva_dosis = Calculos.calcular_dosis(area, cantidad)
                print("Nueva dosis calculada:", nueva_dosis)
                df.at[idx, "DOSIS X HA"] = nueva_dosis
                dosis = nueva_dosis

            # Recalcular ÁREA si falta y DOSIS está disponible
            if pd.isna(area) and not pd.isna(dosis) and not pd.isna(cantidad):
                nueva_area = Calculos.calcular_area(cantidad, dosis)
                print("Nueva área calculada:", nueva_area)
                df.at[idx, "AREA"] = nueva_area

        print("\n================ FIN _calcular_producto ================\n")

        return df

    # ==========================================================
# Realizar cálculos
# ==========================================================
        
def _calcular(self, df):

    df = self._calcular_area(df)

    df = self._calcular_cantidad(df)

    df = self._calcular_dosis(df)

    df = self._calcular_elementos(df)

    df = self._calcular_producto(df)

    # producto pudo completar área o dosis
    df = self._calcular_area(df)

    df = self._calcular_dosis(df)

    # ahora sí todas las unidades quedan completas
    df = self._calcular_elementos(df)

    return df
    # ==========================================================
    # Validar información
    # ==========================================================
    
    def _validar(self, df):

        # ----------------------------
        # Validaciones existentes
        # ----------------------------

        self.validador.validar_hacienda(df)

        self.validador.validar_suerte(df)

        # ----------------------------
        # Validar Área
        # ----------------------------

        if all(col in df.columns for col in ["AREA", "CANTIDAD", "DOSIS X HA"]):

            area = pd.to_numeric(df["AREA"], errors="coerce")

            cantidad = pd.to_numeric(df["CANTIDAD"], errors="coerce")

            dosis = pd.to_numeric(df["DOSIS X HA"], errors="coerce")

            calculada = cantidad / dosis

            diferencia = (area - calculada).abs()

            mascara = (

                area.notna()

                &

                cantidad.notna()

                &

                dosis.notna()

                &

                (diferencia > 0.01)

            )

            for idx in df.index[mascara]:

                self.validador.registrar_error(

                    archivo=df.at[idx, "ARCHIVO_ORIGEN"],

                    fila=idx + 2,

                    campo="AREA",

                    valor=df.at[idx, "AREA"],

                    descripcion="Area no coincide con Cantidad / Dosis"

                )

                df = self.validador.marcar_error(

                    df,

                    idx,

                    "Area inconsistente"

                )

        return pd.DataFrame(self.validador.errores)
