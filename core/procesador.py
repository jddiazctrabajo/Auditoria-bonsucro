# -*- coding: utf-8 -*-

"""
Responsabilidad
---------------

Orquestar todo el proceso de consolidación.

NO realiza cálculos directamente.
NO consulta Excel directamente.
Utiliza ExcelReader, Maestro y Prontuario.

Flujo:

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
"""

from pathlib import Path
import unicodedata
import re

import numpy as np
import pandas as pd

from services.excel_reader import ExcelReader
from core.homologacion import Homologador
from core.maestro import Maestro
from core.prontuario import Prontuario
from core.calculos import Calculos
from core.ciones import dor


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

                print(
                    "\n=============================================="
                )

                print(
                    "ARCHIVO:",
                    archivo.name
                )

                print(
                    "COLUMNAS:",
                    df.columns.tolist()
                )

                # ------------------------------------------
                # Trazabilidad
                # ------------------------------------------

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

        consolidado = consolidado.reset_index(
            drop=True
        )

        consolidado["ESTADO"] = "OK"

        consolidado["OBSERVACIONES"] = ""

        return consolidado

    # ==========================================================
    # Homologar información
    # ==========================================================

    def _homologar(self, df):

        df = self.homologador.dataframe(df)

        # ------------------------------------------------------
        # Normalizar texto
        # ------------------------------------------------------

        def normalizar_texto(texto):

            if pd.isna(texto):

                return ""

            texto = str(texto).upper().strip()

            texto = unicodedata.normalize(
                "NFKD",
                texto
            )

            texto = "".join(
                c
                for c in texto
                if not unicodedata.combining(c)
            )

            texto = " ".join(
                texto.split()
            )

            return texto

        # ------------------------------------------------------
        # Normalizar campos clave
        # ------------------------------------------------------

        columnas_normalizar = [

            "HACIENDA",
            "SUERTE",
            "PRODUCTO"

        ]

        for col in columnas_normalizar:

            if col in df.columns:

                df[col] = df[col].apply(
                    normalizar_texto
                )

        # ------------------------------------------------------
        # Normalizar nombres de columnas
        # ------------------------------------------------------

        def limpiar_nombre(col):

            name = str(col).strip()

            name = re.sub(
                r'[\u200B-\u200D\uFEFF]',
                '',
                name
            )

            name = re.sub(
                r'[\u2010-\u2015]',
                '-',
                name
            )

            name = re.sub(
                r'\s*-\s*',
                ' - ',
                name
            )

            name = re.sub(
                r"[\s\u00A0]+",
                " ",
                name
            )

            name = re.sub(
                r"\s+",
                " ",
                name
            )

            return name

        df.rename(
            columns=lambda c: limpiar_nombre(c),
            inplace=True
        )

        # ------------------------------------------------------
        # Reconocer variantes de columnas de unidades
        # ------------------------------------------------------

        def normalizar_columna(col):

            return re.sub(
                r"[^A-Za-z0-9]",
                "",
                str(col)
            ).upper()

        unidades_esperadas = [

            "UNIDADES - N",
            "UNIDADES - P",
            "UNIDADES - K",
            "UNIDADES - S",
            "UNIDADES - MENORES"

        ]

        for expected in unidades_esperadas:

            if expected in df.columns:

                continue

            elemento = expected.split(
                " - "
            )[1]

            objetivo = normalizar_columna(
                f"UNIDADES{elemento}"
            )

            for col in df.columns:

                if (
                    normalizar_columna(col)
                    == objetivo
                ):

                    df.rename(
                        columns={
                            col: expected
                        },
                        inplace=True
                    )

                    break

        # ------------------------------------------------------
        # Asegurar columnas de unidades
        # ------------------------------------------------------

        for columna in unidades_esperadas:

            if columna not in df.columns:

                df[columna] = pd.NA

        return df

    # ==========================================================
    # Completar desde prontuario
    # ==========================================================

    def _completar_prontuario(self, df):

        prontuario = self.prontuario.df.copy()

        prontuario = prontuario.rename(
            columns={

                "AREA":
                    "AREA_PRONTUARIO",

                "VARIEDAD":
                    "VARIEDAD_PRONTUARIO",

                "ULT_CORTE":
                    "ULT_CORTE_PRONTUARIO",

                "TCH ACTUAL":
                    "TCH_ACTUAL_PRONTUARIO"

            }
        )

        columnas = [

            "HACIENDA",
            "SUERTE",
            "AREA_PRONTUARIO",
            "VARIEDAD_PRONTUARIO",
            "ULT_CORTE_PRONTUARIO",
            "TCH_ACTUAL_PRONTUARIO"

        ]

        prontuario = prontuario[columnas]

        df = df.merge(

            prontuario,

            on=[
                "HACIENDA",
                "SUERTE"
            ],

            how="left"

        )

        # ------------------------------------------------------
        # Completar área
        # ------------------------------------------------------

        if "AREA" in df.columns:

            df["AREA"] = df["AREA"].fillna(
                df["AREA_PRONTUARIO"]
            )

        # ------------------------------------------------------
        # Agregar información del prontuario
        # ------------------------------------------------------

        if "VARIEDAD" not in df.columns:

            df["VARIEDAD"] = (
                df["VARIEDAD_PRONTUARIO"]
            )

        if "ULT_CORTE" not in df.columns:

            df["ULT_CORTE"] = (
                df["ULT_CORTE_PRONTUARIO"]
            )

        if "TCH_ACTUAL" not in df.columns:

            df["TCH_ACTUAL"] = (
                df["TCH_ACTUAL_PRONTUARIO"]
            )

        # ------------------------------------------------------
        # Eliminar auxiliares
        # ------------------------------------------------------

        df.drop(

            columns=[

                "AREA_PRONTUARIO",
                "VARIEDAD_PRONTUARIO",
                "ULT_CORTE_PRONTUARIO",
                "TCH_ACTUAL_PRONTUARIO"

            ],

            errors="ignore",

            inplace=True

        )

        return df

    # ==========================================================
    # Completar desde Maestro
    # ==========================================================

    def _completar_maestro(self, df):

        maestro = self.maestro.df.copy()

        maestro = maestro.rename(
            columns={

                "NOMBRE COMERCIAL":
                    "PRODUCTO",

                "N":
                    "PORC_N",

                "P":
                    "PORC_P",

                "K":
                    "PORC_K",

                "S":
                    "PORC_S",

                "ELEMENTOS MENORES":
                    "PORC_MENORES"

            }
        )

        # ------------------------------------------------------
        # Normalizar producto
        # ------------------------------------------------------

        def normalizar_texto(texto):

            if pd.isna(texto):

                return ""

            texto = str(texto).upper().strip()

            texto = unicodedata.normalize(
                "NFKD",
                texto
            )

            texto = "".join(
                c
                for c in texto
                if not unicodedata.combining(c)
            )

            texto = " ".join(
                texto.split()
            )

            return texto

        maestro["PRODUCTO"] = (
            maestro["PRODUCTO"]
            .apply(normalizar_texto)
        )

        columnas = [

            "PRODUCTO",
            "PORC_N",
            "PORC_P",
            "PORC_K",
            "PORC_S",
            "PORC_MENORES"

        ]

        maestro = maestro[columnas]

        df = df.merge(

            maestro,

            on="PRODUCTO",

            how="left"

        )

        return df

    # ==========================================================
    # Calcular Área
    # ==========================================================

    def _calcular_area(self, df):

        if "AREA" not in df.columns:

            return df

        mascara = (

            df["AREA"].isna()

            & df["CANTIDAD"].notna()

            & df["DOSIS X HA"].notna()

        )

        area = pd.to_numeric(
            df.loc[mascara, "AREA"],
            errors="coerce"
        )

        cantidad = pd.to_numeric(
            df.loc[mascara, "CANTIDAD"],
            errors="coerce"
        )

        dosis = pd.to_numeric(
            df.loc[mascara, "DOSIS X HA"],
            errors="coerce"
        )

        mascara_valida = dosis != 0

        indices = df.loc[
            mascara
        ].index[
            mascara_valida
        ]

        df.loc[
            indices,
            "AREA"
        ] = (

            cantidad.loc[indices]
            /
            dosis.loc[indices]

        ).round(4)

        return df

    # ==========================================================
    # Calcular Cantidad
    # ==========================================================

    def _calcular_cantidad(self, df):

        mascara = (

            df["CANTIDAD"].isna()

            & df["AREA"].notna()

            & df["DOSIS X HA"].notna()

        )

        area = pd.to_numeric(
            df.loc[mascara, "AREA"],
            errors="coerce"
        )

        dosis = pd.to_numeric(
            df.loc[mascara, "DOSIS X HA"],
            errors="coerce"
        )

        mascara_valida = (
            area.notna()
            & dosis.notna()
        )

        indices = df.loc[
            mascara
        ].index[
            mascara_valida
        ]

        df.loc[
            indices,
            "CANTIDAD"
        ] = (

            area.loc[indices]
            *
            dosis.loc[indices]

        ).round(2)

        return df

    # ==========================================================
    # Calcular Dosis
    # ==========================================================

    def _calcular_dosis(self, df):

        mascara = (

            df["DOSIS X HA"].isna()

            & df["AREA"].notna()

            & df["CANTIDAD"].notna()

        )

        area = pd.to_numeric(
            df.loc[mascara, "AREA"],
            errors="coerce"
        )

        cantidad = pd.to_numeric(
            df.loc[mascara, "CANTIDAD"],
            errors="coerce"
        )

        mascara_valida = (
            area.notna()
            & cantidad.notna()
            & (area != 0)
        )

        indices = df.loc[
            mascara
        ].index[
            mascara_valida
        ]

        df.loc[
            indices,
            "DOSIS X HA"
        ] = (

            cantidad.loc[indices]
            /
            area.loc[indices]

        ).round(2)

        return df

    # ==========================================================
    # CALCULAR UNIDADES DE ELEMENTOS
    # ==========================================================

    def _calcular_elementos(self, df):

        print(
            "\n=============================="
        )

        print(
            "INICIO _calcular_elementos"
        )

        print(
            "=============================="
        )

        elementos = {

            "N": "PORC_N",

            "P": "PORC_P",

            "K": "PORC_K",

            "S": "PORC_S",

            "MENORES": "PORC_MENORES"

        }

        # ------------------------------------------------------
        # Verificar CANTIDAD
        # ------------------------------------------------------

        if "CANTIDAD" not in df.columns:

            print(
                "No existe CANTIDAD."
            )

            return df

        # ------------------------------------------------------
        # Asegurar cantidad numérica
        # ------------------------------------------------------

        df["CANTIDAD"] = pd.to_numeric(
            df["CANTIDAD"],
            errors="coerce"
        )

        # ------------------------------------------------------
        # Procesar elementos
        # ------------------------------------------------------

        for elemento, columna_porcentaje in elementos.items():

            nombre_unidades = (
                f"UNIDADES - {elemento}"
            )

            # --------------------------------------------------
            # Crear columna si no existe
            # --------------------------------------------------

            if nombre_unidades not in df.columns:

                df[nombre_unidades] = pd.NA

            # --------------------------------------------------
            # Si no existe porcentaje no podemos calcular
            # --------------------------------------------------

            if columna_porcentaje not in df.columns:

                continue

            df[columna_porcentaje] = pd.to_numeric(
                df[columna_porcentaje],
                errors="coerce"
            )

            # --------------------------------------------------
            # MUY IMPORTANTE
            #
            # Convertimos solamente para evaluar.
            # NO reemplazamos todavía la columna.
            # --------------------------------------------------

            unidades_actuales = pd.to_numeric(
                df[nombre_unidades],
                errors="coerce"
            )

            # --------------------------------------------------
            # Calcular únicamente donde:
            #
            # 1. No existe unidad
            # 2. Existe cantidad
            # 3. Existe porcentaje
            # 4. Porcentaje > 0
            # --------------------------------------------------

            mascara = (

                unidades_actuales.isna()

                & df["CANTIDAD"].notna()

                & df[columna_porcentaje].notna()

                & (
                    df[columna_porcentaje] > 0
                )

            )

            cantidad_calcular = (
                df.loc[
                    mascara,
                    "CANTIDAD"
                ]
            )

            porcentaje_calcular = (
                df.loc[
                    mascara,
                    columna_porcentaje
                ]
            )

            df.loc[
                mascara,
                nombre_unidades
            ] = (

                cantidad_calcular
                *
                porcentaje_calcular
                /
                100

            ).round(4)

            print(
                f"{nombre_unidades}: "
                f"{mascara.sum()} calculados"
            )

        print(
            "FIN _calcular_elementos"
        )

        print(
            "==============================\n"
        )

        return df

    # ==========================================================
    # Calcular Producto a partir de unidades suministradas
    # ==========================================================

    def _calcular_producto(self, df):

        print(
            "\n=============================="
        )

        print(
            "INICIO _calcular_producto"
        )

        print(
            "=============================="
        )

        for idx in df.index:

            # --------------------------------------------------
            # Cantidad actual
            # --------------------------------------------------

            cantidad = Calculos.numero(
                df.at[
                    idx,
                    "CANTIDAD"
                ]
            )

            # --------------------------------------------------
            # Si ya tiene cantidad,
            # NO modificarla
            # --------------------------------------------------

            if not pd.isna(cantidad):

                continue

            # --------------------------------------------------
            # Producto
            # --------------------------------------------------

            producto = df.at[
                idx,
                "PRODUCTO"
            ]

            if pd.isna(producto):

                continue

            # --------------------------------------------------
            # Buscar producto en Maestro
            # --------------------------------------------------

            aportes = (
                self.maestro.obtener_aportes(
                    producto
                )
            )

            if aportes is None:

                continue

            # --------------------------------------------------
            # Buscar unidades suministradas
            # --------------------------------------------------

            for elemento in [

                "N",
                "P",
                "K",
                "S",
                "MENORES"

            ]:

                columna_unidades = (
                    f"UNIDADES - {elemento}"
                )

                if (
                    columna_unidades
                    not in df.columns
                ):

                    continue

                unidades = Calculos.numero(
                    df.at[
                        idx,
                        columna_unidades
                    ]
                )

                porcentaje = Calculos.numero(
                    aportes.get(
                        elemento,
                        0
                    )
                )

                # --------------------------------------------------
                # Si existen unidades y porcentaje,
                # calcular cantidad
                # --------------------------------------------------

                if Calculos.puede_calcular_cantidad_por_elemento(

                    cantidad,
                    unidades,
                    porcentaje

                ):

                    cantidad_calculada = (
                        Calculos.calcular_producto(
                            unidades,
                            porcentaje
                        )
                    )

                    if not pd.isna(
                        cantidad_calculada
                    ):

                        df.at[
                            idx,
                            "CANTIDAD"
                        ] = cantidad_calculada

                        cantidad = (
                            cantidad_calculada
                        )

                        break

            # --------------------------------------------------
            # Después de calcular cantidad,
            # intentar calcular dosis
            # --------------------------------------------------

            area = Calculos.numero(
                df.at[
                    idx,
                    "AREA"
                ]
            )

            dosis = Calculos.numero(
                df.at[
                    idx,
                    "DOSIS X HA"
                ]
            )

            # --------------------------------------------------
            # AREA + CANTIDAD → DOSIS
            # --------------------------------------------------

            if (

                pd.isna(dosis)

                and not pd.isna(area)

                and not pd.isna(cantidad)

            ):

                nueva_dosis = (
                    Calculos.calcular_dosis(
                        area,
                        cantidad
                    )
                )

                df.at[
                    idx,
                    "DOSIS X HA"
                ] = nueva_dosis

                dosis = nueva_dosis

            # --------------------------------------------------
            # DOSIS + CANTIDAD → AREA
            # --------------------------------------------------

            if (

                pd.isna(area)

                and not pd.isna(dosis)

                and not pd.isna(cantidad)

            ):

                nueva_area = (
                    Calculos.calcular_area(
                        cantidad,
                        dosis
                    )
                )

                df.at[
                    idx,
                    "AREA"
                ] = nueva_area

        print(
            "FIN _calcular_producto"
        )

        print(
            "==============================\n"
        )

        return df

    # ==========================================================
    # ORQUESTADOR DE CÁLCULOS
    # ==========================================================

    def _calcular(self, df):

        print(
            "\n########################################"
        )

        print(
            "INICIO PROCESO DE CÁLCULOS"
        )

        print(
            "########################################"
        )

        # ------------------------------------------------------
        # 1. AREA desde CANTIDAD / DOSIS
        # ------------------------------------------------------

        df = self._calcular_area(df)

        # ------------------------------------------------------
        # 2. CANTIDAD desde AREA × DOSIS
        # ------------------------------------------------------

        df = self._calcular_cantidad(df)

        # ------------------------------------------------------
        # 3. DOSIS desde CANTIDAD / AREA
        # ------------------------------------------------------

        df = self._calcular_dosis(df)

        # ------------------------------------------------------
        # 4. UNIDADES desde CANTIDAD × porcentaje
        #
        # Solo llena unidades vacías.
        # ------------------------------------------------------

        df = self._calcular_elementos(df)

        # ------------------------------------------------------
        # 5. Si todavía falta CANTIDAD,
        # buscar unidades suministradas por usuario.
        # ------------------------------------------------------

        df = self._calcular_producto(df)

        # ------------------------------------------------------
        # 6. Volver a completar AREA
        # ------------------------------------------------------

        df = self._calcular_area(df)

        # ------------------------------------------------------
        # 7. Volver a completar DOSIS
        # ------------------------------------------------------

        df = self._calcular_dosis(df)

        # ------------------------------------------------------
        # 8. Ahora que CANTIDAD pudo haber sido calculada,
        # completar las unidades que sigan vacías.
        #
        # Las unidades existentes se conservan.
        # ------------------------------------------------------

        df = self._calcular_elementos(df)

        print(
            "\n########################################"
        )

        print(
            "FIN PROCESO DE CÁLCULOS"
        )

        print(
            "########################################\n"
        )

        return df

    # ==========================================================
    # Validar información
    # ==========================================================

    def _validar(self, df):

        # ------------------------------------------------------
        # Validaciones generales
        # ------------------------------------------------------

        self.validador.validar_hacienda(df)

        self.validador.validar_suerte(df)

        self._validar_unidades_producto(df)

        # ------------------------------------------------------
        # Validar AREA
        # ------------------------------------------------------

        columnas = [
            "AREA",
            "CANTIDAD",
            "DOSIS X HA"
        ]

        if all(
            col in df.columns
            for col in columnas
        ):

            area = pd.to_numeric(
                df["AREA"],
                errors="coerce"
            )

            cantidad = pd.to_numeric(
                df["CANTIDAD"],
                errors="coerce"
            )

            dosis = pd.to_numeric(
                df["DOSIS X HA"],
                errors="coerce"
            )

            # --------------------------------------------------
            # Evitar división por cero
            # --------------------------------------------------

            mascara_base = (

                area.notna()

                & cantidad.notna()

                & dosis.notna()

                & (dosis != 0)

            )

            calculada = pd.Series(
                np.nan,
                index=df.index
            )

            calculada.loc[
                mascara_base
            ] = (

                cantidad.loc[
                    mascara_base
                ]

                /

                dosis.loc[
                    mascara_base
                ]

            )

            diferencia = (
                area - calculada
            ).abs()

            mascara_error = (

                mascara_base

                & calculada.notna()

                & (diferencia > 0.01)

            )

            for idx in df.index[
                mascara_error
            ]:

                self.validador.registrar_error(

                    archivo=df.at[
                        idx,
                        "ARCHIVO_ORIGEN"
                    ],

                    fila=idx + 2,

                    campo="AREA",

                    valor=df.at[
                        idx,
                        "AREA"
                    ],

                    descripcion=(
                        "Area no coincide "
                        "con Cantidad / Dosis"
                    )

                )

                df = self.validador.marcar_error(

                    df,

                    idx,

                    "Area inconsistente"

                )

        return pd.DataFrame(
            self.validador.errores
        )

    def _validar_unidades_producto(self, df):
    
        elementos = [
            "N",
            "P",
            "K",
            "S",
            "MENORES"
        ]
    
        for idx in df.index:
    
            producto = df.at[
                idx,
                "PRODUCTO"
            ]
    
            if pd.isna(producto):
                continue
    
            producto = str(producto).strip()
    
            aportes = self.maestro.obtener_aportes(
                producto
            )
    
            if aportes is None:
                continue
    
            for elemento in elementos:
    
                columna = f"UNIDADES - {elemento}"
    
                if columna not in df.columns:
                    continue
    
                unidades = Calculos.numero(
                    df.at[
                        idx,
                        columna
                    ]
                )
    
                # ----------------------------------------------
                # Si no registraron unidades, no hay error
                # ----------------------------------------------
    
                if pd.isna(unidades):
                    continue
    
                # ----------------------------------------------
                # Aporte real del producto
                # ----------------------------------------------
    
                porcentaje = Calculos.numero(
                    aportes.get(
                        elemento,
                        0
                    )
                )
    
                if pd.isna(porcentaje):
                    porcentaje = 0
    
                # ----------------------------------------------
                # ERROR:
                # Se registraron unidades de un elemento
                # que el producto no contiene
                # ----------------------------------------------
    
                if unidades != 0 and porcentaje == 0:
    
                    descripcion = (
                        f"El producto '{producto}' "
                        f"no aporta {elemento}, "
                        f"pero se registraron "
                        f"{unidades} unidades de {elemento}."
                    )
    
                    # ------------------------------------------
                    # Guardar en self.validador.errores
                    # ------------------------------------------
    
                    self.validador.registrar_error(
    
                        archivo=df.at[
                            idx,
                            "ARCHIVO_ORIGEN"
                        ],
    
                        fila=idx + 2,
    
                        campo=columna,
    
                        valor=unidades,
    
                        descripcion=descripcion
    
                    )
    
                    # ------------------------------------------
                    # Marcar registro en el consolidado
                    # ------------------------------------------
    
                    df = self.validador.marcar_error(
    
                        df,
    
                        idx,
    
                        descripcion
    
                    )
    
        return df
