# -*- coding: utf-8 -*-

"""
Responsabilidad
---------------

Orquestar todo el proceso de consolidación.

NO realiza cálculos directamente.
NO consulta Excel directamente.

Utiliza:
- ExcelReader
- Homologador
- Maestro
- Prontuario
- Calculos
- Validador

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
from core.validaciones import Validador


class Procesador:

    # ==========================================================
    # CONSTRUCTOR
    # ==========================================================

    def __init__(self, ruta_maestro, ruta_prontuario):

        self.reader = ExcelReader()

        self.homologador = Homologador(
            ruta_maestro
        )

        self.maestro = Maestro(
            ruta_maestro
        )

        self.prontuario = Prontuario(
            ruta_prontuario
        )

        self.calculos = Calculos()

        self.validador = Validador()

        # ------------------------------------------------------
        # Columnas mínimas esperadas
        # ------------------------------------------------------

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
    # MÉTODO PRINCIPAL
    # ==========================================================

    def procesar(self, carpeta):

        # ------------------------------------------------------
        # 1. Leer archivos
        # ------------------------------------------------------

        df = self._leer_archivos(
            carpeta
        )

        # ------------------------------------------------------
        # 2. Homologar
        # ------------------------------------------------------

        df = self._homologar(
            df
        )

        # ------------------------------------------------------
        # 3. Completar desde prontuario
        # ------------------------------------------------------

        df = self._completar_prontuario(
            df
        )

        # ------------------------------------------------------
        # 4. Completar desde maestro
        # ------------------------------------------------------

        df = self._completar_maestro(
            df
        )

        # ------------------------------------------------------
        # 5. Ejecutar cálculos
        # ------------------------------------------------------

        df = self._calcular(
            df
        )

        # ------------------------------------------------------
        # 6. Validar
        # ------------------------------------------------------

        errores = self._validar(
            df
        )

        return df, errores

    # ==========================================================
    # LEER ARCHIVOS
    # ==========================================================

    def _leer_archivos(self, carpeta):

        carpeta = Path(carpeta)

        archivos = list(
            carpeta.rglob("*.xlsx")
        )

        archivos.extend(
            carpeta.rglob("*.xlsm")
        )

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

                    columnas_esperadas=(
                        self.columnas_esperadas
                    )

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

                # --------------------------------------------------
                # Trazabilidad
                # --------------------------------------------------

                df["RAZON_SOCIAL"] = (
                    archivo.parent.name
                )

                df["HACIENDA_ARCHIVO"] = (
                    archivo.stem
                )

                df["ARCHIVO_ORIGEN"] = (
                    archivo.name
                )

                lista_df.append(
                    df
                )

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

        # ------------------------------------------------------
        # Estado
        # ------------------------------------------------------

        consolidado["ESTADO"] = "OK"

        # ------------------------------------------------------
        # Observaciones
        #
        # Mantenemos ambas columnas porque algunas partes
        # de la aplicación pueden utilizar cualquiera de las
        # dos variantes.
        # ------------------------------------------------------

        consolidado["OBSERVACION"] = ""

        consolidado["OBSERVACIONES"] = ""

        return consolidado

    # ==========================================================
    # HOMOLOGAR INFORMACIÓN
    # ==========================================================

    def _homologar(self, df):

        df = self.homologador.dataframe(
            df
        )

        # ------------------------------------------------------
        # Normalizar texto
        # ------------------------------------------------------

        def normalizar_texto(texto):

            if pd.isna(texto):

                return ""

            texto = str(
                texto
            ).upper().strip()

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

                df[col] = (
                    df[col]
                    .apply(normalizar_texto)
                )

        # ------------------------------------------------------
        # Normalizar nombres de columnas
        # ------------------------------------------------------

        def limpiar_nombre(col):

            name = str(
                col
            ).strip()

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

            columns=lambda c:
                limpiar_nombre(c),

            inplace=True

        )

        # ------------------------------------------------------
        # Reconocer variantes de columnas
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
        # Asegurar columnas
        # ------------------------------------------------------

        for columna in unidades_esperadas:

            if columna not in df.columns:

                df[columna] = pd.NA

        return df

    # ==========================================================
    # COMPLETAR DESDE PRONTUARIO
    # ==========================================================

    def _completar_prontuario(self, df):

        prontuario = (
            self.prontuario.df.copy()
        )

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

        prontuario = prontuario[
            columnas
        ]

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

            df["AREA"] = (
                df["AREA"]
                .fillna(
                    df["AREA_PRONTUARIO"]
                )
            )

        # ------------------------------------------------------
        # Agregar información
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
    # COMPLETAR DESDE MAESTRO
    # ==========================================================

    def _completar_maestro(self, df):

        maestro = (
            self.maestro.df.copy()
        )

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

            texto = str(
                texto
            ).upper().strip()

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

        maestro = maestro[
            columnas
        ]

        df = df.merge(

            maestro,

            on="PRODUCTO",

            how="left"

        )

        return df

    # ==========================================================
    # CALCULAR ÁREA
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

            df.loc[
                mascara,
                "AREA"
            ],

            errors="coerce"

        )

        cantidad = pd.to_numeric(

            df.loc[
                mascara,
                "CANTIDAD"
            ],

            errors="coerce"

        )

        dosis = pd.to_numeric(

            df.loc[
                mascara,
                "DOSIS X HA"
            ],

            errors="coerce"

        )

        mascara_valida = (

            dosis.notna()
            & (dosis != 0)

        )

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
    # CALCULAR CANTIDAD
    # ==========================================================

    def _calcular_cantidad(self, df):

        mascara = (

            df["CANTIDAD"].isna()

            & df["AREA"].notna()

            & df["DOSIS X HA"].notna()

        )

        area = pd.to_numeric(

            df.loc[
                mascara,
                "AREA"
            ],

            errors="coerce"

        )

        dosis = pd.to_numeric(

            df.loc[
                mascara,
                "DOSIS X HA"
            ],

            errors="coerce"

        )

        mascara_valida = (

            area.notna()
            & dosis.notna()
            & (area != 0)

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
    # CALCULAR DOSIS
    # ==========================================================

    def _calcular_dosis(self, df):

        mascara = (

            df["DOSIS X HA"].isna()

            & df["AREA"].notna()

            & df["CANTIDAD"].notna()

        )

        area = pd.to_numeric(

            df.loc[
                mascara,
                "AREA"
            ],

            errors="coerce"

        )

        cantidad = pd.to_numeric(

            df.loc[
                mascara,
                "CANTIDAD"
            ],

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

        if "CANTIDAD" not in df.columns:

            print(
                "No existe CANTIDAD."
            )

            return df

        df["CANTIDAD"] = pd.to_numeric(

            df["CANTIDAD"],
            errors="coerce"

        )

        for elemento, columna_porcentaje in elementos.items():

            nombre_unidades = (
                f"UNIDADES - {elemento}"
            )

            if nombre_unidades not in df.columns:

                df[nombre_unidades] = pd.NA

            if columna_porcentaje not in df.columns:

                continue

            df[columna_porcentaje] = pd.to_numeric(

                df[columna_porcentaje],
                errors="coerce"

            )

            unidades_actuales = pd.to_numeric(

                df[nombre_unidades],
                errors="coerce"

            )

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
    # CALCULAR UNIDADES POR HECTÁREA
    # ==========================================================

    def _calcular_unidades_hectarea(self, df):

        print(
            "\n=============================="
        )

        print(
            "INICIO _calcular_unidades_hectarea"
        )

        print(
            "=============================="
        )

        elementos = [

            "N",
            "P",
            "K",
            "S",
            "MENORES"

        ]

        if "AREA" not in df.columns:

            print(
                "No existe AREA."
            )

            return df

        area = pd.to_numeric(

            df["AREA"],
            errors="coerce"

        )

        for elemento in elementos:

            columna_unidades = (
                f"UNIDADES - {elemento}"
            )

            columna_unidades_ha = (
                f"UNIDADES/HA - {elemento}"
            )

            # --------------------------------------------------
            # Crear columna
            # --------------------------------------------------

            if (
                columna_unidades_ha
                not in df.columns
            ):

                df[
                    columna_unidades_ha
                ] = pd.NA

            if (
                columna_unidades
                not in df.columns
            ):

                continue

            unidades = pd.to_numeric(

                df[
                    columna_unidades
                ],

                errors="coerce"

            )

            mascara = (

                unidades.notna()

                & area.notna()

                & (area > 0)

            )

            df.loc[
                mascara,
                columna_unidades_ha
            ] = (

                unidades.loc[
                    mascara
                ]

                /

                area.loc[
                    mascara
                ]

            ).round(2)

            print(

                f"{columna_unidades_ha}: "
                f"{mascara.sum()} calculados"

            )

        print(
            "FIN _calcular_unidades_hectarea"
        )

        print(
            "==============================\n"
        )

        return df

    # ==========================================================
    # CALCULAR PRODUCTO A PARTIR DE UNIDADES
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

            cantidad = Calculos.numero(

                df.at[
                    idx,
                    "CANTIDAD"
                ]

            )

            # --------------------------------------------------
            # Si ya existe cantidad,
            # no modificarla
            # --------------------------------------------------

            if not pd.isna(cantidad):

                continue

            producto = df.at[
                idx,
                "PRODUCTO"
            ]

            if pd.isna(producto):

                continue

            aportes = (
                self.maestro.obtener_aportes(
                    producto
                )
            )

            if aportes is None:

                continue

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
                        ] = (
                            cantidad_calculada
                        )

                        cantidad = (
                            cantidad_calculada
                        )

                        break

            # --------------------------------------------------
            # AREA + CANTIDAD → DOSIS
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

        df = self._calcular_area(
            df
        )

        # ------------------------------------------------------
        # 2. CANTIDAD desde AREA × DOSIS
        # ------------------------------------------------------

        df = self._calcular_cantidad(
            df
        )

        # ------------------------------------------------------
        # 3. DOSIS desde CANTIDAD / AREA
        # ------------------------------------------------------

        df = self._calcular_dosis(
            df
        )

        # ------------------------------------------------------
        # 4. UNIDADES desde CANTIDAD × porcentaje
        # ------------------------------------------------------

        df = self._calcular_elementos(
            df
        )

        # ------------------------------------------------------
        # 5. Si falta cantidad,
        # calcularla desde unidades
        # ------------------------------------------------------

        df = self._calcular_producto(
            df
        )

        # ------------------------------------------------------
        # 6. Volver a completar AREA
        # ------------------------------------------------------

        df = self._calcular_area(
            df
        )

        # ------------------------------------------------------
        # 7. Volver a completar DOSIS
        # ------------------------------------------------------

        df = self._calcular_dosis(
            df
        )

        # ------------------------------------------------------
        # 8. Completar unidades nuevamente
        # ------------------------------------------------------

        df = self._calcular_elementos(
            df
        )

        # ------------------------------------------------------
        # 9. CALCULAR UNIDADES POR HECTÁREA
        #
        # IMPORTANTE:
        # Este cálculo debe hacerse DESPUÉS de tener:
        #
        # AREA
        # CANTIDAD
        # UNIDADES - N
        # UNIDADES - P
        # UNIDADES - K
        # UNIDADES - S
        # UNIDADES - MENORES
        #
        # De esta forma las columnas quedan listas para
        # enviarse al consolidado.
        # ------------------------------------------------------

        df = self._calcular_unidades_hectarea(
            df
        )

        # ------------------------------------------------------
        # 10. Segundo cálculo de unidades/ha
        #
        # Esto protege el caso en que alguna cantidad o unidad
        # haya sido calculada durante el paso anterior.
        # ------------------------------------------------------

        df = self._calcular_unidades_hectarea(
            df
        )

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
    # VALIDAR INFORMACIÓN
    # ==========================================================

    def _validar(self, df):

        # ------------------------------------------------------
        # ASEGURAR ESTADO
        # ------------------------------------------------------

        if "ESTADO" not in df.columns:

            df["ESTADO"] = "OK"

        else:

            df["ESTADO"] = (

                df["ESTADO"]
                .fillna("OK")
                .astype(str)

            )

        # ------------------------------------------------------
        # ASEGURAR OBSERVACION
        # ------------------------------------------------------

        if "OBSERVACION" not in df.columns:

            df["OBSERVACION"] = ""

        else:

            df["OBSERVACION"] = (

                df["OBSERVACION"]
                .fillna("")
                .astype(str)

            )

        # ------------------------------------------------------
        # NO BORRAR OBSERVACIONES EXISTENTES
        #
        # Antes se hacía:
        #
        # df["OBSERVACION"] = ""
        #
        # Eso podía eliminar información.
        # ------------------------------------------------------

        # ------------------------------------------------------
        # VALIDAR HACIENDA
        # ------------------------------------------------------

        df = self._validar_hacienda(
            df
        )

        # ------------------------------------------------------
        # VALIDAR SUERTE
        # ------------------------------------------------------

        df = self._validar_suerte(
            df
        )

        # ------------------------------------------------------
        # VALIDAR UNIDADES
        # ------------------------------------------------------

        df = self._validar_unidades_producto(
            df
        )

        # ------------------------------------------------------
        # VALIDAR ÁREA
        # ------------------------------------------------------

        df = self._validar_area(
            df
        )

        # ------------------------------------------------------
        # VALIDAR UNIDADES/HA
        # ------------------------------------------------------

        df = self._validar_unidades_hectarea(
            df
        )

        # ------------------------------------------------------
        # PASAR ERRORES A OBSERVACION
        # ------------------------------------------------------

        df = self._agregar_observaciones(
            df
        )

        # ------------------------------------------------------
        # TABLA DE ERRORES
        # ------------------------------------------------------

        return pd.DataFrame(
            self.validador.errores
        )

    # ==========================================================
    # VALIDAR HACIENDA
    # ==========================================================

    def _validar_hacienda(self, df):

        if "HACIENDA" not in df.columns:

            return df

        prontuario = (
            self.prontuario.df.copy()
        )

        if "HACIENDA" not in prontuario.columns:

            return df

        haciendas_prontuario = set(

            prontuario["HACIENDA"]
            .dropna()
            .astype(str)
            .str.strip()
            .str.upper()

        )

        for idx in df.index:

            hacienda_original = df.at[
                idx,
                "HACIENDA"
            ]

            if (

                pd.isna(hacienda_original)

                or

                str(
                    hacienda_original
                ).strip() == ""

            ):

                descripcion = (
                    "Código de Hacienda vacío."
                )

                self.validador.registrar_error(

                    archivo=df.at[
                        idx,
                        "ARCHIVO_ORIGEN"
                    ],

                    fila=idx + 2,

                    campo="HACIENDA",

                    valor=hacienda_original,

                    descripcion=descripcion

                )

                df = self.validador.marcar_error(

                    df,
                    idx,
                    descripcion

                )

                continue

            hacienda = str(
                hacienda_original
            ).strip().upper()

            if not re.fullmatch(
                r"\d{6}",
                hacienda
            ):

                descripcion = (

                    f"Código de Hacienda inválido: "
                    f"'{hacienda}'. "
                    f"Debe contener exactamente 6 dígitos."

                )

                self.validador.registrar_error(

                    archivo=df.at[
                        idx,
                        "ARCHIVO_ORIGEN"
                    ],

                    fila=idx + 2,

                    campo="HACIENDA",

                    valor=hacienda,

                    descripcion=descripcion

                )

                df = self.validador.marcar_error(

                    df,
                    idx,
                    descripcion

                )

                continue

            if hacienda not in haciendas_prontuario:

                descripcion = (

                    f"El código de Hacienda "
                    f"'{hacienda}' no existe "
                    f"en el prontuario."

                )

                self.validador.registrar_error(

                    archivo=df.at[
                        idx,
                        "ARCHIVO_ORIGEN"
                    ],

                    fila=idx + 2,

                    campo="HACIENDA",

                    valor=hacienda,

                    descripcion=descripcion

                )

                df = self.validador.marcar_error(

                    df,
                    idx,
                    descripcion

                )

        return df

    # ==========================================================
    # VALIDAR SUERTE
    # ==========================================================

    def _validar_suerte(self, df):

        if "SUERTE" not in df.columns:

            return df

        prontuario = (
            self.prontuario.df.copy()
        )

        tiene_prontuario = (

            "HACIENDA" in prontuario.columns

            and

            "SUERTE" in prontuario.columns

        )

        combinaciones_validas = set()

        if tiene_prontuario:

            for _, fila in prontuario.iterrows():

                hacienda = fila.get(
                    "HACIENDA"
                )

                suerte = fila.get(
                    "SUERTE"
                )

                if (

                    pd.isna(hacienda)

                    or

                    pd.isna(suerte)

                ):

                    continue

                hacienda = str(
                    hacienda
                ).strip().upper()

                suerte = str(
                    suerte
                ).strip().upper()

                combinaciones_validas.add(

                    (
                        hacienda,
                        suerte
                    )

                )

        for idx in df.index:

            suerte_original = df.at[
                idx,
                "SUERTE"
            ]

            hacienda_original = df.at[
                idx,
                "HACIENDA"
            ]

            if (

                pd.isna(suerte_original)

                or

                str(
                    suerte_original
                ).strip() == ""

            ):

                descripcion = (
                    "Código de Suerte vacío."
                )

                self.validador.registrar_error(

                    archivo=df.at[
                        idx,
                        "ARCHIVO_ORIGEN"
                    ],

                    fila=idx + 2,

                    campo="SUERTE",

                    valor=suerte_original,

                    descripcion=descripcion

                )

                df = self.validador.marcar_error(

                    df,
                    idx,
                    descripcion

                )

                continue

            suerte = str(
                suerte_original
            ).strip().upper()

            hacienda = str(
                hacienda_original
            ).strip().upper()

            if not re.fullmatch(
                r"\d{3}[A-Z]?",
                suerte
            ):

                descripcion = (

                    f"Código de Suerte inválido: "
                    f"'{suerte}'. "
                    f"Debe tener 3 dígitos "
                    f"o 3 dígitos y una letra."

                )

                self.validador.registrar_error(

                    archivo=df.at[
                        idx,
                        "ARCHIVO_ORIGEN"
                    ],

                    fila=idx + 2,

                    campo="SUERTE",

                    valor=suerte,

                    descripcion=descripcion

                )

                df = self.validador.marcar_error(

                    df,
                    idx,
                    descripcion

                )

                continue

            if (

                re.fullmatch(
                    r"\d{6}",
                    hacienda
                )

                and

                tiene_prontuario

            ):

                combinacion = (

                    hacienda,
                    suerte

                )

                if (
                    combinacion
                    not in combinaciones_validas
                ):

                    descripcion = (

                        f"La combinación "
                        f"Hacienda '{hacienda}' "
                        f"+ Suerte '{suerte}' "
                        f"no existe en el prontuario."

                    )

                    self.validador.registrar_error(

                        archivo=df.at[
                            idx,
                            "ARCHIVO_ORIGEN"
                        ],

                        fila=idx + 2,

                        campo="SUERTE",

                        valor=suerte,

                        descripcion=descripcion

                    )

                    df = self.validador.marcar_error(

                        df,
                        idx,
                        descripcion

                    )

        return df

    # ==========================================================
    # VALIDAR UNIDADES DEL PRODUCTO
    # ==========================================================

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

            producto = str(
                producto
            ).strip()

            aportes = (
                self.maestro.obtener_aportes(
                    producto
                )
            )

            if aportes is None:

                continue

            for elemento in elementos:

                columna = (
                    f"UNIDADES - {elemento}"
                )

                if columna not in df.columns:

                    continue

                unidades = Calculos.numero(

                    df.at[
                        idx,
                        columna
                    ]

                )

                if pd.isna(unidades):

                    continue

                porcentaje = Calculos.numero(

                    aportes.get(
                        elemento,
                        0
                    )

                )

                if pd.isna(porcentaje):

                    porcentaje = 0

                if (

                    unidades != 0

                    and

                    porcentaje == 0

                ):

                    descripcion = (

                        f"El producto '{producto}' "
                        f"no aporta {elemento}, "
                        f"pero se registraron "
                        f"{unidades} unidades de "
                        f"{elemento}."

                    )

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

                    df = self.validador.marcar_error(

                        df,
                        idx,
                        descripcion

                    )

        return df

    # ==========================================================
    # VALIDAR ÁREA
    # ==========================================================

    def _validar_area(self, df):

        columnas = [

            "AREA",
            "CANTIDAD",
            "DOSIS X HA"

        ]

        if not all(

            col in df.columns

            for col in columnas

        ):

            return df

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

            area
            -
            calculada

        ).abs()

        mascara_error = (

            mascara_base

            & calculada.notna()

            & (
                diferencia > 0.01
            )

        )

        for idx in df.index[
            mascara_error
        ]:

            descripcion = (

                "Área inconsistente: "
                "el área registrada "
                f"({area.loc[idx]:.4f}) "
                "no coincide con "
                "Cantidad / Dosis "
                f"({calculada.loc[idx]:.4f})."

            )

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

                descripcion=descripcion

            )

            df = self.validador.marcar_error(

                df,
                idx,
                descripcion

            )

        return df

    # ==========================================================
    # VALIDAR UNIDADES POR HECTÁREA
    # ==========================================================

    def _validar_unidades_hectarea(self, df):

        rangos = {

            "N": {
                "min": 140,
                "max": 180
            },

            "P": {
                "min": 25,
                "max": 50
            },

            "K": {
                "min": 30,
                "max": 90
            }

        }

        if "AREA" not in df.columns:

            return df

        area = pd.to_numeric(

            df["AREA"],
            errors="coerce"

        )

        for idx in df.index:

            area_valor = area.loc[idx]

            if pd.isna(area_valor):

                continue

            if area_valor == 0:

                continue

            for elemento, rango in rangos.items():

                # --------------------------------------------------
                # IMPORTANTE:
                # Usar directamente la columna calculada
                # UNIDADES/HA.
                # --------------------------------------------------

                columna_unidades_ha = (

                    f"UNIDADES/HA - {elemento}"

                )

                if (
                    columna_unidades_ha
                    not in df.columns
                ):

                    continue

                unidades_ha = Calculos.numero(

                    df.at[
                        idx,
                        columna_unidades_ha
                    ]

                )

                if pd.isna(unidades_ha):

                    continue

                minimo = rango["min"]

                maximo = rango["max"]

                if (

                    unidades_ha < minimo

                    or

                    unidades_ha > maximo

                ):

                    descripcion = (

                        f"Unidades de {elemento} "
                        f"fuera de rango: "
                        f"{unidades_ha} unidades/ha. "
                        f"Rango permitido: "
                        f"{minimo} - {maximo} unidades/ha."

                    )

                    self.validador.registrar_error(

                        archivo=df.at[
                            idx,
                            "ARCHIVO_ORIGEN"
                        ],

                        fila=idx + 2,

                        campo=columna_unidades_ha,

                        valor=unidades_ha,

                        descripcion=descripcion

                    )

                    df = self.validador.marcar_error(

                        df,
                        idx,
                        descripcion

                    )

        return df

    # ==========================================================
    # AGREGAR TODAS LAS OBSERVACIONES
    # ==========================================================

    def _agregar_observaciones(self, df):

        # ------------------------------------------------------
        # Asegurar OBSERVACION
        # ------------------------------------------------------

        if "OBSERVACION" not in df.columns:

            df["OBSERVACION"] = ""

        else:

            df["OBSERVACION"] = (

                df["OBSERVACION"]
                .fillna("")
                .astype(str)

            )

        # ------------------------------------------------------
        # Agregar errores
        # ------------------------------------------------------

        for error in self.validador.errores:

            fila = error.get(
                "fila"
            )

            descripcion = (

                error.get(
                    "descripcion"
                )

                or

                error.get(
                    "DESCRIPCION"
                )

                or

                error.get(
                    "ERROR"
                )

                or

                ""

            )

            if not descripcion:

                continue

            if fila is None:

                continue

            try:

                indice = (
                    int(fila) - 2
                )

                if indice not in df.index:

                    continue

                observacion_actual = str(

                    df.at[
                        indice,
                        "OBSERVACION"
                    ]

                )

                if observacion_actual in [

                    "",
                    "nan",
                    "None"

                ]:

                    df.at[
                        indice,
                        "OBSERVACION"
                    ] = descripcion

                else:

                    observaciones_existentes = (

                        observacion_actual
                        .split(" | ")

                    )

                    if (
                        descripcion
                        not in
                        observaciones_existentes
                    ):

                        df.at[
                            indice,
                            "OBSERVACION"
                        ] = (

                            observacion_actual
                            + " | "
                            + descripcion

                        )

            except Exception:

                continue

        # ------------------------------------------------------
        # Determinar estado
        # ------------------------------------------------------

        mascara_con_error = (

            df["OBSERVACION"]
            .fillna("")
            .astype(str)
            .str.strip()
            .ne("")

        )

        df.loc[
            mascara_con_error,
            "ESTADO"
        ] = "ERROR"

        # ------------------------------------------------------
        # Mantener ambas variantes
        # ------------------------------------------------------

        df["OBSERVACIONES"] = (
            df["OBSERVACION"]
        )

        return df
