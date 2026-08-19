# -*- coding: utf-8 -*-
"""
Auditoria Bonsucro
Aplicación Streamlit

@author: jddiazc
"""

import streamlit as st
import pandas as pd
import io
import tempfile

from pathlib import Path
from datetime import datetime

from core.validaciones import Validador
from core.procesador import Procesador


# ===========================================================
# CONFIGURACIÓN
# ===========================================================

st.set_page_config(
    page_title="Trazabilidad fertilización",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ===========================================================
# CSS
# ===========================================================

st.markdown(
    """
    <style>

    /* -------------------------------------------------------
       Fuente
       ------------------------------------------------------- */

    @import url(
        'https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap'
    );

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }


    /* -------------------------------------------------------
       Header
       ------------------------------------------------------- */

    .main-header {
        background: linear-gradient(
            135deg,
            #0d4b3c 0%,
            #1a7a5e 50%,
            #2dba8e 100%
        );

        padding: 2rem 2.5rem;
        border-radius: 16px;
        margin-bottom: 1.5rem;

        box-shadow:
            0 8px 32px rgba(13, 75, 60, 0.25);
    }

    .main-header h1 {
        color: white;
        font-size: 1.8rem;
        font-weight: 700;
        margin: 0;
    }

    .main-header p {
        color: rgba(255,255,255,0.85);
        font-size: 0.95rem;
        margin: 0.3rem 0 0 0;
    }


    /* -------------------------------------------------------
       Sidebar
       ------------------------------------------------------- */

    [data-testid="stSidebar"] {
        background: linear-gradient(
            180deg,
            #0d4b3c 0%,
            #0a3a2e 100%
        );
    }

    [data-testid="stSidebar"] * {
        color: white !important;
    }


    /* -------------------------------------------------------
       Date input
       ------------------------------------------------------- */

    [data-testid="stSidebar"] [data-testid="stDateInput"] input {
        background-color: white !important;
        color: #222222 !important;
        border: 1px solid #cccccc !important;
        border-radius: 8px !important;
    }

    [data-testid="stSidebar"] [data-testid="stDateInput"] input::placeholder {
        color: #666666 !important;
    }

    [data-testid="stSidebar"] [data-testid="stDateInput"] button {
        background-color: white !important;
        color: #222222 !important;
    }


    /* -------------------------------------------------------
       Secciones
       ------------------------------------------------------- */

    .section-title {
        font-size: 1.1rem;
        font-weight: 600;
        color: #0d4b3c;

        margin-top: 1.2rem;
        margin-bottom: 0.8rem;

        padding-bottom: 0.4rem;

        border-bottom: 2px solid #2dba8e;

        display: inline-block;
    }


    /* -------------------------------------------------------
       Footer
       ------------------------------------------------------- */

    .footer {
        text-align: center;

        padding: 1.5rem 0;

        color: #a0b5ad;

        font-size: 0.8rem;

        margin-top: 2rem;

        border-top: 1px solid #e8efe9;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ===========================================================
# CONSTANTES
# ===========================================================

COLUMNAS_ESPERADAS = [
    "HACIENDA",
    "SUERTE",
    "ÁREA",
    "PRODUCTO",
    "CANTIDAD",
    "UNIDAD",
    "DOSIS X HA",
    "UNIDADES - N",
    "UNIDADES - P",
    "UNIDADES - K",
    "MENORES"
]


COLUMNAS_CLAVE = [
    "HACIENDA",
    "FECHAS",
    "SUERTE",
    "AREA",
    "PRODUCTO",
    "DOSIS X HA",
    "UNIDAD",
    "CANTIDAD",
    "UNIDADES - N",
    "UNIDADES - P",
    "UNIDADES - K",
    "MENORES"
]


# ===========================================================
# RUTAS
# ===========================================================

RUTA_MAESTRO = (
    Path(__file__).parent
    / "config"
    / "maestro.xlsx"
)

RUTA_PRONTUARIO = (
    Path(__file__).parent
    / "config"
    / "prontuario.xls"
)


# ===========================================================
# SESSION STATE
# ===========================================================

if "consolidado" not in st.session_state:
    st.session_state["consolidado"] = pd.DataFrame()

if "errores" not in st.session_state:
    st.session_state["errores"] = []

if "detalles" not in st.session_state:
    st.session_state["detalles"] = []


# ===========================================================
# CARGAR MAESTRO
# ===========================================================

@st.cache_data
def cargar_maestro():

    ruta = (
        Path(__file__).parent
        / "config"
        / "maestro.xlsx"
    )

    if not ruta.exists():
        return None, None

    try:

        xls = pd.ExcelFile(ruta)

        fertilizantes = None
        herbicidas = None

        if "APORTE FERTILIZANTES" in xls.sheet_names:

            fertilizantes = pd.read_excel(
                xls,
                "APORTE FERTILIZANTES"
            )

        if "APORTE HERBICIDAS" in xls.sheet_names:

            herbicidas = pd.read_excel(
                xls,
                "APORTE HERBICIDAS"
            )

        return fertilizantes, herbicidas

    except Exception:

        return None, None


# ===========================================================
# PROCESADOR
# ===========================================================

def _ejecutar_procesador(carpeta):

    procesador = Procesador(
        ruta_maestro=str(RUTA_MAESTRO),
        ruta_prontuario=str(RUTA_PRONTUARIO)
    )

    df_consolidado, errores_df = (
        procesador.procesar(carpeta)
    )


    # -------------------------------------------------------
    # Errores
    # -------------------------------------------------------

    if isinstance(
        errores_df,
        pd.DataFrame
    ):

        errores_list = (
            errores_df.to_dict("records")
        )

    else:

        errores_list = (
            list(errores_df)
            if errores_df
            else []
        )


    # -------------------------------------------------------
    # Detalles
    # -------------------------------------------------------

    detalles_list = []


    if (
        not df_consolidado.empty
        and "ARCHIVO_ORIGEN"
        in df_consolidado.columns
    ):

        conteo = (
            df_consolidado
            .groupby("ARCHIVO_ORIGEN")
            .size()
        )

        for archivo_nombre, registros in conteo.items():

            detalles_list.append({

                "Archivo": archivo_nombre,

                "Registros": int(registros),

                "Estado": "OK"

            })


    # -------------------------------------------------------
    # Archivos que fallaron
    # -------------------------------------------------------

    archivos_en_consolidado = set(
        d["Archivo"]
        for d in detalles_list
    )


    for err in errores_list:

        nombre_archivo = (
            err.get("archivo")
            or err.get("ARCHIVO")
        )

        if (
            nombre_archivo
            and nombre_archivo
            not in archivos_en_consolidado
        ):

            detalles_list.append({

                "Archivo": nombre_archivo,

                "Registros": 0,

                "Estado": (
                    err.get("descripcion")
                    or err.get("ERROR")
                    or "Error"
                )

            })

            archivos_en_consolidado.add(
                nombre_archivo
            )


    return (
        df_consolidado,
        errores_list,
        detalles_list
    )


# ===========================================================
# PROCESAR ARCHIVOS SUBIDOS
# ===========================================================

def procesar_archivos(
    archivos_subidos,
    maestro=None,
    prontuario=None,
    periodo=None
):

    with tempfile.TemporaryDirectory() as carpeta_temp:

        carpeta_temp = Path(carpeta_temp)

        for archivo in archivos_subidos:

            destino = (
                carpeta_temp
                / archivo.name
            )

            with open(
                destino,
                "wb"
            ) as f:

                f.write(
                    archivo.getbuffer()
                )


        return _ejecutar_procesador(
            carpeta_temp
        )


# ===========================================================
# PROCESAR CARPETA
# ===========================================================

def cargar_desde_carpeta():

    carpeta_entrada = (
        Path(__file__).parent
        / "data"
        / "entrada"
    )

    return _ejecutar_procesador(
        carpeta_entrada
    )


# ===========================================================
# DATAFRAME → EXCEL
# ===========================================================

def to_excel_bytes(df):

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Datos"
        )

    return output.getvalue()


# ===========================================================
# GUARDAR RESULTADOS
# ===========================================================

def guardar_resultados(
    df_consolidado,
    errores_list,
    detalles_list
):

    st.session_state["consolidado"] = (
        df_consolidado
    )

    st.session_state["errores"] = (
        errores_list
    )

    st.session_state["detalles"] = (
        detalles_list
    )

# ===========================================================
# HEADER
# ===========================================================

st.title("🌿 Auditoría Bonsucro")

st.caption(
    "Procesamiento, validación y consolidación de aplicación de productos agrícolas"
)
# ===========================================================
# CARGAR MAESTRO
# ===========================================================

fertilizantes_df, herbicidas_df = (
    cargar_maestro()
)


# ===========================================================
# SIDEBAR
# ===========================================================

with st.sidebar:

    st.markdown(
        "### ⚙️ Panel de Control"
    )

    st.markdown("---")


    # -------------------------------------------------------
    # Fuente
    # -------------------------------------------------------

    modo = st.radio(
        "📂 Fuente de datos",

        [
            "Subir archivos",
            "Carpeta del proyecto"
        ]
    )


    st.markdown("---")


    # -------------------------------------------------------
    # Maestro
    # -------------------------------------------------------

    st.markdown(
        "### 📊 Maestro de Configuración"
    )


    if fertilizantes_df is not None:

        st.success(
            f"🧪 {len(fertilizantes_df)} "
            f"fertilizantes"
        )


    if herbicidas_df is not None:

        st.info(
            f"🌿 {len(herbicidas_df)} "
            f"herbicidas"
        )


    st.markdown("---")


    # -------------------------------------------------------
    # Periodo
    # -------------------------------------------------------

    st.markdown(
        "### 📅 Período de análisis"
    )


    col_fecha1, col_fecha2 = (
        st.columns(2)
    )


    with col_fecha1:

        fecha_inicio = st.date_input(
            "Desde",
            value=datetime(
                2025,
                1,
                1
            ),
            key="fecha_inicio"
        )


    with col_fecha2:

        fecha_fin = st.date_input(
            "Hasta",
            value=datetime(
                2026,
                12,
                31
            ),
            key="fecha_fin"
        )


    periodo = {

        "fecha_inicio": fecha_inicio,

        "fecha_fin": fecha_fin

    }


    # -------------------------------------------------------
    # Estado
    # -------------------------------------------------------

    st.markdown("---")

    st.markdown(
        "### ⚠️ Estado"
    )


    errores_sidebar = (
        st.session_state
        .get(
            "errores",
            []
        )
    )


    if errores_sidebar:

        df_errores_sidebar = (
            pd.DataFrame(
                errores_sidebar
            )
        )


        st.warning(
            f"Se encontraron "
            f"{len(df_errores_sidebar)} "
            f"errores"
        )


        st.download_button(

            label="📥 Descargar errores",

            data=to_excel_bytes(
                df_errores_sidebar
            ),

            file_name=(
                "Errores_Lectura_"
                f"{datetime.now().strftime('%Y%m%d_%H%M')}"
                ".xlsx"
            ),

            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            ),

            use_container_width=True

        )

    else:

        st.success(
            "✅ Sin errores de lectura"
        )


    st.markdown("---")


    st.caption(
        f"📅 {datetime.now().strftime('%d/%m/%Y')}"
    )


# ===========================================================
# PESTAÑAS
# ===========================================================

tab_visualizador, tab_procesar = st.tabs(
    [
        "⚙️ Procesar",
        "📊 Visualizador"
    ]
)


# ===========================================================
# ===========================================================
# PROCESAR
# ===========================================================
# ===========================================================

with tab_procesar:

    st.header(
        "⚙️ Procesamiento de datos"
    )


    st.write(
        "Cargue y procese los archivos de fertilización."
    )


    # =======================================================
    # SUBIR ARCHIVOS
    # =======================================================

    if modo == "Subir archivos":

        st.subheader(
            "📤 Cargar archivos Excel"
        )


        archivos = st.file_uploader(

            "Seleccione los archivos Excel",

            type=[
                "xlsx",
                "xlsm"
            ],

            accept_multiple_files=True,

            help=(
                "Los archivos deben contener "
                "la hoja FORMATO FERT"
            )
        )


        if archivos:

            st.info(
                f"📄 Archivos seleccionados: "
                f"{len(archivos)}"
            )


            if st.button(
                "🚀 Procesar archivos",
                type="primary",
                use_container_width=True,
                key="procesar_subidos"
            ):

                with st.spinner(
                    "⏳ Procesando archivos..."
                ):

                    (
                        df_consolidado,
                        errores_list,
                        detalles_list
                    ) = procesar_archivos(

                        archivos_subidos=archivos,

                        maestro=fertilizantes_df,

                        prontuario=None,

                        periodo=periodo

                    )


                    guardar_resultados(

                        df_consolidado,

                        errores_list,

                        detalles_list

                    )


                st.success(
                    "✅ Procesamiento terminado correctamente."
                )


    # =======================================================
    # CARPETA
    # =======================================================

    else:

        st.subheader(
            "📁 Datos desde carpeta del proyecto"
        )


        carpeta_entrada = (
            Path(__file__).parent
            / "data"
            / "entrada"
        )


        archivos_encontrados = list(
            carpeta_entrada.rglob(
                "*.xlsx"
            )
        )


        st.info(
            f"📂 Carpeta: `data/entrada`\n\n"
            f"📄 Archivos encontrados: "
            f"**{len(archivos_encontrados)}**"
        )


        if archivos_encontrados:

            if st.button(
                "🚀 Procesar archivos",
                type="primary",
                use_container_width=True,
                key="procesar_carpeta"
            ):

                with st.spinner(
                    "⏳ Procesando archivos..."
                ):

                    (
                        df_consolidado,
                        errores_list,
                        detalles_list
                    ) = cargar_desde_carpeta()


                    guardar_resultados(

                        df_consolidado,

                        errores_list,

                        detalles_list

                    )


                st.success(
                    "✅ Procesamiento terminado correctamente."
                )

        else:

            st.warning(
                "⚠️ No se encontraron archivos "
                "Excel en data/entrada."
            )


    # =======================================================
    # RECUPERAR RESULTADOS
    # =======================================================

    df_consolidado = (
        st.session_state
        .get(
            "consolidado",
            pd.DataFrame()
        )
    )


    errores_list = (
        st.session_state
        .get(
            "errores",
            []
        )
    )


    detalles_list = (
        st.session_state
        .get(
            "detalles",
            []
        )
    )


    # =======================================================
    # RESULTADO POR ARCHIVO
    # =======================================================

    if detalles_list:

        st.divider()

        st.subheader(
            "📋 Resultado del procesamiento"
        )


        df_detalles = pd.DataFrame(
            detalles_list
        )


        st.dataframe(
            df_detalles,
            use_container_width=True,
            hide_index=True
        )


    # =======================================================
    # RESUMEN DEL PROCESAMIENTO
    # =======================================================

    if not df_consolidado.empty:

        st.divider()

        st.subheader(
            "Resumen"
        )


        n_archivos = (

            df_consolidado[
                "ARCHIVO_ORIGEN"
            ].nunique()

            if "ARCHIVO_ORIGEN"
            in df_consolidado.columns

            else 0
        )


        n_registros = len(
            df_consolidado
        )


        n_productos = (

            df_consolidado[
                "PRODUCTO"
            ].nunique()

            if "PRODUCTO"
            in df_consolidado.columns

            else 0
        )


        n_haciendas = (

            df_consolidado[
                "HACIENDA"
            ]
            .dropna()
            .nunique()

            if "HACIENDA"
            in df_consolidado.columns

            else 0
        )


        n_errores = len(
            errores_list
        )


        c1, c2, c3, c4, c5 = (
            st.columns(5)
        )


        with c1:
            st.metric(
                "📄 Archivos",
                n_archivos
            )


        with c2:
            st.metric(
                "📋 Registros",
                n_registros
            )


        with c3:
            st.metric(
                "🧪 Productos",
                n_productos
            )


        with c4:
            st.metric(
                "🏡 Haciendas",
                n_haciendas
            )


        with c5:
            st.metric(
                "⚠️ Errores",
                n_errores
            )


    # =======================================================
    # VALIDACIÓN
    # =======================================================

    if not df_consolidado.empty:

        st.divider()

        st.subheader(
            "✅ Validación de datos"
        )


        with st.expander(
            "🔎 Ejecutar validaciones"
        ):

            if st.button(
                "Validar datos",
                key="validar_datos"
            ):

                validador = Validador()


                validador.validar_hacienda(
                    df_consolidado
                )


                validador.validar_suerte(
                    df_consolidado
                )


                if validador.errores:

                    st.warning(
                        f"Se encontraron "
                        f"{len(validador.errores)} "
                        f"observaciones"
                    )


                    df_errores_val = (
                        pd.DataFrame(
                            validador.errores
                        )
                    )


                    st.dataframe(
                        df_errores_val,
                        use_container_width=True,
                        hide_index=True
                    )

                else:

                    st.success(
                        "✅ Todos los datos "
                        "pasaron la validación"
                    )


    # =======================================================
    # DESCARGAS
    # =======================================================

    if not df_consolidado.empty:

        st.divider()

        st.subheader(
            "💾 Descargas"
        )


        col_d1, col_d2 = (
            st.columns(2)
        )


        # ---------------------------------------------------
        # CONSOLIDADO
        # ---------------------------------------------------

        with col_d1:

            columnas_exportar = [

                "ULT_CORTE",
                "RAZON_SOCIAL",
                "HACIENDA",
                "SUERTE",
                "AREA",
                "PRODUCTO",
                "DOSIS X HA",
                "UNIDAD",
                "CANTIDAD",
                "UNIDAD/HA",
                "UNIDADES - N",
                "UNIDADES - P",
                "UNIDADES - K",
                "UNIDADES - S",
                "UNIDADES - MENORES",
                "HACIENDA_ARCHIVO",
                "ARCHIVO_ORIGEN",
                "ESTADO",
                "OBSERVACIONES"

            ]


            columnas_exportar = [

                c

                for c in columnas_exportar

                if c in df_consolidado.columns

            ]


            df_exportar = (
                df_consolidado[
                    columnas_exportar
                ].copy()
            )


            st.download_button(

                label=(
                    "📥 Descargar consolidado"
                ),

                data=to_excel_bytes(
                    df_exportar
                ),

                file_name=(
                    "Consolidado_Fertilizacion_"
                    f"{datetime.now().strftime('%Y%m%d_%H%M')}"
                    ".xlsx"
                ),

                mime=(
                    "application/vnd.openxmlformats-officedocument"
                    ".spreadsheetml.sheet"
                ),

                use_container_width=True,

                type="primary"

            )


        # ---------------------------------------------------
        # ERRORES
        # ---------------------------------------------------

        with col_d2:

            if errores_list:

                df_errores = (
                    pd.DataFrame(
                        errores_list
                    )
                )


                st.download_button(

                    label=(
                        "📥 Descargar errores"
                    ),

                    data=to_excel_bytes(
                        df_errores
                    ),

                    file_name=(
                        "Errores_Lectura_"
                        f"{datetime.now().strftime('%Y%m%d_%H%M')}"
                        ".xlsx"
                    ),

                    mime=(
                        "application/vnd.openxmlformats-officedocument"
                        ".spreadsheetml.sheet"
                    ),

                    use_container_width=True

                )

            else:

                st.success(
                    "✅ No hubo errores de lectura"
                )


    # =======================================================
    # REFERENCIA MAESTRO
    # =======================================================

    if (
        not df_consolidado.empty
        and fertilizantes_df is not None
        and "PRODUCTO"
        in df_consolidado.columns
    ):

        st.divider()

        st.subheader(
            "🔗 Referencia cruzada con maestro"
        )


        with st.expander(
            "📋 Ver productos del maestro"
        ):

            if (
                "NOMBRE COMERCIAL"
                in fertilizantes_df.columns
            ):

                nombres_maestro = set(

                    fertilizantes_df[
                        "NOMBRE COMERCIAL"
                    ]

                    .dropna()

                    .astype(str)

                    .str.upper()

                    .str.strip()

                )


                productos_datos = set(

                    df_consolidado[
                        "PRODUCTO"
                    ]

                    .dropna()

                    .astype(str)

                    .str.upper()

                    .str.strip()

                )


                en_maestro = (
                    productos_datos
                    &
                    nombres_maestro
                )


                no_en_maestro = (
                    productos_datos
                    -
                    nombres_maestro
                )


                col_m1, col_m2 = (
                    st.columns(2)
                )


                with col_m1:

                    st.markdown(
                        "### ✅ Encontrados"
                    )


                    if en_maestro:

                        for producto in sorted(
                            en_maestro
                        ):

                            st.write(
                                f"• {producto}"
                            )

                    else:

                        st.info(
                            "Ninguno"
                        )


                with col_m2:

                    st.markdown(
                        "### ⚠️ No encontrados"
                    )


                    if no_en_maestro:

                        for producto in sorted(
                            no_en_maestro
                        ):

                            st.write(
                                f"• {producto}"
                            )

                    else:

                        st.success(
                            "Todos los productos "
                            "coinciden con el maestro."
                        )


# ===========================================================
# ===========================================================
# VISUALIZADOR
# ===========================================================
# ===========================================================

with tab_visualizador:

    st.header(
        "📊 Visualizador"
    )


    df_consolidado = (
        st.session_state
        .get(
            "consolidado",
            pd.DataFrame()
        )
    )


    # =======================================================
    # SIN DATOS
    # =======================================================

    if df_consolidado.empty:

        st.info(
            "🌿 No hay datos procesados todavía."
        )

        st.write(
            "Vaya a la pestaña "
            "**⚙️ Procesar** para cargar "
            "y procesar los archivos."
        )


    # =======================================================
    # CON DATOS
    # =======================================================

    else:

        # ===================================================
        # MÉTRICAS
        # ===================================================

        st.subheader(
            "📈 Resumen general"
        )


        n_archivos = (

            df_consolidado[
                "ARCHIVO_ORIGEN"
            ].nunique()

            if "ARCHIVO_ORIGEN"
            in df_consolidado.columns

            else 0
        )


        n_registros = len(
            df_consolidado
        )


        n_productos = (

            df_consolidado[
                "PRODUCTO"
            ].nunique()

            if "PRODUCTO"
            in df_consolidado.columns

            else 0
        )


        n_haciendas = (

            df_consolidado[
                "HACIENDA"
            ]
            .dropna()
            .nunique()

            if "HACIENDA"
            in df_consolidado.columns

            else 0
        )


        errores_actuales = (
            st.session_state
            .get(
                "errores",
                []
            )
        )


        n_errores = len(
            errores_actuales
        )


        c1, c2, c3, c4, c5 = (
            st.columns(5)
        )


        with c1:

            st.metric(
                "📄 Archivos",
                n_archivos
            )


        with c2:

            st.metric(
                "📋 Registros",
                n_registros
            )


        with c3:

            st.metric(
                "🧪 Productos",
                n_productos
            )


        with c4:

            st.metric(
                "🏡 Haciendas",
                n_haciendas
            )


        with c5:

            st.metric(
                "⚠️ Errores",
                n_errores
            )


        # ===================================================
        # FILTROS
        # ===================================================

        st.divider()

        st.subheader(
            "🔍 Filtros"
        )


        col_f1, col_f2, col_f3 = (
            st.columns(3)
        )


        df_filtrado = (
            df_consolidado.copy()
        )


        # ---------------------------------------------------
        # ARCHIVO
        # ---------------------------------------------------

        with col_f1:

            if (
                "ARCHIVO_ORIGEN"
                in df_filtrado.columns
            ):

                archivos_unicos = sorted(

                    df_filtrado[
                        "ARCHIVO_ORIGEN"
                    ]

                    .dropna()

                    .astype(str)

                    .unique()

                )


                sel_archivos = (
                    st.multiselect(

                        "📄 Archivo",

                        options=archivos_unicos,

                        default=archivos_unicos,

                        key="viz_archivos"

                    )
                )


                df_filtrado = (
                    df_filtrado[
                        df_filtrado[
                            "ARCHIVO_ORIGEN"
                        ]
                        .astype(str)
                        .isin(
                            sel_archivos
                        )
                    ]
                )


        # ---------------------------------------------------
        # PRODUCTO
        # ---------------------------------------------------

        with col_f2:

            if (
                "PRODUCTO"
                in df_filtrado.columns
            ):

                productos_unicos = sorted(

                    df_filtrado[
                        "PRODUCTO"
                    ]

                    .dropna()

                    .astype(str)

                    .unique()

                )


                sel_productos = (
                    st.multiselect(

                        "🧪 Producto",

                        options=productos_unicos,

                        default=productos_unicos,

                        key="viz_productos"

                    )
                )


                df_filtrado = (
                    df_filtrado[
                        df_filtrado[
                            "PRODUCTO"
                        ]
                        .astype(str)
                        .isin(
                            sel_productos
                        )
                    ]
                )


        # ---------------------------------------------------
        # HACIENDA
        # ---------------------------------------------------

        with col_f3:

            if (
                "HACIENDA"
                in df_filtrado.columns
            ):

                haciendas_unicas = sorted(

                    df_filtrado[
                        "HACIENDA"
                    ]

                    .dropna()

                    .astype(str)

                    .unique()

                )


                sel_haciendas = (
                    st.multiselect(

                        "🏡 Hacienda",

                        options=haciendas_unicas,

                        default=haciendas_unicas,

                        key="viz_haciendas"

                    )
                )


                df_filtrado = (
                    df_filtrado[
                        df_filtrado[
                            "HACIENDA"
                        ]
                        .astype(str)
                        .isin(
                            sel_haciendas
                        )
                    ]
                )


        # ===================================================
        # TABLA
        # ===================================================

        st.divider()

        st.subheader(
            "📋 Datos consolidados"
        )


        columnas_mostrar = [

            c

            for c in (
                COLUMNAS_CLAVE
                +
                ["ARCHIVO_ORIGEN"]
            )

            if c in df_filtrado.columns

        ]


        if columnas_mostrar:

            st.dataframe(

                df_filtrado[
                    columnas_mostrar
                ],

                use_container_width=True,

                hide_index=True,

                height=450

            )

        else:

            st.dataframe(

                df_filtrado,

                use_container_width=True,

                hide_index=True,

                height=450

            )


        st.caption(
            f"Mostrando "
            f"{len(df_filtrado)} "
            f"de "
            f"{len(df_consolidado)} "
            f"registros"
        )


        # ===================================================
        # ANÁLISIS VISUAL
        # ===================================================

        st.divider()

        st.subheader(
            "📊 Análisis visual"
        )


        tab_producto, tab_cantidad, tab_dosis = (
            st.tabs(
                [
                    "🧪 Por producto",
                    "📦 Cantidad por producto",
                    "🌱 Dosis por producto"
                ]
            )
        )


        # ---------------------------------------------------
        # PRODUCTOS
        # ---------------------------------------------------

        with tab_producto:

            if (
                "PRODUCTO"
                in df_filtrado.columns
            ):

                conteo = (

                    df_filtrado[
                        "PRODUCTO"
                    ]

                    .dropna()

                    .astype(str)

                    .value_counts()

                    .head(15)

                )


                if not conteo.empty:

                    st.bar_chart(
                        conteo
                    )

                else:

                    st.info(
                        "No hay datos disponibles."
                    )

            else:

                st.info(
                    "No existe la columna PRODUCTO."
                )


        # ---------------------------------------------------
        # CANTIDAD
        # ---------------------------------------------------

        with tab_cantidad:

            if (
                "PRODUCTO"
                in df_filtrado.columns
                and
                "CANTIDAD"
                in df_filtrado.columns
            ):

                cantidad_prod = (

                    df_filtrado

                    .groupby(
                        "PRODUCTO"
                    )["CANTIDAD"]

                    .sum()

                    .dropna()

                    .sort_values(
                        ascending=False
                    )

                    .head(15)

                )


                if not cantidad_prod.empty:

                    st.bar_chart(
                        cantidad_prod
                    )

                else:

                    st.info(
                        "No hay datos disponibles."
                    )

            else:

                st.info(
                    "No existen las columnas "
                    "necesarias para este gráfico."
                )


        # ---------------------------------------------------
        # DOSIS
        # ---------------------------------------------------

        with tab_dosis:

            if (
                "PRODUCTO"
                in df_filtrado.columns
                and
                "DOSIS X HA"
                in df_filtrado.columns
            ):

                dosis_prod = (

                    df_filtrado

                    .groupby(
                        "PRODUCTO"
                    )["DOSIS X HA"]

                    .mean()

                    .dropna()

                    .sort_values(
                        ascending=False
                    )

                    .head(15)

                )


                if not dosis_prod.empty:

                    st.bar_chart(
                        dosis_prod
                    )

                else:

                    st.info(
                        "No hay datos disponibles."
                    )

            else:

                st.info(
                    "No existen las columnas "
                    "necesarias para este gráfico."
                )


# ===========================================================
# FOOTER
# ===========================================================

st.markdown(
    "---"
)

st.caption(
    "🌿 Módulo fertilización · "
    "Auditoría Bonsucro · 2026"
)
