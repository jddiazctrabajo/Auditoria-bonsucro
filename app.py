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

from services.excel_reader import ExcelReader
from core.validaciones import Validador
from core.procesador import Procesador


# ===========================================================
# Configuración de página
# ===========================================================

st.set_page_config(
    page_title="Trazabilidad fertilización",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ===========================================================
# Estilos CSS
# ===========================================================

st.markdown("""
<style>

    /* =======================================================
       Fuente y fondo
       ======================================================= */

    @import url(
        'https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap'
    );

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }


    /* =======================================================
       Header principal
       ======================================================= */

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
        color: #ffffff;
        font-size: 1.8rem;
        font-weight: 700;
        margin: 0 0 0.3rem 0;
        letter-spacing: -0.5px;
    }

    .main-header p {
        color: rgba(255,255,255,0.85);
        font-size: 0.95rem;
        margin: 0;
        font-weight: 300;
    }


    /* =======================================================
       Tarjetas métricas
       ======================================================= */

    .metric-card {
        background: linear-gradient(
            145deg,
            #ffffff 0%,
            #f7faf9 100%
        );

        border: 1px solid #e0ece8;
        border-radius: 14px;

        padding: 1.3rem 1.5rem;

        text-align: center;

        box-shadow:
            0 2px 12px rgba(0,0,0,0.04);

        transition:
            transform 0.2s ease,
            box-shadow 0.2s ease;
    }

    .metric-card:hover {
        transform: translateY(-2px);

        box-shadow:
            0 6px 20px rgba(0,0,0,0.08);
    }

    .metric-value {
        font-size: 2.2rem;
        font-weight: 700;
        color: #0d4b3c;

        line-height: 1;

        margin-bottom: 0.3rem;
    }

    .metric-label {
        font-size: 0.78rem;
        color: #6b8f83;

        text-transform: uppercase;
        letter-spacing: 0.8px;

        font-weight: 600;
    }


    /* =======================================================
       Status badges
       ======================================================= */

    .badge-ok {
        display: inline-block;

        background: #d4edda;
        color: #155724;

        padding: 0.25rem 0.75rem;

        border-radius: 20px;

        font-size: 0.8rem;
        font-weight: 600;
    }

    .badge-error {
        display: inline-block;

        background: #f8d7da;
        color: #721c24;

        padding: 0.25rem 0.75rem;

        border-radius: 20px;

        font-size: 0.8rem;
        font-weight: 600;
    }

    .badge-info {
        display: inline-block;

        background: #d1ecf1;
        color: #0c5460;

        padding: 0.25rem 0.75rem;

        border-radius: 20px;

        font-size: 0.8rem;
        font-weight: 600;
    }


    /* =======================================================
       Sidebar
       ======================================================= */

    [data-testid="stSidebar"] {
        background: linear-gradient(
            180deg,
            #0d4b3c 0%,
            #0a3a2e 100%
        );
    }

    [data-testid="stSidebar"] * {
        color: #ffffff !important;
    }


    /* =======================================================
       DATE INPUT - CORRECCIÓN
       ======================================================= */

    [data-testid="stSidebar"] [data-testid="stDateInput"] input {
        background-color: #ffffff !important;
        color: #222222 !important;

        border: 1px solid #d0d0d0 !important;

        border-radius: 8px !important;
    }

    [data-testid="stSidebar"] [data-testid="stDateInput"] input::placeholder {
        color: #666666 !important;
    }

    [data-testid="stSidebar"] [data-testid="stDateInput"] button {
        background-color: #ffffff !important;
        color: #333333 !important;
    }

    [data-testid="stSidebar"] [data-testid="stDateInput"] svg {
        color: #333333 !important;
        fill: #333333 !important;
    }


    /* =======================================================
       Selectbox / Multiselect
       ======================================================= */

    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stMultiSelect label {
        color: rgba(255,255,255,0.9) !important;
        font-weight: 500;
    }


    /* =======================================================
       Sección
       ======================================================= */

    .section-title {
        font-size: 1.1rem;

        font-weight: 600;

        color: #0d4b3c;

        margin: 1.5rem 0 0.8rem 0;

        padding-bottom: 0.4rem;

        border-bottom: 2px solid #2dba8e;

        display: inline-block;
    }


    /* =======================================================
       Tabla de datos
       ======================================================= */

    .stDataFrame {
        border-radius: 12px;
        overflow: hidden;
    }


    /* =======================================================
       File uploader
       ======================================================= */

    [data-testid="stFileUploader"] {
        border: 2px dashed #2dba8e;

        border-radius: 14px;

        padding: 1rem;

        background: #f7fdf9;
    }


    /* =======================================================
       Tabs
       ======================================================= */

    button[data-baseweb="tab"] {
        font-size: 1rem;
        font-weight: 600;
    }


    /* =======================================================
       Footer
       ======================================================= */

    .footer {
        text-align: center;

        padding: 1.5rem 0;

        color: #a0b5ad;

        font-size: 0.8rem;

        margin-top: 2rem;

        border-top: 1px solid #e8efe9;
    }

</style>
""", unsafe_allow_html=True)


# ===========================================================
# Constantes
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
# Rutas
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
# Inicializar session_state
# ===========================================================

if "consolidado" not in st.session_state:
    st.session_state["consolidado"] = pd.DataFrame()

if "errores" not in st.session_state:
    st.session_state["errores"] = []

if "detalles" not in st.session_state:
    st.session_state["detalles"] = []


# ===========================================================
# Funciones auxiliares
# ===========================================================

@st.cache_data
def cargar_maestro():
    """
    Carga el archivo maestro de configuración.
    """

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
# Procesador
# ===========================================================

def _ejecutar_procesador(carpeta):
    """
    Instancia el Procesador con las rutas del maestro
    y el prontuario.

    Procesa todos los archivos de la carpeta.
    """

    procesador = Procesador(
        ruta_maestro=str(RUTA_MAESTRO),
        ruta_prontuario=str(RUTA_PRONTUARIO)
    )

    df_consolidado, errores_df = (
        procesador.procesar(carpeta)
    )


    # -------------------------------------------------------
    # Convertir errores a lista
    # -------------------------------------------------------

    if isinstance(
        errores_df,
        pd.DataFrame
    ):

        errores_list = (
            errores_df
            .to_dict("records")
        )

    else:

        errores_list = (
            list(errores_df)
            if errores_df
            else []
        )


    # -------------------------------------------------------
    # Detalles por archivo
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
    # Archivos con error completo
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
# Procesar archivos subidos
# ===========================================================

def procesar_archivos(
    archivos_subidos,
    maestro,
    prontuario,
    periodo
):
    """
    Guarda temporalmente los archivos subidos
    y los procesa mediante Procesador.
    """

    with tempfile.TemporaryDirectory() as carpeta_temp:

        carpeta_temp = Path(carpeta_temp)

        for archivo in archivos_subidos:

            destino = (
                carpeta_temp
                / archivo.name
            )

            with open(destino, "wb") as f:

                f.write(
                    archivo.getbuffer()
                )

        return _ejecutar_procesador(
            carpeta_temp
        )


# ===========================================================
# Procesar desde carpeta
# ===========================================================

def cargar_desde_carpeta():
    """
    Procesa los archivos existentes en:
    data/entrada
    """

    carpeta_entrada = (
        Path(__file__).parent
        / "data"
        / "entrada"
    )

    return _ejecutar_procesador(
        carpeta_entrada
    )


# ===========================================================
# DataFrame → Excel
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
# Función para guardar resultados
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
# Header
# ===========================================================

st.markdown("""
<div class="main-header">

    <h1>
        🌿 Auditoria Bonsucro
    </h1>

    <p>
        Procesamiento, validación y consolidación
        de datos de fertilización agrícola
    </p>

</div>
""", unsafe_allow_html=True)


# ===========================================================
# Sidebar
# ===========================================================

with st.sidebar:

    st.markdown(
        "### ⚙️ Panel de Control"
    )

    st.markdown("---")


    # -------------------------------------------------------
    # Fuente de datos
    # -------------------------------------------------------

    modo = st.radio(
        "📂 Fuente de datos",

        [
            "Subir archivos",
            "Carpeta del proyecto"
        ],

        help=(
            "Seleccione cómo desea cargar "
            "los datos"
        )
    )


    st.markdown("---")


    # -------------------------------------------------------
    # Maestro
    # -------------------------------------------------------

    st.markdown(
        "### 📊 Maestro de Configuración"
    )


    fertilizantes_df, herbicidas_df = (
        cargar_maestro()
    )


    if fertilizantes_df is not None:

        st.markdown(
            f"""
            <span class="badge-ok">
                {len(fertilizantes_df)}
                fertilizantes
            </span>
            """,
            unsafe_allow_html=True
        )


    if herbicidas_df is not None:

        st.markdown(
            f"""
            <span class="badge-info">
                {len(herbicidas_df)}
                herbicidas
            </span>
            """,
            unsafe_allow_html=True
        )


    st.markdown("---")


    # -------------------------------------------------------
    # Periodo
    # -------------------------------------------------------

    st.markdown(
        "### 📅 Período de análisis"
    )


    col1, col2 = st.columns(2)


    with col1:

        fecha_inicio = st.date_input(
            "Desde",
            value=datetime(
                2025,
                1,
                1
            )
        )


    with col2:

        fecha_fin = st.date_input(
            "Hasta",
            value=datetime(
                2026,
                12,
                31
            )
        )


    periodo = {

        "fecha_inicio": fecha_inicio,

        "fecha_fin": fecha_fin

    }


    # -------------------------------------------------------
    # Estado de errores
    # -------------------------------------------------------

    errores_sidebar = (
        st.session_state
        .get("errores", [])
    )


    st.markdown("---")


    st.markdown(
        "### ⚠️ Estado del procesamiento"
    )


    if errores_sidebar:

        df_errores_sidebar = (
            pd.DataFrame(
                errores_sidebar
            )
        )


        st.warning(
            f"Se encontraron "
            f"**{len(df_errores_sidebar)}** "
            f"errores."
        )


        st.download_button(

            label=(
                "📥 Descargar errores (.xlsx)"
            ),

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
            "✅ No hay errores de lectura"
        )


    st.markdown("---")


    # -------------------------------------------------------
    # Fecha
    # -------------------------------------------------------

    st.markdown(
        f"📅 {datetime.now().strftime('%d/%m/%Y')}"
    )


# ===========================================================
# PESTAÑAS PRINCIPALES
# ===========================================================

tab_visualizador, tab_procesar = st.tabs(
    [
        "📊 Visualizador",
        "⚙️ Procesar"
    ]
)


# ===========================================================
# ===========================================================
# PESTAÑA PROCESAR
# ===========================================================
# ===========================================================

with tab_procesar:

    st.markdown(
        '<p class="section-title">'
        '⚙️ Procesamiento de datos'
        '</p>',
        unsafe_allow_html=True
    )


    # =======================================================
    # Fuente: subir archivos
    # =======================================================

    if modo == "Subir archivos":

        st.markdown(
            '<p class="section-title">'
            '📤 Cargar archivos Excel'
            '</p>',
            unsafe_allow_html=True
        )


        archivos = st.file_uploader(

            "Seleccione los archivos Excel "
            "con el formato de fertilización",

            type=[
                "xlsx",
                "xlsm"
            ],

            accept_multiple_files=True,

            help=(
                "Los archivos deben contener "
                "la hoja 'FORMATO FERT'"
            )
        )


        if archivos:

            st.info(
                f"📄 {len(archivos)} archivo(s) "
                "seleccionado(s)"
            )


            if st.button(
                "🚀 Procesar archivos",
                type="primary",
                use_container_width=True
            ):

                with st.spinner(
                    "⏳ Procesando archivos..."
                ):

                    df_consolidado, errores_list, detalles_list = (

                        procesar_archivos(

                            archivos_subidos=archivos,

                            maestro=fertilizantes_df,

                            prontuario=None,

                            periodo=periodo

                        )

                    )


                    guardar_resultados(

                        df_consolidado,

                        errores_list,

                        detalles_list

                    )


                st.success(
                    "✅ Procesamiento terminado"
                )

                st.rerun()


    # =======================================================
    # Fuente: carpeta
    # =======================================================

    else:

        st.markdown(
            '<p class="section-title">'
            '📁 Datos desde carpeta del proyecto'
            '</p>',
            unsafe_allow_html=True
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
            f"📂 Carpeta: `data/entrada`  \n"
            f"📄 Archivos encontrados: "
            f"**{len(archivos_encontrados)}**"
        )


        if archivos_encontrados:

            if st.button(
                "🚀 Procesar archivos",
                type="primary",
                use_container_width=True
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
                    "✅ Procesamiento terminado"
                )

                st.rerun()

        else:

            st.warning(
                "⚠️ No se encontraron archivos "
                "Excel en data/entrada."
            )


    # =======================================================
    # Recuperar resultados
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
    # Resultado del procesamiento
    # =======================================================

    if detalles_list:

        st.markdown("---")


        st.markdown(
            '<p class="section-title">'
            '📋 Resultado del procesamiento'
            '</p>',
            unsafe_allow_html=True
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
    # Métricas del procesamiento
    # =======================================================

    if not df_consolidado.empty:

        st.markdown("---")


        st.markdown(
            '<p class="section-title">'
            '📈 Resumen del procesamiento'
            '</p>',
            unsafe_allow_html=True
        )


        col1, col2, col3, col4, col5 = (
            st.columns(5)
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


        with col1:

            st.markdown(
                f"""
                <div class="metric-card">

                    <div class="metric-value">
                        {n_archivos}
                    </div>

                    <div class="metric-label">
                        Archivos
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )


        with col2:

            st.markdown(
                f"""
                <div class="metric-card">

                    <div class="metric-value">
                        {n_registros}
                    </div>

                    <div class="metric-label">
                        Registros
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )


        with col3:

            st.markdown(
                f"""
                <div class="metric-card">

                    <div class="metric-value">
                        {n_productos}
                    </div>

                    <div class="metric-label">
                        Productos
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )


        with col4:

            st.markdown(
                f"""
                <div class="metric-card">

                    <div class="metric-value">
                        {n_haciendas}
                    </div>

                    <div class="metric-label">
                        Haciendas
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )


        with col5:

            st.markdown(
                f"""
                <div class="metric-card">

                    <div class="metric-value">
                        {n_errores}
                    </div>

                    <div class="metric-label">
                        Errores
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )


    # =======================================================
    # Validaciones
    # =======================================================

    if not df_consolidado.empty:

        st.markdown("---")


        st.markdown(
            '<p class="section-title">'
            '✅ Validación de datos'
            '</p>',
            unsafe_allow_html=True
        )


        with st.expander(
            "🔎 Ejecutar validaciones",
            expanded=False
        ):


            if st.button(
                "Validar datos",
                type="secondary"
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
    # Descargas
    # =======================================================

    if not df_consolidado.empty:

        st.markdown("---")


        st.markdown(
            '<p class="section-title">'
            '💾 Descargas'
            '</p>',
            unsafe_allow_html=True
        )


        col_d1, col_d2 = (
            st.columns(2)
        )


        # ---------------------------------------------------
        # Consolidado
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
                    "📥 Descargar consolidado (.xlsx)"
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
        # Errores
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
                        "📥 Descargar errores (.xlsx)"
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
    # Referencia cruzada con maestro
    # =======================================================

    if (
        not df_consolidado.empty
        and fertilizantes_df is not None
        and "PRODUCTO"
        in df_consolidado.columns
    ):

        st.markdown("---")


        st.markdown(
            '<p class="section-title">'
            '🔗 Referencia cruzada con maestro'
            '</p>',
            unsafe_allow_html=True
        )


        with st.expander(
            "📋 Ver productos del maestro "
            "de fertilizantes",
            expanded=False
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
                    & nombres_maestro
                )


                no_en_maestro = (
                    productos_datos
                    - nombres_maestro
                )


                col_m1, col_m2 = (
                    st.columns(2)
                )


                with col_m1:

                    st.markdown(
                        "**✅ Encontrados en maestro:**"
                    )


                    if en_maestro:

                        for p in sorted(
                            en_maestro
                        ):

                            st.markdown(
                                f"- {p}"
                            )

                    else:

                        st.caption(
                            "Ninguno"
                        )


                with col_m2:

                    st.markdown(
                        "**⚠️ NO encontrados en maestro:**"
                    )


                    if no_en_maestro:

                        for p in sorted(
                            no_en_maestro
                        ):

                            st.markdown(
                                f"- {p}"
                            )

                    else:

                        st.caption(
                            "Todos coinciden"
                        )


# ===========================================================
# ===========================================================
# PESTAÑA VISUALIZADOR
# ===========================================================
# ===========================================================

with tab_visualizador:

    st.markdown(
        '<p class="section-title">'
        '📊 Visualizador de fertilización'
        '</p>',
        unsafe_allow_html=True
    )


    # =======================================================
    # Recuperar consolidado
    # =======================================================

    df_consolidado = (
        st.session_state
        .get(
            "consolidado",
            pd.DataFrame()
        )
    )


    if df_consolidado.empty:

        st.markdown(
            """
            <div style="
                text-align: center;
                padding: 3rem 2rem;
                background:
                    linear-gradient(
                        135deg,
                        #f7fdf9,
                        #edf7f2
                    );

                border-radius: 16px;

                border:
                    2px dashed #2dba8e;
            ">

                <div style="
                    font-size: 4rem;
                    margin-bottom: 1rem;
                ">
                    🌿
                </div>

                <h3 style="
                    color: #0d4b3c;
                    margin-bottom: 0.5rem;
                ">
                    No hay datos procesados
                </h3>

                <p style="
                    color: #6b8f83;
                    font-size: 0.95rem;
                ">
                    Vaya a la pestaña
                    <strong>⚙️ Procesar</strong>
                    para cargar y procesar los archivos.
                </p>

            </div>
            """,
            unsafe_allow_html=True
        )


    else:

        # ===================================================
        # Métricas
        # ===================================================

        st.markdown(
            '<p class="section-title">'
            '📈 Resumen General'
            '</p>',
            unsafe_allow_html=True
        )


        col1, col2, col3, col4, col5 = (
            st.columns(5)
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


        with col1:

            st.markdown(
                f"""
                <div class="metric-card">

                    <div class="metric-value">
                        {n_archivos}
                    </div>

                    <div class="metric-label">
                        Archivos
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )


        with col2:

            st.markdown(
                f"""
                <div class="metric-card">

                    <div class="metric-value">
                        {n_registros}
                    </div>

                    <div class="metric-label">
                        Registros
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )


        with col3:

            st.markdown(
                f"""
                <div class="metric-card">

                    <div class="metric-value">
                        {n_productos}
                    </div>

                    <div class="metric-label">
                        Productos
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )


        with col4:

            st.markdown(
                f"""
                <div class="metric-card">

                    <div class="metric-value">
                        {n_haciendas}
                    </div>

                    <div class="metric-label">
                        Haciendas
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )


        with col5:

            st.markdown(
                f"""
                <div class="metric-card">

                    <div class="metric-value">
                        {n_errores}
                    </div>

                    <div class="metric-label">
                        Errores
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )


        # ===================================================
        # Filtros
        # ===================================================

        st.markdown("---")


        st.markdown(
            '<p class="section-title">'
            '🔍 Filtros'
            '</p>',
            unsafe_allow_html=True
        )


        col_f1, col_f2, col_f3 = (
            st.columns(3)
        )


        df_filtrado = (
            df_consolidado.copy()
        )


        # ---------------------------------------------------
        # Filtro archivo
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
                        ].isin(
                            sel_archivos
                        )
                    ]
                )


        # ---------------------------------------------------
        # Filtro producto
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
                        ].isin(
                            sel_productos
                        )
                    ]
                )


        # ---------------------------------------------------
        # Filtro hacienda
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
        # Tabla consolidada
        # ===================================================

        st.markdown(
            '<p class="section-title">'
            '📊 Datos consolidados'
            '</p>',
            unsafe_allow_html=True
        )


        columnas_mostrar = [

            c

            for c in (
                COLUMNAS_CLAVE
                + ["ARCHIVO_ORIGEN"]
            )

            if c in df_filtrado.columns

        ]


        st.dataframe(

            (
                df_filtrado[
                    columnas_mostrar
                ]

                if columnas_mostrar

                else df_filtrado
            ),

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
        # Gráficos
        # ===================================================

        st.markdown("---")


        st.markdown(
            '<p class="section-title">'
            '📊 Análisis Visual'
            '</p>',
            unsafe_allow_html=True
        )


        tab1, tab2, tab3 = st.tabs(

            [
                "🧪 Por Producto",
                "📦 Cantidad por Producto",
                "🌱 Dosis por Producto"
            ]

        )


        # ---------------------------------------------------
        # Gráfico productos
        # ---------------------------------------------------

        with tab1:

            if (
                "PRODUCTO"
                in df_filtrado.columns
            ):

                conteo = (

                    df_filtrado[
                        "PRODUCTO"
                    ]

                    .dropna()

                    .value_counts()

                    .head(15)

                )


                st.bar_chart(
                    conteo
                )

            else:

                st.info(
                    "No hay datos de producto disponibles"
                )


        # ---------------------------------------------------
        # Gráfico cantidad
        # ---------------------------------------------------

        with tab2:

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


                st.bar_chart(
                    cantidad_prod
                )

            else:

                st.info(
                    "No hay datos de cantidad disponibles"
                )


        # ---------------------------------------------------
        # Gráfico dosis
        # ---------------------------------------------------

        with tab3:

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


                st.bar_chart(
                    dosis_prod
                )

            else:

                st.info(
                    "No hay datos de dosis disponibles"
                )


# ===========================================================
# Estado vacío general
# ===========================================================

if (
    st.session_state[
        "consolidado"
    ].empty
    and
    not st.session_state[
        "detalles"
    ]
):

    st.markdown(
        "<br>",
        unsafe_allow_html=True
    )


    col_empty1, col_empty2, col_empty3 = (
        st.columns(
            [1, 2, 1]
        )
    )


    with col_empty2:

        st.markdown(
            """
            <div style="
                text-align: center;

                padding: 3rem 2rem;

                background:
                    linear-gradient(
                        135deg,
                        #f7fdf9,
                        #edf7f2
                    );

                border-radius: 16px;

                border:
                    2px dashed #2dba8e;
            ">

                <div style="
                    font-size: 4rem;
                    margin-bottom: 1rem;
                ">
                    🌿
                </div>

                <h3 style="
                    color: #0d4b3c;
                    margin-bottom: 0.5rem;
                ">
                    Módulo fertilización
                </h3>

                <p style="
                    color: #6b8f83;
                    font-size: 0.95rem;
                ">
                    Vaya a la pestaña
                    <strong>⚙️ Procesar</strong>
                    para cargar o procesar los archivos.
                </p>

            </div>
            """,
            unsafe_allow_html=True
        )


# ===========================================================
# Footer
# ===========================================================

st.markdown(
    """
    <div class="footer">

        Módulo fertilización
        · Auditoría Bonsucro · 2026

    </div>
    """,
    unsafe_allow_html=True
)
