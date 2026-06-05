import os
import re
import html
import textwrap
from datetime import datetime

import pandas as pd
import streamlit as st
import plotly.graph_objects as go

# CONFIGURACIÓN GENERAL
st.set_page_config(
    page_title="Q-Find",
    page_icon="",
    layout="wide",
    initial_sidebar_state="collapsed"
)

EXCEL_PATH = os.getenv(
    "EXCEL_PATH",
    r"C:\Users\prac.ccgestion\Prebel S.A BIC\Validaciones - Documentos\SEGUIMIENTO ACT CAL-VAL -2024.xlsx"
)

SHEET_NAME = "INDICADORES"

# RENDER HTML
def render_html(markup: str):
    clean_markup = textwrap.dedent(markup).strip()

    # Si es CSS, lo dejamos normal
    if clean_markup.startswith("<style"):
        st.markdown(clean_markup, unsafe_allow_html=True)
        return

    # Si es HTML, lo compactamos para que Streamlit no lo interprete como código
    clean_markup = "".join(
        line.strip()
        for line in clean_markup.splitlines()
        if line.strip()
    )

    st.markdown(clean_markup, unsafe_allow_html=True)

# ESTILOS CSS
render_html("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap');

    :root {
        --fs-hero: clamp(28px, 3vw, 42px);
        --fs-card-title: clamp(16px, 1.2vw, 20px);
        --fs-normal: clamp(11px, 0.9vw, 13px);
        --fs-small: clamp(9px, 0.78vw, 12px);
        --fs-tooltip: clamp(9px, 0.72vw, 11px);
        --card-padding-y: clamp(16px, 1.4vw, 22px);
        --card-padding-x: clamp(14px, 1.4vw, 22px);
    }

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    .stApp {
        background-color: #f7f9fc;
    }

    .block-container {
        padding-top: 1.2rem;
        padding-left: 1.6rem;
        padding-right: 1.6rem;
        max-width: 100%;
    }

    header[data-testid="stHeader"] {
        display: none;
    }

    div[data-testid="stToolbar"] {
        display: none;
    }

    .topbar {
        width: calc(100% + 3.2rem);
        background: #ffffff;
        border-bottom: 1px solid #dfe6f0;
        padding: 18px 26px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin: -1.2rem -1.6rem 2rem -1.6rem;
        min-width: 0;
    }

    .brand {
        display: flex;
        align-items: center;
        gap: 14px;
        min-width: 0;
    }

    .brand-icon {
        width: 44px;
        height: 44px;
        min-width: 44px;
        border-radius: 12px;
        background: #1769ff;
        display: flex;
        align-items: center;
        justify-content: center;
        color: white;
        font-weight: 900;
        font-size: 22px;
    }

    .brand-text {
        min-width: 0;
    }

    .brand-text h1 {
        margin: 0;
        padding: 0;
        font-size: 25px;
        line-height: 1;
        font-weight: 900;
        color: #101828;
        white-space: nowrap;
    }

    .brand-text span {
        color: #1769ff;
    }

    .brand-text p {
        margin: 5px 0 0 0;
        padding: 0;
        font-size: 11px;
        letter-spacing: 1.4px;
        color: #8a96ac;
        font-weight: 800;
    }

    .pro-badge {
        background: #edf2f7;
        color: #64748b !important;
        border-radius: 6px;
        padding: 3px 7px;
        font-size: 10px;
        margin-left: 6px;
        vertical-align: middle;
        font-weight: 800;
    }

    .hero-title {
        font-size: var(--fs-hero);
        line-height: 1.08;
        font-weight: 900;
        color: #0f172a;
        max-width: 670px;
        margin-top: 8px;
        margin-bottom: 14px;
        letter-spacing: -1.2px;
    }

    .hero-title span {
        color: #1769ff;
    }

    .hero-description {
        color: #667085;
        font-size: clamp(13px, 1vw, 15px);
        max-width: 620px;
        margin-bottom: 26px;
        line-height: 1.5;
    }

    div[data-testid="stTextInput"] {
        max-width: 760px;
        margin-bottom: 22px;
    }

    div[data-testid="stTextInput"] label {
        display: none !important;
    }

    div[data-testid="stTextInput"] > div {
        min-height: 72px !important;
    }

    div[data-testid="stTextInput"] > div > div {
        min-height: 72px !important;
        border-radius: 18px !important;
    }

    div[data-testid="stTextInput"] input {
        height: 72px !important;
        min-height: 72px !important;
        padding: 0 24px !important;
        border-radius: 18px !important;
        border: 2px solid #dce4ef !important;
        font-size: 17px !important;
        color: #1e293b !important;
        background: #ffffff !important;
        box-shadow: 0 2px 5px rgba(15, 23, 42, 0.06) !important;
    }

    div[data-testid="stTextInput"] input:focus {
        border-color: #1769ff !important;
        box-shadow: 0 0 0 2px rgba(23, 105, 255, 0.12) !important;
    }

    div[data-testid="stTextInput"] input::placeholder {
        color: #7c8ba3 !important;
        opacity: 1 !important;
    }

    div[data-testid="stSelectbox"] label {
        color: #667085 !important;
        font-size: 12px !important;
        font-weight: 900 !important;
        letter-spacing: 0.9px !important;
        text-transform: uppercase !important;
    }

    div[data-testid="stSelectbox"] > div > div {
        border-radius: 12px !important;
    }

    .filters-space {
        margin-bottom: 8px;
    }

    .result-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-top: 24px;
        margin-bottom: 22px;
        gap: 20px;
        min-width: 0;
    }

    .result-count {
        font-size: var(--fs-small);
        color: #98a2b3;
        letter-spacing: 1.1px;
        font-weight: 900;
        text-transform: uppercase;
    }

    .result-count strong {
        color: #0f172a;
        margin-left: 6px;
    }

    .legend {
        display: flex;
        gap: 18px;
        align-items: center;
        justify-content: flex-end;
        font-size: var(--fs-small);
        font-weight: 900;
        letter-spacing: 1.2px;
        text-transform: uppercase;
        white-space: nowrap;
        flex-wrap: wrap;
        min-width: 0;
    }

    .dot-green,
    .dot-red,
    .dot-orange {
        width: 8px;
        height: 8px;
        border-radius: 999px;
        display: inline-block;
        margin-right: 7px;
    }

    .dot-green {
        background: #10b981;
    }

    .dot-red {
        background: #f43f5e;
    }

    .dot-orange {
        background: #f97316;
    }

    .legend-green {
        color: #008f5d;
    }

    .legend-red {
        color: #e11d48;
    }

    .legend-orange {
        color: #f97316;
    }

    .cards-grid {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: clamp(14px, 1.5vw, 22px);
        width: 100%;
        align-items: stretch;
    }

    .card {
        background: #ffffff;
        border: 1px solid #dfe6f0;
        border-radius: 18px;
        padding: var(--card-padding-y) var(--card-padding-x) 18px var(--card-padding-x);
        min-height: 310px;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.03);
        overflow: visible;
        height: 100%;
        min-width: 0;
    }

    .badges {
        display: flex;
        gap: 8px;
        margin-bottom: 13px;
        flex-wrap: wrap;
        min-width: 0;
    }

    .badge {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        background: #edf3fb;
        color: #334155;
        border-radius: 5px;
        font-size: clamp(9px, 0.72vw, 11px);
        font-weight: 900;
        padding: 4px 8px;
        max-width: 100%;
        overflow-wrap: anywhere;
        line-height: 1.25;
    }

    .badge-prebel {
        background: #e7f0ff;
        color: #1769ff;
    }

    .card-title {
        font-size: var(--fs-card-title);
        color: #020617;
        font-weight: 900;
        line-height: 1.2;
        min-height: auto;
        margin-bottom: 18px;
        text-transform: uppercase;
        overflow-wrap: anywhere;
        word-break: normal;
    }

    .status-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        border-top: 1px solid #edf2f7;
        border-bottom: 1px solid #edf2f7;
        margin-bottom: 0;
        min-width: 0;
    }

    .status-box {
        padding: clamp(12px, 1.2vw, 18px) clamp(8px, 1vw, 16px);
        min-width: 0;
    }

    .status-box:first-child {
        border-right: 1px solid #edf2f7;
    }

    .status-title-blue,
    .status-title-green {
        font-size: clamp(9px, 0.72vw, 12px);
        font-weight: 900;
        letter-spacing: clamp(1px, 0.15vw, 2px);
        line-height: 1.35;
        text-transform: uppercase;
        margin-bottom: 20px;
        color: #0f172a;
        overflow-wrap: anywhere;
    }

    .metric-row {
        display: grid;
        grid-template-columns: minmax(0, 1fr) auto;
        align-items: center;
        margin-bottom: 16px;
        min-height: 28px;
        gap: 8px;
        min-width: 0;
    }

    .metric-label {
        font-size: clamp(9px, 0.72vw, 12px);
        color: #98a2b3;
        font-weight: 900;
        letter-spacing: 0.8px;
        text-transform: uppercase;
        min-width: 0;
        overflow-wrap: anywhere;
    }

    .tooltip-label {
        position: relative;
        display: inline-flex;
        align-items: center;
        gap: 5px;
        cursor: help;
        min-width: 0;
        max-width: 100%;
        overflow-wrap: anywhere;
    }

    .title-tooltip {
        align-items: flex-start;
    }

    .tooltip-label::after {
        content: attr(data-tooltip);
        position: absolute;
        left: 0;
        bottom: calc(100% + 10px);
        width: min(210px, calc(100vw - 48px));
        max-width: min(210px, calc(100vw - 48px));
        background: #0f172a;
        color: #ffffff;
        padding: 9px 11px;
        border-radius: 9px;
        font-size: var(--fs-tooltip);
        font-weight: 700;
        letter-spacing: 0;
        line-height: 1.3;
        text-transform: none;
        opacity: 0;
        visibility: hidden;
        transform: translateY(4px);
        transition: all 0.18s ease;
        z-index: 9999;
        box-shadow: 0 8px 20px rgba(15, 23, 42, 0.18);
        white-space: normal;
        overflow-wrap: anywhere;
        pointer-events: none;
    }

    .tooltip-label::before {
        content: "";
        position: absolute;
        left: 14px;
        bottom: calc(100% + 4px);
        border-width: 6px 6px 0 6px;
        border-style: solid;
        border-color: #0f172a transparent transparent transparent;
        opacity: 0;
        visibility: hidden;
        transition: all 0.18s ease;
        z-index: 10000;
        pointer-events: none;
    }

    .tooltip-label:hover::after,
    .tooltip-label:hover::before {
        opacity: 1;
        visibility: visible;
        transform: translateY(0);
    }

    .status-box:nth-child(2) .tooltip-label::after {
        left: auto;
        right: 0;
    }

    .status-box:nth-child(2) .tooltip-label::before {
        left: auto;
        right: 14px;
    }

    .tooltip-icon {
        width: 15px;
        height: 15px;
        min-width: 15px;
        border-radius: 999px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        background: #eef4ff;
        color: #1769ff;
        font-size: 10px;
        font-weight: 900;
        letter-spacing: 0;
        text-transform: none;
    }

    .pill {
        min-width: 54px;
        max-width: 100%;
        text-align: center;
        padding: 6px 7px;
        border-radius: 9px;
        font-size: clamp(10px, 0.78vw, 13px);
        font-weight: 900;
        border: 1px solid #dce4ef;
        color: #16325c;
        background: #f8fafc;
        white-space: normal;
        line-height: 1.15;
        overflow-wrap: anywhere;
    }

    .pill-blue {
        background: #f1f5f9;
        color: #16325c;
    }

    .pill-green {
        background: #ecfdf5;
        color: #008f5d;
        border-color: #bbf7d0;
    }

    .pill-red {
        background: #fff1f2;
        color: #e11d48;
        border-color: #fecdd3;
    }

    .pill-orange {
        background: #fff7ed;
        color: #f97316;
        border-color: #fed7aa;
    }

    .card-footer {
        border-top: 1px solid #edf2f7;
        padding-top: 14px;
        margin-top: 14px;
        display: flex;
        flex-direction: column;
        gap: 6px;
    }

    .footer-label {
        color: #98a2b3;
        font-size: clamp(9px, 0.72vw, 11px);
        font-weight: 900;
        letter-spacing: 0.9px;
        text-transform: uppercase;
    }

    .footer-value {
        color: #334155;
        font-size: clamp(11px, 0.9vw, 13px);
        font-weight: 800;
        line-height: 1.35;
        overflow-wrap: anywhere;
    }

    .empty-state {
        background: white;
        border: 1px dashed #cbd5e1;
        border-radius: 18px;
        padding: 36px;
        color: #64748b;
        text-align: center;
        font-weight: 700;
    }

    .side-panel {
        padding-top: 10px;
    }

    .side-kpi-label {
        color: #98a2b3;
        font-size: clamp(9px, 0.75vw, 11px);
        font-weight: 900;
        letter-spacing: 1px;
        text-transform: uppercase;
        margin-bottom: 8px;
    }

    .side-kpi-value {
        color: #0f172a;
        font-size: clamp(26px, 2.8vw, 34px);
        font-weight: 900;
        line-height: 1;
        margin-bottom: 8px;
    }

    .side-kpi-subtitle {
        color: #667085;
        font-size: clamp(10px, 0.82vw, 12px);
        font-weight: 700;
        line-height: 1.4;
    }

    .side-card-space {
        margin-bottom: 14px;
    }

    @media (max-width: 1200px) {
        .cards-grid {
            grid-template-columns: repeat(2, minmax(0, 1fr));
        }

        .status-grid {
            grid-template-columns: 1fr;
        }

        .status-box:first-child {
            border-right: none;
            border-bottom: 1px solid #edf2f7;
        }

        .status-box:nth-child(2) .tooltip-label::after {
            left: 0;
            right: auto;
        }

        .status-box:nth-child(2) .tooltip-label::before {
            left: 14px;
            right: auto;
        }
    }

    @media (max-width: 900px) {
        .hero-title {
            font-size: 32px;
        }

        .legend {
            flex-direction: column;
            align-items: flex-start;
            gap: 8px;
            white-space: normal;
        }

        .result-row {
            flex-direction: column;
            align-items: flex-start;
            gap: 14px;
        }
    }

    @media (max-width: 760px) {
        .block-container {
            padding-top: 1rem;
            padding-left: 1rem;
            padding-right: 1rem;
        }

        .topbar {
            width: calc(100% + 2rem);
            margin: -1rem -1rem 1.5rem -1rem;
            padding: 16px;
        }

        .brand {
            gap: 10px;
        }

        .brand-icon {
            width: 38px;
            height: 38px;
            min-width: 38px;
            font-size: 18px;
        }

        .brand-text h1 {
            font-size: 21px;
            white-space: normal;
        }

        .brand-text p {
            font-size: 9px;
            letter-spacing: 1px;
        }

        .pro-badge {
            font-size: 9px;
            padding: 2px 6px;
        }

        .hero-title {
            font-size: 30px;
            line-height: 1.12;
            letter-spacing: -0.8px;
        }

        .hero-description {
            font-size: 14px;
            margin-bottom: 22px;
        }

        div[data-testid="stTextInput"] {
            max-width: 100%;
            margin-bottom: 18px;
        }

        div[data-testid="stTextInput"] input {
            height: 60px !important;
            min-height: 60px !important;
            font-size: 14px !important;
            padding: 0 16px !important;
        }

        div[data-testid="stTextInput"] > div {
            min-height: 60px !important;
        }

        div[data-testid="stTextInput"] > div > div {
            min-height: 60px !important;
        }

        .cards-grid {
            grid-template-columns: 1fr;
            gap: 16px;
        }

        .card {
            padding: 18px 16px;
            border-radius: 16px;
            min-height: auto;
        }

        .card-title {
            font-size: 17px;
            min-height: auto;
            margin-bottom: 14px;
        }

        .status-box {
            padding: 16px 0;
        }

        .metric-row {
            gap: 12px;
        }

        .metric-label {
            font-size: 11px;
        }

        .pill {
            min-width: 72px;
            font-size: 12px;
            padding: 7px 8px;
        }

        .tooltip-label::after {
            width: min(190px, calc(100vw - 40px));
            max-width: min(190px, calc(100vw - 40px));
            left: 0;
        }
    }

    @media (max-width: 480px) {
        .hero-title {
            font-size: 26px;
        }

        .hero-title br {
            display: none;
        }

        .result-count {
            font-size: 11px;
        }

        .legend {
            font-size: 11px;
            gap: 10px;
        }

        .badges {
            gap: 6px;
        }

        .badge {
            font-size: 10px;
            padding: 4px 7px;
        }

        .status-title-blue,
        .status-title-green {
            font-size: 11px;
            letter-spacing: 1.5px;
        }

        .metric-row {
            align-items: flex-start;
        }

        .tooltip-icon {
            width: 14px;
            height: 14px;
            min-width: 14px;
            font-size: 9px;
        }

        .tooltip-label::after {
            width: min(175px, calc(100vw - 32px));
            max-width: min(175px, calc(100vw - 32px));
            font-size: 10px;
        }
    }

    @media (max-width: 380px) {
        .metric-row {
            grid-template-columns: 1fr;
            align-items: flex-start;
        }

        .pill {
            justify-self: flex-start;
        }
    }
</style>
""")

# FUNCIONES DE LIMPIEZA Y LECTURA

def normalize_text(value):
    if pd.isna(value):
        return ""

    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)

    replacements = {
        "á": "a",
        "é": "e",
        "í": "i",
        "ó": "o",
        "ú": "u",
        "ñ": "n",
    }

    for original, replacement in replacements.items():
        text = text.replace(original, replacement)

    return text

def normalize_column_name(value):
    text = normalize_text(value)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text

def get_column(df, column_name):
    target = normalize_column_name(column_name)

    for col in df.columns:
        if normalize_column_name(col) == target:
            return col

    return None

def get_required_column(df, column_name):
    col = get_column(df, column_name)

    if col is None:
        st.error(f"No se encontró la columna requerida: {column_name}")
        return None

    return col

@st.cache_data(ttl=60)
def load_data():
    if not os.path.exists(EXCEL_PATH):
        st.error(f"No se encontró el archivo Excel en la ruta: {EXCEL_PATH}")
        return pd.DataFrame()

    try:
        df = pd.read_excel(
            EXCEL_PATH,
            sheet_name=SHEET_NAME,
            engine="openpyxl"
        )

        df = df.dropna(how="all")
        df.columns = [str(col).strip() for col in df.columns]

        return df

    except Exception as error:
        st.error(f"Error leyendo el Excel: {error}")
        return pd.DataFrame()

def prepare_data(df):
    if df.empty:
        return df

    col_equipo = get_required_column(df, "Equipo")
    col_clasificacion = get_required_column(df, "Campo de clasificación")
    col_denominacion = get_required_column(df, "Denominación de objeto técnico")
    col_ubicacion = get_required_column(df, "Denominacion de la ubicacion tecnica")

    col_estado_calificacion = get_required_column(df, "ESTADO CALIFICACION (VIGENTE/VENCIDO)")
    col_dias_recal = get_required_column(df, "DIAS CONTROL  RECAL")

    col_estado_limpieza = get_required_column(df, "ESTADO VALIDACION")
    col_dias_reval = get_required_column(df, "DIAS CONTROL  REVAL")

    col_nombre_proceso_validacion = get_required_column(
        df,
        "NOMBRE DE PROCESO DE VALIDACION (INGRESO MANUAL)"
    )

    required_columns = [
        col_equipo,
        col_clasificacion,
        col_denominacion,
        col_ubicacion,
        col_estado_calificacion,
        col_dias_recal,
        col_estado_limpieza,
        col_dias_reval,
        col_nombre_proceso_validacion,
    ]

    if any(col is None for col in required_columns):
        st.info("Revisa que los nombres de las columnas en el Excel estén escritos igual o muy parecido a los indicados.")
        return pd.DataFrame()

    result = pd.DataFrame()

    result["equipo"] = df[col_equipo]
    result["campo_clasificacion"] = df[col_clasificacion]
    result["denominacion"] = df[col_denominacion]
    result["ubicacion_tecnica"] = df[col_ubicacion]

    result["estado_calificacion"] = df[col_estado_calificacion]
    result["dias_calificacion"] = df[col_dias_recal]

    result["estado_limpieza"] = df[col_estado_limpieza]
    result["dias_limpieza"] = df[col_dias_reval]

    result["nombre_proceso_validacion"] = df[col_nombre_proceso_validacion]

    for col in result.columns:
        result[col] = result[col].fillna("")

    result["search_text"] = (
        result["equipo"].astype(str) + " " +
        result["campo_clasificacion"].astype(str) + " " +
        result["denominacion"].astype(str) + " " +
        result["ubicacion_tecnica"].astype(str)
    ).apply(normalize_text)

    return result

def format_value(value, default="N/A"):
    if pd.isna(value):
        return default

    value_str = str(value).strip()

    if value_str == "":
        return default

    if value_str.lower() in ["nan", "none", "nat"]:
        return default

    if value_str.endswith(".0"):
        value_str = value_str[:-2]

    return value_str

def safe_text(value):
    return html.escape(str(value))

DIAS_PROXIMO_A_VENCER = 30

def to_number(value):
    try:
        value_str = str(value).strip().replace(",", ".")
        return float(value_str)
    except Exception:
        return None

def get_semaforo_class(estado, dias=None):
    estado_normalizado = normalize_text(estado)
    dias_numero = to_number(dias)

    if (
        "vencido" in estado_normalizado
        or "pendiente" in estado_normalizado
        or "no vigente" in estado_normalizado
        or "rechazado" in estado_normalizado
        or "no validado" in estado_normalizado
        or "no calificado" in estado_normalizado
        or (dias_numero is not None and dias_numero < 0)
    ):
        return "pill-red"

    if (
        "proximo" in estado_normalizado
        or "próximo" in estado_normalizado
        or (
            dias_numero is not None
            and 0 <= dias_numero <= DIAS_PROXIMO_A_VENCER
        )
    ):
        return "pill-orange"

    if (
        "vigente" in estado_normalizado
        or "validado" in estado_normalizado
        or "calificado" in estado_normalizado
        or "aprobado" in estado_normalizado
        or "ok" in estado_normalizado
        or "cumple" in estado_normalizado
    ):
        return "pill-green"

    return "pill-blue"

def get_dias_label(estado, dias):
    estado_normalizado = normalize_text(estado)
    dias_numero = to_number(dias)

    if "vencido" in estado_normalizado or (dias_numero is not None and dias_numero < 0):
        return "Días<br>vencido"

    return "Días de<br>vigencia"

def es_no_calificado(estado):
    return "no calificado" in normalize_text(estado)

def es_calificado(estado):
    estado_normalizado = normalize_text(estado)

    if estado_normalizado in ["", "n/a", "na"]:
        return False

    return not es_no_calificado(estado)

def es_no_validado(estado):
    return "no validado" in normalize_text(estado)

def es_validado(estado):
    estado_normalizado = normalize_text(estado)

    if estado_normalizado in ["", "n/a", "na"]:
        return False

    return not es_no_validado(estado)

def es_calificado_vigente(estado, dias):
    estado_normalizado = normalize_text(estado)
    dias_numero = to_number(dias)

    if estado_normalizado in ["", "n/a", "na"]:
        return False

    if (
        "no calificado" in estado_normalizado
        or "no vigente" in estado_normalizado
        or "vencido" in estado_normalizado
        or "pendiente" in estado_normalizado
        or "rechazado" in estado_normalizado
        or (dias_numero is not None and dias_numero < 0)
    ):
        return False

    return (
        "vigente" in estado_normalizado
        or "calificado" in estado_normalizado
        or "aprobado" in estado_normalizado
        or "ok" in estado_normalizado
        or "cumple" in estado_normalizado
        or (dias_numero is not None and dias_numero >= 0)
    )

def es_validacion_vigente(estado, dias):
    estado_normalizado = normalize_text(estado)
    dias_numero = to_number(dias)

    if estado_normalizado in ["", "n/a", "na"]:
        return False

    if (
        "no validado" in estado_normalizado
        or "no vigente" in estado_normalizado
        or "vencido" in estado_normalizado
        or "pendiente" in estado_normalizado
        or "rechazado" in estado_normalizado
        or (dias_numero is not None and dias_numero < 0)
    ):
        return False

    return (
        "vigente" in estado_normalizado
        or "validado" in estado_normalizado
        or "aprobado" in estado_normalizado
        or "ok" in estado_normalizado
        or "cumple" in estado_normalizado
        or (dias_numero is not None and dias_numero >= 0)
    )

def has_content(value):
    if pd.isna(value):
        return False

    text = str(value).strip()

    if text == "":
        return False

    if normalize_text(text) in ["nan", "none", "nat", "n/a", "na"]:
        return False

    return True

def calculate_percentage(value, total):
    if total <= 0:
        return 0.0
    return round((value / total) * 100, 1)

def build_gauge_chart(percent_value):
    percent_value = max(0, min(100, float(percent_value)))

    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=percent_value,
            number={
                "suffix": "%",
                "font": {"size": 24, "color": "#0f172a"}
            },
            gauge={
                "axis": {
                    "range": [0, 100],
                    "showticklabels": False,
                    "ticks": ""
                },
                "bar": {
                    "color": "#1769ff",
                    "thickness": 0.28
                },
                "bgcolor": "white",
                "borderwidth": 0,
                "steps": [
                    {"range": [0, 50], "color": "#fee2e2"},
                    {"range": [50, 80], "color": "#fef3c7"},
                    {"range": [80, 100], "color": "#dcfce7"},
                ],
            }
        )
    )

    fig.update_layout(
        height=160,
        margin=dict(l=10, r=10, t=10, b=10),
        paper_bgcolor="white"
    )

    return fig

def get_dashboard_kpis(df):
    if df.empty:
        return {
            "total_equipos": 0,
            "equipos_calificados": 0,
            "pct_equipos": 0.0,
            "total_validaciones": 0,
            "validaciones_validadas": 0,
            "pct_validaciones": 0.0,
        }

    # KPI 1: cantidad de equipos
    total_equipos = int(df["equipo"].apply(has_content).sum())

    equipos_calificados = int(
        df["estado_calificacion"].apply(es_calificado).sum()
    )

    pct_equipos = calculate_percentage(
        equipos_calificados,
        total_equipos
    )

    # KPI 2: cantidad de validaciones únicas
    procesos_df = df[df["nombre_proceso_validacion"].apply(has_content)].copy()

    if not procesos_df.empty:
        procesos_df["proceso_key"] = (
            procesos_df["nombre_proceso_validacion"]
            .astype(str)
            .str.strip()
            .apply(normalize_text)
        )

        procesos_unicos = procesos_df.drop_duplicates(
            subset=["proceso_key"]
        ).copy()
    else:
        procesos_unicos = pd.DataFrame()

    total_validaciones = int(len(procesos_unicos))

    if not procesos_unicos.empty:
        validaciones_validadas = int(
            procesos_unicos["estado_limpieza"].apply(es_validado).sum()
        )
    else:
        validaciones_validadas = 0

    pct_validaciones = calculate_percentage(
        validaciones_validadas,
        total_validaciones
    )

    return {
        "total_equipos": total_equipos,
        "equipos_calificados": equipos_calificados,
        "pct_equipos": pct_equipos,
        "total_validaciones": total_validaciones,
        "validaciones_validadas": validaciones_validadas,
        "pct_validaciones": pct_validaciones,
    }

def render_card(row):
    equipo = safe_text(format_value(row.get("equipo", ""), ""))
    campo_clasificacion = safe_text(format_value(row.get("campo_clasificacion", ""), ""))
    denominacion = safe_text(format_value(row.get("denominacion", ""), "SIN NOMBRE").upper())
    ubicacion_tecnica = safe_text(format_value(row.get("ubicacion_tecnica", ""), "N/A"))

    estado_calificacion_raw = format_value(row.get("estado_calificacion", ""), "N/A")
    dias_calificacion_raw = format_value(row.get("dias_calificacion", ""), "0")

    estado_limpieza_raw = format_value(row.get("estado_limpieza", ""), "N/A")
    dias_limpieza_raw = format_value(row.get("dias_limpieza", ""), "0")

    if "no calificado" in normalize_text(estado_calificacion_raw):
        dias_calificacion_raw = "N/A"

    if "no validado" in normalize_text(estado_limpieza_raw):
        dias_limpieza_raw = "N/A"

    estado_calificacion = safe_text(estado_calificacion_raw)
    dias_calificacion = safe_text(dias_calificacion_raw)

    estado_limpieza = safe_text(estado_limpieza_raw)
    dias_limpieza = safe_text(dias_limpieza_raw)

    texto_dias_calificacion = dias_calificacion if dias_calificacion == "N/A" else f"{dias_calificacion}<br>días"
    texto_dias_limpieza = dias_limpieza if dias_limpieza == "N/A" else f"{dias_limpieza}<br>días"

    equipo_badge = f'<span class="badge">Código SAP: {equipo}</span>' if equipo else ""
    clasificacion_badge = f'<span class="badge badge-prebel">Código Prebel: {campo_clasificacion}</span>' if campo_clasificacion else ""

    estado_calificacion_class = get_semaforo_class(
        estado_calificacion_raw,
        dias_calificacion_raw
    )

    estado_limpieza_class = get_semaforo_class(
        estado_limpieza_raw,
        dias_limpieza_raw
    )

    label_dias_calificacion = get_dias_label(
        estado_calificacion_raw,
        dias_calificacion_raw
    )

    label_dias_limpieza = get_dias_label(
        estado_limpieza_raw,
        dias_limpieza_raw
    )

    card_html = f"""
    <div class="card">
        <div class="badges">
            {equipo_badge}
            {clasificacion_badge}
        </div>

        <div class="card-title">{denominacion}</div>

        <div class="status-grid">
            <div class="status-box">
                <div class="status-title-blue">
                    <span class="tooltip-label title-tooltip" data-tooltip="Las calificaciones cuentan con 4 años de vigencia">
                        Calificación de<br>equipo <span class="tooltip-icon">?</span>
                    </span>
                </div>

                <div class="metric-row">
                    <div class="metric-label tooltip-label" data-tooltip="Indica la condición actual del equipo con respecto a su estado calificado">
                        Estado <span class="tooltip-icon">?</span>
                    </div>
                    <div class="pill {estado_calificacion_class}">{estado_calificacion}</div>
                </div>

                <div class="metric-row">
                    <div class="metric-label tooltip-label" data-tooltip="Días restantes para su vencimiento">
                        {label_dias_calificacion} <span class="tooltip-icon">?</span>
                    </div>
                    <div class="pill {estado_calificacion_class}">{texto_dias_calificacion}</div>
                </div>
            </div>

            <div class="status-box">
                <div class="status-title-green">
                    <span class="tooltip-label title-tooltip" data-tooltip="Las validaciones de limpieza y sanitización cuentan con 6 años de vigencia">
                        Validación de limpieza y<br>sanitización <span class="tooltip-icon">?</span>
                    </span>
                </div>

                <div class="metric-row">
                    <div class="metric-label tooltip-label" data-tooltip="Indica la condición actual del proceso con respecto a su estado validado">
                        Estado <span class="tooltip-icon">?</span>
                    </div>
                    <div class="pill {estado_limpieza_class}">{estado_limpieza}</div>
                </div>

                <div class="metric-row">
                    <div class="metric-label tooltip-label" data-tooltip="Días restantes para su vencimiento">
                        {label_dias_limpieza} <span class="tooltip-icon">?</span>
                    </div>
                    <div class="pill {estado_limpieza_class}">{texto_dias_limpieza}</div>
                </div>
            </div>
        </div>

        <div class="card-footer">
            <div class="footer-label">Ubicación técnica</div>
            <div class="footer-value">{ubicacion_tecnica}</div>
        </div>
    </div>
    """

    return card_html

# INTERFAZ
render_html("""
<div class="topbar">
    <div class="brand">
        <div class="brand-icon">▣</div>
        <div class="brand-text">
            <h1><span> Q-Find </span> <span class="pro-badge">PRO</span></h1>
            <p>GESTIÓN DE ACTIVOS PREBEL</p>
        </div>
    </div>
</div>
""")

df_raw = load_data()
df = prepare_data(df_raw)

dashboard_kpis = get_dashboard_kpis(df)

hero_col, side_col = st.columns([1.7, 1.1], gap="large")

with hero_col:
    render_html("""
    <div class="hero-title">
        Consulta el estado de tus<br>
        equipos en <span>tiempo real.</span>
    </div>

    <div class="hero-description">
        Busca por código SAP, código Prebel o nombre del equipo para verificar el estado calificado del equipo y estado validado del proceso de limpieza y sanitización.
    </div>
    """)

    search = st.text_input(
        label="Buscar",
        placeholder="Ingresa código SAP, código Prebel o nombre del equipo..."
    )

with side_col:
    render_html('<div class="side-panel"></div>')

    with st.container(border=True):
        kpi_1_col, gauge_1_col = st.columns([1.1, 1.3])

        with kpi_1_col:
            render_html(f"""
            <div class="side-kpi-label">Cantidad de equipos</div>
            <div class="side-kpi-value">{dashboard_kpis["total_equipos"]}</div>
            <div class="side-kpi-subtitle">
                {dashboard_kpis["equipos_calificados"]} calificados
            </div>
            """)

        with gauge_1_col:
            st.plotly_chart(
                build_gauge_chart(dashboard_kpis["pct_equipos"]),
                use_container_width=True,
                config={"displayModeBar": False}
            )

    render_html('<div class="side-card-space"></div>')

    with st.container(border=True):
        kpi_2_col, gauge_2_col = st.columns([1.1, 1.3])

        with kpi_2_col:
            render_html(f"""
            <div class="side-kpi-label">Cantidad de validaciones L&S</div>
            <div class="side-kpi-value">{dashboard_kpis["total_validaciones"]}</div>
            <div class="side-kpi-subtitle">
                {dashboard_kpis["validaciones_validadas"]} validadas
            </div>
            """)

        with gauge_2_col:
            st.plotly_chart(
                build_gauge_chart(dashboard_kpis["pct_validaciones"]),
                use_container_width=True,
                config={"displayModeBar": False}
            )

filter_col_1, filter_col_2 = st.columns(2)

with filter_col_1:
    filtro_calificaciones = st.selectbox(
        "Filtro de calificaciones",
        ["Todos", "Calificados", "Calificados vigentes", "No calificados"]
    )

with filter_col_2:
    filtro_validaciones = st.selectbox(
        "Filtro de validaciones",
        ["Todos", "Validados", "Validaciones vigentes", "No validados"]
    )

render_html('<div class="filters-space"></div>')

filtered_df = df.copy()

if not filtered_df.empty and search.strip():
    query = normalize_text(search)
    filtered_df = filtered_df[filtered_df["search_text"].str.contains(query, na=False, regex=False)]

if not filtered_df.empty:
    if filtro_calificaciones == "Calificados":
        filtered_df = filtered_df[
            filtered_df["estado_calificacion"].apply(es_calificado)
        ]
    elif filtro_calificaciones == "Calificados vigentes":
        filtered_df = filtered_df[
            filtered_df.apply(
                lambda row: es_calificado_vigente(
                    row["estado_calificacion"],
                    row["dias_calificacion"]
                ),
                axis=1
            )
        ]
    elif filtro_calificaciones == "No calificados":
        filtered_df = filtered_df[
            filtered_df["estado_calificacion"].apply(es_no_calificado)
        ]

if not filtered_df.empty:
    if filtro_validaciones == "Validados":
        filtered_df = filtered_df[
            filtered_df["estado_limpieza"].apply(es_validado)
        ]
    elif filtro_validaciones == "Validaciones vigentes":
        filtered_df = filtered_df[
            filtered_df.apply(
                lambda row: es_validacion_vigente(
                    row["estado_limpieza"],
                    row["dias_limpieza"]
                ),
                axis=1
            )
        ]
    elif filtro_validaciones == "No validados":
        filtered_df = filtered_df[
            filtered_df["estado_limpieza"].apply(es_no_validado)
        ]

render_html(f"""
<div class="result-row">
    <div class="result-count">
        ⚙ RESULTADOS ENCONTRADOS:
        <strong>{len(filtered_df)}</strong>
    </div>

    <div class="legend">
        <div class="legend-green">
            <span class="dot-green"></span> Calificado / Validado
        </div>
        <div class="legend-orange">
            <span class="dot-orange"></span> Vigente / Próximo a vencer
        </div>
        <div class="legend-red">
            <span class="dot-red"></span> Vencido
        </div>
    </div>
</div>
""")

if df.empty:
    render_html("""
<div class="empty-state">
    No hay información disponible. Verifica la ruta del Excel, el nombre de la hoja y los nombres de las columnas.
</div>
""")

elif filtered_df.empty:
    render_html("""
<div class="empty-state">
    No se encontraron equipos con ese criterio de búsqueda o filtros aplicados.
</div>
""")

else:
    cards_html = '<div class="cards-grid">'

    for _, row in filtered_df.iterrows():
        cards_html += render_card(row)

    cards_html += '</div>'

    render_html(cards_html)