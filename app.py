import os
import re
import html
import textwrap
from datetime import datetime

import pandas as pd
import streamlit as st


# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================

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


# =========================================================
# RENDER HTML
# =========================================================

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


# =========================================================
# ESTILOS CSS
# =========================================================

render_html("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap');

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
        font-size: 42px;
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
        font-size: 15px;
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

    .result-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-top: 24px;
        margin-bottom: 22px;
        gap: 20px;
    }

    .result-count {
        font-size: 12px;
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
        font-size: 12px;
        font-weight: 900;
        letter-spacing: 1.2px;
        text-transform: uppercase;
        white-space: nowrap;
    }

    .dot-green,
    .dot-red {
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
        width: 8px;
        height: 8px;
        border-radius: 999px;
        display: inline-block;
        margin-right: 7px;
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
        gap: 22px;
        width: 100%;
        align-items: stretch;
    }

    .card {
        background: #ffffff;
        border: 1px solid #dfe6f0;
        border-radius: 18px;
        padding: 22px 22px 18px 22px;
        min-height: 310px;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.03);
        overflow: visible;
        height: 100%;
    }

    .badges {
        display: flex;
        gap: 8px;
        margin-bottom: 13px;
        flex-wrap: wrap;
    }

    .badge {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        background: #edf3fb;
        color: #334155;
        border-radius: 5px;
        font-size: 11px;
        font-weight: 900;
        padding: 4px 8px;
        max-width: 100%;
        overflow-wrap: anywhere;
    }

    .badge-prebel {
        background: #e7f0ff;
        color: #1769ff;
    }

    .card-title {
        font-size: 20px;
        color: #020617;
        font-weight: 900;
        line-height: 1.2;
        min-height: 48px;
        margin-bottom: 18px;
        text-transform: uppercase;
        overflow-wrap: anywhere;
    }

    .status-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        border-top: 1px solid #edf2f7;
        border-bottom: 1px solid #edf2f7;
        margin-bottom: 16px;
    }

    .status-box {
        padding: 18px 16px;
        min-width: 0;
    }

    .status-box:first-child {
        border-right: 1px solid #edf2f7;
    }

    .status-title-blue,
    .status-title-green {
        font-size: 12px;
        font-weight: 900;
        letter-spacing: 2px;
        line-height: 1.35;
        text-transform: uppercase;
        margin-bottom: 20px;
        color: #0f172a;
        overflow-wrap: anywhere;
    }

    .metric-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 16px;
        min-height: 28px;
        gap: 10px;
    }

    .metric-label {
        font-size: 12px;
        color: #98a2b3;
        font-weight: 900;
        letter-spacing: 0.8px;
        text-transform: uppercase;
        min-width: 0;
    }

    .tooltip-label {
        position: relative;
        display: inline-flex;
        align-items: center;
        gap: 5px;
        cursor: help;
    }

    .tooltip-label::after {
        content: attr(data-tooltip);
        position: absolute;
        left: 0;
        bottom: calc(100% + 10px);
        width: 230px;
        background: #0f172a;
        color: #ffffff;
        padding: 9px 11px;
        border-radius: 9px;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0;
        line-height: 1.35;
        text-transform: none;
        opacity: 0;
        visibility: hidden;
        transform: translateY(4px);
        transition: all 0.18s ease;
        z-index: 9999;
        box-shadow: 0 8px 20px rgba(15, 23, 42, 0.18);
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
    }

    .tooltip-label:hover::after,
    .tooltip-label:hover::before {
        opacity: 1;
        visibility: visible;
        transform: translateY(0);
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
    }

    .pill {
        min-width: 66px;
        text-align: center;
        padding: 7px 9px;
        border-radius: 9px;
        font-size: 13px;
        font-weight: 900;
        border: 1px solid #dce4ef;
        color: #16325c;
        background: #f8fafc;
        white-space: nowrap;
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
        display: flex;
        justify-content: space-between;
        align-items: center;
        color: #7185aa;
        font-size: 12px;
        padding-top: 4px;
        gap: 12px;
    }

    .history-link {
        color: #1769ff;
        font-weight: 900;
        letter-spacing: 0.4px;
        text-transform: uppercase;
        text-decoration: none;
        white-space: nowrap;
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

        .card-footer {
            flex-direction: column;
            align-items: flex-start;
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
            width: 190px;
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
            width: 175px;
            font-size: 10px;
        }
    }
</style>
""")


# =========================================================
# FUNCIONES DE LIMPIEZA Y LECTURA
# =========================================================

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

    col_estado_calificacion = get_required_column(df, "ESTADO CALIFICACION (VIGENTE/VENCIDO)")
    col_dias_recal = get_required_column(df, "DIAS CONTROL  RECAL")

    col_estado_limpieza = get_required_column(df, "ESTADO VALIDACION")
    col_dias_reval = get_required_column(df, "DIAS CONTROL  REVAL")

    required_columns = [
        col_equipo,
        col_clasificacion,
        col_denominacion,
        col_estado_calificacion,
        col_dias_recal,
        col_estado_limpieza,
        col_dias_reval,
    ]

    if any(col is None for col in required_columns):
        st.info("Revisa que los nombres de las columnas en el Excel estén escritos igual o muy parecido a los indicados.")
        return pd.DataFrame()

    result = pd.DataFrame()

    result["equipo"] = df[col_equipo]
    result["campo_clasificacion"] = df[col_clasificacion]
    result["denominacion"] = df[col_denominacion]

    result["estado_calificacion"] = df[col_estado_calificacion]
    result["dias_calificacion"] = df[col_dias_recal]

    result["estado_limpieza"] = df[col_estado_limpieza]
    result["dias_limpieza"] = df[col_dias_reval]

    for col in result.columns:
        result[col] = result[col].fillna("")

    result["search_text"] = (
        result["equipo"].astype(str) + " " +
        result["campo_clasificacion"].astype(str) + " " +
        result["denominacion"].astype(str)
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


def render_card(row):
    equipo = safe_text(format_value(row.get("equipo", ""), ""))
    campo_clasificacion = safe_text(format_value(row.get("campo_clasificacion", ""), ""))
    denominacion = safe_text(format_value(row.get("denominacion", ""), "SIN NOMBRE").upper())

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
                    Calificación de<br>equipo
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
                    Validación de limpieza y<br>sanitización
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
    </div>
    """

    return card_html


# =========================================================
# INTERFAZ
# =========================================================

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

render_html("""
<div class="hero-title">
    Consulta el estado de tus<br>
    equipos en <span>tiempo real.</span>
</div>

<div class="hero-description">
    Busca por código Prebel, código SAP o nombre del equipo para verificar el estado calificado del equipo y estado validado del proceso de limpieza y sanitización.
</div>
""")

df_raw = load_data()
df = prepare_data(df_raw)

search = st.text_input(
    label="Buscar",
    placeholder="Ingresa código Prebel, código SAP o nombre del equipo..."
)

if not df.empty and search.strip():
    query = normalize_text(search)
    filtered_df = df[df["search_text"].str.contains(query, na=False, regex=False)]
else:
    filtered_df = df.copy()

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
            <span class="dot-orange"></span> Vigente (Próximo a vencer)
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
    No se encontraron equipos con ese criterio de búsqueda.
</div>
""")

else:
    cards_html = '<div class="cards-grid">'

    for _, row in filtered_df.iterrows():
        cards_html += render_card(row)

    cards_html += '</div>'

    render_html(cards_html)