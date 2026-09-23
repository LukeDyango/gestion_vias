import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import sqlite3
import os
import io
import json
import base64
import uuid
import zipfile
import time
import numbers
import calendar
from datetime import datetime, date, timedelta
import openpyxl
import openpyxl.styles
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage, ImageOps
import gspread
from google.oauth2.service_account import Credentials as GoogleCredentials
from streamlit_geolocation import streamlit_geolocation
import altair as alt
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage

st.set_page_config(page_title="efe · Trenes & CHILE", page_icon="🚆", layout="centered", initial_sidebar_state="collapsed")

MOBILE_CSS = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header[data-testid="stHeader"] {background: transparent; height: 0.5rem;}

.block-container {padding-top: 0.5rem; padding-bottom: 5.5rem; max-width: 460px;}

/* ---- Header (blue app bar) ---- */
.app-header {
    background: #0B3B8A; color: white; display: flex; align-items: center;
    justify-content: space-between; height: 46px; padding: 0 16px; box-sizing: border-box;
}
.app-header.header-solo { border-radius: 14px; }
.app-header.header-right { border-radius: 0 14px 14px 0; }
.app-header .title { font-weight: 700; font-size: 18px; }
.app-header .right-icon { font-size: 17px; opacity: .95; }

.st-key-header_row div[data-testid="stHorizontalBlock"] { gap: 0 !important; }
.st-key-btn_back button {
    background: #0B3B8A !important; color: white !important; border: none !important;
    border-radius: 14px 0 0 14px !important; font-size: 20px !important; font-weight: 700 !important;
    height: 46px !important; padding: 0 !important; box-shadow: none !important;
}

/* ---- Generic buttons ---- */
div.stButton > button, div.stFormSubmitButton > button, div.stDownloadButton > button {
    border-radius: 10px; padding: 0.6rem 1rem; font-weight: 600; width: 100%; border: 1px solid #dcdfe4;
}

/* ---- Responsive: en pantallas angostas, columnas van una debajo de otra ---- */
@media (max-width: 480px) {
    div[data-testid="stHorizontalBlock"] { flex-wrap: wrap; }
    div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] { min-width: 100% !important; flex: 1 1 100% !important; }
    .st-key-bottom_nav div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] { min-width: 0 !important; flex: 1 1 0 !important; }
    .st-key-header_row div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] { min-width: 0 !important; }
}

/* ---- Home action cards (colored) ---- */
.st-key-btn_crear_aviso button, .st-key-btn_guardar_aviso button,
.st-key-btn_iniciar_ot button, .st-key-btn_aprobar button,
.st-key-btn_guardar_reporte button {
    background: #2FA84F !important; color: #fff !important; border: none !important;
}
.st-key-btn_generar_reporte button { background: #0B3B8A !important; color: #fff !important; border: none !important; }
.st-key-btn_add_trabajo button, .st-key-btn_add_equipo_otro button, .st-key-btn_add_material button, .st-key-btn_add_collera button,
.st-key-btn_add_trabajador button { background: #F1F3F6 !important; color: #0B3B8A !important; border: 1px dashed #0B3B8A !important; }
.st-key-btn_backlog_avisos button { background: #0B3B8A !important; color: #fff !important; border: none !important; text-align: left !important; }
.st-key-btn_backlog_ots button { background: #3D7DD9 !important; color: #fff !important; border: none !important; text-align: left !important; }
.st-key-btn_reportes button, .st-key-btn_backup_reportes button, .st-key-btn_preparar_respaldo button { background: #F1F3F6 !important; color: #333 !important; border: 1px solid #dcdfe4 !important; text-align: left !important; }
.st-key-btn_mis_avisos button { background: #0B3B8A !important; color: #fff !important; border: none !important; text-align: left !important; }
.st-key-btn_mis_reportes button { background: #3D7DD9 !important; color: #fff !important; border: none !important; text-align: left !important; }
.st-key-btn_rechazar button { background: #DC3545 !important; color: #fff !important; border: none !important; }
.st-key-btn_observar button { background: #E0A458 !important; color: #fff !important; border: none !important; }
.st-key-btn_cerrar button, .st-key-btn_cerrar_ot button { background: #6c757d !important; color: #fff !important; border: none !important; }

/* ---- Pills / badges ---- */
.pill { display: inline-block; padding: 3px 12px; border-radius: 999px; font-size: 12px; font-weight: 700; }

/* ---- List rows ---- */
.aviso-row { padding: 10px 0; border-bottom: 1px solid #eee; }
.aviso-title { font-weight: 700; font-size: 15px; margin-bottom: 2px; }
.aviso-desc { color: #444; font-size: 13.5px; }
.aviso-meta { color: #888; font-size: 12px; margin-top: 4px; }

/* ---- Login ---- */
.st-key-btn_login button { background: #0B3B8A !important; color: #fff !important; border: none !important; }

/* ---- Fixed bottom nav ---- */
.st-key-bottom_nav {
    position: fixed; bottom: 0; left: 0; right: 0; margin: 0 auto; max-width: 460px;
    background: white; border-top: 1px solid #e3e5e8; padding: 4px 4px 12px 4px; z-index: 999;
}
.st-key-bottom_nav div.stButton > button {
    border: none !important; background: transparent !important; color: #6c757d;
    font-weight: 600; font-size: 12px; box-shadow: none !important; padding: 6px 2px !important;
}
.nav-brand { text-align: center; color: #b6bcc4; font-size: 11px; letter-spacing: 1px; padding-top: 2px; }

/* ---- Pie "Desarrollado por" ---- */
.pie-dev { margin: 22px 0 10px 0; padding: 14px 16px; border-radius: 12px; text-align: center;
    background: linear-gradient(135deg, #0B3B8A 0%, #3D7DD9 100%); color: #fff; }
.pie-dev .pie-dev-nombre { font-weight: 700; font-size: 13px; }
.pie-dev .pie-dev-contacto { font-size: 11.5px; opacity: .9; margin-top: 2px; }
.pie-dev .pie-dev-contacto a { color: #fff; text-decoration: underline; }
/* ---- Componente invisible del historial del navegador (botón Atrás) ---- */
.st-key-nav_historial_box { position: absolute; height: 0; overflow: hidden; margin: 0; }
.st-key-btn_footer_whatsapp button {
    background: #25D366 !important; color: #fff !important; border: none !important;
    font-weight: 700 !important; margin-top: 8px !important;
}
</style>
"""
st.markdown(MOBILE_CSS, unsafe_allow_html=True)

# -------------------------
# Base de datos SQLite (independiente de Prueba_app2.py)
# -------------------------
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "registros_app3.db")

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

# Columnas del tercer nivel de validación, específicas por rol (EFE o Supervisor Subcontrato)
AVISO_COLUMNS = [
    "AvisoID", "Activo", "Ubicacion", "Prioridad", "Tipificacion",
    "FechaHora", "Descripcion", "Adjunto", "CreadoPor", "RolCreador",
    "Estado", "OT_Creada", "ValidacionNivel", "RolFinal", "ReporteID",
    "AprobadoSacyr", "ObsSacyr", "FechaValSacyr", "RespValSacyr",
    "AprobadoADI", "ObsADI", "FechaValADI", "RespValADI",
    "AprobadoEFE", "ObsEFE", "FechaValEFE", "RespValEFE",
    "AprobadoSubcontrato", "ObsSubcontrato", "FechaValSubcontrato", "RespValSubcontrato",
]

OTS_COLUMNS = ["OTID", "AvisoID", "Activo", "Responsable", "FechaProgramada", "EstadoOT"]

# Control de Durmientes por Collera (Norma NS-01-01-00)
COLLERAS_CSV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "colleras_catalogo.csv")
DURMIENTES_CSV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "durmientes_estado_completo.csv")
DURMIENTES_COLUMNS = ["ColleraID", "Posicion", "Estado", "FechaActualizacion", "ReporteID"]
ESTADOS_DURMIENTE = ["Bueno", "Malo", "Nuevo", "Reemplazado"]

def init_db():
    conn = get_db()
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS activos (
        CodigoActivo TEXT PRIMARY KEY,
        NombreActivo TEXT, Ubicacion TEXT, TipoActivo TEXT
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS avisos (
        AvisoID TEXT PRIMARY KEY,
        Activo TEXT, Ubicacion TEXT, Prioridad TEXT, Tipificacion TEXT,
        FechaHora TEXT, Descripcion TEXT, Adjunto TEXT, CreadoPor TEXT, RolCreador TEXT,
        Estado TEXT, OT_Creada INTEGER,
        ValidacionNivel TEXT, RolFinal TEXT, ReporteID TEXT,
        AprobadoSacyr INTEGER, ObsSacyr TEXT, FechaValSacyr TEXT, RespValSacyr TEXT,
        AprobadoADI INTEGER, ObsADI TEXT, FechaValADI TEXT, RespValADI TEXT,
        AprobadoEFE INTEGER, ObsEFE TEXT, FechaValEFE TEXT, RespValEFE TEXT,
        AprobadoSubcontrato INTEGER, ObsSubcontrato TEXT, FechaValSubcontrato TEXT, RespValSubcontrato TEXT
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS ots (
        OTID TEXT PRIMARY KEY,
        AvisoID TEXT, Activo TEXT, Responsable TEXT, FechaProgramada TEXT, EstadoOT TEXT
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS colleras (
        ColleraID TEXT PRIMARY KEY, PK INTEGER, Collera INTEGER,
        Ubicacion TEXT, KmDesde REAL, KmHasta REAL
    )""")
    # Sin PRIMARY KEY compuesta a propósito: es un historial de cambios (una fila por
    # modificación real), no "última foto" -- cada inspección que cambia un durmiente
    # agrega una fila nueva en vez de sobreescribir la anterior.
    c.execute("""CREATE TABLE IF NOT EXISTS durmientes_estado (
        ColleraID TEXT, Posicion INTEGER, Estado TEXT,
        FechaActualizacion TEXT, ReporteID TEXT
    )""")

    # Catálogo de colleras: geometría prácticamente inmutable, se siembra una sola vez desde
    # el CSV que viaja junto a la app (así sobrevive a que Streamlit Cloud reinicie el disco
    # local, sin depender de Google Sheets para 2404 filas de datos estáticos).
    c.execute("SELECT COUNT(*) FROM colleras")
    if c.fetchone()[0] == 0 and os.path.exists(COLLERAS_CSV_PATH):
        import csv as _csv
        with open(COLLERAS_CSV_PATH, encoding="utf-8") as f:
            filas_colleras = [
                (r["ColleraID"], int(r["PK"]), int(r["Collera"]), r["Ubicacion"],
                 float(r["Km_Desde"]), float(r["Km_Hasta"]))
                for r in _csv.DictReader(f)
            ]
        c.executemany("INSERT OR REPLACE INTO colleras VALUES (?,?,?,?,?,?)", filas_colleras)

    # Migración: agrega columnas nuevas si la tabla avisos ya existía con el esquema anterior (sin romper datos previos)
    c.execute("PRAGMA table_info(avisos)")
    cols_existentes = {row[1] for row in c.fetchall()}
    columnas_nuevas = {
        "RolFinal": "TEXT DEFAULT 'EFE'",
        "ReporteID": "TEXT DEFAULT ''",
        "AprobadoSubcontrato": "INTEGER DEFAULT 0",
        "ObsSubcontrato": "TEXT DEFAULT ''",
        "FechaValSubcontrato": "TEXT DEFAULT ''",
        "RespValSubcontrato": "TEXT DEFAULT ''",
    }
    for col, decl in columnas_nuevas.items():
        if col not in cols_existentes:
            c.execute(f"ALTER TABLE avisos ADD COLUMN {col} {decl}")

    conn.commit()
    conn.close()

def _valores_hoja_sheets_directo(nombre_hoja: str) -> list:
    """Lee todos los valores de una hoja de Google Sheets en UNA sola llamada a la API.
    Lanza excepción si Sheets no está configurado o si la lectura falla: nunca devuelve
    una hoja 'vacía' por un error transitorio (cuota, red), porque el llamador podría
    tomarla como dato real y luego sobreescribir el Sheet con esa tabla vacía."""
    sh = _cliente_sheets()
    if sh is None:
        raise RuntimeError("Google Sheets no está configurado")
    try:
        # UNFORMATTED_VALUE: sin esto, Sheets entrega los números ya "formateados" como
        # texto según el locale de la hoja (chileno: coma decimal) -- ej. "46,644" -- y esa
        # coma se interpreta como separador de miles (asumiendo locale inglés), quedando
        # 46644. UNFORMATTED_VALUE trae el número real de vuelta, sin pasar por texto.
        resp = sh.values_get(f"'{nombre_hoja}'", params={"valueRenderOption": "UNFORMATTED_VALUE"})
    except gspread.exceptions.APIError:
        # Solo si la pestaña de verdad no existe se crea vacía; cualquier otro error se propaga.
        try:
            sh.worksheet(nombre_hoja)
        except gspread.WorksheetNotFound:
            _hoja_sheets(sh, nombre_hoja)
            return [REPORTES_SHEETS[nombre_hoja]]
        raise
    return resp.get("values", [])

@st.cache_data(ttl=60, show_spinner=False)
def _valores_hoja_sheets(nombre_hoja: str) -> list:
    """Versión con caché (compartida entre usuarios, 60 s) de _valores_hoja_sheets_directo:
    evita volver a pedirle a Google cada hoja en cada interacción de cada usuario (la API
    permite ~60 lecturas por minuto). Se invalida apenas la app escribe algo en Sheets."""
    return _valores_hoja_sheets_directo(nombre_hoja)

def _invalidar_cache_sheets():
    _valores_hoja_sheets.clear()

def _fecha_desde_serial(valor):
    """Las filas guardadas por versiones antiguas de la app (modo USER_ENTERED) quedaron
    con las fechas convertidas por Google a número de serie (46235 = 2026-08-01). Con
    UNFORMATTED_VALUE llegan como número: se devuelven al texto 'AAAA-MM-DD[ HH:MM]' que
    usa el resto de la app. Cualquier otro valor se deja tal cual."""
    if isinstance(valor, bool) or not isinstance(valor, (int, float)) or not 20000 <= valor <= 80000:
        return valor
    dt = datetime(1899, 12, 30) + timedelta(days=float(valor))
    if float(valor).is_integer():
        return dt.strftime("%Y-%m-%d")
    return (dt + timedelta(seconds=30)).strftime("%Y-%m-%d %H:%M")

def _valores_a_df(valores: list, columnas: list) -> pd.DataFrame:
    """Convierte la respuesta cruda de Sheets (lista de filas, la primera = encabezados) en
    DataFrame. Sheets omite las celdas vacías al final de cada fila, así que se rellenan."""
    if len(valores) <= 1:
        return pd.DataFrame(columns=columnas)
    headers = [str(h) for h in valores[0]]
    n = len(headers)
    filas = [(list(f) + [""] * n)[:n] for f in valores[1:]]
    df = pd.DataFrame(filas, columns=headers)
    for col in columnas:
        if col not in df.columns:
            df[col] = ""
    for col in df.columns:
        if col.startswith("Fecha"):
            df[col] = df[col].map(_fecha_desde_serial)
    return df

def _leer_hoja_sheets_o_none(nombre_hoja: str, columnas: list, estricto: bool = False, fresco: bool = False):
    """Lee una hoja desde Google Sheets. Devuelve None si Sheets no está configurado.
    Si está configurado pero la lectura falla: con estricto=True lanza la excepción (para
    que el llamador conserve lo que ya tenía), si no devuelve None (respaldo local).
    fresco=True salta la caché (p.ej. para calcular el próximo correlativo)."""
    if _cliente_sheets() is None:
        return None
    try:
        lector = _valores_hoja_sheets_directo if fresco else _valores_hoja_sheets
        return _valores_a_df(lector(nombre_hoja), columnas)
    except Exception:
        if estricto:
            raise
        return None

def load_avisos(estricto: bool = False):
    """Avisos: Google Sheets es la fuente de verdad cuando está configurado (así los avisos
    creados en Streamlit Cloud sobreviven a que el hosting reinicie el disco local); si no
    está configurado, o falla la conexión, se usa el SQLite local como respaldo."""
    df = _leer_hoja_sheets_o_none("Avisos", AVISO_COLUMNS, estricto)
    if df is None:
        conn = get_db()
        df = pd.read_sql("SELECT * FROM avisos", conn)
        conn.close()
    for col in ["OT_Creada", "AprobadoSacyr", "AprobadoADI", "AprobadoEFE", "AprobadoSubcontrato"]:
        df[col] = df[col].apply(_a_bool)
    df["RolFinal"] = df["RolFinal"].fillna("EFE").replace("", "EFE")
    return df

def load_activos():
    conn = get_db()
    df = pd.read_sql("SELECT * FROM activos", conn)
    conn.close()
    return df

def load_ots(estricto: bool = False):
    """OTs: mismo criterio que load_avisos() -- Sheets primero si está configurado, si no SQLite."""
    df = _leer_hoja_sheets_o_none("OTs", OTS_COLUMNS, estricto)
    if df is None:
        conn = get_db()
        df = pd.read_sql("SELECT * FROM ots", conn)
        conn.close()
    return df

def save_all_avisos():
    df = st.session_state.avisos.copy()
    for col in ["OT_Creada", "AprobadoSacyr", "AprobadoADI", "AprobadoEFE", "AprobadoSubcontrato"]:
        df[col] = df[col].astype(int)
    df = df[AVISO_COLUMNS]
    conn = get_db()
    df.to_sql("avisos", conn, if_exists="replace", index=False)
    conn.commit()
    conn.close()
    _sincronizar_tabla_sheets("Avisos", df.values.tolist())

def save_all_ots():
    df = st.session_state.ots.copy()
    if not df.empty:
        df = df[OTS_COLUMNS]
    conn = get_db()
    df.to_sql("ots", conn, if_exists="replace", index=False)
    conn.commit()
    conn.close()
    _sincronizar_tabla_sheets("OTs", df.values.tolist())

# -------------------------
# Control de Durmientes por Collera (catálogo geométrico + estado real)
# -------------------------
def cargar_colleras_de_pk(pk: int) -> list:
    conn = get_db()
    filas = conn.execute(
        "SELECT ColleraID, PK, Collera, Ubicacion, KmDesde, KmHasta FROM colleras WHERE PK=? ORDER BY Collera",
        (pk,)).fetchall()
    conn.close()
    return [dict(f) for f in filas]

def cargar_todas_colleras() -> pd.DataFrame:
    conn = get_db()
    df = pd.read_sql("SELECT * FROM colleras", conn)
    conn.close()
    return df

def load_durmientes_estado(estricto: bool = False) -> pd.DataFrame:
    """Estado real de los durmientes: Google Sheets es la fuente de verdad si está
    configurado (así las inspecciones de terreno sobreviven a un reinicio de Streamlit
    Cloud), si no se usa el SQLite local -- mismo criterio que load_avisos()/load_ots()."""
    df = _leer_hoja_sheets_o_none("DurmientesEstado", DURMIENTES_COLUMNS, estricto)
    if df is None:
        conn = get_db()
        try:
            df = pd.read_sql("SELECT * FROM durmientes_estado", conn)
        except Exception:
            df = pd.DataFrame(columns=DURMIENTES_COLUMNS)
        conn.close()
    if not df.empty:
        df["Posicion"] = df["Posicion"].apply(lambda v: int(_num(v)))
    return df

def _seed_durmientes_estado_inicial() -> pd.DataFrame:
    """La primera vez que corre la app (tabla/hoja de estado todavía vacía), siembra las
    pocas colleras que el catastro ya trae con datos reales -- igual criterio que colleras,
    pero acá NO se reintenta en cada init: solo corre mientras el estado esté vacío, para
    no pisar inspecciones de terreno ya guardadas con el dato original del Excel."""
    if not os.path.exists(DURMIENTES_CSV_PATH):
        return pd.DataFrame(columns=DURMIENTES_COLUMNS)
    import csv as _csv
    filas = []
    with open(DURMIENTES_CSV_PATH, encoding="utf-8") as f:
        for r in _csv.DictReader(f):
            for pos in range(1, 24):
                val = r.get(f"D{pos}")
                if val:
                    filas.append({
                        "ColleraID": r["ColleraID"], "Posicion": pos, "Estado": val,
                        "FechaActualizacion": "", "ReporteID": "SEED-CATASTRO",
                    })
    return pd.DataFrame(filas, columns=DURMIENTES_COLUMNS)

def save_all_durmientes_estado():
    df = st.session_state.durmientes_estado.copy()
    if not df.empty:
        df = df[DURMIENTES_COLUMNS]
    conn = get_db()
    conn.execute("DELETE FROM durmientes_estado")
    if not df.empty:
        conn.executemany("INSERT INTO durmientes_estado VALUES (?,?,?,?,?)", df.values.tolist())
    conn.commit()
    conn.close()
    _sincronizar_tabla_sheets("DurmientesEstado", df.values.tolist() if not df.empty else [])

def _estado_vigente(df_estado: pd.DataFrame) -> pd.DataFrame:
    """durmientes_estado es un historial de cambios (una fila por modificación real), no
    'última foto'. Esta función deja solo la fila más reciente de cada (ColleraID,
    Posicion) -- el estado vigente -- para usar en cumplimiento normativo y en el
    formulario. Desempate estable por orden de inserción si dos cambios cayeron en el
    mismo minuto (FechaActualizacion solo tiene precisión de minuto)."""
    if df_estado.empty:
        return df_estado
    df = df_estado.copy()
    df["_orden"] = range(len(df))
    df = df.sort_values(["FechaActualizacion", "_orden"])
    return df.groupby(["ColleraID", "Posicion"], as_index=False).tail(1).drop(columns="_orden")

def _estados_por_collera(df_estado_vigente: pd.DataFrame) -> dict:
    """{ColleraID: {posición: estado}} a partir del estado vigente, en una sola pasada
    (en vez de filtrar el DataFrame una vez por cada una de las 2404 colleras)."""
    resultado = {}
    if df_estado_vigente.empty:
        return resultado
    for cid, pos, est in zip(df_estado_vigente["ColleraID"], df_estado_vigente["Posicion"], df_estado_vigente["Estado"]):
        if est:
            resultado.setdefault(str(cid), {})[int(_num(pos))] = est
    return resultado

def cargar_estado_collera(collera_id: str) -> dict:
    return _estados_por_collera(_estado_vigente(st.session_state.durmientes_estado)).get(str(collera_id), {})

# Estados generales que cuentan como incumplimiento / como falta de datos (igual que el
# Dashboard del Excel: "No Cumple" + "No Cumple (racha)"; "Sin Inspeccionar" + "Inspección Incompleta").
ESTADOS_NO_CUMPLE = ("No Cumple", "No Cumple (racha)")
ESTADOS_SIN_DATOS = ("Sin Inspeccionar", "Inspección Incompleta")

def es_curva(ubicacion) -> bool:
    """Igual que el Excel de control (columna AJ: IF(C12="Curva", ...)): solo la ubicación
    exactamente 'Curva' usa el límite de racha de curva (máx. 2 'Malo' seguidos). Las demás
    ('Curva < 1000', 'Principio Curva > 1000', etc.) se evalúan como recta (máx. 3), tal
    como lo hace el Excel -- decisión confirmada por el usuario. La comparación de Excel
    no distingue mayúsculas."""
    return str(ubicacion or "").lower() == "curva"

def evaluar_secuencia(secuencia: list, ubicacion: str) -> dict:
    """Misma lógica que las columnas AA:AK de la hoja 'Control Durmientes' del Excel
    (Norma NS-01-01-00, puntos 6.5.3 y 6.5.4). `secuencia` = estados de D1..D23
    (None o "" = posición no inspeccionada; NO se asume 'Bueno')."""
    secuencia = [(e or None) for e in secuencia]
    total_registrado = sum(1 for e in secuencia if e)                    # AA
    n_buenos = secuencia.count("Bueno")                                   # AB
    n_malos = secuencia.count("Malo")                                     # AC
    n_nuevos = secuencia.count("Nuevo")                                   # AD
    n_reempl = secuencia.count("Reemplazado")                             # AE
    efectivos = n_buenos + n_nuevos + n_reempl                            # AF
    pct_renov = (n_nuevos + n_reempl) / total_registrado if total_registrado else 0.0  # AG

    # Racha de 'Malo' consecutivos: una posición en blanco corta la racha (AH del Excel).
    racha, racha_max = 0, 0
    for e in secuencia:
        racha = racha + 1 if e == "Malo" else 0
        racha_max = max(racha_max, racha)
    limite_racha = 2 if es_curva(ubicacion) else 3

    if total_registrado == 0:                                             # AI
        min_efectivos = "Sin datos"
    elif total_registrado < 10:
        min_efectivos = "Incompleto"
    else:
        min_efectivos = "Sí" if efectivos >= 10 else "No"

    if total_registrado == 0:                                             # AJ
        racha_consec = "Sin datos"
    else:
        racha_consec = "No" if racha_max > limite_racha else "Sí"

    if racha_consec == "No":                                              # AK
        estado_general = "No Cumple (racha)"
    elif total_registrado == 0:
        estado_general = "Sin Inspeccionar"
    elif total_registrado < 10:
        estado_general = "Inspección Incompleta"
    elif min_efectivos == "Sí":
        estado_general = "Cumple"
    else:
        estado_general = "No Cumple"

    return {
        "total_registrado": total_registrado, "n_buenos": n_buenos, "n_malos": n_malos,
        "n_nuevos": n_nuevos, "n_reempl": n_reempl, "efectivos": efectivos,
        "pct_renovacion": pct_renov, "racha_max": racha_max, "limite_racha": limite_racha,
        "min_efectivos": min_efectivos, "racha_consec": racha_consec, "estado_general": estado_general,
    }

def evaluar_collera(collera_id: str, ubicacion: str, df_estado_vigente: pd.DataFrame | None = None,
                    estados_por_collera: dict | None = None) -> dict:
    """Evalúa una collera con su estado vigente. Para loops grandes (las 2404 colleras)
    pasar `estados_por_collera` ya calculado con _estados_por_collera()."""
    if estados_por_collera is None:
        if df_estado_vigente is None:
            df_estado_vigente = _estado_vigente(st.session_state.durmientes_estado)
        estados_por_collera = _estados_por_collera(df_estado_vigente)
    estados = estados_por_collera.get(str(collera_id), {})
    return evaluar_secuencia([estados.get(p) for p in range(1, 24)], ubicacion)

def _evaluar_todas_colleras() -> list:
    """[(fila del catálogo, evaluación)] de las 2404 colleras con el estado vigente."""
    df_colleras = cargar_todas_colleras()
    estados = _estados_por_collera(_estado_vigente(st.session_state.durmientes_estado))
    return [(row, evaluar_collera(row["ColleraID"], row["Ubicacion"], estados_por_collera=estados))
            for _, row in df_colleras.iterrows()]

def resumen_global_durmientes() -> dict:
    resultados = [ev for _, ev in _evaluar_todas_colleras()]
    cumplen = sum(1 for r in resultados if r["estado_general"] == "Cumple")
    no_cumplen = sum(1 for r in resultados if r["estado_general"] in ESTADOS_NO_CUMPLE)
    sin_datos = sum(1 for r in resultados if r["estado_general"] in ESTADOS_SIN_DATOS)
    evaluables = cumplen + no_cumplen
    pct_cumplimiento = (cumplen / evaluables * 100) if evaluables else 0.0
    tot = {k: sum(r[k] for r in resultados) for k in ("total_registrado", "n_buenos", "n_malos", "n_nuevos", "n_reempl")}
    pct_renov_global = ((tot["n_nuevos"] + tot["n_reempl"]) / tot["total_registrado"] * 100) if tot["total_registrado"] else 0.0
    return {"total": len(resultados), "cumplen": cumplen, "no_cumplen": no_cumplen,
            "sin_datos": sin_datos, "pct_cumplimiento": pct_cumplimiento,
            "total_buenos": tot["n_buenos"], "total_malos": tot["n_malos"], "total_nuevos": tot["n_nuevos"],
            "total_reempl": tot["n_reempl"], "pct_renov_global": pct_renov_global}

def resumen_por_pk() -> list:
    """Una fila por PK: n° colleras, evaluadas, estado del km (Bueno/Regular/Malo/Sin Datos)."""
    por_pk = {}
    for row, ev in _evaluar_todas_colleras():
        por_pk.setdefault(int(row["PK"]), []).append(ev)
    filas = []
    for pk in sorted(por_pk):
        evals = por_pk[pk]
        n_total = len(evals)
        n_con_datos = sum(1 for e in evals if e["total_registrado"] > 0)
        n_no_cumplen = sum(1 for e in evals if e["estado_general"] in ESTADOS_NO_CUMPLE)
        if n_con_datos == 0:
            estado_km = "Sin Datos"
        elif n_no_cumplen > 0:
            estado_km = "Malo"
        elif n_con_datos < n_total / 2:
            estado_km = "Regular"
        else:
            estado_km = "Bueno"
        tot_reg = sum(e["total_registrado"] for e in evals)
        n_nue_reem = sum(e["n_nuevos"] + e["n_reempl"] for e in evals)
        filas.append({"PK": int(pk), "N° Colleras": n_total, "Con datos": n_con_datos,
                      "Total Durm.": tot_reg,
                      "Buenos": sum(e["n_buenos"] for e in evals), "Malos": sum(e["n_malos"] for e in evals),
                      "Nuevos": sum(e["n_nuevos"] for e in evals), "Reempl.": sum(e["n_reempl"] for e in evals),
                      "% Renov.": round(n_nue_reem / tot_reg * 100, 1) if tot_reg else 0.0,
                      "No cumplen": n_no_cumplen, "Estado KM": estado_km})
    return filas

# Columnas de la foto de cada inspección de collera hecha en un Reporte Diario: la misma
# fila que la hoja 'Control Durmientes' del Excel (D1..D23 + columnas calculadas).
INSPECCION_COLUMNS = (
    ["ReporteID", "Fecha", "ColleraID", "PK", "Collera", "Ubicacion"]
    + [f"D{p}" for p in range(1, 24)]
    + ["TotalRegistrado", "Buenos", "Malos", "Nuevos", "Reemplazados", "Efectivos",
       "PctRenovacion", "MinEfectivos", "RachaConsec", "EstadoGeneral"]
)

def fila_inspeccion(reporte_id: str, fecha_str: str, collera: dict, secuencia: list) -> list:
    ev = evaluar_secuencia(secuencia, collera["Ubicacion"])
    return ([reporte_id, fecha_str, collera["ColleraID"], int(collera["PK"]), int(collera["Collera"]), collera["Ubicacion"]]
            + [e or "" for e in secuencia]
            + [ev["total_registrado"], ev["n_buenos"], ev["n_malos"], ev["n_nuevos"], ev["n_reempl"],
               ev["efectivos"], round(ev["pct_renovacion"], 4), ev["min_efectivos"], ev["racha_consec"],
               ev["estado_general"]])

def filas_cambios_durmientes(reporte_id: str, fecha_str: str, secuencias: dict) -> list:
    """Filas nuevas para el historial DurmientesEstado: {ColleraID: [23 estados]}. Solo
    agrega las posiciones que cambian respecto al estado vigente (no ensucia el historial
    con 'confirmaciones' sin cambio) y nunca sobreescribe filas anteriores, para poder
    auditar cuándo y en qué reporte cambió cada durmiente. En blanco = no tocar."""
    vigente = _estados_por_collera(_estado_vigente(st.session_state.durmientes_estado))
    filas = []
    for collera_id, secuencia in secuencias.items():
        previos = vigente.get(str(collera_id), {})
        for pos, estado in enumerate(secuencia, start=1):
            if estado and estado != previos.get(pos):
                filas.append([collera_id, pos, estado, fecha_str, reporte_id])
    return filas

def persistir_cambios_durmientes_local(filas: list):
    """Tras guardar en Sheets (con guardar_bloques), refleja los cambios en la sesión y en
    el SQLite local. Se AGREGAN filas en vez de reescribir todo el historial."""
    if not filas:
        return
    st.session_state.durmientes_estado = pd.concat(
        [st.session_state.durmientes_estado, pd.DataFrame(filas, columns=DURMIENTES_COLUMNS)], ignore_index=True)
    conn = get_db()
    conn.executemany("INSERT INTO durmientes_estado VALUES (?,?,?,?,?)", filas)
    conn.commit()
    conn.close()

# -------------------------
# Reporte Diario (Personal Terreno) - base de datos en Excel
# -------------------------
REPORTES_XLSX_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_diarios.xlsx")
FOTOS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fotos_reportes")

REPORTES_SHEETS = {
    "Reportes": ["ReporteID", "Fecha", "GrupoVia", "Usuario", "Observaciones", "FechaCreacion", "Ubicacion"],
    "Trabajos": ["ReporteID", "Actividad", "ColleraDesdeID", "ColleraHastaID", "KmDesde", "KmHasta", "Unidad", "Cantidad", "Hombres", "HH"],
    "Equipos": ["ReporteID", "Equipo", "Cantidad"],
    "Materiales": ["ReporteID", "Material", "Cantidad", "Estado"],
    "Asistencia": ["ReporteID", "Trabajador", "Cargo", "Estado", "HoraIngreso", "HoraSalida", "HorasExtras"],
    "Fotos": ["ReporteID", "NombreArchivo", "RutaArchivo"],
    "PlanMensual": ["Actividad", "Unidad", "Anio", "Mes", "CantidadPlanificada"],
    "Avisos": AVISO_COLUMNS,
    "OTs": OTS_COLUMNS,
    "DurmientesEstado": DURMIENTES_COLUMNS,
    "InspeccionColleras": INSPECCION_COLUMNS,
}

DIAS_ES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

def horas_jornada_por_fecha(fecha) -> int:
    wd = fecha.weekday()  # Lunes=0 ... Domingo=6
    if wd in (0, 1):
        return 9
    if wd in (2, 3, 4):
        return 8
    return 6

ACTIVIDADES_TRABAJO = [
    "Inspección Vía",
    "Inspección de Desviador",
    "Reemplazo de Durmientes Común",
    "Mantenimiento de geometría vía",
    "Reposición elementos enrrieladura - Eclisas",
    "Reposición elementos enrrieladura - Pernos",
    "Reposición elementos enrrieladura - Tirafondos",
    "Mantenimiento de Desviador",
    "Mantenimiento de Desvíos",
    "Mantenimiento Cruce a Nivel",
    "Mantenimiento de Cruce peatonal",
    "Mantenimiento de alcantarillas",
    "Mantenimiento de cunetas",
    "Reparación/Reposición de Señalización Pare",
    "Reparación/Reposición de Señal Cruce Pito",
    "Reparación/Reposición de Baliza PK",
    "Control vegetación y retiro de desechos",
    "Otras actividades",
]

UNIDADES_TRABAJO = ["Kmv", "Und", "mlv", "H-H"]

MESES_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

EQUIPOS_CATALOGO = [
    "Cortadora de Rieles", "Generador Eléctrico", "Llave de Impacto", "Motosierra", "Taladro Industrial",
    "Bandeja", "Regla Nivel Trocha", "Gatas Alza Vías", "Grupo Bateo Universal", "Perforadora de Rieles",
    "Equipo de Iluminación", "Desbrozadora", "Desbrozadora de Altura", "Arriendo Radio comunicaciones",
    "Vehículo de transporte de personal", "Camión Pluma HiRail",
]

MATERIALES_CATALOGO = [
    "Sujeciones clip", "Silla tipo Pandroll", "Silla re 115 doble pestaña riel", "Tirafondo un2",
    "Tirafondo un5", "Eclisas planchuela tipo X", "Pernos rieleros tipo X", "Perno 19x300mm",
    "Señalética Vial de Cruces a Nivel y ferroviaria", "Grasa grafitada", "Asfalto en frío en cruce a nivel",
    "Durmientes de madera impregnada", "Otro (especificar)",
]

ESTADOS_MATERIAL = ["Buen Estado", "Mal Estado"]

CARGOS_TRABAJADOR = ["Jefe de Grupo", "Ayudante", "Chofer", "P5.1", "Operario Vía", "Otro"]

GRUPOS_VIA = ["01", "02", "03", "04", "05"]

# -------------------------
# Google Sheets: histórico durable (sobrevive reinicios del hosting).
# El Excel local se mantiene igual, para las descargas del momento; Sheets es
# la fuente de verdad cuando está configurada en Secrets ([google_sheets] ...).
# -------------------------
GOOGLE_SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

@st.cache_resource(show_spinner=False)
def _cliente_sheets():
    """Conecta con el Google Sheet configurado en Secrets. Si no está configurado
    o falla la conexión, devuelve None y la app sigue funcionando solo con el Excel local."""
    try:
        creds_info = dict(st.secrets["google_sheets"]["credentials"])
        sheet_id = st.secrets["google_sheets"]["sheet_id"]
    except Exception:
        return None
    try:
        creds = GoogleCredentials.from_service_account_info(creds_info, scopes=GOOGLE_SHEETS_SCOPES)
        gc = gspread.authorize(creds)
        return gc.open_by_key(sheet_id)
    except Exception:
        return None

def _asegurar_hoja(spreadsheet, nombre_hoja: str, headers: list | None = None):
    """Devuelve (worksheet, encabezados reales de la fila 1), creando la pestaña si no
    existe. Si el esquema creció (p.ej. se agregó una columna nueva) y esta hoja ya
    existía en el Sheet real, agrega los encabezados faltantes al final -- ampliando
    primero la grilla: antes no se ampliaba, Google rechazaba la columna nueva con
    'exceeds grid limits' y ese error, silenciado, hacía perder filas enteras (así se
    perdieron las actividades del RPT-00014 en la hoja Trabajos)."""
    headers = headers or REPORTES_SHEETS[nombre_hoja]
    try:
        ws = spreadsheet.worksheet(nombre_hoja)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=nombre_hoja, rows=2000, cols=max(len(headers), 10))
        ws.update(values=[headers], range_name="A1", value_input_option="RAW")
        return ws, list(headers)
    header_row = [str(h) for h in ws.row_values(1)]
    faltantes = [h for h in headers if h not in header_row]
    if faltantes:
        nuevo_header = header_row + faltantes
        if len(nuevo_header) > ws.col_count:
            ws.resize(cols=len(nuevo_header))
        ws.update(values=[nuevo_header], range_name="A1", value_input_option="RAW")
        header_row = nuevo_header
    return ws, header_row

def _hoja_sheets(spreadsheet, nombre_hoja: str, headers: list | None = None):
    return _asegurar_hoja(spreadsheet, nombre_hoja, headers)[0]

def _preparar_hojas(spreadsheet, nombres: list) -> dict:
    """{nombre: (worksheet, encabezados reales)} para varias hojas con 2 llamadas a la API
    en total (en vez de 2 por hoja). Solo si falta una hoja o columnas usa _asegurar_hoja."""
    existentes = {ws.title: ws for ws in spreadsheet.worksheets()}
    presentes = [n for n in nombres if n in existentes]
    encabezados = {}
    if presentes:
        resp = spreadsheet.values_batch_get([f"'{n}'!1:1" for n in presentes])
        for n, vr in zip(presentes, resp.get("valueRanges", [])):
            encabezados[n] = [str(h) for h in (vr.get("values") or [[]])[0]]
    resultado = {}
    for n in nombres:
        header = encabezados.get(n)
        if n in existentes and header is not None and all(h in header for h in REPORTES_SHEETS[n]):
            resultado[n] = (existentes[n], header)
        else:
            resultado[n] = _asegurar_hoja(spreadsheet, n)
    return resultado

def _fila_por_encabezado(fila: list, headers_app: list, header_real: list) -> list:
    """Reordena una fila (en el orden de REPORTES_SHEETS) según el orden REAL de las
    columnas de la hoja. Hojas creadas por versiones antiguas de la app tienen otro orden
    (p.ej. Trabajos: KmDesde antes que ColleraDesdeID); escribir por posición dejaba los
    datos en columnas equivocadas."""
    valores = dict(zip(headers_app, fila))
    return [valores.get(h, "") for h in header_real]

def _celda_api(v) -> dict:
    """Celda para la API batchUpdate. stringValue se guarda literal (como RAW): no se
    reinterpreta con el locale chileno ni como fórmula."""
    v = _celda_sheets(v)
    if isinstance(v, bool):
        return {"userEnteredValue": {"boolValue": v}}
    if isinstance(v, numbers.Number):
        return {"userEnteredValue": {"numberValue": float(v)}}
    if v == "":
        return {}
    return {"userEnteredValue": {"stringValue": str(v)}}

def _agregar_filas_excel_local(bloques: dict):
    init_reportes_excel()
    wb = openpyxl.load_workbook(REPORTES_XLSX_PATH)
    for nombre_hoja, filas in bloques.items():
        ws_local = wb[nombre_hoja]
        header_real = [c.value for c in ws_local[1]]
        for fila in filas:
            ws_local.append(_fila_por_encabezado(fila, REPORTES_SHEETS[nombre_hoja], header_real))
    wb.save(REPORTES_XLSX_PATH)

def _reporte_ya_guardado(reporte_id: str) -> bool:
    try:
        df = _leer_hoja_sheets_o_none("Reportes", REPORTES_SHEETS["Reportes"], estricto=True, fresco=True)
    except Exception:
        return False
    return df is not None and str(reporte_id) in set(df["ReporteID"].astype(str))

def guardar_bloques(bloques: dict) -> bool:
    """Agrega filas a varias hojas a la vez: {nombre_hoja: [filas en orden REPORTES_SHEETS]}.
    En Google Sheets van TODAS en un solo batchUpdate, que Google aplica completo o no
    aplica nada: un reporte ya no puede quedar a medias (cabecera guardada y actividades
    perdidas). Reintenta 3 veces. Devuelve False si no quedó guardado de forma durable,
    para que el llamador avise y NO limpie el formulario (así no se pierde lo ingresado).
    El Excel local se escribe después, solo como copia para descargas del momento."""
    bloques = {h: f for h, f in bloques.items() if f}
    if not bloques:
        return True
    sh = _cliente_sheets()
    if sh is not None:
        reporte_id = bloques["Reportes"][0][0] if "Reportes" in bloques else None
        ok = False
        for intento in range(3):
            if intento and reporte_id and _reporte_ya_guardado(reporte_id):
                ok = True  # el intento anterior sí llegó a Google, solo falló la respuesta
                break
            try:
                requests = []
                hojas = _preparar_hojas(sh, list(bloques))
                for nombre_hoja, filas in bloques.items():
                    ws, header_real = hojas[nombre_hoja]
                    requests.append({"appendCells": {
                        "sheetId": ws.id,
                        "rows": [{"values": [_celda_api(v) for v in _fila_por_encabezado(f, REPORTES_SHEETS[nombre_hoja], header_real)]}
                                 for f in filas],
                        "fields": "userEnteredValue",
                    }})
                sh.batch_update({"requests": requests})
                ok = True
                break
            except Exception:
                time.sleep(1.5 * (intento + 1))
        _invalidar_cache_sheets()
        if not ok:
            return False
    try:
        _agregar_filas_excel_local(bloques)
    except Exception:
        if sh is None:
            return False  # sin Sheets, el Excel local es el único registro
    return True

def _leer_hoja_df(nombre_hoja: str, fresco: bool = False) -> pd.DataFrame:
    """Lee una hoja completa como DataFrame: desde Google Sheets si está configurado
    (fuente de verdad, sobrevive reinicios), si no desde el Excel local."""
    columnas = REPORTES_SHEETS.get(nombre_hoja, [])
    df = _leer_hoja_sheets_o_none(nombre_hoja, columnas, fresco=fresco)
    if df is not None:
        return df
    if not os.path.exists(REPORTES_XLSX_PATH):
        return pd.DataFrame(columns=columnas)
    try:
        return pd.read_excel(REPORTES_XLSX_PATH, sheet_name=nombre_hoja)
    except Exception:
        return pd.DataFrame(columns=columnas)

def _celda_sheets(v):
    """Valor apto para enviar a Sheets: None/NaN (celdas vacías de pandas) no son JSON válido."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return v

def _num(val, default=0.0):
    """Convierte a número de forma segura; Sheets a veces entrega '' en celdas vacías."""
    try:
        if val is None or val == "":
            return default
        return float(val)
    except (TypeError, ValueError):
        return default

def _a_bool(val) -> bool:
    """Convierte a booleano de forma segura, venga de SQLite (0/1) o de Sheets (texto)."""
    return str(val).strip().lower() in ("1", "true", "verdadero", "si", "sí")

def _sincronizar_tabla_sheets(nombre_hoja: str, filas: list):
    """Reemplaza por completo el contenido de una hoja en Google Sheets con las filas dadas.
    Se usa para tablas chicas que se reescriben enteras cada vez que cambian -Avisos, OTs,
    PlanMensual- en vez de ir agregando filas como en Reportes/Trabajos.
    Primero escribe los datos nuevos encima y recién después limpia las filas sobrantes:
    antes se borraba la pestaña y se recreaba, y si la API fallaba entre medio la tabla
    completa se perdía en Sheets."""
    sh = _cliente_sheets()
    if sh is None:
        return
    headers = REPORTES_SHEETS[nombre_hoja]
    # RAW (no USER_ENTERED): USER_ENTERED hace que el propio Sheet reinterprete los
    # números según la configuración regional de la hoja -- con locale chileno (coma
    # decimal, punto como separador de miles) "430.56" se leía como "43.056" y quedaba
    # guardado como 43056. RAW escribe el valor literal, sin reinterpretación.
    valores = [headers] + [[_celda_sheets(v) for v in fila] for fila in filas]
    try:
        ws = _hoja_sheets(sh, nombre_hoja)
        filas_grilla = max(ws.row_count, len(valores))
        cols_grilla = max(ws.col_count, len(headers))
        if filas_grilla > ws.row_count or cols_grilla > ws.col_count:
            ws.resize(rows=filas_grilla, cols=cols_grilla)
        ws.update(values=valores, range_name="A1", value_input_option="RAW")
        if filas_grilla > len(valores):
            ws.batch_clear([f"A{len(valores) + 1}:{get_column_letter(cols_grilla)}{filas_grilla}"])
    except Exception:
        pass  # si falla Sheets, los datos igual quedaron guardados localmente
    finally:
        _invalidar_cache_sheets()

# -------------------------
# Respaldo de fotos en Google Sheets
# Las fotos se guardaban solo en el disco local, que Streamlit Cloud borra al reiniciar.
# Se guardan además en la hoja "FotosData" del mismo Sheet: la foto comprimida, en base64,
# partida en trozos de 45.000 caracteres (una celda admite 50.000). Se usa Sheets y no
# Drive porque una cuenta de servicio no tiene cuota para subir archivos a un Drive
# personal. Cuando una foto falta en el disco, se reconstruye desde esa hoja.
# -------------------------
FOTOS_DATA_HOJA = "FotosData"
FOTOS_DATA_COLUMNS = ["ReporteID", "NombreArchivo", "Parte", "TotalPartes", "Datos"]
FOTO_MAX_LADO = 1280   # px del lado mayor: suficiente para el PDF (13 cm de ancho)
FOTO_CALIDAD = 75
FOTO_CHUNK = 45000

def _ruta_foto_local(reporte_id, nombre) -> str:
    return os.path.join(FOTOS_DIR, str(reporte_id), str(nombre))

def _ubicar_foto(reporte_id, nombre, ruta_original="") -> str | None:
    """Ruta local donde está la foto ahora mismo, o None si no está en el disco."""
    ruta = _ruta_foto_local(reporte_id, nombre)
    if os.path.exists(ruta):
        return ruta
    if isinstance(ruta_original, str) and ruta_original and os.path.exists(ruta_original):
        return ruta_original
    return None

def _respaldar_fotos_sheets(reporte_id: str, fotos: list) -> bool | None:
    """Sube a FotosData las fotos [(nombre, ruta)] de un reporte. Devuelve None si Sheets
    no está configurado, True si todas quedaron respaldadas y False si alguna falló."""
    sh = _cliente_sheets()
    if sh is None or not fotos:
        return None
    try:
        ws = _hoja_sheets(sh, FOTOS_DATA_HOJA, FOTOS_DATA_COLUMNS)
    except Exception:
        return False
    todo_ok = True
    for nombre, ruta in fotos:
        try:
            with open(ruta, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")
            partes = [b64[i:i + FOTO_CHUNK] for i in range(0, len(b64), FOTO_CHUNK)]
            filas = [[str(reporte_id), nombre, i + 1, len(partes), p] for i, p in enumerate(partes)]
            # Una llamada por foto, para no mandar varios MB en un solo request.
            ws.append_rows(filas, value_input_option="RAW")
        except Exception:
            todo_ok = False
    _indice_fotos_sheets.clear()
    return todo_ok

@st.cache_data(ttl=300, show_spinner=False)
def _indice_fotos_sheets() -> dict:
    """{(ReporteID, NombreArchivo): {parte: n° de fila}} de la hoja FotosData. Lee solo las
    columnas chicas (A:C), no los datos de las fotos. Lanza excepción si no se puede leer
    (las excepciones no quedan en caché, así que se reintenta en la próxima llamada)."""
    sh = _cliente_sheets()
    if sh is None:
        return {}
    resp = sh.values_get(f"'{FOTOS_DATA_HOJA}'!A2:C", params={"valueRenderOption": "UNFORMATTED_VALUE"})
    indice = {}
    for i, fila in enumerate(resp.get("values", [])):
        if len(fila) < 3:
            continue
        clave = (str(fila[0]), str(fila[1]))
        # Si una foto se subió dos veces, gana la última subida.
        indice.setdefault(clave, {})[int(_num(fila[2]))] = i + 2
    return indice

def _restaurar_fotos_desde_sheets(pares: list):
    """Reconstruye en el disco local las fotos [(ReporteID, NombreArchivo)] que falten,
    descargándolas de FotosData (varias por llamada a la API)."""
    faltantes = [(str(r), str(n)) for r, n in pares if not os.path.exists(_ruta_foto_local(r, n))]
    sh = _cliente_sheets()
    if not faltantes or sh is None:
        return
    try:
        indice = _indice_fotos_sheets()
    except Exception:
        return
    faltantes = [p for p in dict.fromkeys(faltantes) if p in indice]
    for i in range(0, len(faltantes), 10):
        grupo = faltantes[i:i + 10]
        rangos, meta = [], []
        for r, n in grupo:
            partes = indice[(r, n)]
            filas = [partes[k] for k in sorted(partes)]
            meta.append((r, n, len(filas)))
            rangos += [f"'{FOTOS_DATA_HOJA}'!E{fila}" for fila in filas]
        try:
            resp = sh.values_batch_get(rangos, params={"valueRenderOption": "UNFORMATTED_VALUE"})
        except Exception:
            continue
        trozos = [str((vr.get("values") or [[""]])[0][0]) for vr in resp.get("valueRanges", [])]
        pos = 0
        for r, n, k in meta:
            b64 = "".join(trozos[pos:pos + k])
            pos += k
            try:
                datos = base64.b64decode(b64, validate=True)
            except Exception:
                continue
            ruta = _ruta_foto_local(r, n)
            os.makedirs(os.path.dirname(ruta), exist_ok=True)
            with open(ruta, "wb") as f:
                f.write(datos)

def subir_fotos_locales_faltantes() -> int:
    """Sube a FotosData las fotos que existen en el disco local pero todavía no están
    respaldadas (reportes guardados antes de este respaldo). Devuelve cuántas subió."""
    if _cliente_sheets() is None:
        return 0
    df_fot = _leer_hoja_df("Fotos")
    if df_fot.empty:
        return 0
    try:
        indice = _indice_fotos_sheets()
    except Exception:
        indice = {}
    por_reporte = {}
    for _, f in df_fot.iterrows():
        rid, nombre = str(f["ReporteID"]), str(f["NombreArchivo"])
        if (rid, nombre) in indice:
            continue
        ruta = _ubicar_foto(rid, nombre, f.get("RutaArchivo"))
        if ruta:
            por_reporte.setdefault(rid, []).append((nombre, ruta))
    n = 0
    for rid, fotos in por_reporte.items():
        if _respaldar_fotos_sheets(rid, fotos):
            n += len(fotos)
    return n

def migrate_reportes_excel():
    """Agrega hojas/columnas nuevas a un libro ya existente sin tocar los datos previos."""
    wb = openpyxl.load_workbook(REPORTES_XLSX_PATH)
    changed = False
    for sheet_name, headers in REPORTES_SHEETS.items():
        if sheet_name not in wb.sheetnames:
            ws = wb.create_sheet(sheet_name)
            ws.append(headers)
            changed = True
            continue
        ws = wb[sheet_name]
        header_row = [c.value for c in ws[1]] if ws.max_row >= 1 else []
        faltantes = [h for h in headers if h not in header_row]
        if faltantes:
            base_col = len(header_row)
            for i, h in enumerate(faltantes):
                ws.cell(row=1, column=base_col + i + 1, value=h)
            changed = True
    if changed:
        wb.save(REPORTES_XLSX_PATH)

def init_reportes_excel():
    if os.path.exists(REPORTES_XLSX_PATH):
        migrate_reportes_excel()
        return
    wb = openpyxl.Workbook()
    first_sheet = wb.active
    first = True
    for sheet_name, headers in REPORTES_SHEETS.items():
        ws = first_sheet if first else wb.create_sheet(sheet_name)
        if first:
            ws.title = sheet_name
            first = False
        ws.append(headers)
    wb.save(REPORTES_XLSX_PATH)

def next_reporte_id():
    """Próximo correlativo RPT-xxxxx. Con Sheets configurado se lee sin caché y en modo
    estricto: si Google no responde, lanza excepción en vez de calcular el número con el
    Excel local (que en la nube puede estar vacío y repetiría RPT-00001, pisando reportes)."""
    if _cliente_sheets() is not None:
        df = _leer_hoja_sheets_o_none("Reportes", REPORTES_SHEETS["Reportes"], estricto=True, fresco=True)
    else:
        df = _leer_hoja_df("Reportes")
    nums = []
    if not df.empty and "ReporteID" in df.columns:
        for val in df["ReporteID"].astype(str):
            try:
                nums.append(int(val.split("-")[1]))
            except Exception:
                pass
    n = (max(nums) + 1) if nums else 1
    return f"RPT-{n:05d}"

def obtener_actividades_hoy(usuario: str) -> list:
    """Actividades registradas en el/los Reporte(s) Diario(s) de hoy para este usuario."""
    df_rep = _leer_hoja_df("Reportes")
    df_trab = _leer_hoja_df("Trabajos")
    if df_rep.empty or df_trab.empty:
        return []
    hoy = datetime.now().strftime("%Y-%m-%d")
    reportes_hoy = df_rep[(df_rep["Fecha"].astype(str) == hoy) & (df_rep["Usuario"] == usuario)]["ReporteID"].tolist()
    if not reportes_hoy:
        return []
    return df_trab[df_trab["ReporteID"].isin(reportes_hoy)]["Actividad"].dropna().unique().tolist()

# -------------------------
# Planificación mensual (Plan vs Ejecutado)
# -------------------------
def guardar_plan_mensual(filas: list):
    """Reescribe por completo la hoja PlanMensual con las filas dadas. Es una tabla chica
    (actividades x meses), así que reemplazar todo al guardar es más simple y seguro que
    editar celdas puntuales, y evita filas duplicadas para la misma Actividad/Año/Mes."""
    headers = REPORTES_SHEETS["PlanMensual"]
    init_reportes_excel()
    wb = openpyxl.load_workbook(REPORTES_XLSX_PATH)
    if "PlanMensual" in wb.sheetnames:
        del wb["PlanMensual"]
    ws = wb.create_sheet("PlanMensual")
    ws.append(headers)
    for fila in filas:
        ws.append(fila)
    wb.save(REPORTES_XLSX_PATH)

    _sincronizar_tabla_sheets("PlanMensual", filas)

def ejecutado_por_rango(desde: date, hasta: date) -> dict:
    """Suma la Cantidad ejecutada por Actividad entre dos fechas (ambas incluidas), a partir
    de los Reportes Diarios reales (hoja Trabajos, cruzada con Reportes por Fecha)."""
    df_rep = _leer_hoja_df("Reportes")
    df_trab = _leer_hoja_df("Trabajos")
    if df_rep.empty or df_trab.empty:
        return {}
    # Fechas ilegibles quedan como NaT y se descartan (antes NaT hacía caer la comparación).
    fechas = pd.to_datetime(df_rep["Fecha"].map(_fecha_desde_serial).astype(str).str[:10], errors="coerce")
    en_rango = fechas.notna() & (fechas >= pd.Timestamp(desde)) & (fechas <= pd.Timestamp(hasta))
    reportes = set(df_rep[en_rango]["ReporteID"].astype(str))
    if not reportes:
        return {}
    trab = df_trab[df_trab["ReporteID"].astype(str).isin(reportes)].copy()
    trab["Cantidad"] = trab["Cantidad"].apply(_num)
    return trab.groupby("Actividad")["Cantidad"].sum().to_dict()

def ejecutado_por_actividad(anio: int, mes: int) -> dict:
    return ejecutado_por_rango(date(anio, mes, 1), date(anio, mes, calendar.monthrange(anio, mes)[1]))

def _plan_por_mes(df_plan: pd.DataFrame) -> dict:
    """{(año, mes): {actividad: cantidad planificada}}"""
    plan = {}
    for _, r in df_plan.iterrows():
        clave = (int(_num(r["Anio"])), int(_num(r["Mes"])))
        plan.setdefault(clave, {})
        plan[clave][r["Actividad"]] = plan[clave].get(r["Actividad"], 0.0) + _num(r["CantidadPlanificada"])
    return plan

def plan_por_rango(df_plan: pd.DataFrame, desde: date, hasta: date) -> dict:
    """Planificado por actividad entre dos fechas. El plan se define por mes; para una
    semana (o cualquier tramo) se prorratea por día: plan del mes / días del mes × días
    del tramo que caen en ese mes. Para un mes completo da exactamente el plan del mes."""
    if df_plan.empty or hasta < desde:
        return {}
    plan_mes = _plan_por_mes(df_plan)
    total = {}
    dia = desde
    while dia <= hasta:
        fin_mes = date(dia.year, dia.month, calendar.monthrange(dia.year, dia.month)[1])
        tramo_fin = min(fin_mes, hasta)
        fraccion = ((tramo_fin - dia).days + 1) / fin_mes.day
        for act, cant in plan_mes.get((dia.year, dia.month), {}).items():
            total[act] = total.get(act, 0.0) + cant * fraccion
        dia = tramo_fin + timedelta(days=1)
    return total

def semanas_del_mes(anio: int, mes: int) -> list:
    """Semanas (lunes a domingo) del mes, recortadas a los días del mes: [(inicio, fin)]."""
    primero = date(anio, mes, 1)
    ultimo = date(anio, mes, calendar.monthrange(anio, mes)[1])
    semanas = []
    inicio = primero
    while inicio <= ultimo:
        fin = min(inicio + timedelta(days=6 - inicio.weekday()), ultimo)
        semanas.append((inicio, fin))
        inicio = fin + timedelta(days=1)
    return semanas

def tabla_plan_vs_ejecutado(plan: dict, ejecutado: dict, unidades: dict, extra: dict | None = None) -> pd.DataFrame:
    """Una fila por actividad: Planificado, Ejecutado y Avance %. Incluye actividades
    ejecutadas que no estén en el catálogo (p.ej. nombres antiguos) para no esconderlas."""
    actividades = ACTIVIDADES_TRABAJO + [a for a in list(plan) + list(ejecutado) if a not in ACTIVIDADES_TRABAJO]
    filas = []
    for act in dict.fromkeys(actividades):
        plan_val = float(plan.get(act, 0.0))
        ejec_val = float(ejecutado.get(act, 0.0))
        avance = (ejec_val / plan_val * 100) if plan_val > 0 else (100.0 if ejec_val > 0 else 0.0)
        fila = {"Actividad": act, "Unidad": unidades.get(act) or "—"}
        for nombre, valores in (extra or {}).items():
            fila[nombre] = round(float(valores.get(act, 0.0)), 3)
        fila.update({"Planificado": round(plan_val, 3), "Ejecutado": round(ejec_val, 3), "Avance %": round(avance, 1)})
        filas.append(fila)
    return pd.DataFrame(filas)

def grafico_avance(df: pd.DataFrame):
    """Barras de Avance % por actividad (solo las que tienen plan o ejecución), con la
    línea del 100 %. Verde ≥ 100 %, amarillo ≥ 70 %, rojo < 70 %."""
    df = df[(df["Planificado"] > 0) | (df["Ejecutado"] > 0)].copy()
    if df.empty:
        return None
    df["Tramo"] = df["Avance %"].apply(lambda v: "≥ 100%" if v >= 100 else ("70–99%" if v >= 70 else "< 70%"))
    barras = alt.Chart(df).mark_bar().encode(
        x=alt.X("Avance %:Q", title="Avance %"),
        y=alt.Y("Actividad:N", sort="-x", title=None, axis=alt.Axis(labelLimit=180)),
        color=alt.Color("Tramo:N", scale=alt.Scale(domain=["≥ 100%", "70–99%", "< 70%"],
                                                   range=["#2FA84F", "#E0A458", "#DC3545"]),
                        legend=alt.Legend(title=None, orient="bottom")),
        tooltip=["Actividad:N", "Planificado:Q", "Ejecutado:Q", "Avance %:Q"],
    )
    regla = alt.Chart(pd.DataFrame({"x": [100]})).mark_rule(strokeDash=[4, 4], color="#0B3B8A").encode(x="x:Q")
    return (barras + regla).properties(height=max(120, 28 * len(df)))

def _ciclo_anual(anio: int, mes: int) -> list:
    """Los 12 (Año, Mes) del ciclo de mantención Septiembre-Agosto que contiene el mes
    dado (ej. Ene 2027 cae en el ciclo Sep 2026-Ago 2027, igual que el Excel del cliente)."""
    inicio_anio = anio if mes >= 9 else anio - 1
    return [(inicio_anio if m >= 9 else inicio_anio + 1, m) for m in [9, 10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8]]

def meta_anual_por_actividad(df_plan: pd.DataFrame, anio: int, mes: int) -> dict:
    """Suma la Cantidad Planificada de cada actividad en todo el ciclo Sep-Ago que
    contiene (anio, mes) -- la meta anual completa, no solo la del mes seleccionado."""
    if df_plan.empty:
        return {}
    ciclo = set(_ciclo_anual(anio, mes))
    en_ciclo = df_plan.apply(lambda r: (int(_num(r["Anio"])), int(_num(r["Mes"]))) in ciclo, axis=1)
    df_ciclo = df_plan[en_ciclo]
    if df_ciclo.empty:
        return {}
    return df_ciclo.groupby("Actividad")["CantidadPlanificada"].apply(lambda s: s.apply(_num).sum()).to_dict()

def _resolver_nombre(item: dict) -> str:
    if item["nombre"] == "Otro (especificar)":
        return item.get("nombre_custom", "").strip() or "Otro"
    return item["nombre"]

# Encabezados del resumen "plano" (una fila por actividad; Equipo/Materiales/Fotos-Observación/
# Asistencia van cada uno en una sola celda, con sus varios valores separados por coma).
FLAT_HEADERS = [
    "Grupo Vía", "Jefe de Grupo", "Fecha", "Ubicación", "Trabajo Realizado", "Collera Desde", "Collera Hasta",
    "Km Desde", "Km Hasta", "Unidad", "Cantidad", "Horas Trabajadas", "Equipo Utilizado", "Materiales",
    "Fotos / Observación", "Asistencia",
]
FLAT_COL_WIDTHS = [12, 18, 12, 25, 30, 10, 10, 10, 10, 8, 10, 14, 30, 30, 35, 35]

def _texto_equipos(equipos) -> str:
    return ", ".join(f"{e['nombre']} ({e['cantidad']:.0f})" for e in equipos) or "—"

def _texto_materiales(materiales) -> str:
    return ", ".join(
        f"{_resolver_nombre(m)} ({m['cantidad']:.0f}, {m.get('estado', '') or '—'})" for m in materiales
    ) or "—"

def _texto_fotos_obs(fotos, observaciones) -> str:
    partes = []
    if fotos:
        partes.append("Fotos: " + ", ".join(nombre for nombre, _ in fotos))
    if observaciones and str(observaciones).strip():
        partes.append(f"Obs: {observaciones}")
    return " | ".join(partes) or "—"

def _texto_asistencia(asistencia) -> str:
    return ", ".join(f"{a['nombre']} ({a['cargo']}) - {a['estado']}" for a in asistencia) or "—"

def _ajustar_anchos_columnas(ws, anchos):
    for i, ancho in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(i)].width = ancho

PDF_ANCHO_UTIL = A4[0] - 3 * cm  # ancho usable con márgenes de 1.5cm a cada lado

_ESTILO_CELDA_PDF = ParagraphStyle("celda_pdf", fontName="Helvetica", fontSize=8, leading=10, textColor=colors.HexColor("#222222"))
_ESTILO_CELDA_PDF_HEADER = ParagraphStyle("celda_pdf_header", fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=colors.white)

def _celda_pdf(valor, header=False) -> Paragraph:
    texto = "—" if valor in (None, "") else str(valor)
    return Paragraph(texto, _ESTILO_CELDA_PDF_HEADER if header else _ESTILO_CELDA_PDF)

def _tabla_pdf(headers, filas, col_widths):
    """Tabla con anchos de columna fijos y texto envuelto en Paragraph, para que nombres
    largos (actividad, material) se ajusten en vez de desbordar o descuadrar la tabla."""
    data = [[_celda_pdf(h, header=True) for h in headers]]
    if filas:
        for fila in filas:
            data.append([_celda_pdf(v) for v in fila])
    else:
        data.append([_celda_pdf("Sin registros")] + [_celda_pdf("") for _ in headers[1:]])
    t = Table(data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B3B8A")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#C7CCD4")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F1F3F6")]),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
    ]))
    return t

def _seccion_pdf(titulo: str) -> Table:
    """Barra de título de sección (celeste con texto azul), consistente en todo el PDF."""
    t = Table([[titulo]], colWidths=[PDF_ANCHO_UTIL])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#EEF2FA")),
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#0B3B8A")),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))
    return t

def _pie_pagina_pdf(canvas, doc, reporte_id: str):
    canvas.saveState()
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.HexColor("#8A8F98"))
    canvas.drawString(1.5 * cm, 0.9 * cm, f"Reporte Diario {reporte_id} · Generado {datetime.now().strftime('%d-%m-%Y %H:%M')}")
    canvas.drawRightString(A4[0] - 1.5 * cm, 0.9 * cm, f"Página {doc.page}")
    canvas.restoreState()

def generar_excel_reporte(resumen: dict) -> bytes:
    """Genera un Excel individual (no el libro maestro): una fila por actividad, con Equipo,
    Materiales, Fotos/Observación y Asistencia condensados cada uno en una sola celda."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reporte"

    equipo_txt = _texto_equipos(resumen["equipos"])
    material_txt = _texto_materiales(resumen["materiales"])
    fotos_obs_txt = _texto_fotos_obs(resumen["fotos"], resumen["observaciones"])
    asistencia_txt = _texto_asistencia(resumen["asistencia"])

    ws.append(FLAT_HEADERS)
    for t in resumen["trabajos"]:
        ws.append([
            resumen["grupo_via"], resumen["usuario"], resumen["fecha"], resumen.get("ubicacion") or "—",
            t["actividad"], t["collera_desde_id"], t["collera_hasta_id"], t["km_desde"], t["km_hasta"], t["unidad"], t["cantidad"], t["hh"],
            equipo_txt, material_txt, fotos_obs_txt, asistencia_txt,
        ])
    _ajustar_anchos_columnas(ws, FLAT_COL_WIDTHS)

    if resumen["fotos"]:
        fila = ws.max_row + 3
        ws.cell(row=fila - 1, column=1, value="FOTOGRAFÍAS")
        for nombre, ruta in resumen["fotos"]:
            try:
                img = XLImage(ruta)
                img.width, img.height = 260, 195
                ws.add_image(img, f"A{fila}")
                fila += 13
            except Exception:
                ws.cell(row=fila, column=1, value=f"(No se pudo incrustar: {nombre})")
                fila += 1

    if resumen.get("colleras"):
        hoja_excel_durmientes(wb, resumen["colleras"])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# Colores de la hoja 'Control Durmientes' del Excel original (formato condicional).
_RELLENO_ESTADO_DURM = {"Bueno": "C6EFCE", "Malo": "FFC7CE", "Nuevo": "BDD7EE", "Reemplazado": "E4DFEC"}
_RELLENO_ESTADO_GENERAL = {"Cumple": "C6EFCE", "No Cumple": "FFC7CE", "No Cumple (racha)": "FFC7CE",
                           "Inspección Incompleta": "FFEB9C", "Sin Inspeccionar": "D9D9D9"}
_LETRA_ESTADO = {"Bueno": "B", "Malo": "M", "Nuevo": "N", "Reemplazado": "R"}

def _secuencia_en_letras(secuencia: list) -> str:
    """'BMMMB MMBBB ...' en grupos de 5, para que el PDF pueda cortar la línea."""
    letras = "".join(_LETRA_ESTADO.get(e, "-") for e in secuencia)
    return " ".join(letras[i:i + 5] for i in range(0, len(letras), 5))

def hoja_excel_durmientes(wb, colleras: list, con_reporte: bool = False):
    """Agrega una hoja 'Durmientes' con el mismo formato que 'Control Durmientes' del Excel
    (una fila por collera: D1..D23 + columnas calculadas). `colleras` = lista de
    resumen_collera(); con_reporte=True agrega ReporteID y Fecha al inicio (histórico)."""
    ws = wb.create_sheet("Durmientes")
    base = ["REPORTE", "FECHA"] if con_reporte else []
    headers = base + ["PK", "COLLERA", "UBICACIÓN"] + [f"D{p}" for p in range(1, 24)] + [
        "TOTAL REGISTRADO", "N° BUENOS", "N° MALOS", "N° NUEVOS", "N° REEMPL.", "EFECTIVOS (B+N+R)",
        "% RENOV. (N+R)", "MÍN. EFECTIVOS", "RACHA CONSEC.", "ESTADO GENERAL"]
    ws.append(headers)
    for celda in ws[1]:
        celda.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
        celda.fill = openpyxl.styles.PatternFill("solid", fgColor="0B3B8A")
        celda.alignment = openpyxl.styles.Alignment(wrap_text=True, horizontal="center", vertical="center")
    col_d1 = len(base) + 4
    for c in colleras:
        ws.append(([c.get("reporte_id", ""), c.get("fecha", "")] if con_reporte else [])
                  + [c["pk"], c["collera"], c["ubicacion"]] + [e or "" for e in c["secuencia"]]
                  + [c["total_registrado"], c["n_buenos"], c["n_malos"], c["n_nuevos"], c["n_reempl"],
                     c["efectivos"], c["pct_renovacion"], c["min_efectivos"], c["racha_consec"], c["estado_general"]])
        fila = ws.max_row
        for i, e in enumerate(c["secuencia"]):
            if e in _RELLENO_ESTADO_DURM:
                ws.cell(row=fila, column=col_d1 + i).fill = openpyxl.styles.PatternFill("solid", fgColor=_RELLENO_ESTADO_DURM[e])
        ws.cell(row=fila, column=col_d1 + 29).number_format = "0.0%"
        celda_estado = ws.cell(row=fila, column=col_d1 + 32)
        if c["estado_general"] in _RELLENO_ESTADO_GENERAL:
            celda_estado.fill = openpyxl.styles.PatternFill("solid", fgColor=_RELLENO_ESTADO_GENERAL[c["estado_general"]])
    anchos = ([11, 11] if con_reporte else []) + [6, 9, 22] + [11] * 23 + [11, 9, 9, 9, 9, 11, 10, 11, 10, 20]
    _ajustar_anchos_columnas(ws, anchos)
    ws.freeze_panes = ws.cell(row=2, column=col_d1)

def colleras_desde_inspecciones(df_insp: pd.DataFrame) -> list:
    """Convierte filas de la hoja InspeccionColleras en la lista de resumen_collera()."""
    colleras = []
    for _, r in df_insp.iterrows():
        secuencia = [(r.get(f"D{p}") or None) if isinstance(r.get(f"D{p}"), str) else None for p in range(1, 24)]
        info = {"ColleraID": str(r["ColleraID"]), "PK": int(_num(r["PK"])), "Collera": int(_num(r["Collera"])),
                "Ubicacion": str(r["Ubicacion"])}
        colleras.append({**resumen_collera(info, secuencia), "reporte_id": str(r["ReporteID"]), "fecha": str(r["Fecha"])})
    return colleras

def generar_pdf_reporte(resumen: dict) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=1.5 * cm, bottomMargin=1.8 * cm,
                             leftMargin=1.5 * cm, rightMargin=1.5 * cm)
    styles = getSampleStyleSheet()
    el = []

    # Encabezado con color de marca (en vez del título plano de antes)
    estilo_banner = ParagraphStyle("banner_pdf", fontName="Helvetica-Bold", fontSize=14, textColor=colors.white)
    banner = Table([[Paragraph(f"REPORTE DIARIO — {resumen['reporte_id']}", estilo_banner)]], colWidths=[PDF_ANCHO_UTIL])
    banner.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#0B3B8A")),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
    ]))
    el.append(banner)
    el.append(Spacer(1, 10))

    # Ficha de datos generales (etiqueta/valor en grilla, no un párrafo corrido)
    estilo_etiqueta = ParagraphStyle("etiqueta_pdf", fontName="Helvetica-Bold", fontSize=9, textColor=colors.HexColor("#0B3B8A"))
    estilo_valor = ParagraphStyle("valor_pdf", fontName="Helvetica", fontSize=9, textColor=colors.HexColor("#222222"))
    mitad = PDF_ANCHO_UTIL / 2 - 85
    info_data = [
        [Paragraph("Grupo Vía", estilo_etiqueta), Paragraph(str(resumen["grupo_via"]), estilo_valor),
         Paragraph("Fecha", estilo_etiqueta), Paragraph(str(resumen["fecha"]), estilo_valor)],
        [Paragraph("Jefe de Grupo", estilo_etiqueta), Paragraph(str(resumen["usuario"]), estilo_valor),
         Paragraph("Jornada", estilo_etiqueta), Paragraph(f"{resumen['horas_dia']} h/trabajador", estilo_valor)],
        [Paragraph("Ubicación", estilo_etiqueta), Paragraph(str(resumen.get("ubicacion") or "—"), estilo_valor), "", ""],
    ]
    info_table = Table(info_data, colWidths=[85, mitad, 85, mitad])
    info_table.setStyle(TableStyle([
        ("SPAN", (1, 2), (3, 2)),
        ("LINEBELOW", (0, 0), (-1, -2), 0.4, colors.HexColor("#E3E5E8")),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    el.append(info_table)
    el.append(Spacer(1, 14))

    el.append(_seccion_pdf("Trabajos"))
    el.append(Spacer(1, 4))
    headers_trab = ["Actividad", "Collera Desde", "Collera Hasta", "Km Desde", "Km Hasta", "Unidad", "Cant.", "N° Trab.", "HH"]
    anchos_trab = [130, 58, 58, 42, 42, 42, 38, 42, 38]
    filas_trab = [
        [t["actividad"], t["collera_desde_id"], t["collera_hasta_id"], t["km_desde"], t["km_hasta"], t["unidad"], t["cantidad"], t["hombres"], t["hh"]]
        for t in resumen["trabajos"]
    ]
    el.append(_tabla_pdf(headers_trab, filas_trab, anchos_trab))
    el.append(Spacer(1, 12))

    if resumen.get("colleras"):
        el.append(_seccion_pdf("Control de Durmientes por Collera (Norma NS-01-01-00)"))
        el.append(Spacer(1, 4))
        headers_col = ["Collera", "Ubicación", "D1 → D23", "B", "M", "N", "R", "Efect.", "% Renov.", "Mín. Efect.", "Racha", "Estado"]
        anchos_col = [40, 52, 112, 20, 20, 20, 20, 32, 38, 44, 36, 76]
        filas_col = [
            [c["collera_id"], c["ubicacion"], _secuencia_en_letras(c["secuencia"]),
             c["n_buenos"], c["n_malos"], c["n_nuevos"], c["n_reempl"], c["efectivos"],
             f"{c['pct_renovacion'] * 100:.1f}%", c["min_efectivos"], c["racha_consec"], c["estado_general"]]
            for c in resumen["colleras"]
        ]
        el.append(_tabla_pdf(headers_col, filas_col, anchos_col))
        el.append(Paragraph("B = Bueno · M = Malo · N = Nuevo · R = Reemplazado · - = sin inspeccionar", _ESTILO_CELDA_PDF))
        el.append(Spacer(1, 12))

    el.append(_seccion_pdf("Equipos"))
    el.append(Spacer(1, 4))
    filas_equ = [[e["nombre"], e["cantidad"]] for e in resumen["equipos"]]
    el.append(_tabla_pdf(["Equipo", "Cantidad"], filas_equ, [PDF_ANCHO_UTIL - 90, 90]))
    el.append(Spacer(1, 12))

    el.append(_seccion_pdf("Materiales"))
    el.append(Spacer(1, 4))
    filas_mat = [[_resolver_nombre(m), m["cantidad"], m.get("estado", "")] for m in resumen["materiales"]]
    el.append(_tabla_pdf(["Material", "Cantidad", "Estado"], filas_mat, [PDF_ANCHO_UTIL - 190, 80, 110]))
    el.append(Spacer(1, 12))

    el.append(_seccion_pdf("Observaciones"))
    el.append(Spacer(1, 6))
    el.append(Paragraph(resumen["observaciones"] or "Sin observaciones.", styles["Normal"]))
    el.append(Spacer(1, 12))

    el.append(_seccion_pdf("Asistencia"))
    el.append(Spacer(1, 4))
    filas_asi = [[a["nombre"], a["cargo"], a["estado"]] for a in resumen["asistencia"]]
    el.append(_tabla_pdf(["Trabajador", "Cargo", "Estado"], filas_asi, [PDF_ANCHO_UTIL - 220, 130, 90]))

    if resumen["fotos"]:
        el.append(Spacer(1, 12))
        el.append(_seccion_pdf("Fotografías"))
        el.append(Spacer(1, 8))
        for nombre, ruta in resumen["fotos"]:
            try:
                with PILImage.open(ruta) as im:
                    w, h = im.size
                ancho = 13 * cm
                alto = ancho * h / w
                el.append(RLImage(ruta, width=ancho, height=alto))
                el.append(Spacer(1, 8))
            except Exception:
                el.append(Paragraph(f"(No se pudo incrustar: {nombre})", styles["Normal"]))

    pie_pagina = lambda c, d: _pie_pagina_pdf(c, d, resumen["reporte_id"])
    doc.build(el, onFirstPage=pie_pagina, onLaterPages=pie_pagina)
    return buf.getvalue()

def generar_documentos_reporte(resumen: dict):
    """Genera el Excel y PDF individuales. Si una foto hace fallar la incrustación
    (formatos raros de cámara que a veces se cuelan), reintenta sin fotos para que
    el reporte —ya guardado en el libro maestro— nunca se pierda por esto."""
    try:
        excel_bytes = generar_excel_reporte(resumen)
    except Exception:
        excel_bytes = generar_excel_reporte({**resumen, "fotos": []})
    try:
        pdf_bytes = generar_pdf_reporte(resumen)
    except Exception:
        pdf_bytes = generar_pdf_reporte({**resumen, "fotos": []})
    return excel_bytes, pdf_bytes

def _filas_respaldo_plano() -> list:
    """Arma las filas planas del histórico completo (una fila por actividad, con Equipo/
    Materiales/Fotos-Observación/Asistencia condensados cada uno en una sola celda) --
    compartido entre el Excel y el PDF de respaldo, para no duplicar la lógica."""
    df_rep = _leer_hoja_df("Reportes")
    df_trab = _leer_hoja_df("Trabajos")
    df_equ = _leer_hoja_df("Equipos")
    df_mat = _leer_hoja_df("Materiales")
    df_asi = _leer_hoja_df("Asistencia")
    df_fot = _leer_hoja_df("Fotos")
    if df_rep.empty:
        return []

    filas = []
    for _, rep in df_rep.iterrows():
        rid = rep["ReporteID"]
        trabajos_rep = df_trab[df_trab["ReporteID"] == rid]
        if trabajos_rep.empty:
            continue

        equipo_txt = ", ".join(
            f"{r['Equipo']} ({_num(r['Cantidad']):.0f})" for _, r in df_equ[df_equ["ReporteID"] == rid].iterrows()
        ) or "—"
        material_txt = ", ".join(
            f"{r['Material']} ({_num(r['Cantidad']):.0f}, {r.get('Estado') or '—'})"
            for _, r in df_mat[df_mat["ReporteID"] == rid].iterrows()
        ) or "—"
        fotos_lista = df_fot[df_fot["ReporteID"] == rid]["NombreArchivo"].tolist()
        partes_fo = []
        if fotos_lista:
            partes_fo.append("Fotos: " + ", ".join(fotos_lista))
        obs = rep.get("Observaciones")
        if isinstance(obs, str) and obs.strip():
            partes_fo.append(f"Obs: {obs}")
        fotos_obs_txt = " | ".join(partes_fo) or "—"
        asistencia_txt = ", ".join(
            f"{r['Trabajador']} ({r['Cargo']}) - {r['Estado']}" for _, r in df_asi[df_asi["ReporteID"] == rid].iterrows()
        ) or "—"

        ubicacion = rep.get("Ubicacion")
        ubicacion_txt = ubicacion if isinstance(ubicacion, str) and ubicacion.strip() else "—"

        for _, t in trabajos_rep.iterrows():
            filas.append([
                rep["GrupoVia"], rep["Usuario"], rep["Fecha"], ubicacion_txt,
                t["Actividad"], t.get("ColleraDesdeID", ""), t.get("ColleraHastaID", ""), t["KmDesde"], t["KmHasta"], t["Unidad"], t["Cantidad"], t["HH"],
                equipo_txt, material_txt, fotos_obs_txt, asistencia_txt,
            ])
    return filas

def generar_respaldo_plano() -> bytes | None:
    """Resumen legible de TODO el histórico en Excel. El libro maestro (reportes_diarios.xlsx)
    sigue normalizado en hojas separadas para Power BI; esto es solo una vista aparte para leer rápido."""
    filas = _filas_respaldo_plano()
    if not filas:
        return None

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumen"
    ws.append(FLAT_HEADERS)
    for fila in filas:
        ws.append(fila)
    _ajustar_anchos_columnas(ws, FLAT_COL_WIDTHS)

    # Todas las inspecciones de colleras hechas en Reportes Diarios, en el formato del Excel de control.
    df_insp = _leer_hoja_df("InspeccionColleras")
    if not df_insp.empty:
        hoja_excel_durmientes(wb, colleras_desde_inspecciones(df_insp), con_reporte=True)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def generar_respaldo_pdf() -> bytes | None:
    """Misma información que generar_respaldo_plano() pero en PDF (apaisado, para que
    quepan las 16 columnas) -- entrega solicitada además del Excel."""
    filas = _filas_respaldo_plano()
    if not filas:
        return None

    ancho_pagina = landscape(A4)
    ancho_util = ancho_pagina[0] - 2.4 * cm
    anchos = [32, 50, 42, 50, 72, 38, 38, 32, 32, 28, 28, 28, 58, 58, 58, 58]  # suma ~704, cabe en ~704pt

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=ancho_pagina, topMargin=1.2 * cm, bottomMargin=1.6 * cm,
                             leftMargin=1.2 * cm, rightMargin=1.2 * cm)

    estilo_banner = ParagraphStyle("banner_resp", fontName="Helvetica-Bold", fontSize=13, textColor=colors.white)
    banner = Table([[Paragraph("RESPALDO HISTÓRICO — REPORTES DIARIOS", estilo_banner)]], colWidths=[ancho_util])
    banner.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#0B3B8A")),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
    ]))
    estilo_sub = ParagraphStyle("sub_resp", fontName="Helvetica", fontSize=8, textColor=colors.HexColor("#6c757d"))
    el = [
        banner,
        Spacer(1, 8),
        Paragraph(f"{len(filas)} actividades registradas · Generado {datetime.now().strftime('%d-%m-%Y %H:%M')}", estilo_sub),
        Spacer(1, 10),
        _tabla_pdf(FLAT_HEADERS, filas, anchos),
    ]

    def _pie(c, d):
        c.saveState()
        c.setFont("Helvetica", 7)
        c.setFillColor(colors.HexColor("#8A8F98"))
        c.drawString(1.2 * cm, 0.9 * cm, f"Respaldo histórico · Generado {datetime.now().strftime('%d-%m-%Y %H:%M')}")
        c.drawRightString(ancho_pagina[0] - 1.2 * cm, 0.9 * cm, f"Página {d.page}")
        c.restoreState()

    doc.build(el, onFirstPage=_pie, onLaterPages=_pie)
    return buf.getvalue()

def generar_respaldo_fotos_zip() -> bytes | None:
    """ZIP con todas las fotos del histórico, una carpeta por ReporteID. Las fotos que ya
    no están en el disco local (reinicio de Streamlit Cloud) se recuperan desde Sheets."""
    df_fot = _leer_hoja_df("Fotos")
    if df_fot.empty:
        return None
    registros = [(str(f["ReporteID"]), str(f["NombreArchivo"]), f.get("RutaArchivo")) for _, f in df_fot.iterrows()]
    _restaurar_fotos_desde_sheets([(rid, nombre) for rid, nombre, _ in registros])
    buf = io.BytesIO()
    n = 0
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as zf:  # JPEG ya viene comprimido
        for rid, nombre, ruta_original in dict.fromkeys(registros):
            ruta = _ubicar_foto(rid, nombre, ruta_original)
            if ruta:
                zf.write(ruta, arcname=f"{rid}/{nombre}")
                n += 1
    return buf.getvalue() if n else None

def resumen_desde_historico(reporte_id: str) -> dict | None:
    """Arma el mismo 'resumen' que produce guardar_reporte(), pero leyendo un reporte ya
    guardado desde Sheets/Excel -- para volver a generar su Excel y PDF (con fotos)."""
    df_rep = _leer_hoja_df("Reportes")
    if df_rep.empty:
        return None
    fila_rep = df_rep[df_rep["ReporteID"].astype(str) == str(reporte_id)]
    if fila_rep.empty:
        return None
    rep = fila_rep.iloc[0]

    def _de(nombre_hoja):
        df = _leer_hoja_df(nombre_hoja)
        return df[df["ReporteID"].astype(str) == str(reporte_id)] if not df.empty else df

    def _txt(v):
        return "" if v is None or (isinstance(v, float) and pd.isna(v)) else v

    trabajos = [
        {"actividad": t["Actividad"], "collera_desde_id": _txt(t.get("ColleraDesdeID")),
         "collera_hasta_id": _txt(t.get("ColleraHastaID")), "km_desde": _txt(t["KmDesde"]),
         "km_hasta": _txt(t["KmHasta"]), "unidad": _txt(t["Unidad"]), "cantidad": _num(t["Cantidad"]),
         "hombres": int(_num(t["Hombres"])), "hh": _num(t["HH"])}
        for _, t in _de("Trabajos").iterrows()
    ]
    equipos = [{"nombre": e["Equipo"], "cantidad": _num(e["Cantidad"])} for _, e in _de("Equipos").iterrows()]
    materiales = [
        {"nombre": m["Material"], "cantidad": _num(m["Cantidad"]), "estado": _txt(m.get("Estado"))}
        for _, m in _de("Materiales").iterrows()
    ]
    asistencia = [
        {"nombre": a["Trabajador"], "cargo": a["Cargo"], "estado": a["Estado"]}
        for _, a in _de("Asistencia").iterrows()
    ]
    df_fot = _de("Fotos")
    _restaurar_fotos_desde_sheets([(reporte_id, n) for n in df_fot["NombreArchivo"].astype(str)] if not df_fot.empty else [])
    fotos = []
    for _, f in df_fot.iterrows():
        ruta = _ubicar_foto(reporte_id, f["NombreArchivo"], f.get("RutaArchivo"))
        if ruta:
            fotos.append((str(f["NombreArchivo"]), ruta))

    fecha = str(rep["Fecha"])[:10]
    try:
        horas_dia = horas_jornada_por_fecha(datetime.strptime(fecha, "%Y-%m-%d"))
    except ValueError:
        horas_dia = "—"
    return {
        "reporte_id": str(reporte_id), "fecha": fecha, "grupo_via": _txt(rep["GrupoVia"]),
        "usuario": _txt(rep["Usuario"]), "horas_dia": horas_dia,
        "ubicacion": _txt(rep.get("Ubicacion")), "observaciones": _txt(rep.get("Observaciones")),
        "trabajos": trabajos, "equipos": equipos, "materiales": materiales,
        "asistencia": asistencia, "fotos": fotos,
        "colleras": colleras_desde_inspecciones(_de("InspeccionColleras")),
        "total_hh": sum(t["hh"] for t in trabajos),
    }

def mostrar_detalle_reporte_diario(reporte_id: str):
    """Vista de solo lectura de un Reporte Diario, para que el validador vea exactamente
    lo que ingresó Personal de Terreno (sin poder editarlo desde acá)."""
    df_rep = _leer_hoja_df("Reportes")
    df_trab = _leer_hoja_df("Trabajos")
    df_equ = _leer_hoja_df("Equipos")
    df_mat = _leer_hoja_df("Materiales")
    df_asi = _leer_hoja_df("Asistencia")
    df_fot = _leer_hoja_df("Fotos")
    if df_rep.empty:
        st.info("No se encontró el archivo de Reportes Diarios.")
        return

    fila_rep = df_rep[df_rep["ReporteID"] == reporte_id]
    if fila_rep.empty:
        st.info("Este Reporte Diario ya no está disponible (puede haberse limpiado el respaldo).")
        return
    rep = fila_rep.iloc[0]

    st.caption(f"📄 Reporte {reporte_id} · Grupo Vía {rep['GrupoVia']} · {rep['Fecha']} · Jefe de Grupo: {rep['Usuario']}")
    ubicacion = rep.get("Ubicacion")
    if isinstance(ubicacion, str) and ubicacion.strip():
        st.caption(f"📍 Ubicación: {ubicacion}")

    st.markdown("**🛠️ Trabajos**")
    trabajos_rep = df_trab[df_trab["ReporteID"] == reporte_id].drop(columns=["ReporteID"])
    if not trabajos_rep.empty:
        st.dataframe(trabajos_rep, hide_index=True, width="stretch")
    else:
        st.caption("Sin actividades registradas.")

    df_insp = _leer_hoja_df("InspeccionColleras")
    insp_rep = df_insp[df_insp["ReporteID"].astype(str) == str(reporte_id)] if not df_insp.empty else df_insp
    if not insp_rep.empty:
        st.markdown("**🛤️ Control de Durmientes por Collera**")
        for c in colleras_desde_inspecciones(insp_rep):
            with st.container(border=True):
                st.markdown(f"**Collera {c['collera_id']}** · {c['ubicacion']}")
                st.markdown(franja_durmientes(c["secuencia"]))
                badge(c["estado_general"], TIPO_BADGE_ESTADO.get(c["estado_general"], "info"))
                st.dataframe(tabla_indicadores_collera(c), hide_index=True, width="stretch")
        st.caption("🟩 Bueno · 🟥 Malo · 🟦 Nuevo · 🟪 Reemplazado · ⬜ sin inspeccionar")

    st.markdown("**🚜 Equipos**")
    equipos_rep = df_equ[df_equ["ReporteID"] == reporte_id].drop(columns=["ReporteID"])
    if not equipos_rep.empty:
        st.dataframe(equipos_rep, hide_index=True, width="stretch")
    else:
        st.caption("Sin equipos registrados.")

    st.markdown("**📦 Materiales**")
    materiales_rep = df_mat[df_mat["ReporteID"] == reporte_id].drop(columns=["ReporteID"])
    if not materiales_rep.empty:
        st.dataframe(materiales_rep, hide_index=True, width="stretch")
    else:
        st.caption("Sin materiales registrados.")

    st.markdown("**👷 Asistencia**")
    asistencia_rep = df_asi[df_asi["ReporteID"] == reporte_id]
    if not asistencia_rep.empty:
        st.dataframe(asistencia_rep[["Trabajador", "Cargo", "Estado"]], hide_index=True, width="stretch")
    else:
        st.caption("Sin asistencia registrada.")

    st.markdown("**📝 Observaciones**")
    obs = rep.get("Observaciones")
    st.write(obs if isinstance(obs, str) and obs.strip() else "—")

    st.markdown("**📷 Fotografías**")
    fotos_rep = df_fot[df_fot["ReporteID"] == reporte_id]
    if not fotos_rep.empty:
        _restaurar_fotos_desde_sheets([(reporte_id, n) for n in fotos_rep["NombreArchivo"].astype(str)])
        cols = st.columns(2)
        for i, (_, f) in enumerate(fotos_rep.iterrows()):
            ruta = _ubicar_foto(reporte_id, f["NombreArchivo"], f["RutaArchivo"])
            with cols[i % 2]:
                if ruta:
                    st.image(ruta, caption=f["NombreArchivo"], width="stretch")
                else:
                    st.caption(f"(No disponible: {f['NombreArchivo']})")
    else:
        st.caption("Sin fotografías.")

    st.markdown("**⬇️ Descargar este reporte**")
    clave_docs = f"docs_historico_{reporte_id}"
    if st.button("📄 Preparar Excel y PDF", key=f"btn_prep_docs_{reporte_id}", width="stretch"):
        with st.spinner("Generando documentos..."):
            resumen = resumen_desde_historico(reporte_id)
            st.session_state[clave_docs] = generar_documentos_reporte(resumen) if resumen else None
    docs = st.session_state.get(clave_docs)
    if docs:
        excel_bytes, pdf_bytes = docs
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "⬇️ Excel", data=excel_bytes, file_name=f"Reporte_{reporte_id}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"btn_dl_excel_{reporte_id}", width="stretch",
            )
        with d2:
            st.download_button(
                "⬇️ PDF", data=pdf_bytes, file_name=f"Reporte_{reporte_id}.pdf", mime="application/pdf",
                key=f"btn_dl_pdf_{reporte_id}", width="stretch",
            )

# -------------------------
# Perfiles / roles
# -------------------------
ROLES_VALIDADORES = ["Sacyr", "ADI (ITO)", "EFE", "Supervisor Subcontrato"]  # también son niveles de validación, en orden
PERFILES = ["Personal Terreno"] + ROLES_VALIDADORES

init_db()

def init_data():
    if "perfil" not in st.session_state:
        st.session_state.perfil = None
    if "usuario" not in st.session_state:
        st.session_state.usuario = ""
    if "login_form_counter" not in st.session_state:
        st.session_state.login_form_counter = 0
    if "page" not in st.session_state:
        st.session_state.page = "Inicio"
    if "activos" not in st.session_state:
        st.session_state.activos = load_activos()
    # Avisos, OTs y durmientes se refrescan en cada recarga (desde la caché de Sheets, 60 s),
    # para que se vea lo que guardaron otros usuarios sin tener que volver a iniciar sesión.
    # estricto=True: si Sheets falla en un refresco, se conserva lo que ya había en la sesión
    # en vez de caer al SQLite local, que en la nube puede estar vacío o desactualizado.
    for clave, cargador in (("avisos", load_avisos), ("ots", load_ots), ("durmientes_estado", load_durmientes_estado)):
        if clave in st.session_state:
            try:
                st.session_state[clave] = cargador(estricto=True)
            except Exception:
                pass
    if "avisos" not in st.session_state:
        st.session_state.avisos = load_avisos()
    if "ots" not in st.session_state:
        st.session_state.ots = load_ots()
    if "durmientes_estado" not in st.session_state:
        df_estado = load_durmientes_estado()
        if df_estado.empty:
            df_estado = _seed_durmientes_estado_inicial()
            st.session_state.durmientes_estado = df_estado
            if not df_estado.empty:
                save_all_durmientes_estado()
        else:
            st.session_state.durmientes_estado = df_estado
    if "selected_aviso" not in st.session_state:
        st.session_state.selected_aviso = None
    if "reporte_trabajos" not in st.session_state:
        st.session_state.reporte_trabajos = []
    if "reporte_colleras" not in st.session_state:
        st.session_state.reporte_colleras = []
    if "reporte_equipos_catalogo" not in st.session_state:
        st.session_state.reporte_equipos_catalogo = {nombre: {"usado": False, "cantidad": 1.0} for nombre in EQUIPOS_CATALOGO}
    if "reporte_equipos_otros" not in st.session_state:
        st.session_state.reporte_equipos_otros = []
    if "equipo_otro_form_counter" not in st.session_state:
        st.session_state.equipo_otro_form_counter = 0
    if "reporte_materiales" not in st.session_state:
        st.session_state.reporte_materiales = []
    if "reporte_asistencia" not in st.session_state:
        st.session_state.reporte_asistencia = []
    if "asistencia_form_counter" not in st.session_state:
        st.session_state.asistencia_form_counter = 0
    if "reporte_fotos_counter" not in st.session_state:
        st.session_state.reporte_fotos_counter = 0
    if "rep_ubicacion_gps" not in st.session_state:
        st.session_state.rep_ubicacion_gps = ""
    if "ubicacion_texto_counter" not in st.session_state:
        st.session_state.ubicacion_texto_counter = 0
    if "aviso_form_counter" not in st.session_state:
        st.session_state.aviso_form_counter = 0
    if "last_reporte" not in st.session_state:
        st.session_state.last_reporte = None
    if "last_reporte_excel" not in st.session_state:
        st.session_state.last_reporte_excel = None
    if "last_reporte_pdf" not in st.session_state:
        st.session_state.last_reporte_pdf = None

init_data()

def next_aviso_id():
    df = st.session_state.avisos
    nums = []
    for x in df["AvisoID"].astype(str).tolist():
        try:
            nums.append(int(x.split("-")[1]))
        except Exception:
            pass
    n = (max(nums) + 1) if nums else 1
    return f"EFE-{n:05d}"

def next_ot_id():
    df = st.session_state.ots
    nums = []
    for x in df["OTID"].astype(str).tolist():
        try:
            nums.append(int(x.split("-")[1]))
        except Exception:
            pass
    n = (max(nums) + 1) if nums else 1
    return f"OT-{n:05d}"

def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M")

def new_row_id():
    return uuid.uuid4().hex[:8]

def new_trabajo_row():
    return {"id": new_row_id(), "actividad": ACTIVIDADES_TRABAJO[0], "pk": 33,
            "collera_desde_id": "", "collera_hasta_id": "", "km_desde": "", "km_hasta": "",
            "unidad": UNIDADES_TRABAJO[0], "cantidad": 0.0, "hombres": 0}

def new_material_row():
    return {"id": new_row_id(), "nombre": MATERIALES_CATALOGO[0], "nombre_custom": "", "cantidad": 0.0,
            "estado": ESTADOS_MATERIAL[0]}

def reset_reporte_form():
    st.session_state.reporte_trabajos = []
    st.session_state.reporte_colleras = []
    st.session_state.reporte_equipos_catalogo = {nombre: {"usado": False, "cantidad": 1.0} for nombre in EQUIPOS_CATALOGO}
    st.session_state.reporte_equipos_otros = []
    st.session_state.reporte_materiales = []
    st.session_state.reporte_asistencia = []
    st.session_state.reporte_fotos_counter += 1
    st.session_state.rep_ubicacion_gps = ""
    st.session_state.ubicacion_texto_counter += 1

def elapsed_str(fecha_hora_str: str) -> str:
    try:
        dt = datetime.strptime(fecha_hora_str, "%Y-%m-%d %H:%M")
    except Exception:
        return ""
    delta = datetime.now() - dt
    total_min = max(int(delta.total_seconds() // 60), 0)
    h, m = divmod(total_min, 60)
    return f"{h}hr{m}" if h > 0 else f"{m}min"

# -------------------------
# UI helpers
# -------------------------
def _slug(label: str) -> str:
    return label.replace(" ", "_").replace("(", "").replace(")", "")

def badge(text, kind="info"):
    colors = {"ok": "#198754", "warn": "#E0A458", "bad": "#dc3545", "info": "#0d6efd", "muted": "#6c757d"}
    bg = colors.get(kind, "#0d6efd")
    st.markdown(f"<span class='pill' style='background:{bg};color:white;'>{text}</span>", unsafe_allow_html=True)

def priority_kind(p):
    if p in ["Alta", "Crítica"]:
        return "bad"
    if p == "Media":
        return "warn"
    return "ok"

def app_header(title: str, back_page: str | None = None, right_icon: str = ""):
    if back_page:
        with st.container(key="header_row"):
            col1, col2 = st.columns([1, 5])
            with col1:
                if st.button("←", key="btn_back"):
                    st.session_state.page = back_page
                    st.rerun()
            with col2:
                st.markdown(
                    f"""<div class="app-header header-right">
                            <div class="title">{title}</div>
                            <div class="right-icon">{right_icon}</div>
                        </div>""",
                    unsafe_allow_html=True,
                )
    else:
        st.markdown(
            f"""<div class="app-header header-solo">
                    <div class="title">{title}</div>
                    <div class="right-icon">{right_icon}</div>
                </div>""",
            unsafe_allow_html=True,
        )
    st.write("")

def perfil_bar():
    c1, c2 = st.columns([4, 1])
    with c1:
        st.caption(f"Perfil: **{st.session_state.perfil}** · {st.session_state.usuario}")
    with c2:
        if st.button("Salir", key="btn_salir"):
            st.session_state.perfil = None
            st.session_state.page = "Inicio"
            st.rerun()

def render_pie_desarrollador():
    st.markdown(
        """
        <div class="pie-dev">
            <div class="pie-dev-nombre">👨‍💻 Desarrollado por Lukas Dyango Sanhueza Canales</div>
            <div class="pie-dev-contacto">
                <a href="mailto:lukassanhueza2000@gmail.com">lukassanhueza2000@gmail.com</a> · +56 9 2722 1163
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.link_button(
        "💬 Contactar por WhatsApp",
        "https://wa.me/56927221163",
        key="btn_footer_whatsapp", width="stretch",
    )

def bottom_nav(items):
    active_slug = _slug(st.session_state.page)
    st.markdown(
        f"<style>.st-key-navbtn_{active_slug} button {{ color:#0B3B8A !important; }}</style>",
        unsafe_allow_html=True,
    )
    with st.container(key="bottom_nav"):
        st.markdown("<div class='nav-brand'>SACYR</div>", unsafe_allow_html=True)
        cols = st.columns(len(items))
        for col, (label, icon) in zip(cols, items):
            with col:
                if st.button(f"{icon} {label}", key=f"navbtn_{_slug(label)}"):
                    st.session_state.page = label
                    st.rerun()

def can_validate(role: str, aviso_row: dict) -> bool:
    if aviso_row["Estado"] in ["Rechazado", "Cerrado", "Validado", "OT creada"]:
        return False
    return aviso_row["ValidacionNivel"] == role

def final_cols(rol_final: str):
    """Columnas del tercer nivel de validación, según si el aviso lo cierra EFE o el Supervisor Subcontrato."""
    if rol_final == "Supervisor Subcontrato":
        return "AprobadoSubcontrato", "ObsSubcontrato", "FechaValSubcontrato", "RespValSubcontrato"
    return "AprobadoEFE", "ObsEFE", "FechaValEFE", "RespValEFE"

def move_to_next_level(aviso_id: str, current_level: str):
    df = st.session_state.avisos
    rol_final = df.loc[df["AvisoID"] == aviso_id, "RolFinal"].iloc[0] or "EFE"
    next_map = {"Sacyr": "ADI (ITO)", "ADI (ITO)": rol_final, rol_final: "Final"}
    nxt = next_map.get(current_level, "Final")
    if nxt == "Final":
        st.session_state.avisos.loc[df["AvisoID"] == aviso_id, "ValidacionNivel"] = "Final"
        st.session_state.avisos.loc[df["AvisoID"] == aviso_id, "Estado"] = "Validado"
    else:
        st.session_state.avisos.loc[df["AvisoID"] == aviso_id, "ValidacionNivel"] = nxt
        st.session_state.avisos.loc[df["AvisoID"] == aviso_id, "Estado"] = "En validación"

# -------------------------
# LOGIN
# -------------------------
def cargar_usuarios() -> dict:
    """Lee las 6 cuentas de acceso desde Secrets de Streamlit (nunca desde el código ni GitHub).
    En Streamlit Cloud: panel de la app -> Settings -> Secrets.
    En local: archivo .streamlit/secrets.toml (no se sube al repositorio)."""
    try:
        usuarios_raw = dict(st.secrets["usuarios"])
    except Exception:
        return {}
    usuarios = {}
    for datos in usuarios_raw.values():
        clave = str(datos.get("usuario", "")).strip().lower()
        if clave:
            usuarios[clave] = {
                "password": str(datos.get("password", "")),
                "nombre": datos.get("nombre", clave),
                "perfil": datos.get("perfil", "Personal Terreno"),
            }
    return usuarios

def render_login():
    app_header("Bienvenido")
    st.caption("Ingresa con tu usuario y contraseña.")
    counter = st.session_state.login_form_counter
    usuario_input = st.text_input("Usuario", key=f"login_usuario_{counter}")
    password_input = st.text_input("Contraseña", type="password", key=f"login_password_{counter}")

    if st.button("Ingresar", key="btn_login"):
        usuarios = cargar_usuarios()
        if not usuarios:
            st.error("Todavía no se configuraron los usuarios de esta app. Avisa al administrador.")
        else:
            cuenta = usuarios.get(usuario_input.strip().lower())
            if cuenta and password_input == cuenta["password"] and cuenta["perfil"] in PERFILES:
                st.session_state.perfil = cuenta["perfil"]
                st.session_state.usuario = cuenta["nombre"]
                st.session_state.page = "Inicio"
                st.rerun()
            else:
                st.session_state.login_form_counter += 1
                st.error("Usuario o contraseña incorrectos.")

    st.markdown("<div style='text-align:center;opacity:.5;padding-top:10px;'>sacyr</div>", unsafe_allow_html=True)

if st.session_state.perfil is None:
    render_login()
    st.stop()

# -------------------------
# Pantalla: Crear Aviso (común a todos los perfiles)
# -------------------------
def page_crear_aviso(title="Crear Aviso"):
    app_header(title, back_page="Inicio")
    perfil_bar()

    counter = st.session_state.aviso_form_counter
    es_terreno = st.session_state.perfil == "Personal Terreno"

    if es_terreno:
        actividades_hoy = obtener_actividades_hoy(st.session_state.usuario)
        opciones_activo = actividades_hoy + ["Otro"]
        activo_sel = st.selectbox(
            "Actividad relacionada (reporte de hoy)", opciones_activo, key=f"aviso_activo_{counter}",
            help="Si la incidencia ocurrió durante una actividad ya registrada hoy en tu Reporte Diario, selecciónala. Si no, usa 'Otro'.",
        )
        if activo_sel == "Otro":
            activo_otro = st.text_input(
                "Especifica a qué corresponde la incidencia", key=f"aviso_activo_otro_{counter}",
                placeholder="Ej: Bus MF-177 / Vía km 23...",
            )
            activo = activo_otro.strip() or "Otro"
        else:
            activo = activo_sel
    else:
        activos_df = st.session_state.activos
        activos_list = (activos_df["CodigoActivo"] + " - " + activos_df["NombreActivo"]).tolist()
        if activos_list:
            activo = st.selectbox("Activo", activos_list, key=f"aviso_activo_{counter}")
        else:
            st.caption("Aún no hay activos registrados en el catálogo.")
            activo = st.text_input("Activo", key=f"aviso_activo_{counter}", placeholder="Escribe el activo o ubicación...")

    ubicacion = st.text_input("Ubicación", placeholder="Ej: Andén sur / Túnel interior...", key=f"aviso_ubic_{counter}")
    prioridad = st.selectbox("Prioridad", ["Baja", "Media", "Alta", "Crítica"], index=2, key=f"aviso_prio_{counter}")
    tipificacion = st.selectbox("Tipificación", ["Falla eléctrica", "Falla mecánica", "Frenos", "Puertas",
                                                   "Balasto", "Agujas/Desviador", "Rieles", "Vegetación", "Otro"],
                                 key=f"aviso_tipi_{counter}")
    descripcion = st.text_area("Descripción", height=100, placeholder="Describe el hallazgo o avería...", key=f"aviso_desc_{counter}")
    adjunto = st.file_uploader("Adjuntar foto o PDF (opcional)", type=["png", "jpg", "jpeg", "pdf"], key=f"aviso_adj_{counter}")

    if st.button("Guardar", key="btn_guardar_aviso"):
        aviso_id = next_aviso_id()
        new_row = {
            "AvisoID": aviso_id, "Activo": activo, "Ubicacion": ubicacion or "—",
            "Prioridad": prioridad, "Tipificacion": tipificacion,
            "FechaHora": now_str(), "Descripcion": descripcion or "—",
            "Adjunto": adjunto.name if adjunto else "",
            "CreadoPor": st.session_state.usuario, "RolCreador": st.session_state.perfil,
            "Estado": "Nuevo", "OT_Creada": False,
            "ValidacionNivel": "Sacyr", "RolFinal": "EFE", "ReporteID": "",
            "AprobadoSacyr": False, "ObsSacyr": "", "FechaValSacyr": "", "RespValSacyr": "",
            "AprobadoADI": False, "ObsADI": "", "FechaValADI": "", "RespValADI": "",
            "AprobadoEFE": False, "ObsEFE": "", "FechaValEFE": "", "RespValEFE": "",
            "AprobadoSubcontrato": False, "ObsSubcontrato": "", "FechaValSubcontrato": "", "RespValSubcontrato": "",
        }
        st.session_state.avisos = pd.concat([st.session_state.avisos, pd.DataFrame([new_row])], ignore_index=True)
        save_all_avisos()
        st.session_state.aviso_form_counter = counter + 1
        st.success(f"Aviso {aviso_id} creado ✅")
        if st.session_state.perfil == "Personal Terreno":
            st.session_state.page = "Mis Avisos"
        else:
            st.session_state.selected_aviso = aviso_id
            st.session_state.page = "Detalle"
        st.rerun()

def render_aviso_row(row):
    st.markdown(
        f"""<div class="aviso-row">
                <div class="aviso-title">{row['Activo']}</div>
                <div class="aviso-desc">{row['Descripcion'][:90]}{'...' if len(row['Descripcion']) > 90 else ''}</div>
                <div class="aviso-meta">📄 {row['AvisoID']} · {row['FechaHora']} &nbsp; 🕐 {elapsed_str(row['FechaHora'])}</div>
            </div>""",
        unsafe_allow_html=True,
    )

# =========================================================
# Reporte Diario (Personal Terreno)
# =========================================================
def render_captura_ubicacion() -> str:
    """Ubicación donde se hace el reporte: por GPS del celular o escrita a mano. Devuelve el texto final."""
    modo = st.radio(
        "¿Cómo quieres registrar la ubicación?",
        ["📍 Usar GPS del celular", "✏️ Escribir descripción"],
        key="rep_ubicacion_modo", horizontal=True,
    )

    if modo == "📍 Usar GPS del celular":
        if st.session_state.rep_ubicacion_gps:
            st.success(f"📍 {st.session_state.rep_ubicacion_gps}")
            if st.button("🔄 Volver a capturar", key="btn_gps_recapturar"):
                st.session_state.rep_ubicacion_gps = ""
                st.rerun()
            return st.session_state.rep_ubicacion_gps

        st.caption("Toca el botón 📍 y acepta el permiso de ubicación si tu navegador lo pide.")
        loc = streamlit_geolocation()
        lat = loc.get("latitude") if loc else None
        lon = loc.get("longitude") if loc else None
        if lat is not None and lon is not None:
            st.session_state.rep_ubicacion_gps = f"{lat:.6f}, {lon:.6f} (GPS)"
            st.rerun()
        return ""
    else:
        return st.text_input(
            "Descripción de la ubicación", key=f"rep_ubicacion_texto_{st.session_state.ubicacion_texto_counter}",
            placeholder="Ej: Km 23.400, cerca de la estación Malloco...",
        )

def render_trabajos_section():
    if not st.session_state.reporte_trabajos:
        st.caption("Aún no hay actividades agregadas.")
    for row in st.session_state.reporte_trabajos:
        rid = row["id"]
        with st.container(border=True):
            row["actividad"] = st.selectbox(
                "Actividad", ACTIVIDADES_TRABAJO,
                index=ACTIVIDADES_TRABAJO.index(row["actividad"]) if row["actividad"] in ACTIVIDADES_TRABAJO else 0,
                key=f"trab_act_{rid}",
            )
            row["pk"] = st.number_input(
                "PK", min_value=33, max_value=61, step=1, value=int(row.get("pk", 33)), key=f"trab_pk_{rid}",
            )
            colleras_pk = cargar_colleras_de_pk(row["pk"])
            idx_desde = idx_hasta = None
            if not colleras_pk:
                st.warning(f"No hay colleras registradas para el PK {row['pk']}.")
                row["km_desde"], row["km_hasta"] = "", ""
                row["collera_desde_id"], row["collera_hasta_id"] = "", ""
            else:
                opciones = [f"Collera {c['Collera']} (km {c['KmDesde']:.3f}–{c['KmHasta']:.3f})" for c in colleras_pk]
                cA, cB = st.columns(2)
                with cA:
                    collera_desde = st.selectbox("Collera Desde", opciones, key=f"trab_coldesde_{rid}")
                with cB:
                    collera_hasta = st.selectbox(
                        "Collera Hasta", opciones, index=len(opciones) - 1, key=f"trab_colhasta_{rid}",
                    )
                idx_desde = opciones.index(collera_desde)
                idx_hasta = opciones.index(collera_hasta)
                row["km_desde"] = f"{colleras_pk[idx_desde]['KmDesde']:.3f}"
                row["km_hasta"] = f"{colleras_pk[idx_hasta]['KmHasta']:.3f}"
                row["collera_desde_id"] = colleras_pk[idx_desde]["ColleraID"]
                row["collera_hasta_id"] = colleras_pk[idx_hasta]["ColleraID"]

            c3, c4 = st.columns(2)
            with c3:
                row["unidad"] = st.selectbox(
                    "Unidad", UNIDADES_TRABAJO,
                    index=UNIDADES_TRABAJO.index(row["unidad"]) if row["unidad"] in UNIDADES_TRABAJO else 0,
                    key=f"trab_uni_{rid}",
                )
            with c4:
                row["cantidad"] = st.number_input("Cantidad", min_value=0.0, value=float(row["cantidad"]), step=1.0, key=f"trab_cant_{rid}")
            row["hombres"] = st.number_input("N° Trabajadores (Hombre)", min_value=0, value=int(row["hombres"]), step=1, key=f"trab_hom_{rid}")

            if row["actividad"] in ("Inspección Vía", "Reemplazo de Durmientes Común"):
                st.caption("🛤️ Registra el estado de los durmientes en la sección **Control de Durmientes por Collera** (más abajo).")

            if st.button("🗑 Eliminar actividad", key=f"trab_del_{rid}"):
                st.session_state.reporte_trabajos = [r for r in st.session_state.reporte_trabajos if r["id"] != rid]
                st.rerun()

# -------------------------
# Control de Durmientes por Collera (en el Reporte Diario, lo ingresa el Jefe de Grupo)
# -------------------------
ETIQUETA_ESTADO = {"Bueno": "Bueno", "Malo": "Malo", "Nuevo": "Nuevo", "Reemplazado": "Reempl."}
CUADRO_ESTADO = {"Bueno": "🟩", "Malo": "🟥", "Nuevo": "🟦", "Reemplazado": "🟪"}
TIPO_BADGE_ESTADO = {"Cumple": "ok", "No Cumple": "bad", "No Cumple (racha)": "bad",
                     "Inspección Incompleta": "warn", "Sin Inspeccionar": "muted"}

def new_collera_row():
    return {"id": new_row_id(), "pk": 33}

def _clave_durmiente(rid: str, collera_id: str, pos: int) -> str:
    return f"durm_{rid}_{collera_id}_{pos}"

def _completar_vacios(rid: str, collera_id: str, estado: str):
    for pos in range(1, 24):
        clave = _clave_durmiente(rid, collera_id, pos)
        if not st.session_state.get(clave):
            st.session_state[clave] = estado

def _restaurar_vigente(rid: str, collera_id: str):
    vigente = cargar_estado_collera(collera_id)
    for pos in range(1, 24):
        st.session_state[_clave_durmiente(rid, collera_id, pos)] = vigente.get(pos)

def tabla_indicadores_collera(ev: dict) -> pd.DataFrame:
    """Los mismos indicadores de la hoja 'Control Durmientes' del Excel (columnas AA:AK)."""
    return pd.DataFrame([
        ("TOTAL REGISTRADO", ev["total_registrado"]),
        ("N° BUENOS", ev["n_buenos"]),
        ("N° MALOS", ev["n_malos"]),
        ("N° NUEVOS", ev["n_nuevos"]),
        ("N° REEMPL.", ev["n_reempl"]),
        ("EFECTIVOS (B+N+R)", ev["efectivos"]),
        ("% RENOV. (N+R)", f"{ev['pct_renovacion'] * 100:.1f}%"),
        ("MÍN. EFECTIVOS (≥10)", ev["min_efectivos"]),
        (f"RACHA CONSEC. (máx. {ev['limite_racha']} Malo)", f"{ev['racha_consec']} (racha {ev['racha_max']})"),
        ("ESTADO GENERAL", ev["estado_general"]),
    ], columns=["Indicador", "Valor"]).astype(str)

def franja_durmientes(secuencia: list) -> str:
    return "".join(CUADRO_ESTADO.get(e, "⬜") for e in secuencia)

def render_colleras_section():
    st.caption(
        "Registra el estado de cada durmiente (D1 a D23) de las colleras que inspeccionaste o "
        "intervenidas hoy. Se precarga el último estado conocido; deja sin marcar lo que no "
        "inspeccionaste (no se asume 'Bueno'). Norma NS-01-01-00: mínimo 10 efectivos y "
        "máximo 3 'Malo' seguidos en recta / 2 en curva."
    )
    if not st.session_state.reporte_colleras:
        st.caption("Aún no hay colleras agregadas.")
    for item in st.session_state.reporte_colleras:
        rid = item["id"]
        with st.container(border=True):
            item["pk"] = st.number_input("PK", min_value=33, max_value=61, step=1, value=int(item["pk"]), key=f"col_pk_{rid}")
            colleras_pk = cargar_colleras_de_pk(item["pk"])
            if not colleras_pk:
                st.warning(f"No hay colleras registradas para el PK {item['pk']}.")
                item.pop("collera", None)
                continue
            por_id = {c["ColleraID"]: c for c in colleras_pk}
            collera_id = st.selectbox(
                "Collera", list(por_id), key=f"col_sel_{rid}",
                format_func=lambda cid: (f"Collera {por_id[cid]['Collera']} · km {por_id[cid]['KmDesde']:.3f}–"
                                         f"{por_id[cid]['KmHasta']:.3f} · {por_id[cid]['Ubicacion']}"),
            )
            collera = por_id[collera_id]

            vigente = cargar_estado_collera(collera_id)
            for pos in range(1, 24):
                clave = _clave_durmiente(rid, collera_id, pos)
                if clave not in st.session_state:
                    st.session_state[clave] = vigente.get(pos)

            b1, b2 = st.columns(2)
            with b1:
                st.button("✅ Marcar vacíos como Bueno", key=f"col_llenar_{rid}",
                          on_click=_completar_vacios, args=(rid, collera_id, "Bueno"))
            with b2:
                st.button("↩️ Volver al último estado", key=f"col_restaurar_{rid}",
                          on_click=_restaurar_vigente, args=(rid, collera_id))

            secuencia = []
            for pos in range(1, 24):
                secuencia.append(st.segmented_control(
                    f"D{pos}", ESTADOS_DURMIENTE, format_func=ETIQUETA_ESTADO.get,
                    key=_clave_durmiente(rid, collera_id, pos),
                ))
            item["collera"] = collera
            item["secuencia"] = secuencia

            ev = evaluar_secuencia(secuencia, collera["Ubicacion"])
            st.markdown(f"**D1 → D23:** {franja_durmientes(secuencia)}")
            st.caption("🟩 Bueno · 🟥 Malo · 🟦 Nuevo · 🟪 Reemplazado · ⬜ sin inspeccionar")
            badge(ev["estado_general"], TIPO_BADGE_ESTADO.get(ev["estado_general"], "info"))
            st.dataframe(tabla_indicadores_collera(ev), hide_index=True, width="stretch")

            if st.button("🗑 Quitar collera", key=f"col_del_{rid}"):
                st.session_state.reporte_colleras = [c for c in st.session_state.reporte_colleras if c["id"] != rid]
                st.rerun()

def render_equipos_section():
    st.caption("Marca los equipos utilizados durante la jornada e indica la cantidad.")
    for nombre in EQUIPOS_CATALOGO:
        item = st.session_state.reporte_equipos_catalogo[nombre]
        slug = _slug(nombre)
        c1, c2, c3 = st.columns([3, 1, 1.3])
        with c1:
            st.markdown(nombre)
        with c2:
            item["usado"] = st.toggle("Usado", value=item["usado"], key=f"equipo_usado_{slug}", label_visibility="collapsed")
        with c3:
            if item["usado"]:
                item["cantidad"] = st.number_input(
                    "Cantidad", min_value=0.0, value=float(item["cantidad"] or 1.0), step=1.0,
                    key=f"equipo_cant_{slug}", label_visibility="collapsed",
                )
            else:
                st.caption("—")

    st.divider()
    st.markdown("**Otro equipo no listado**")
    counter = st.session_state.equipo_otro_form_counter
    c1, c2 = st.columns([3, 2])
    with c1:
        nombre_nuevo = st.text_input("Nombre del equipo", key=f"equipo_otro_nombre_{counter}", placeholder="Ej: Compactadora manual")
    with c2:
        cantidad_nueva = st.number_input("Cantidad", min_value=0.0, value=1.0, step=1.0, key=f"equipo_otro_cant_{counter}")
    if st.button("➕ Agregar", key="btn_add_equipo_otro"):
        nombre_limpio = nombre_nuevo.strip()
        if not nombre_limpio:
            st.warning("Escribe el nombre del equipo antes de agregar.")
        else:
            st.session_state.reporte_equipos_otros.append({"id": new_row_id(), "nombre": nombre_limpio, "cantidad": cantidad_nueva})
            st.session_state.equipo_otro_form_counter += 1
            st.rerun()

    for row in st.session_state.reporte_equipos_otros:
        rid = row["id"]
        c1, c2, c3 = st.columns([3, 2, 1])
        with c1:
            st.markdown(f"**{row['nombre']}**")
        with c2:
            st.caption(f"Cantidad: {row['cantidad']:.0f}")
        with c3:
            if st.button("🗑", key=f"equipo_otro_del_{rid}"):
                st.session_state.reporte_equipos_otros = [r for r in st.session_state.reporte_equipos_otros if r["id"] != rid]
                st.rerun()

def render_materiales_section():
    if not st.session_state.reporte_materiales:
        st.caption("Aún no hay materiales agregados.")
    for row in st.session_state.reporte_materiales:
        rid = row["id"]
        with st.container(border=True):
            c1, c2 = st.columns([3, 2])
            with c1:
                row["nombre"] = st.selectbox(
                    "Material", MATERIALES_CATALOGO,
                    index=MATERIALES_CATALOGO.index(row["nombre"]) if row["nombre"] in MATERIALES_CATALOGO else 0,
                    key=f"material_nombre_{rid}",
                )
            with c2:
                row["cantidad"] = st.number_input("Cantidad", min_value=0.0, value=float(row["cantidad"]), step=1.0, key=f"material_cant_{rid}")
            if row["nombre"] == "Otro (especificar)":
                row["nombre_custom"] = st.text_input("Especificar material", value=row["nombre_custom"], key=f"material_custom_{rid}")
            row["estado"] = st.selectbox(
                "Estado del material", ESTADOS_MATERIAL,
                index=ESTADOS_MATERIAL.index(row["estado"]) if row.get("estado") in ESTADOS_MATERIAL else 0,
                key=f"material_estado_{rid}",
            )
            if st.button("🗑 Eliminar material", key=f"material_del_{rid}"):
                st.session_state.reporte_materiales = [r for r in st.session_state.reporte_materiales if r["id"] != rid]
                st.rerun()

def render_asistencia_section():
    counter = st.session_state.asistencia_form_counter
    c1, c2 = st.columns([3, 2])
    with c1:
        nombre_nuevo = st.text_input("Nombre del trabajador", key=f"asist_nombre_{counter}", placeholder="Escribe el nombre...")
    with c2:
        cargo_nuevo = st.selectbox("Cargo", CARGOS_TRABAJADOR, key=f"asist_cargo_{counter}")

    if st.button("➕ Agregar", key="btn_add_trabajador"):
        nombre_limpio = nombre_nuevo.strip()
        if not nombre_limpio:
            st.warning("Escribe el nombre del trabajador antes de agregar.")
        else:
            st.session_state.reporte_asistencia.append({"nombre": nombre_limpio, "cargo": cargo_nuevo, "estado": "Asistió"})
            st.session_state.asistencia_form_counter += 1
            st.rerun()

    st.write("")
    if not st.session_state.reporte_asistencia:
        st.caption("Aún no hay trabajadores agregados a la asistencia.")
    for idx, a in enumerate(st.session_state.reporte_asistencia):
        c1, c2, c3 = st.columns([3, 2, 1])
        with c1:
            st.markdown(f"**{a['nombre']}**")
            st.caption(a["cargo"])
        with c2:
            a["estado"] = st.selectbox(
                "Estado", ["Asistió", "No asistió"],
                index=0 if a["estado"] == "Asistió" else 1,
                key=f"asist_estado_{idx}_{a['nombre']}", label_visibility="collapsed",
            )
        with c3:
            if st.button("🗑", key=f"asist_del_{idx}_{a['nombre']}"):
                st.session_state.reporte_asistencia.pop(idx)
                st.rerun()

def crear_aviso_desde_reporte(reporte_id: str, grupo_via: str, ubicacion: str, resumen_texto: str, creado_por: str | None = None) -> str:
    """Al guardar un Reporte Diario, se levanta automáticamente un aviso para que lo validen
    Sacyr, la ITO y el Supervisor del Subcontrato (esta app es de uso exclusivo para Icafal)."""
    aviso_id = next_aviso_id()
    new_row = {
        "AvisoID": aviso_id, "Activo": f"Reporte Diario · Grupo Vía {grupo_via}",
        "Ubicacion": ubicacion or f"Grupo Vía {grupo_via}", "Prioridad": "Media", "Tipificacion": "Reporte Diario",
        "FechaHora": now_str(), "Descripcion": resumen_texto, "Adjunto": "",
        "CreadoPor": creado_por or st.session_state.usuario, "RolCreador": "Personal Terreno",
        "Estado": "Nuevo", "OT_Creada": False,
        "ValidacionNivel": "Sacyr", "RolFinal": "Supervisor Subcontrato", "ReporteID": reporte_id,
        "AprobadoSacyr": False, "ObsSacyr": "", "FechaValSacyr": "", "RespValSacyr": "",
        "AprobadoADI": False, "ObsADI": "", "FechaValADI": "", "RespValADI": "",
        "AprobadoEFE": False, "ObsEFE": "", "FechaValEFE": "", "RespValEFE": "",
        "AprobadoSubcontrato": False, "ObsSubcontrato": "", "FechaValSubcontrato": "", "RespValSubcontrato": "",
    }
    st.session_state.avisos = pd.concat([st.session_state.avisos, pd.DataFrame([new_row])], ignore_index=True)
    save_all_avisos()
    return aviso_id

def reconciliar_avisos_faltantes() -> int:
    """Recrea Avisos para Reportes Diarios que ya existen (en Sheets o excel local) pero no
    tienen Aviso asociado. Pasa con reportes guardados ANTES de que Avisos también quedara
    guardado en Sheets: el Aviso vivía solo en el SQLite efímero de Streamlit Cloud y se
    perdió en un reinicio, aunque el Reporte real (hojas Reportes/Trabajos) siguió intacto.
    Quedan como 'Nuevo' -- el estado de aprobación que hubieran tenido antes no es
    recuperable porque nunca quedó guardado de forma durable."""
    df_rep = _leer_hoja_df("Reportes")
    if df_rep.empty:
        return 0
    existentes = set(st.session_state.avisos["ReporteID"].astype(str)) if not st.session_state.avisos.empty else set()
    faltantes = df_rep[~df_rep["ReporteID"].astype(str).isin(existentes)]
    for _, rep in faltantes.iterrows():
        ubicacion = rep.get("Ubicacion")
        ubicacion = ubicacion if isinstance(ubicacion, str) and ubicacion.strip() else ""
        crear_aviso_desde_reporte(
            str(rep["ReporteID"]), str(rep.get("GrupoVia") or ""), ubicacion,
            f"Reporte Diario recuperado del histórico ({rep.get('Fecha')}).",
            creado_por=rep.get("Usuario") or "Personal Terreno",
        )
    return len(faltantes)

def equipos_seleccionados():
    seleccion = [
        {"nombre": nombre, "cantidad": item["cantidad"]}
        for nombre, item in st.session_state.reporte_equipos_catalogo.items()
        if item["usado"]
    ]
    seleccion += [{"nombre": o["nombre"], "cantidad": o["cantidad"]} for o in st.session_state.reporte_equipos_otros]
    return seleccion

def resumen_collera(collera: dict, secuencia: list) -> dict:
    """Datos de una collera inspeccionada, para el PDF/Excel del reporte."""
    return {"collera_id": collera["ColleraID"], "pk": int(collera["PK"]), "collera": int(collera["Collera"]),
            "ubicacion": collera["Ubicacion"], "secuencia": list(secuencia),
            **evaluar_secuencia(secuencia, collera["Ubicacion"])}

def guardar_reporte(grupo_via, fecha_reporte, ubicacion, observaciones, fotos_subidas):
    trabajos = st.session_state.reporte_trabajos
    if not trabajos:
        return {"ok": False, "msg": "Agrega al menos una actividad en la sección Trabajos antes de guardar."}

    colleras_insp = [c for c in st.session_state.reporte_colleras if c.get("collera")]
    ids_colleras = [c["collera"]["ColleraID"] for c in colleras_insp]
    repetidas = sorted({cid for cid in ids_colleras if ids_colleras.count(cid) > 1})
    if repetidas:
        return {"ok": False, "msg": f"La collera {', '.join(repetidas)} está agregada más de una vez. Deja solo una."}
    vacias = [c["collera"]["ColleraID"] for c in colleras_insp if not any(c["secuencia"])]
    if vacias:
        return {"ok": False, "msg": f"La collera {', '.join(vacias)} no tiene ningún durmiente marcado. Márcalos o quítala."}

    try:
        reporte_id = next_reporte_id()
    except Exception:
        return {"ok": False, "msg": "No se pudo conectar con Google Sheets para numerar el reporte. "
                                    "Revisa la señal e intenta de nuevo: lo que ingresaste sigue en el formulario."}
    horas_dia = horas_jornada_por_fecha(fecha_reporte)
    fecha_str = fecha_reporte.strftime("%Y-%m-%d")
    reporte_row = [reporte_id, fecha_str, grupo_via, st.session_state.usuario, observaciones or "", now_str(), ubicacion or ""]

    trabajos_calc = [{**t, "hh": t["hombres"] * horas_dia} for t in trabajos]
    trabajos_rows = [
        [reporte_id, t["actividad"], t["collera_desde_id"], t["collera_hasta_id"],
         t["km_desde"], t["km_hasta"], t["unidad"], t["cantidad"], t["hombres"], t["hh"]]
        for t in trabajos_calc
    ]
    inspeccion_rows = [fila_inspeccion(reporte_id, fecha_str, c["collera"], c["secuencia"]) for c in colleras_insp]
    cambios_durmientes = filas_cambios_durmientes(
        reporte_id, fecha_str, {c["collera"]["ColleraID"]: c["secuencia"] for c in colleras_insp})
    equipos_usados = equipos_seleccionados()
    equipos_rows = [[reporte_id, e["nombre"], e["cantidad"]] for e in equipos_usados]
    materiales_rows = [
        [reporte_id, _resolver_nombre(m), m["cantidad"], m.get("estado", "")]
        for m in st.session_state.reporte_materiales
    ]
    asistencia_rows = [
        [reporte_id, a["nombre"], a["cargo"], a["estado"], "", "", ""]
        for a in st.session_state.reporte_asistencia
    ]

    fotos_guardadas = []
    if fotos_subidas:
        carpeta_fotos = os.path.join(FOTOS_DIR, reporte_id)
        os.makedirs(carpeta_fotos, exist_ok=True)
        for idx, foto in enumerate(fotos_subidas, start=1):
            nombre_final = f"foto_{idx:02d}.jpg"
            ruta = os.path.join(carpeta_fotos, nombre_final)
            try:
                # Normaliza a JPEG estándar: los celulares suelen entregar formatos
                # (MPO, HEIC, etc.) que PIL abre pero openpyxl no sabe incrustar al guardar.
                # Se achica a FOTO_MAX_LADO para que el respaldo en Sheets sea liviano.
                img = PILImage.open(foto)
                img = ImageOps.exif_transpose(img)
                img = img.convert("RGB")
                img.thumbnail((FOTO_MAX_LADO, FOTO_MAX_LADO))
                img.save(ruta, format="JPEG", quality=FOTO_CALIDAD, optimize=True)
            except Exception:
                nombre_final = foto.name
                ruta = os.path.join(carpeta_fotos, nombre_final)
                with open(ruta, "wb") as f:
                    f.write(foto.getbuffer())
            fotos_guardadas.append((nombre_final, ruta))
    fotos_rows = [[reporte_id, nombre, ruta] for nombre, ruta in fotos_guardadas]

    guardado = guardar_bloques({
        "Reportes": [reporte_row], "Trabajos": trabajos_rows, "Equipos": equipos_rows,
        "Materiales": materiales_rows, "Asistencia": asistencia_rows, "Fotos": fotos_rows,
        "InspeccionColleras": inspeccion_rows, "DurmientesEstado": cambios_durmientes,
    })
    if not guardado:
        return {"ok": False, "msg": "No se pudo guardar el reporte en Google Sheets (sin conexión o Google no respondió). "
                                    "No se guardó nada a medias: lo que ingresaste sigue en el formulario, intenta de nuevo."}
    persistir_cambios_durmientes_local(cambios_durmientes)
    fotos_respaldo = _respaldar_fotos_sheets(reporte_id, fotos_guardadas)

    resumen = {
        "reporte_id": reporte_id, "fecha": fecha_str, "grupo_via": grupo_via,
        "usuario": st.session_state.usuario, "horas_dia": horas_dia,
        "ubicacion": ubicacion or "", "observaciones": observaciones or "",
        "trabajos": trabajos_calc,
        "equipos": equipos_usados,
        "materiales": list(st.session_state.reporte_materiales),
        "asistencia": list(st.session_state.reporte_asistencia),
        "fotos": fotos_guardadas,
        "fotos_respaldo": fotos_respaldo,
        "colleras": [resumen_collera(c["collera"], c["secuencia"]) for c in colleras_insp],
        "total_hh": sum(t["hh"] for t in trabajos_calc),
    }

    texto_colleras = ""
    if resumen["colleras"]:
        n_no = sum(1 for c in resumen["colleras"] if c["estado_general"] in ESTADOS_NO_CUMPLE)
        texto_colleras = f" {len(resumen['colleras'])} collera(s) inspeccionada(s), {n_no} no cumple(n) la norma."
    resumen_texto = (
        f"Reporte Diario {reporte_id} · {fecha_str} · {len(trabajos_calc)} actividad(es), "
        f"{len(resumen['equipos'])} equipo(s), {len(resumen['materiales'])} material(es), "
        f"{len(resumen['asistencia'])} trabajador(es) en asistencia, {resumen['total_hh']:.0f} HH totales."
        + texto_colleras
    )
    resumen["aviso_id"] = crear_aviso_desde_reporte(reporte_id, grupo_via, ubicacion, resumen_texto)

    excel_bytes, pdf_bytes = generar_documentos_reporte(resumen)

    return {"ok": True, "msg": f"Reporte {reporte_id} guardado ✅", "resumen": resumen,
            "excel_bytes": excel_bytes, "pdf_bytes": pdf_bytes}

def page_generar_reporte():
    app_header("Generar Reporte", back_page="Inicio")
    perfil_bar()

    st.markdown("#### Datos generales")
    c1, c2 = st.columns(2)
    with c1:
        grupo_via = st.selectbox("Grupo Vía", GRUPOS_VIA, key="rep_grupo_via")
    with c2:
        fecha_reporte = st.date_input("Fecha", datetime.now().date(), key="rep_fecha")
    horas_dia = horas_jornada_por_fecha(fecha_reporte)
    st.caption(f"📅 {DIAS_ES[fecha_reporte.weekday()]} → jornada estándar: **{horas_dia} h por trabajador** "
               f"(se usará para calcular las Horas Hombre de cada actividad).")

    st.markdown("**Ubicación**")
    ubicacion = render_captura_ubicacion()

    st.divider()
    st.markdown("#### 🛠️ Trabajos")
    render_trabajos_section()
    if st.button("➕ Agregar actividad", key="btn_add_trabajo"):
        st.session_state.reporte_trabajos.append(new_trabajo_row())
        st.rerun()

    st.divider()
    st.markdown("#### 🛤️ Control de Durmientes por Collera")
    render_colleras_section()
    if st.button("➕ Agregar collera", key="btn_add_collera"):
        st.session_state.reporte_colleras.append(new_collera_row())
        st.rerun()

    st.divider()
    st.markdown("#### 🚜 Equipos")
    render_equipos_section()

    st.divider()
    st.markdown("#### 📦 Materiales")
    render_materiales_section()
    if st.button("➕ Agregar material", key="btn_add_material"):
        st.session_state.reporte_materiales.append(new_material_row())
        st.rerun()

    st.divider()
    st.markdown("#### 📷 Fotografías")
    fotos_subidas = st.file_uploader(
        "Adjuntar foto(s) del trabajo realizado", type=["png", "jpg", "jpeg"], accept_multiple_files=True,
        key=f"rep_fotos_{st.session_state.reporte_fotos_counter}",
    )

    st.divider()
    st.markdown("#### 📝 Observaciones")
    observaciones = st.text_area("Observaciones", key="rep_observaciones", height=100, label_visibility="collapsed",
                                  placeholder="Comentarios, novedades o información relevante de la jornada...")

    st.divider()
    st.markdown("#### 👷 Control de Asistencia")
    render_asistencia_section()

    st.divider()
    if st.button("💾 Guardar Reporte", key="btn_guardar_reporte"):
        resultado = guardar_reporte(grupo_via, fecha_reporte, ubicacion, observaciones, fotos_subidas)
        if resultado["ok"]:
            st.session_state.last_reporte = resultado["resumen"]
            st.session_state.last_reporte_excel = resultado["excel_bytes"]
            st.session_state.last_reporte_pdf = resultado["pdf_bytes"]
            reset_reporte_form()
            st.session_state.page = "Reporte Guardado"
            st.rerun()
        else:
            st.error(resultado["msg"])

def render_boton_compartir_excel(excel_bytes: bytes, reporte_id: str, resumen: dict):
    """Boton que abre el panel nativo de 'Compartir' del celular (Correo, Gmail, WhatsApp, etc.)
    con el Excel del reporte ya adjunto. No requiere configurar ningun correo ni servidor propio."""
    b64 = base64.b64encode(excel_bytes).decode("utf-8")
    file_name = f"Reporte_{reporte_id}.xlsx"
    asunto = f"Reporte Diario {reporte_id} - Grupo Via {resumen['grupo_via']} - {resumen['fecha']}"
    cuerpo = (
        f"Se adjunta el Reporte Diario {reporte_id} del Grupo Via {resumen['grupo_via']}, "
        f"correspondiente al {resumen['fecha']}, generado por {resumen['usuario']}."
    )

    html_template = """
    <div style="font-family: -apple-system, sans-serif;">
      <button id="btnCompartirReporte" style="
          width:100%; padding:0.65rem 1rem; border-radius:10px; font-weight:600;
          background:#2FA84F; color:#fff; border:none; font-size:15px; cursor:pointer;">
        📤 Enviar / Compartir Excel
      </button>
      <p id="compartirStatus" style="font-size:12.5px; color:#888; margin-top:6px; min-height:16px;"></p>
    </div>
    <script>
    (function () {
      const b64Data = "__B64DATA__";
      const fileName = __FILENAME_JSON__;
      const subject = __SUBJECT_JSON__;
      const bodyText = __BODY_JSON__;

      const btn = document.getElementById('btnCompartirReporte');
      const statusEl = document.getElementById('compartirStatus');

      btn.addEventListener('click', async function () {
        statusEl.textContent = 'Preparando archivo...';
        try {
          const byteChars = atob(b64Data);
          const byteNumbers = new Array(byteChars.length);
          for (let i = 0; i < byteChars.length; i++) {
            byteNumbers[i] = byteChars.charCodeAt(i);
          }
          const byteArray = new Uint8Array(byteNumbers);
          const file = new File([byteArray], fileName, {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          });

          if (navigator.canShare && navigator.canShare({ files: [file] })) {
            await navigator.share({ files: [file], title: subject, text: bodyText });
            statusEl.textContent = 'Compartido';
          } else {
            statusEl.textContent = 'Este navegador no permite adjuntar directo. Se abrira tu correo: adjunta el Excel a mano (boton Descargar Excel de arriba).';
            window.location.href = 'mailto:?subject=' + encodeURIComponent(subject) + '&body=' + encodeURIComponent(bodyText);
          }
        } catch (err) {
          if (err && err.name === 'AbortError') {
            statusEl.textContent = 'Cancelado.';
          } else {
            statusEl.textContent = 'No se pudo compartir: ' + (err && err.message ? err.message : err);
          }
        }
      });
    })();
    </script>
    """

    html_code = (
        html_template
        .replace("__B64DATA__", b64)
        .replace("__FILENAME_JSON__", json.dumps(file_name))
        .replace("__SUBJECT_JSON__", json.dumps(asunto))
        .replace("__BODY_JSON__", json.dumps(cuerpo))
    )
    st.iframe(html_code, height=90)

def page_reporte_guardado():
    app_header("Reporte Guardado", back_page="Inicio")
    perfil_bar()
    resumen = st.session_state.last_reporte
    if not resumen:
        st.info("No hay un reporte reciente para mostrar.")
        return

    st.success(f"Reporte **{resumen['reporte_id']}** guardado correctamente ✅")
    st.markdown("#### Resumen de lo reportado")
    c1, c2 = st.columns(2)
    with c1:
        st.metric("Actividades", len(resumen["trabajos"]))
        st.metric("Materiales", len(resumen["materiales"]))
        st.metric("Fotografías", len(resumen["fotos"]))
    with c2:
        st.metric("Equipos", len(resumen["equipos"]))
        st.metric("Asistencia", len(resumen["asistencia"]))
        st.metric("HH totales", f"{resumen['total_hh']:.0f}")
    st.caption(f"Grupo Vía {resumen['grupo_via']} · {resumen['fecha']} · Reportado por {resumen['usuario']}")
    if resumen.get("fotos_respaldo") is False:
        st.warning(
            "El reporte quedó guardado, pero alguna foto no se pudo respaldar en Google Sheets. "
            "Descarga el PDF ahora para no perderla, o avisa al administrador."
        )

    st.divider()
    st.markdown("#### Descargar reporte")
    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            "⬇️ Descargar Excel", data=st.session_state.last_reporte_excel,
            file_name=f"Reporte_{resumen['reporte_id']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_descargar_excel", width="stretch",
        )
    with dl2:
        st.download_button(
            "⬇️ Descargar PDF", data=st.session_state.last_reporte_pdf,
            file_name=f"Reporte_{resumen['reporte_id']}.pdf", mime="application/pdf",
            key="btn_descargar_pdf", width="stretch",
        )

    st.divider()
    st.markdown("#### Enviar por correo")
    render_boton_compartir_excel(st.session_state.last_reporte_excel, resumen["reporte_id"], resumen)
    st.caption(
        "Abre el panel de Compartir del celular (Correo, Gmail, WhatsApp, etc.) con el Excel ya adjunto. "
        "Si el navegador no lo permite, se abrirá el correo sin adjunto — usa el botón Descargar Excel de arriba y adjúntalo a mano."
    )

    st.divider()
    st.info(
        f"Se generó el aviso **{resumen['aviso_id']}** en el Backlog de Avisos para su validación por "
        "Sacyr, la ITO y el Supervisor del Subcontrato."
    )
    st.caption("Esta aplicación es de uso exclusivo para Icafal.")

    st.divider()
    if st.button("🏠 Volver al Inicio", key="btn_volver_inicio_reporte"):
        st.session_state.last_reporte = None
        st.session_state.last_reporte_excel = None
        st.session_state.last_reporte_pdf = None
        st.session_state.page = "Inicio"
        st.rerun()

# =========================================================
# FLUJO: PERSONAL TERRENO (solo reporte)
# =========================================================
def terreno_inicio():
    app_header("Inicio", right_icon="💬")
    perfil_bar()
    df = st.session_state.avisos
    mios = df[df["CreadoPor"] == st.session_state.usuario]
    pend = int(mios[mios["Estado"].isin(["Nuevo", "En validación", "Observado"])].shape[0])

    st.markdown("### Bienvenido")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("➕  Crear Incidencia", key="btn_crear_aviso"):
            st.session_state.page = "Crear Incidencia"
            st.rerun()
    with c2:
        if st.button("📝  Generar Reporte", key="btn_generar_reporte"):
            st.session_state.page = "Generar Reporte"
            st.rerun()
    if st.button(f"📋  Mis Avisos · {pend} Pendientes", key="btn_mis_avisos"):
        st.session_state.page = "Mis Avisos"
        st.rerun()
    if st.button("🗂️  Mis Reportes (histórico)", key="btn_mis_reportes"):
        st.session_state.page = "Mis Reportes"
        st.rerun()

def terreno_mis_reportes():
    """Los Reportes Diarios ya enviados por este usuario, leídos desde Google Sheets: quedan
    guardados aunque la app se reinicie, y se pueden abrir y volver a descargar."""
    app_header("Mis Reportes", back_page="Inicio")
    perfil_bar()
    df_rep = _leer_hoja_df("Reportes")
    mios = df_rep[df_rep["Usuario"].astype(str) == str(st.session_state.usuario)] if not df_rep.empty else df_rep
    if mios.empty:
        st.info("Aún no has enviado Reportes Diarios.")
        return
    mios = mios.sort_values(["Fecha", "ReporteID"], ascending=False, key=lambda s: s.astype(str))
    st.caption(f"{len(mios)} reporte(s) enviados.")
    for _, rep in mios.iterrows():
        rid = str(rep["ReporteID"])
        colA, colB = st.columns([4, 1])
        with colA:
            st.markdown(f"**{rid}** · {rep['Fecha']}")
            st.caption(f"Grupo Vía {rep['GrupoVia']} · creado {rep.get('FechaCreacion', '')}")
        with colB:
            if st.button("Ver ›", key=f"ver_rep_{rid}"):
                st.session_state.selected_reporte = rid
                st.session_state.page = "Detalle Reporte"
                st.rerun()

def terreno_detalle_reporte():
    app_header("Detalle Reporte", back_page="Mis Reportes")
    perfil_bar()
    reporte_id = st.session_state.get("selected_reporte")
    if not reporte_id:
        st.info("Selecciona un reporte desde Mis Reportes.")
        return
    mostrar_detalle_reporte_diario(reporte_id)

def terreno_mis_avisos():
    app_header("Mis Avisos", back_page="Inicio")
    perfil_bar()
    df = st.session_state.avisos
    mios = df[df["CreadoPor"] == st.session_state.usuario].sort_values("FechaHora", ascending=False)
    if mios.empty:
        st.info("Aún no has creado avisos.")
        return
    estado_kind = {"Nuevo": "info", "En validación": "info", "Observado": "warn",
                   "Validado": "ok", "OT creada": "ok", "Cerrado": "muted", "Rechazado": "bad"}
    for _, row in mios.iterrows():
        colA, colB = st.columns([5, 1])
        with colA:
            render_aviso_row(row)
        with colB:
            badge(row["Estado"], estado_kind.get(row["Estado"], "info"))
        if row["Estado"] == "Observado":
            obs = row["ObsSacyr"] or row["ObsADI"] or row["ObsEFE"] or row.get("ObsSubcontrato", "")
            if obs:
                st.warning(f"Observación: {obs}")

def flujo_terreno():
    page = st.session_state.page
    if page == "Inicio":
        terreno_inicio()
    elif page == "Crear Incidencia":
        page_crear_aviso(title="Crear Incidencia")
    elif page == "Generar Reporte":
        page_generar_reporte()
    elif page == "Reporte Guardado":
        page_reporte_guardado()
    elif page == "Mis Avisos":
        terreno_mis_avisos()
    elif page == "Mis Reportes":
        terreno_mis_reportes()
    elif page == "Detalle Reporte":
        terreno_detalle_reporte()
    else:
        st.session_state.page = "Inicio"
        st.rerun()
    render_pie_desarrollador()
    bottom_nav([("Inicio", "🏠"), ("Crear Incidencia", "➕"), ("Generar Reporte", "📝"), ("Mis Avisos", "📋")])

# =========================================================
# FLUJO: VALIDADORES (Sacyr / ADI (ITO) / EFE)
# =========================================================
def validador_inicio():
    app_header("Inicio", right_icon="💬")
    perfil_bar()
    dfA = st.session_state.avisos
    pend = int(dfA[~dfA["Estado"].isin(["OT creada", "Cerrado", "Rechazado"])].shape[0]) if not dfA.empty else 0
    ots_prog = int((st.session_state.ots["EstadoOT"] != "Cerrada").sum()) if not st.session_state.ots.empty else 0

    st.markdown("### Bienvenido")
    if st.button("➕  Crear Aviso", key="btn_crear_aviso"):
        st.session_state.page = "Crear Aviso"
        st.rerun()
    if st.button(f"ℹ️  Backlog de Avisos\n{pend} Pendientes", key="btn_backlog_avisos"):
        st.session_state.page = "Backlog"
        st.rerun()
    if st.button(f"🔧  Backlog de OTs\n{ots_prog} Programadas", key="btn_backlog_ots"):
        st.session_state.page = "OTs"
        st.rerun()
    if st.button("📅  Planificación (Plan vs Ejecutado)", key="btn_planificacion"):
        st.session_state.page = "Planificación"
        st.rerun()
    if st.button("🛤️  Durmientes (Norma NS-01-01-00)", key="btn_durmientes"):
        st.session_state.page = "Durmientes"
        st.rerun()
    if st.button("🗂️  Control de Colleras (administración)", key="btn_control_colleras"):
        st.session_state.page = "Control Colleras"
        st.rerun()
    if st.button("📊  Reportes (Demo)", key="btn_reportes"):
        st.info("Próximamente: dashboard de KPIs.")

    st.markdown("##### Descargar histórico completo")
    # Se genera a pedido (no en cada recarga de Inicio): arma el Excel y el PDF de todo el
    # histórico y junta las fotos, recuperando desde Sheets las que falten en el disco.
    if st.button("📦  Preparar descarga del histórico", key="btn_preparar_respaldo"):
        with st.spinner("Leyendo el histórico desde Google Sheets..."):
            st.session_state.respaldo_excel = generar_respaldo_plano()
            st.session_state.respaldo_pdf = generar_respaldo_pdf()
            st.session_state.respaldo_fotos_zip = generar_respaldo_fotos_zip()
            st.session_state.respaldo_generado = now_str()
    if st.session_state.get("respaldo_generado"):
        st.caption(f"Histórico preparado el {st.session_state.respaldo_generado}.")
        if not (st.session_state.respaldo_excel or st.session_state.respaldo_fotos_zip):
            st.info("Todavía no hay Reportes Diarios en el histórico.")
        col_excel, col_pdf, col_fotos = st.columns(3)
        with col_excel:
            if st.session_state.respaldo_excel:
                st.download_button(
                    "💾  Excel",
                    data=st.session_state.respaldo_excel,
                    file_name="reportes_diarios_respaldo.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_backup_reportes", width="stretch",
                )
        with col_pdf:
            if st.session_state.respaldo_pdf:
                st.download_button(
                    "📄  PDF",
                    data=st.session_state.respaldo_pdf,
                    file_name="reportes_diarios_respaldo.pdf",
                    mime="application/pdf",
                    key="btn_backup_reportes_pdf", width="stretch",
                )
        with col_fotos:
            if st.session_state.respaldo_fotos_zip:
                st.download_button(
                    "🖼️  Fotos (ZIP)",
                    data=st.session_state.respaldo_fotos_zip,
                    file_name="reportes_diarios_fotos.zip",
                    mime="application/zip",
                    key="btn_backup_fotos", width="stretch",
                )

    if st.session_state.perfil == "Sacyr":  # herramientas de recuperación: solo Sacyr (no ADI, EFE ni Subcontrato)
        render_avanzado_sacyr()

    st.markdown("<div style='text-align:center;opacity:.5;padding-top:10px;'>sacyr</div>", unsafe_allow_html=True)

def render_avanzado_sacyr():
    with st.expander("⚙️ Avanzado"):
        st.caption(
            "Si un Reporte Diario aparece en el respaldo pero no en el Backlog de Avisos, "
            "es porque el Aviso quedó atrapado antes de que Avisos también quedara guardado "
            "en Google Sheets. Este botón recrea los avisos que falten (quedan como 'Nuevo', "
            "listos para validar; el estado de aprobación anterior no es recuperable)."
        )
        if st.button("🔄 Recuperar avisos faltantes desde Reportes Diarios", key="btn_reconciliar_avisos"):
            n = reconciliar_avisos_faltantes()
            if n:
                st.success(f"Se recrearon {n} aviso(s) que faltaban ✅")
            else:
                st.info("No hay avisos faltantes por recuperar.")
            st.rerun()

        st.divider()
        st.caption(
            "Las fotos de reportes guardados antes del respaldo en Google Sheets existen solo "
            "en el disco del servidor (se pierden si Streamlit Cloud reinicia). Este botón sube "
            "a Sheets las que todavía estén en el disco."
        )
        if st.button("🖼️ Respaldar fotos locales en Google Sheets", key="btn_subir_fotos_locales"):
            with st.spinner("Subiendo fotos..."):
                n = subir_fotos_locales_faltantes()
            if n:
                st.success(f"Se respaldaron {n} foto(s) ✅")
            else:
                st.info("No hay fotos locales pendientes de respaldar (o Google Sheets no está configurado).")

def validador_backlog():
    app_header("Backlog de Avisos", back_page="Inicio", right_icon="➕")
    perfil_bar()
    df = st.session_state.avisos.copy()
    tab1, tab2, tab3 = st.tabs(["Pendientes", "OT Creada", "Cerrados"])

    def render_list(filtered_df):
        if filtered_df.empty:
            st.info("No hay registros para mostrar.")
            return
        for _, row in filtered_df.iterrows():
            colA, colB = st.columns([5, 1])
            with colA:
                render_aviso_row(row)
            with colB:
                badge(row["Prioridad"], priority_kind(row["Prioridad"]))
                if st.button("Ver ›", key=f"ver_{row['AvisoID']}"):
                    st.session_state.selected_aviso = row["AvisoID"]
                    st.session_state.page = "Detalle"
                    st.rerun()

    with tab1:
        render_list(df[~df["Estado"].isin(["OT creada", "Cerrado", "Rechazado"])])
    with tab2:
        render_list(df[df["Estado"] == "OT creada"])
    with tab3:
        render_list(df[df["Estado"].isin(["Cerrado", "Rechazado"])])

def validador_detalle():
    app_header("Detalle Aviso / Crear OT", back_page="Backlog", right_icon="⋮")
    perfil_bar()
    aviso_id = st.session_state.selected_aviso
    if not aviso_id:
        st.info("Selecciona un aviso desde el Backlog primero.")
        return

    df = st.session_state.avisos
    row = df[df["AvisoID"] == aviso_id].iloc[0].to_dict()

    tiene_reporte = bool(row.get("ReporteID"))
    if tiene_reporte:
        tab_detalle, tab_reporte = st.tabs(["📋 Detalle del Aviso", "👁️ Ver Reporte Diario"])
    else:
        tab_detalle = st.container()
        tab_reporte = None

    with tab_detalle:
        estado_kind = {"Nuevo": "bad", "En validación": "info", "Observado": "warn",
                       "Validado": "ok", "OT creada": "ok", "Cerrado": "muted", "Rechazado": "bad"}
        c1, c2 = st.columns([3, 2])
        with c1:
            badge(row["Estado"], estado_kind.get(row["Estado"], "info"))
        with c2:
            st.markdown(f"<div style='text-align:right;color:#888;font-size:13px;'>ID: {row['AvisoID']}</div>", unsafe_allow_html=True)

        st.write("")
        st.markdown(
            f"""
            **Detectado:** {row['CreadoPor']} ({row['RolCreador']})
            **Activo:** {row['Activo']}
            **Ubicación:** {row['Ubicacion']}
            **Prioridad:** {row['Prioridad']}
            **Tipificación:** {row['Tipificacion']}
            """
        )
        if row["Adjunto"]:
            st.caption(f"📎 Adjunto: {row['Adjunto']}")
        if tiene_reporte:
            st.caption(f"🔗 Vinculado al Reporte Diario {row['ReporteID']} · ver pestaña 'Ver Reporte Diario'")
        st.caption(f"📄 {row['AvisoID']} · {row['FechaHora']} &nbsp; 🕐 {elapsed_str(row['FechaHora'])}")

        st.divider()

        rol_final = row.get("RolFinal") or "EFE"
        col_aprob, col_obs_name, col_fecha, col_resp = final_cols(rol_final)

        st.markdown(f"## Validación (Sacyr / ADI / {rol_final})")
        c1, c2, c3 = st.columns(3)
        with c1:
            badge("Sacyr", "ok" if row["AprobadoSacyr"] else "muted")
            st.caption(row["FechaValSacyr"] or "—")
        with c2:
            badge("ADI (ITO)", "ok" if row["AprobadoADI"] else "muted")
            st.caption(row["FechaValADI"] or "—")
        with c3:
            badge(rol_final, "ok" if row[col_aprob] else "muted")
            st.caption(row[col_fecha] or "—")

        st.write("")
        role = st.session_state.perfil
        can_do = can_validate(role, row)
        obs = st.text_area("Observación / comentario", height=80, key=f"obs_{aviso_id}_{role}")

        obs_col_map = {"Sacyr": "ObsSacyr", "ADI (ITO)": "ObsADI", rol_final: col_obs_name}

        colA, colB, colC = st.columns(3)
        with colA:
            if st.button("✅ Aprobar", disabled=not can_do, key="btn_aprobar"):
                if role == "Sacyr":
                    st.session_state.avisos.loc[df["AvisoID"] == aviso_id, ["AprobadoSacyr", "ObsSacyr", "FechaValSacyr", "RespValSacyr"]] = [True, obs, now_str(), st.session_state.usuario]
                elif role == "ADI (ITO)":
                    st.session_state.avisos.loc[df["AvisoID"] == aviso_id, ["AprobadoADI", "ObsADI", "FechaValADI", "RespValADI"]] = [True, obs, now_str(), st.session_state.usuario]
                elif role == rol_final:
                    st.session_state.avisos.loc[df["AvisoID"] == aviso_id, [col_aprob, col_obs_name, col_fecha, col_resp]] = [True, obs, now_str(), st.session_state.usuario]
                move_to_next_level(aviso_id, role)
                save_all_avisos()
                st.success("Aprobación registrada ✅")
                st.rerun()
        with colB:
            if st.button("🟨 Observar", disabled=not can_do, key="btn_observar"):
                st.session_state.avisos.loc[df["AvisoID"] == aviso_id, "Estado"] = "Observado"
                st.session_state.avisos.loc[df["AvisoID"] == aviso_id, "ValidacionNivel"] = "Sacyr"
                st.session_state.avisos.loc[df["AvisoID"] == aviso_id, obs_col_map[role]] = obs
                save_all_avisos()
                st.warning("Aviso observado. Vuelve a Sacyr para completar información.")
                st.rerun()
        with colC:
            if st.button("⛔ Rechazar", disabled=not can_do, key="btn_rechazar"):
                st.session_state.avisos.loc[df["AvisoID"] == aviso_id, "Estado"] = "Rechazado"
                st.session_state.avisos.loc[df["AvisoID"] == aviso_id, obs_col_map[role]] = obs
                save_all_avisos()
                st.error("Aviso rechazado.")
                st.rerun()

        st.divider()
        st.markdown("## Crear OT")
        df = st.session_state.avisos
        row = df[df["AvisoID"] == aviso_id].iloc[0].to_dict()

        if row["OT_Creada"]:
            st.info("Este aviso ya tiene una OT asociada. Ve a Backlog de OTs.")
        elif row["Estado"] != "Validado":
            st.info(f"⚠️ Para crear OT, el Aviso debe estar **Validado** (Sacyr → ADI → {rol_final}).")
        else:
            with st.form("form_ot"):
                responsable = st.selectbox("Responsable", ["Camila Pérez", "Jefe Mantenimiento", "Planner", "Técnico 1"])
                fecha_prog_d = st.date_input("Fecha Programada (día)", datetime.now().date())
                fecha_prog_t = st.time_input("Hora Programada", datetime.now().time().replace(second=0, microsecond=0))
                submitted_ot = st.form_submit_button("Iniciar OT", key="btn_iniciar_ot")
            if submitted_ot:
                fecha_prog = datetime.combine(fecha_prog_d, fecha_prog_t)
                ot_id = next_ot_id()
                new_ot = {"OTID": ot_id, "AvisoID": row["AvisoID"], "Activo": row["Activo"],
                          "Responsable": responsable, "FechaProgramada": fecha_prog.strftime("%Y-%m-%d %H:%M"), "EstadoOT": "Programada"}
                st.session_state.ots = pd.concat([st.session_state.ots, pd.DataFrame([new_ot])], ignore_index=True)
                st.session_state.avisos.loc[df["AvisoID"] == row["AvisoID"], "OT_Creada"] = True
                st.session_state.avisos.loc[df["AvisoID"] == row["AvisoID"], "Estado"] = "OT creada"
                save_all_avisos()
                save_all_ots()
                st.success(f"OT {ot_id} creada ✅")
                st.session_state.page = "OTs"
                st.rerun()

        st.write("")
        if st.button("✅ Cerrar Aviso", key="btn_cerrar"):
            st.session_state.avisos.loc[st.session_state.avisos["AvisoID"] == aviso_id, "Estado"] = "Cerrado"
            save_all_avisos()
            st.success("Aviso cerrado.")

    if tab_reporte is not None:
        with tab_reporte:
            st.caption("Solo lectura — esto es exactamente lo que ingresó Personal de Terreno.")
            mostrar_detalle_reporte_diario(row["ReporteID"])

def validador_ots():
    app_header("Backlog de OTs", back_page="Inicio", right_icon="📊")
    perfil_bar()
    df = st.session_state.ots.copy()
    if df.empty:
        st.info("Aún no hay OTs. Crea una desde un Aviso validado.")
        return
    estado_kind = {"Programada": "info", "En ejecución": "warn", "Cerrada": "muted"}
    for _, r in df.iterrows():
        st.markdown(f"**{r['OTID']}**")
        badge(r["EstadoOT"], estado_kind.get(r["EstadoOT"], "info"))
        st.write(f"Activo: {r['Activo']}")
        st.caption(f"Aviso: {r['AvisoID']} · Responsable: {r['Responsable']} · Programada: {r['FechaProgramada']}")
        cols = st.columns(3)
        with cols[0]:
            if st.button("▶ Ejecutar", key=f"exec_{r['OTID']}"):
                st.session_state.ots.loc[st.session_state.ots["OTID"] == r["OTID"], "EstadoOT"] = "En ejecución"
                save_all_ots()
                st.rerun()
        with cols[1]:
            if st.button("✅ Cerrar", key=f"close_{r['OTID']}"):
                st.session_state.ots.loc[st.session_state.ots["OTID"] == r["OTID"], "EstadoOT"] = "Cerrada"
                save_all_ots()
                st.rerun()
        with cols[2]:
            if st.button("🧾 Ver", key=f"seeav_{r['OTID']}"):
                st.session_state.selected_aviso = r["AvisoID"]
                st.session_state.page = "Detalle"
                st.rerun()
        st.divider()

def page_planificacion():
    app_header("Planificación", back_page="Inicio")
    perfil_bar()

    with st.expander("📤 Cargar planificación completa desde CSV (reemplaza TODO lo guardado)"):
        st.caption(
            "El archivo debe tener las columnas Actividad, Unidad, Anio, Mes, CantidadPlanificada "
            "(una fila por actividad y mes). Esto reemplaza toda la hoja PlanMensual — todos los "
            "años y meses guardados hasta ahora — no solo el mes que se ve más abajo."
        )
        archivo_csv = st.file_uploader("Archivo CSV", type=["csv"], key="plan_csv_uploader")
        if archivo_csv is not None:
            try:
                df_csv = pd.read_csv(archivo_csv)
                columnas_esperadas = {"Actividad", "Unidad", "Anio", "Mes", "CantidadPlanificada"}
                faltantes = columnas_esperadas - set(df_csv.columns)
                if faltantes:
                    st.error(f"Faltan columnas en el archivo: {', '.join(sorted(faltantes))}")
                else:
                    st.dataframe(df_csv.head(10), hide_index=True, width="stretch")
                    st.caption(f"{len(df_csv)} filas detectadas.")
                    if st.button("⚠️ Reemplazar toda la planificación con este archivo", key="btn_cargar_plan_csv"):
                        filas_csv = [
                            [r["Actividad"], r["Unidad"], int(r["Anio"]), int(r["Mes"]),
                             float(str(r["CantidadPlanificada"]).replace(",", "."))]
                            for _, r in df_csv.iterrows()
                        ]
                        guardar_plan_mensual(filas_csv)
                        st.success(f"Planificación cargada: {len(filas_csv)} filas ✅")
                        st.rerun()
            except Exception as e:
                st.error(f"No se pudo leer el archivo: {e}")

    hoy = datetime.now()
    c1, c2 = st.columns(2)
    with c1:
        anio = st.selectbox("Año", [2025, 2026, 2027], index=1, key="plan_anio")
    with c2:
        mes = st.selectbox(
            "Mes", list(range(1, 13)), index=hoy.month - 1,
            format_func=lambda m: MESES_ES[m - 1], key="plan_mes",
        )

    df_plan = _leer_hoja_df("PlanMensual")
    plan_mes_df = pd.DataFrame(columns=REPORTES_SHEETS["PlanMensual"])
    if not df_plan.empty:
        plan_mes_df = df_plan[
            (df_plan["Anio"].astype(str) == str(anio)) & (df_plan["Mes"].astype(str) == str(mes))
        ]
    plan_actual = dict(zip(plan_mes_df["Actividad"], plan_mes_df["CantidadPlanificada"].apply(_num)))
    unidad_actual = dict(zip(plan_mes_df["Actividad"], plan_mes_df["Unidad"]))

    st.markdown(f"#### Meta planificada — {MESES_ES[mes - 1]} {anio}")
    st.caption("Edita la cantidad planificada de cada actividad para este mes y guarda. Estos valores los define Sacyr según la programación real (no vienen precargados).")

    filas_editor = [
        {
            "Actividad": act,
            "Unidad": unidad_actual.get(act) or UNIDADES_TRABAJO[0],
            "Planificado": float(plan_actual.get(act, 0.0)),
        }
        for act in ACTIVIDADES_TRABAJO
    ]
    df_editado = st.data_editor(
        pd.DataFrame(filas_editor),
        hide_index=True, width="stretch", key=f"plan_editor_{anio}_{mes}",
        column_config={
            "Actividad": st.column_config.TextColumn(disabled=True),
            "Unidad": st.column_config.SelectboxColumn(options=UNIDADES_TRABAJO),
            "Planificado": st.column_config.NumberColumn(min_value=0.0, step=0.001, format="%.3f"),
        },
    )

    if st.button("💾 Guardar planificación del mes", key="btn_guardar_plan", width="stretch"):
        filas_mes = [
            [row["Actividad"], row["Unidad"], anio, mes, row["Planificado"]]
            for _, row in df_editado.iterrows()
        ]
        otras_filas = []
        if not df_plan.empty:
            resto = df_plan[~((df_plan["Anio"].astype(str) == str(anio)) & (df_plan["Mes"].astype(str) == str(mes)))]
            otras_filas = resto.values.tolist()
        guardar_plan_mensual(otras_filas + filas_mes)
        st.success("Planificación guardada ✅")
        st.rerun()

    st.divider()
    st.markdown("#### Plan vs Ejecutado")
    st.caption("«Ejecutado» se calcula automáticamente sumando lo registrado en los Reportes Diarios del período.")
    vista = st.segmented_control(
        "Período", ["Semanal", "Mensual", "Anual (Sep–Ago)"], default="Mensual", key="plan_vista",
    ) or "Mensual"

    if vista == "Semanal":
        semanas = semanas_del_mes(anio, mes)
        hoy_d = hoy.date()
        idx_hoy = next((i for i, (a, b) in enumerate(semanas) if a <= hoy_d <= b), 0)
        idx = st.selectbox(
            "Semana", list(range(len(semanas))), index=idx_hoy, key=f"plan_semana_{anio}_{mes}",
            format_func=lambda i: f"Semana {i + 1} · {semanas[i][0]:%d/%m} al {semanas[i][1]:%d/%m}",
        )
        desde, hasta = semanas[idx]
        st.caption(
            f"Planificado de la semana = plan de {MESES_ES[mes - 1]} prorrateado por día "
            f"({(hasta - desde).days + 1} de {calendar.monthrange(anio, mes)[1]} días)."
        )
        df_comp = tabla_plan_vs_ejecutado(plan_por_rango(df_plan, desde, hasta), ejecutado_por_rango(desde, hasta), unidad_actual)
        st.dataframe(df_comp, hide_index=True, width="stretch")

        st.markdown(f"##### Ejecutado por semana — {MESES_ES[mes - 1]} {anio}")
        st.caption("Cada celda: ejecutado / planificado de esa semana.")
        pivote = {"Actividad": []}
        datos_semanas = [(plan_por_rango(df_plan, a, b), ejecutado_por_rango(a, b)) for a, b in semanas]
        actividades = [a for a in df_comp["Actividad"]
                       if any(p.get(a, 0) or e.get(a, 0) for p, e in datos_semanas)]
        pivote["Actividad"] = actividades
        for i, (a, b) in enumerate(semanas):
            p, e = datos_semanas[i]
            pivote[f"S{i + 1} ({a:%d}–{b:%d})"] = [f"{e.get(act, 0.0):.2f} / {p.get(act, 0.0):.2f}" for act in actividades]
        if actividades:
            st.dataframe(pd.DataFrame(pivote), hide_index=True, width="stretch")
        else:
            st.caption("Sin planificación ni ejecución registradas en este mes.")

    elif vista == "Mensual":
        ejecutado = ejecutado_por_actividad(anio, mes)
        meta_anual = meta_anual_por_actividad(df_plan, anio, mes)
        df_comp = tabla_plan_vs_ejecutado(plan_actual, ejecutado, unidad_actual, extra={"Meta Anual": meta_anual})
        st.dataframe(df_comp, hide_index=True, width="stretch")

    else:
        ciclo = _ciclo_anual(anio, mes)
        desde = date(ciclo[0][0], 9, 1)
        hasta = date(ciclo[-1][0], 8, 31)
        corte = min(max(hoy.date(), desde), hasta)
        st.caption(
            f"Ciclo {desde:%m/%Y} – {hasta:%m/%Y}. «Planificado a la fecha» = plan acumulado hasta el "
            f"{corte:%d/%m/%Y}; el Avance % compara lo ejecutado contra eso. «% de la Meta» = ejecutado / meta anual."
        )
        meta_anual = plan_por_rango(df_plan, desde, hasta)
        ejecutado = ejecutado_por_rango(desde, corte)
        unidades_ciclo = {}
        if not df_plan.empty:
            unidades_ciclo = dict(zip(df_plan["Actividad"], df_plan["Unidad"]))
        df_comp = tabla_plan_vs_ejecutado(plan_por_rango(df_plan, desde, corte), ejecutado,
                                          unidades_ciclo, extra={"Meta Anual": meta_anual})
        df_comp = df_comp.rename(columns={"Planificado": "Planificado a la fecha"})
        df_comp["% de la Meta"] = [
            round(e / m * 100, 1) if m > 0 else 0.0 for e, m in zip(df_comp["Ejecutado"], df_comp["Meta Anual"])
        ]
        st.dataframe(df_comp, hide_index=True, width="stretch")
        df_comp = df_comp.rename(columns={"Planificado a la fecha": "Planificado"})

    grafico = grafico_avance(df_comp)
    if grafico is not None:
        st.markdown("##### Avance por actividad")
        st.altair_chart(grafico, width="stretch")

def page_dashboard_durmientes():
    app_header("Control de Durmientes", back_page="Inicio")
    perfil_bar()
    st.caption("Cumplimiento de la Norma NS-01-01-00: mínimo 10 durmientes efectivos por collera y racha máxima de 'Malo' consecutivos (2 en curva, 3 en recta).")

    resumen = resumen_global_durmientes()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Colleras registradas", resumen["total"])
    c2.metric("Cumplen", resumen["cumplen"])
    c3.metric("No cumplen", resumen["no_cumplen"])
    c4.metric("% cumplimiento", f"{resumen['pct_cumplimiento']:.0f}%")
    st.caption("% cumplimiento sobre colleras evaluables (Cumple + No Cumple). 'No cumplen' incluye 'No Cumple (racha)'.")
    d1, d2, d3, d4, d5 = st.columns(5)
    d1.metric("Buenos", resumen["total_buenos"])
    d2.metric("Malos", resumen["total_malos"])
    d3.metric("Nuevos", resumen["total_nuevos"])
    d4.metric("Reempl.", resumen["total_reempl"])
    d5.metric("% renov.", f"{resumen['pct_renov_global']:.1f}%")

    st.divider()
    st.markdown("#### Distribución de colleras")
    df_torta = pd.DataFrame({
        "Estado": ["Cumple", "No Cumple", "Sin datos suficientes"],
        "Colleras": [resumen["cumplen"], resumen["no_cumplen"], resumen["sin_datos"]],
    })
    grafico_torta = alt.Chart(df_torta).mark_arc(innerRadius=60).encode(
        theta=alt.Theta("Colleras:Q"),
        color=alt.Color("Estado:N", scale=alt.Scale(
            domain=["Cumple", "No Cumple", "Sin datos suficientes"],
            range=["#2FA84F", "#DC3545", "#B0B5BB"],
        ), legend=alt.Legend(title=None)),
        tooltip=["Estado:N", "Colleras:Q"],
    ).properties(height=260)
    st.altair_chart(grafico_torta, width="stretch")

    st.divider()
    if st.button("🗂️ Ver / editar colleras por PK", key="btn_ir_control_colleras", width="stretch"):
        st.session_state.page = "Control Colleras"
        st.rerun()

    st.divider()
    st.markdown("#### Mapa de condición por PK")
    st.caption("Toca un PK para ir directo a sus colleras en Control de Colleras.")
    color = {"Bueno": "#2FA84F", "Regular": "#E0A458", "Malo": "#DC3545", "Sin Datos": "#B0B5BB"}
    filas = resumen_por_pk()
    css_mapa_pk = "".join(
        f".st-key-mapa_pk_{f['PK']} button {{ background:{color[f['Estado KM']]} !important; "
        f"color:#fff !important; border:none !important; padding:4px 0 !important; "
        f"font-size:11px !important; font-weight:700 !important; }}\n"
        for f in filas
    )
    st.markdown(f"<style>{css_mapa_pk}</style>", unsafe_allow_html=True)
    cols = st.columns(len(filas))
    for col, f in zip(cols, filas):
        with col:
            if st.button(str(f["PK"]), key=f"mapa_pk_{f['PK']}", help=f"PK {f['PK']} · {f['Estado KM']}"):
                st.session_state.control_col_pk = f["PK"]
                st.session_state.page = "Control Colleras"
                st.rerun()

    st.divider()
    st.markdown("#### Detalle por PK")
    st.dataframe(pd.DataFrame(filas), hide_index=True, width="stretch")

def page_control_colleras():
    """Pantalla de administración: edita el estado de durmientes de cualquier collera
    directamente (sin pasar por un Reporte Diario). Separada a propósito del flujo de
    terreno -- el Reporte Diario sigue registrando una collera puntual, ligada a su
    ReporteID, para mantener esa trazabilidad. Los cambios de acá quedan en el mismo
    historial (durmientes_estado), solo que con un 'ReporteID' de ajuste manual."""
    app_header("Control de Colleras", back_page="Durmientes")
    perfil_bar()
    st.caption("Edición directa del estado de durmientes por collera. Los cambios quedan en el historial igual que los reportados en terreno.")

    df_colleras = cargar_todas_colleras()
    if df_colleras.empty:
        st.info("El catálogo de colleras está vacío.")
        return
    pks_disponibles = sorted(int(p) for p in df_colleras["PK"].unique())
    pk_default = st.session_state.get("control_col_pk", pks_disponibles[0])
    if pk_default not in pks_disponibles:
        pk_default = pks_disponibles[0]
    if "control_col_pk" not in st.session_state:
        st.session_state.control_col_pk = pk_default
    pk_sel = st.selectbox("PK", pks_disponibles, key="control_col_pk")

    colleras_pk = df_colleras[df_colleras["PK"] == pk_sel].sort_values("Collera")
    df_vigente = _estado_vigente(st.session_state.durmientes_estado)

    opciones_estado = ["", "Bueno", "Malo", "Nuevo", "Reemplazado"]
    filas_editor = []
    for _, c in colleras_pk.iterrows():
        collera_id = c["ColleraID"]
        estados = {}
        if not df_vigente.empty:
            filas_c = df_vigente[df_vigente["ColleraID"] == collera_id]
            estados = {int(r["Posicion"]): r["Estado"] for _, r in filas_c.iterrows()}
        fila = {"Collera": collera_id, "N°": int(c["Collera"]), "Ubicación": c["Ubicacion"]}
        for pos in range(1, 24):
            fila[f"D{pos}"] = estados.get(pos, "")
        filas_editor.append(fila)

    column_config = {
        "Collera": st.column_config.TextColumn(disabled=True, width="small"),
        "N°": st.column_config.NumberColumn(disabled=True, width="small"),
        "Ubicación": st.column_config.TextColumn(disabled=True, width="small"),
    }
    for pos in range(1, 24):
        column_config[f"D{pos}"] = st.column_config.SelectboxColumn(options=opciones_estado, width="small")

    st.caption(f"{len(filas_editor)} colleras en el PK {pk_sel}. Deja en blanco lo no inspeccionado.")
    df_editado = st.data_editor(
        pd.DataFrame(filas_editor), hide_index=True, width="stretch",
        key=f"control_colleras_editor_{pk_sel}", column_config=column_config,
    )

    if st.button("💾 Guardar cambios de este PK", key="btn_guardar_control_colleras", width="stretch"):
        fecha_str = datetime.now().strftime("%Y-%m-%d")
        reporte_ajuste = f"AJUSTE-MANUAL-{st.session_state.usuario}"
        secuencias = {
            fila["Collera"]: [v if isinstance(v, str) and v else None for v in (fila.get(f"D{p}") for p in range(1, 24))]
            for _, fila in df_editado.iterrows()
        }
        filas_nuevas = filas_cambios_durmientes(reporte_ajuste, fecha_str, secuencias)
        if not filas_nuevas:
            st.info("No hay cambios nuevos que guardar.")
        elif guardar_bloques({"DurmientesEstado": filas_nuevas}):
            persistir_cambios_durmientes_local(filas_nuevas)
            st.success(f"{len(filas_nuevas)} cambio(s) guardado(s) ✅")
            st.rerun()
        else:
            st.error("No se pudo guardar en Google Sheets. Tus cambios siguen en la tabla: intenta de nuevo.")

    st.divider()
    st.markdown("#### Resumen del PK")
    st.caption("Mismas columnas que la hoja 'Control Durmientes' del Excel.")
    estados = _estados_por_collera(_estado_vigente(st.session_state.durmientes_estado))
    filas_resumen = []
    for _, c in colleras_pk.iterrows():
        ev = evaluar_collera(c["ColleraID"], c["Ubicacion"], estados_por_collera=estados)
        filas_resumen.append({
            "Collera": c["ColleraID"], "Ubicación": c["Ubicacion"],
            "TOTAL REGISTRADO": ev["total_registrado"], "N° BUENOS": ev["n_buenos"], "N° MALOS": ev["n_malos"],
            "N° NUEVOS": ev["n_nuevos"], "N° REEMPL.": ev["n_reempl"], "EFECTIVOS (B+N+R)": ev["efectivos"],
            "% RENOV. (N+R)": round(ev["pct_renovacion"] * 100, 1),
            "MÍN. EFECTIVOS": ev["min_efectivos"], "RACHA CONSEC.": ev["racha_consec"],
            "Racha máx. Malo": ev["racha_max"], "ESTADO GENERAL": ev["estado_general"],
        })
    st.dataframe(pd.DataFrame(filas_resumen), hide_index=True, width="stretch")

def flujo_validador():
    page = st.session_state.page
    if page == "Inicio":
        validador_inicio()
    elif page == "Crear Aviso":
        page_crear_aviso()
    elif page == "Backlog":
        validador_backlog()
    elif page == "Detalle":
        validador_detalle()
    elif page == "OTs":
        validador_ots()
    elif page == "Planificación":
        page_planificacion()
    elif page == "Durmientes":
        page_dashboard_durmientes()
    elif page == "Control Colleras":
        page_control_colleras()
    else:
        st.session_state.page = "Inicio"
        st.rerun()
    render_pie_desarrollador()
    bottom_nav([("Inicio", "🏠"), ("Crear Aviso", "➕"), ("Backlog", "📋"), ("OTs", "🔧")])

# -------------------------
# Botón Atrás/Adelante del navegador
# -------------------------
_nav_historial = components.declare_component(
    "nav_historial", path=os.path.join(os.path.dirname(os.path.abspath(__file__)), "nav_historial"))

def sincronizar_historial_navegador():
    """Cada pantalla queda como una entrada del historial del navegador, así el botón
    Atrás/Adelante (o el gesto de volver del celular) lleva a la pantalla anterior en vez
    de salir de la app. El componente avisa la pantalla elegida con un valor {pagina, t};
    se lee desde session_state ANTES de dibujar nada, para que el componente reciba ya la
    pantalla nueva y no la vuelva a apilar en el historial."""
    valor = st.session_state.get("nav_historial")
    if isinstance(valor, dict) and valor.get("pagina") and valor.get("t") != st.session_state.get("nav_historial_t"):
        st.session_state.nav_historial_t = valor.get("t")
        st.session_state.page = valor["pagina"]
    # Contenedor fijo al inicio: el componente conserva siempre la misma posición y no se
    # vuelve a crear en cada cambio de pantalla (si se recreara, perdería el historial).
    with st.container(key="nav_historial_box"):
        try:
            _nav_historial(pagina=st.session_state.page, key="nav_historial", default=None)
        except Exception:
            pass  # sin el componente la app funciona igual, solo sin botón Atrás

sincronizar_historial_navegador()

# -------------------------
# Enrutamiento por perfil
# -------------------------
if st.session_state.perfil == "Personal Terreno":
    flujo_terreno()
else:
    flujo_validador()
