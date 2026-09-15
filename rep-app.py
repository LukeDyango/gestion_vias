"""
============================================================
 REPORTE DIARIO — Navegación jerárquica tipo WBS (Primavera P6)
------------------------------------------------------------
 La aplicación NO es un formulario largo: es un árbol de
 decisiones. Cada pantalla muestra una sola decisión y entre
 5 y 15 opciones grandes. El formulario del reporte aparece
 recién cuando se llega a un Step (nodo hoja).

     Login    Usuario y contraseña. El perfil Supervisor entra
              directo a SUS actividades: él es el supervisor
              del reporte, no hay a quién elegir. Solo los
              perfiles que ven todo (Administración, Oficina
              Técnica, Inspector) eligen a nombre de quién
              registran.
     Nivel 1  Obra          -> raíz del árbol
     Nivel 2  Tipo de obra
     Nivel 3  Obra individual
     Nivel 4  Actividad
     Nivel 5  Steps (padre/hijo, profundidad libre)
     -------- Formulario del Reporte Diario

 El árbol se guarda como lista de adyacencia
 (id, nombre, padre_id) igual que el WBS de Primavera, por lo
 que admite cualquier profundidad sin tocar el código. Del
 reporte se guarda el Step final Y la ruta completa que se
 recorrió para llegar a él (trazabilidad).

 Cada supervisor solo ve lo que se le asignó, y la asignación
 puede hacerse a cualquier altura del árbol —una obra entera,
 una actividad o un Step suelto— desde 👥 Asignar. El id de
 cada nodo es su ruta, así que la visibilidad se resuelve por
 prefijo.

 El reporte devuelve, además de lo que escribe el supervisor,
 el ID ACT y el COD STEP de la BD de actividades del programa
 y la HH calculada (jornada × dotación). Esos códigos no se
 piden en pantalla: se deducen de la rama recorrida.

 Ejecutar:  streamlit run rep-app.py
============================================================
"""


import base64
import hashlib
import io
import json
import os
import re
import sqlite3
import unicodedata
import uuid
from datetime import date, datetime

import streamlit as st

# Pillow llega junto con Streamlit, pero el sello de la foto no debe
# tumbar la app si el entorno viene recortado: sin Pillow se guarda la
# foto tal cual y se avisa en pantalla.
try:
    from PIL import Image, ImageDraw, ImageFont, ImageOps
    HAY_PILLOW = True
except ImportError:                                  # pragma: no cover
    HAY_PILLOW = False

# Google Sheets es el histórico durable cuando la app corre en
# Streamlit Cloud: allá el disco es efímero y la base SQLite se pierde
# en cada reinicio. Sin estas librerías la app funciona igual, solo que
# guardando en local (que es lo que corresponde en el PC de la oficina).
try:
    import gspread
    from google.oauth2.service_account import Credentials as CredencialesGoogle
    HAY_GSPREAD = True
except ImportError:                                  # pragma: no cover
    HAY_GSPREAD = False

st.set_page_config(
    page_title="Reporte Diario · SACYR",
    page_icon="🏗️",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ============================================================
# 1. Estilos — pensados para uso en terreno con guantes:
#    objetivos táctiles grandes, alto contraste, una columna.
# ============================================================

MOBILE_CSS = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header[data-testid="stHeader"] {background: transparent; height: 0.5rem;}

.block-container {padding-top: 0.5rem; padding-bottom: 4rem; max-width: 480px;}

/* ---- Barra superior ---- */
.app-header {
    background: #10263F; color: #fff; display: flex; align-items: center;
    justify-content: space-between; min-height: 50px; padding: 0 14px;
    box-sizing: border-box; border-radius: 14px;
}
.app-header.header-right { border-radius: 0 14px 14px 0; }
.app-header .title { font-weight: 700; font-size: 17px; line-height: 1.15; }
.app-header .sub { font-size: 11.5px; opacity: .75; letter-spacing: .3px; }
.app-header .brand { font-size: 11px; font-weight: 700; letter-spacing: 1px; opacity: .9; }

/* La flecha y el título forman una sola barra: sin separación entre
   columnas y sin relleno propio, para que no quede un hueco blanco
   que además roba superficie táctil al botón de volver. */
.st-key-header_row div[data-testid="stHorizontalBlock"] { gap: 0 !important; }
.st-key-header_row div[data-testid="stColumn"] { padding: 0 !important; }
.st-key-btn_back, .st-key-btn_back div.stButton { width: 100% !important; }
.st-key-btn_back button {
    background: #10263F !important; color: #fff !important; border: none !important;
    border-radius: 14px 0 0 14px !important; font-size: 22px !important; font-weight: 700 !important;
    min-height: 50px !important; width: 100% !important; padding: 0 !important;
    box-shadow: none !important; margin: 0 !important;
}
.st-key-btn_back button:hover { background: #1D3D63 !important; color: #fff !important; }

/* ---- Migas de pan: la ruta recorrida dentro del proyecto ---- */
.breadcrumb {
    background: #F1F4F8; border-left: 4px solid #E8501E; border-radius: 8px;
    padding: 8px 12px; margin: 10px 0 4px 0; font-size: 12.5px; color: #33475B;
    line-height: 1.5; word-break: break-word;
}
.breadcrumb b { color: #10263F; }
.pregunta { font-size: 15.5px; font-weight: 700; color: #10263F; margin: 12px 0 6px 0; }

/* ---- Botones generales ---- */
div.stButton > button, div.stFormSubmitButton > button, div.stDownloadButton > button {
    border-radius: 10px; padding: 0.6rem 1rem; font-weight: 600; border: 1px solid #DCDFE4;
}

/* ---- Opciones del árbol: filas grandes, texto a la izquierda ---- */
div[class*="st-key-nodo_"] button {
    min-height: 58px !important; text-align: left !important; justify-content: flex-start !important;
    background: #fff !important; border: 1px solid #DCDFE4 !important; color: #10263F !important;
    font-size: 15px !important; font-weight: 600 !important; border-left: 5px solid #10263F !important;
}
div[class*="st-key-nodo_"] button:hover { border-color: #E8501E !important; background: #FFF7F3 !important; }
/* Los Steps (hojas) se distinguen: son los que abren el formulario */
div[class*="st-key-step_"] button {
    min-height: 58px !important; text-align: left !important; justify-content: flex-start !important;
    background: #FFF7F3 !important; border: 1px solid #F3C4AF !important; color: #8A2B00 !important;
    font-size: 15px !important; font-weight: 600 !important; border-left: 5px solid #E8501E !important;
}
div[class*="st-key-login_"] button {
    min-height: 56px !important; text-align: left !important; justify-content: flex-start !important;
}
div[class*="st-key-sup_"] button {
    min-height: 56px !important; text-align: left !important; justify-content: flex-start !important;
    border-left: 5px solid #2FA84F !important;
}
.st-key-btn_guardar button, .st-key-btn_entrar button {
    background: #2FA84F !important; color: #fff !important; border: none !important;
    min-height: 52px !important; font-size: 16px !important;
}
.st-key-btn_salir button { background: #F1F3F6 !important; color: #6C757D !important; }

/* ---- Tarjeta de contexto del formulario ---- */
.ctx-card {
    background: #10263F; color: #fff; border-radius: 12px; padding: 12px 14px; margin-bottom: 10px;
}
.ctx-card .step { font-size: 17px; font-weight: 700; margin-bottom: 4px; }
.ctx-card .ruta { font-size: 11.5px; opacity: .8; line-height: 1.5; word-break: break-word; }
.ctx-card .meta { font-size: 11.5px; opacity: .95; margin-top: 6px; }
/* Códigos del programa (ID ACT · COD STEP): referencia, no dato a llenar */
.ctx-card .codigos {
    font-family: ui-monospace, Menlo, Consolas, monospace; font-size: 11px;
    opacity: .8; margin-top: 4px; letter-spacing: .3px;
}
/* Estado en que está la actividad, junto a los códigos. Se muestra, no
   se edita: cerrar una actividad será un paso aparte. */
.ctx-card .estado {
    display: inline-block; margin-left: 8px; border-radius: 999px;
    padding: 1px 10px; font-size: 10.5px; font-weight: 700;
    letter-spacing: .4px; text-transform: uppercase; vertical-align: middle;
    font-family: system-ui, sans-serif;
}
/* Día y horas de la jornada: explica de dónde sale el número de horas */
.ctx-card .jornada {
    display: inline-block; margin-top: 8px; background: #E8501E; color: #fff;
    border-radius: 999px; padding: 3px 12px; font-size: 11.5px; font-weight: 700;
    letter-spacing: .4px; text-transform: uppercase;
}

/* Hoja (Step) en la pantalla de asignaciones: no se navega, solo se marca */
.asig-hoja {
    padding: 10px 12px; border: 1px dashed #DCDFE4; border-radius: 10px;
    font-size: 14px; color: #8A2B00; background: #FFF7F3; word-break: break-word;
}

.pill { display: inline-block; padding: 3px 12px; border-radius: 999px; font-size: 12px; font-weight: 700; }
.rep-row { padding: 10px 0; border-bottom: 1px solid #EEE; }
.rep-title { font-weight: 700; font-size: 14.5px; color: #10263F; }
.rep-meta { color: #888; font-size: 12px; margin-top: 2px; }
.hint { color: #6C757D; font-size: 12.5px; }

/* Resultado de la búsqueda por ID: la ruta y los identificadores van
   debajo del nombre, en chico, porque lo que se busca es la actividad */
.busq-ruta {
    color: #6C757D; font-size: 11.5px; line-height: 1.5;
    margin: -6px 0 10px 6px; word-break: break-word;
}
.busq-ruta code { font-size: 10.5px; color: #10263F; background: #F1F4F8;
                  padding: 1px 5px; border-radius: 4px; }
div[class*="st-key-busq_"] button {
    min-height: 44px !important; text-align: left !important;
    justify-content: flex-start !important; font-size: 14px !important;
    border-left: 4px solid #E8501E !important;
}

/* ---- Responsive: en pantallas angostas todo va en una columna ---- */
@media (max-width: 480px) {
    div[data-testid="stHorizontalBlock"] { flex-wrap: wrap; }
    div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] { min-width: 100% !important; flex: 1 1 100% !important; }
    .st-key-header_row div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] { min-width: 0 !important; }
    .st-key-bottom_nav div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] { min-width: 0 !important; flex: 1 1 0 !important; }
}
</style>
"""
st.markdown(MOBILE_CSS, unsafe_allow_html=True)

# ============================================================
# 2. Constantes de negocio
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "rep_app.db")
FOTOS_DIR = os.path.join(BASE_DIR, "fotos_reportes", "rep-app")

EMPRESAS = ["Sacyr", "Subcontrato A", "Subcontrato B", "Subcontrato C"]

# Los permisos quedan definidos desde el login (Nivel 0).
#   ve_todo      : el perfil no está limitado a obras asignadas
#   puede_reportar: puede crear reportes diarios
PERFILES = {
    "Supervisor":      {"ve_todo": False, "puede_reportar": True,  "color": "#2FA84F"},
    "Administrador":   {"ve_todo": True,  "puede_reportar": True,  "color": "#10263F"},
    "Oficina Técnica": {"ve_todo": True,  "puede_reportar": True,  "color": "#7B4FBE"},
    "Inspector":       {"ve_todo": True,  "puede_reportar": False, "color": "#E0A458"},
}

ESTADOS = ["En ejecución", "Terminada", "Paralizada", "No ejecutada"]

# ============================================================
#  MARCHA BLANCA — campos que hoy no se usan
#  ------------------------------------------------------------
#  Se ocultan del formulario, no se borran: la columna sigue en la
#  base y en el envío a SharePoint, así que activar cualquiera de
#  estos vuelve a mostrarlo sin migrar nada.
# ============================================================
MOSTRAR_NOMBRES_PERSONAL = False   # textarea "Nombres del personal"
MOSTRAR_LINEA_BASE = False         # "· línea base X HH (N% consumido)"
MOSTRAR_ESTADO_EDITABLE = False    # selector "Estado de la actividad"
MOSTRAR_MATERIALES = False         # textarea "Materiales"
MOSTRAR_SEGURIDAD = False          # bloque "Seguridad": incidencia y riesgos

# Con el selector oculto, todo reporte nuevo se guarda con este estado.
# Cerrar una actividad será un paso aparte, no una casilla del reporte.
ESTADO_POR_DEFECTO = "En ejecución"
ESTADO_SIN_REPORTES = "Sin iniciar"

# Color de cada estado en la tarjeta de contexto.
COLOR_ESTADO = {
    "En ejecución": "#2FA84F",
    "Terminada": "#0B6BCB",
    "Paralizada": "#DC3545",
    "No ejecutada": "#6C757D",
    ESTADO_SIN_REPORTES: "#6C757D",
}

# ---------- Maquinaria por subcontrato ----------
# Catálogo inicial del parque de equipos: (EMPRESA, PATENTE, TIPO).
# Es solo la semilla —vive en la tabla `maquinaria`, no en el código—:
# desde el propio formulario se agrega la máquina que falte, sin tocar
# este archivo. La patente es el identificador: se pide en el formato
# de la planilla, en mayúsculas.
#
# El supervisor declara primero QUÉ SUBCONTRATO ejecutó el Step y recién
# entonces la lista de maquinaria se acota a la de esa empresa. Elegir
# entre 6 máquinas propias es una decisión; entre 70 de todo el
# proyecto, una fuente de errores.
MAQUINARIA_INICIAL = [
    ("CAMAN",       "VTTY-67", "EXCAVADORA"),
    ("TOLMOR",      "TSCG-69", "EXCAVADORA"),
    ("ICAFAL",      "SVZR-29", "CAMION"),
    ("ICAFAL",      "HWHS-54", "CAMION PLUMA"),
    ("ICAFAL",      "PLFL-22", "CAMION"),
    ("TOLMOR",      "HJGG-45", "RETROEXCAVADORA"),
    ("ICAFAL",      "RKKP-71", "RETROEXCAVADORA"),
    ("TOLMOR",      "KKTV-18", "CAMION PLUMA"),
    ("VHV",         "TSCP-60", "EXCAVADORA"),
    ("VHV",         "FHFR-11", "CAMION TOLVA"),
    ("VHV",         "FHGP-54", "CAMION TOLVA"),
    ("VHV",         "GYBW97",  "CAMION TOLVA"),
    ("VHV",         "GYPS-86", "CAMION TOLVA"),
    ("VHV",         "JHWL-91", "CAMION TOLVA"),
    ("VHV",         "JJWH-63", "CAMION TOLVA"),
    ("VHV",         "RDJD-99", "MOTONIVELADORA"),
    ("TOLMOR",      "SXGG-16", "RETROEXCAVADORA"),
    ("VHV",         "VHTJ-98", "EXCAVADORA"),
    ("VHV",         "VKFP-50", "EXCAVADORA"),
    ("VHV",         "VPWH-82", "RODILLO"),
    ("ICAFAL",      "JDCY-69", "BIVIAL MECALAC"),
    ("VHV",         "TZWK-29", "CAMION ALJIBE"),
    ("ICAFAL",      "JJGX60",  "RETROEXCAVADORA"),
    ("ICAFAL",      "PKPX-56", "CAMION 3/4"),
    ("ICAFAL",      "RCDB-64", "CAMION 3/4"),
    ("INGEO",       "CVGL-71", "CAMION"),
    ("INGEO",       "DDZV-30", "RETROEXCAVADORA"),
    ("VHV",         "PLVX-47", "CAMION TOLVA"),
    ("CO-OL",       "KJHW-40", "CAMION"),
    ("CO-OL",       "KJTS-67", "CAMION"),
    ("VHV",         "VDCW-50", "RODILLO"),
    ("TOLMOR",      "SZFT-54", "RETROEXCAVADORA"),
    ("ICAFAL",      "TZVD-74", "EQUIPO BIVIAL"),
    ("CAMAN",       "VRCC-81", "RETROEXCAVADORA"),
    ("CAMAN",       "TPRG-15", "CAMION"),
    ("CAMAN",       "VHZT-50", "CAMION TOLVA"),
    ("VHV",         "TWTR-29", "CAMION ALJIBE"),
    ("VHV",         "VGCP-68", "MOTONIVELADORA"),
    ("ICAFAL",      "VJHL-21", "EXCAVADORA"),
    ("ICAFAL",      "VJGV-76", "EXCAVADORA"),
    ("TOLMOR",      "TRYW-83", "RETROEXCAVADORA"),
    ("SETRASA",     "VSKP-26", "CAMION ALJIBE"),
    ("VHV",         "JGZV-48", "CAMION TOLVA"),
    ("VHV",         "LCTH-46", "CAMION TOLVA"),
    ("SETRASA",     "LBXF-67", "CAMION PLUMA"),
    ("SACYR",       "LRTP-80", "RODILLO"),
    ("VHV",         "TDCX-73", "RETROEXCAVADORA"),
    ("DOMATI",      "SPYD-48", "CAMION PLUMA"),
    ("CPL",         "SLWR-41", "CAMION PLANO"),
    ("SMI",         "SZBL-37", "CAMION"),
    ("SETRASA",     "LGYJ-43", "CAMION ALJIBE"),
    ("CPL",         "RWCC99",  "MINICARGADOR"),
    ("CPL",         "VGJD-26", "MINICARGADOR"),
    ("CPL",         "GVDB-63", "MINIEXCAVADORA"),
    ("ICAFAL",      "VJDX-23", "CAMION PLUMA"),
    ("CRISAN",      "VRLS-46", "CAMION"),
    ("TITAN",       "FSRX-45", "GRUA"),
    ("DOMATI",      "TTPS-73", "EXCAVADORA"),
    ("DOMATI",      "RJXS-80", "CAMION PLUMA"),
    ("TITAN",       "XN8324",  "CAMION PLUMA"),
    ("CHILERETROS", "VXPB-74", "EXCAVADORA"),
    ("CPL",         "TLXW-53", "CAMION PLUMA"),
    ("CHILERETROS", "VSSH-16", "RODILLO DYNAPAC"),
    ("CHILERETROS", "SPZY-17", "CAMION TOLVA"),
    ("SACYR",       "VTHV-44", "MANIPULADOR"),
    ("CRISAN",      "RHYV-60", "RETROEXCAVADORA"),
    ("CHILERETROS", "PFZF-38", "CAMION TOLVA"),
    ("SERSAN",      "LPST-42", "CARGADOR FRONTAL"),
    ("CHILERETROS", "LTJK-22", "EXCAVADORA"),
    ("WOMAN",       "LCSP-28", "CAMION PLUMA"),
    ("ENGINEX",     "TLXW-53", "CAMION PLUMA"),
]

# Jornada contractual: las horas NO se escriben ni se eligen. El reporte
# es del día en que se hace, y el día de la semana fija las horas. Es lo
# que permite que el formulario no tenga apartado de jornada: un dato
# menos que llenar en terreno y uno menos que corregir en oficina.
#   0 = lunes … 6 = domingo
HORAS_POR_DIA = {0: 9.0, 1: 9.0,            # lunes y martes
                 2: 8.0, 3: 8.0, 4: 8.0,    # miércoles a viernes
                 5: 6.0, 6: 6.0}            # sábado y domingo
DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves",
               "Viernes", "Sábado", "Domingo"]

# ---------- Sello de la foto ----------
# La evidencia se marca sobre la imagen como lo hace una cámara GPS:
# el dato viaja pegado a la foto y ya no depende de que alguien lo
# transcriba bien en la planilla.
SELLO_POR_DEFECTO = {
    "camino": "Ruta G-78",
    "comuna": "Melipilla",
    "region": "Región Metropolitana de Santiago",
    "pais": "Chile",
    "altitud": 0.0,
}
ORIENTACIONES = ["—", "N", "NE", "E", "SE", "S", "SO", "O", "NO"]

# Umbral a partir del cual la pantalla ofrece buscador en vez de
# obligar a leer la lista completa (regla: 5 a 15 opciones visibles).
MAX_OPCIONES_SIN_BUSCADOR = 12

# ---------- Envío a SharePoint ----------
# Mismo enfoque que la app Reporte-diario: la app NUNCA habla directo
# con SharePoint ni guarda credenciales. Hace POST de un JSON plano a
# un webhook —el disparador HTTP de un flujo de Power Automate, o el
# Web App de un Google Apps Script que escribe en una hoja— y es ese
# flujo el que crea el elemento en la lista. La URL del webhook actúa
# como firma de acceso, así que el subcontratista sin cuenta SACYR
# puede reportar igual.
CONFIG_PATH = os.path.join(BASE_DIR, "rep_app_config.json")

CONFIG_SYNC_POR_DEFECTO = {
    "modo": "local",        # "local" (solo SQLite) | "webhook"
    "url": "",              # URL del disparador HTTP / Web App
    "url_login": "",        # webhook que valida contra la lista Usuarios App
    "token": "",            # opcional: viaja como cabecera X-Token
    "enviar_fotos": True,   # incrusta las fotos en base64 dentro del JSON
    "max_fotos": 4,
    "tiempo_limite": 30,    # segundos
    # Con las fotos de Drive públicas, la hoja las muestra con =IMAGE().
    # Privadas, la celda queda como enlace: quien lo abra necesita
    # permiso. Es una decisión de la obra, por eso empieza apagada.
    "fotos_publicas": False,
}

# Estructura de la lista de SharePoint. **Es el mismo juego de datos que
# exporta el Excel**: una columna de la lista por cada columna de la
# planilla, ni una más. Lo único que se traduce aquí es el encabezado
# —que lleva espacios y acentos— al nombre interno que usa SharePoint.
#
# Derivarlo en vez de repetirlo es lo que impide que se separen: si
# mañana cambia una columna del Excel, cambia sola la lista, la hoja y
# lo que recibe el flujo. Al agregar una columna hay que crearla también
# en la lista de SharePoint y mapearla en Power Automate.
NOMBRE_SP = {
    # columna del Excel  ->  (nombre en SharePoint, tipo de la columna)
    "FOLIO":         ("Title",         "Texto"),
    "COD ACTIVIDAD": ("CodActividad",  "Texto"),
    "COD STEPS":     ("CodSteps",      "Texto"),
    "COD OBRA":      ("CodObra",       "Texto"),
    "TRAMO":         ("Tramo",         "Texto"),
    "SECTOR":        ("Sector",        "Texto"),
    "TIPO_OBRA":     ("TipoObra",      "Texto"),
    "OBRA":          ("Obra",          "Texto"),
    "ACTIVIDAD":     ("Actividad",     "Texto largo"),
    "FECHA":         ("Fecha",         "Fecha"),
    "HH":            ("HH",            "Número"),
    "HM":            ("HM",            "Número"),
    "CANT":          ("Cantidad",      "Número"),
    "UND":           ("Unidad",        "Texto"),
    "PERSONAL":      ("Personal",      "Número"),
    "SUBCONTRATO":   ("Subcontrato",   "Texto"),
    "SUPERVISOR":    ("Supervisor",    "Texto"),
    "OBS":           ("Observaciones", "Texto largo"),
    "FOTO":          ("Foto",          "Texto"),
}


def columna_sp(columna):
    """Nombre y tipo en SharePoint de una columna de la planilla."""
    if columna.startswith("MAQUINA USADA "):
        return "Maquina" + columna.rsplit(" ", 1)[-1], "Texto"
    return NOMBRE_SP[columna]


def columnas_sp():
    """El contrato de la lista: (nombre, tipo, de qué columna sale)."""
    columnas = []
    for excel in columnas_entrega():
        nombre, tipo = columna_sp(excel)
        nota = ("Folio del reporte: la clave que evita duplicados"
                if nombre == "Title" else f"Columna «{excel}» de la planilla")
        columnas.append((nombre, tipo, nota))
    return columnas


# Lista "Usuarios App": la fuente de verdad del login cuando hay
# SharePoint. La contraseña NUNCA se guarda: solo su hash SHA-256, el
# mismo que calcula hash_clave(). El flujo compara hashes, así que ni
# el flujo ni la lista ven la contraseña real.
COLUMNAS_SP_USUARIOS = [
    ("Title",           "Texto",  "Nombre de usuario (jperez). Debe ser único"),
    ("NombreCompleto",  "Texto",  "Juan Pérez"),
    ("ClaveHash",       "Texto",  "SHA-256 de la contraseña (64 caracteres hex)"),
    ("Empresa",         "Texto",  "Sacyr · Subcontrato A · …"),
    ("Perfil",          "Texto",  "Supervisor · Administrador · Oficina Técnica · Inspector"),
    ("ObrasAsignadas",  "Texto",  "Separadas por ';'. Nombre de obra raíz o id de nodo "
                                  "(TIPO OBRA / OBRA / Actividad / Step). Sin la columna, "
                                  "manda lo asignado en 👥 Asignar"),
    ("Activo",          "Sí/No",  "No = no puede entrar"),
]

COLUMNAS_SP_SUPERVISORES = [
    ("Title",    "Texto", "Nombre del supervisor"),
    ("Empresa",  "Texto", "Empresa a la que pertenece"),
    ("Activo",   "Sí/No", "No = no aparece en el Nivel 1"),
]


# ============================================================
# 3. Base de datos
# ============================================================

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def hash_clave(clave: str) -> str:
    return hashlib.sha256(clave.encode("utf-8")).hexdigest()


def init_db():
    conn = get_db()
    c = conn.cursor()

    c.execute("""CREATE TABLE IF NOT EXISTS usuarios (
        id       INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario  TEXT UNIQUE,
        clave    TEXT,
        nombre   TEXT,
        empresa  TEXT,
        perfil   TEXT,
        activo   INTEGER DEFAULT 1
    )""")

    c.execute("""CREATE TABLE IF NOT EXISTS supervisores (
        id      INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre  TEXT,
        empresa TEXT,
        activo  INTEGER DEFAULT 1
    )""")

    # Parque de maquinaria por subcontrato. La patente identifica a la
    # máquina dentro de su empresa —en la planilla hay una patente
    # repetida entre dos empresas distintas, así que la clave es el par—.
    c.execute("""CREATE TABLE IF NOT EXISTS maquinaria (
        id        INTEGER PRIMARY KEY AUTOINCREMENT,
        empresa   TEXT,
        patente   TEXT,
        tipo      TEXT,
        activo    INTEGER DEFAULT 1,
        creado_en TEXT,
        creado_por TEXT
    )""")
    c.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_maquinaria "
              "ON maquinaria(empresa, patente)")

    # Árbol del proyecto (WBS). Lista de adyacencia: el mismo
    # principio de Primavera P6. La profundidad no está limitada.
    #   tipo: obra | grupo | especialidad | actividad | step
    #   pregunta: texto que encabeza la pantalla de sus hijos
    #
    # id_actividad / id_step / hh_lb / cantidad_lb son las columnas de
    # la BD de actividades del programa (ID ACT · COD STEP · HH LB ·
    # CANTIDAD LB). Viven en el nodo, no en el reporte: el reporte solo
    # las copia al momento de guardarse (ver resolver_ids()).
    c.execute("""CREATE TABLE IF NOT EXISTS nodos (
        id        TEXT PRIMARY KEY,
        nombre    TEXT,
        padre_id  TEXT,
        tipo      TEXT,
        codigo    TEXT,
        unidad    TEXT,
        pregunta  TEXT,
        orden     INTEGER DEFAULT 0,
        activo    INTEGER DEFAULT 1,
        latitud   REAL,
        longitud  REAL,
        id_actividad TEXT,
        id_step      TEXT,
        hh_lb        REAL,
        cantidad_lb  REAL,
        peso         REAL
    )""")
    c.execute("CREATE INDEX IF NOT EXISTS idx_nodos_padre ON nodos(padre_id)")

    # Migración: coordenadas de la planilla de obras e identificadores
    # de la BD de actividades.
    columnas = {f["name"] for f in c.execute("PRAGMA table_info(nodos)").fetchall()}
    for columna, tipo in (("latitud", "REAL"), ("longitud", "REAL"),
                          ("id_actividad", "TEXT"), ("id_step", "TEXT"),
                          ("hh_lb", "REAL"), ("cantidad_lb", "REAL"),
                          # Cuánto de su actividad representa el paso: los
                          # pesos de P01..P10 suman 1.
                          ("peso", "REAL")):
        if columna not in columnas:
            c.execute(f"ALTER TABLE nodos ADD COLUMN {columna} {tipo}")

    # Nodos del árbol asignados a cada usuario. Antes solo se guardaban
    # raíces (la obra completa); ahora el nodo puede estar a cualquier
    # altura —TIPO OBRA, OBRA, Actividad o Step— y lo asignado incluye
    # todo lo que cuelga de él. Ver nodos_asignados() / es_visible().
    c.execute("""CREATE TABLE IF NOT EXISTS permisos (
        usuario_id INTEGER,
        nodo_id    TEXT,
        PRIMARY KEY (usuario_id, nodo_id)
    )""")

    c.execute("""CREATE TABLE IF NOT EXISTS reportes (
        id             TEXT PRIMARY KEY,
        creado_en      TEXT,
        usuario        TEXT,
        empresa        TEXT,
        perfil         TEXT,
        supervisor     TEXT,
        nodo_id        TEXT,
        step_nombre    TEXT,
        ruta_ids       TEXT,
        ruta_texto     TEXT,
        obra           TEXT,
        especialidad   TEXT,
        actividad      TEXT,
        id_actividad   TEXT,
        id_step        TEXT,
        hh_calculada   REAL,
        hm_calculada   REAL,
        hh_lb          REAL,
        cantidad_lb    REAL,
        subcontrato    TEXT,
        equipos_patentes TEXT,
        fecha          TEXT,
        hora_inicio    TEXT,
        hora_termino   TEXT,
        horas          REAL,
        cantidad       REAL,
        unidad         TEXT,
        avance_pct     REAL,
        estado         TEXT,
        personal       INTEGER,
        personal_nombres TEXT,
        equipos        TEXT,
        materiales     TEXT,
        clima          TEXT,
        temperatura    REAL,
        observaciones  TEXT,
        incidencias    TEXT,
        riesgos        TEXT,
        firma_nombre   TEXT,
        firma_ok       INTEGER,
        fotos          TEXT,
        latitud        REAL,
        longitud       REAL,
        tipo_obra      TEXT,
        nombre_obra    TEXT,
        camino         TEXT,
        comuna         TEXT,
        region         TEXT,
        pais           TEXT,
        altitud        REAL,
        orientacion    TEXT,
        pk_desde       TEXT,
        pk_hasta       TEXT,
        sincronizado   INTEGER DEFAULT 0,
        sheets_ok      INTEGER DEFAULT 0,
        sp_ok          INTEGER DEFAULT 0,
        sync_fecha     TEXT,
        sync_intentos  INTEGER DEFAULT 0,
        sync_error     TEXT
    )""")

    # Migración: bases creadas antes del envío a SharePoint.
    columnas_rep = {f["name"] for f in c.execute("PRAGMA table_info(reportes)").fetchall()}
    for columna, tipo in (("latitud", "REAL"), ("longitud", "REAL"),
                          ("tipo_obra", "TEXT"), ("nombre_obra", "TEXT"),
                          ("sincronizado", "INTEGER DEFAULT 0"), ("sync_fecha", "TEXT"),
                          ("sync_intentos", "INTEGER DEFAULT 0"), ("sync_error", "TEXT"),
                          # Sello de la foto: dónde se tomó la evidencia.
                          ("camino", "TEXT"), ("comuna", "TEXT"), ("region", "TEXT"),
                          ("pais", "TEXT"), ("altitud", "REAL"), ("orientacion", "TEXT"),
                          ("pk_desde", "TEXT"), ("pk_hasta", "TEXT"),
                          # Trazabilidad contra la BD de actividades.
                          ("actividad", "TEXT"), ("id_actividad", "TEXT"),
                          ("id_step", "TEXT"), ("hh_calculada", "REAL"),
                          ("hm_calculada", "REAL"),
                          ("hh_lb", "REAL"), ("cantidad_lb", "REAL"),
                          # Quién ejecutó y con qué máquinas.
                          ("subcontrato", "TEXT"), ("equipos_patentes", "TEXT"),
                          # Estado de cada destino por separado: un reporte
                          # puede estar en Sheets y no en SharePoint.
                          ("sheets_ok", "INTEGER DEFAULT 0"), ("sp_ok", "INTEGER DEFAULT 0"),
                          # Enlaces de las fotos ya subidas a Drive (JSON).
                          # En la nube el disco es efímero: esto es lo que
                          # sobrevive, y evita volver a subirlas al reintentar.
                          ("fotos_url", "TEXT"),
                          # Dónde se ejecutó, en la rama lineal. Sin tramo y
                          # sector, T1-G-VC.MOV y T1-E-VC.MOV se confunden.
                          ("cod_obra", "TEXT"), ("tramo", "TEXT"), ("sector", "TEXT")):
        if columna not in columnas_rep:
            c.execute(f"ALTER TABLE reportes ADD COLUMN {columna} {tipo}")

    conn.commit()
    conn.close()


# ---------- Semilla: usuarios, supervisores y árbol de ejemplo ----------

def _slug(texto: str) -> str:
    """Identificador legible y estable a partir del nombre."""
    limpio = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    limpio = re.sub(r"[^a-zA-Z0-9]+", "-", limpio).strip("-").lower()
    return limpio[:40] or "x"


def clave_widget(prefijo: str, nodo_id: str) -> str:
    """Clave única y estable para el botón de un nodo.

    No se usa el slug del id: las rutas profundas se truncan y dos
    nodos hermanos terminarían compartiendo clave. El prefijo se
    conserva porque el CSS lo usa para distinguir grupos de Steps.
    """
    return f"{prefijo}_{hashlib.md5(nodo_id.encode('utf-8')).hexdigest()[:12]}"


def _step(nombre, unidad="und"):
    return {"nombre": nombre, "tipo": "step", "unidad": unidad}


def sin_actividades():
    """Lo que cuelga de una obra a la que aún no se le cargó el programa.

    Nada. Antes colgaba un catálogo genérico de obra civil —«Excavación
    por bancos de 2 m», «Perfilado de talud»…— igual para las 156 obras,
    y era peor que dejarla vacía: esos Steps no tienen ID de actividad
    ni existen en el programa, así que un supervisor podía reportar
    contra una actividad inventada y nadie podía cruzar ese reporte con
    P6. La obra vacía se anuncia como tal en pantalla.
    """
    return []


def _tramos(cantidad, prefijo="Tramo"):
    return [
        {"nombre": f"{prefijo} {i}", "tipo": "grupo",
         "pregunta": "Seleccione la especialidad", "hijos": sin_actividades()}
        for i in range(1, cantidad + 1)
    ]


# Tipos de obra puntual, tal como vienen en la planilla de obras
# (columnas ID TIPO OBRA / TIPO OBRA). El Nivel 2 —las obras
# individuales con su ID OBRA y coordenadas— se carga con el
# importador: ver importar_obras_puntuales().
#   (ID TIPO OBRA, TIPO OBRA, [(ID OBRA, OBRA, Latitud, Longitud), ...])
# Las obras sin coordenadas en la planilla quedan con None.
TIPOS_OBRA_PUNTUAL = [
    ("EST", "ESTACIONES", [
        ("EST.TAL", "Estación Talagante", -33.65908824, -70.92669716),
        ("EST.ELM", "Estación El Monte", -33.67935077, -70.98970227),
        ("EST.MEL", "Estación Melipilla", -33.67976338, -71.21010338),
    ]),
    ("SER", "SER", [
        ("SER.ELM", "SER EL MONTE", None, None),
        ("SER.MEL", "SER MELIPILLA", None, None),
    ]),
    ("PVD", "PASOS VEHICULARES DESNIVELADOS", [
        ("PVD.LUC", "PVD Lucas Pacheco", None, None),
        ("PVD.CAR", "PVD Anibal Pinto Los Carrera", -33.6783325, -70.98755219),
        ("PVD.DOM", "PVD Domingo Santa María", -33.67985228, -70.99102976),
        ("PVD.MAR", "PVD El Marco", -33.67552694, -71.16104279),
        ("PVD.ALC", "PVD Alcalde", -33.67993056, -71.20775545),
    ]),
    ("PMP", "PASOS MULTI PROPOSITO", [
        ("PMP.LUC", "PMP Lucas Pacheco", -33.65803701, -70.92405282),
        ("PMP.TAL", "PMP Talagante", None, None),
        ("PMP.ANP", "PMP Anibal Pinto", -33.67759237, -70.98573397),
        ("PMP.CAR", "PMP Los Carrera", -33.67841656, -70.98754459),
        ("PMP.DOM", "PMP Domingo Santa María", -33.68040236, -70.99226102),
        ("PMP.MON", "PMP El Montino", -33.68608468, -71.10123777),
        ("PMP.MAR", "PMP Camino El Marco", -33.67558749, -71.16059565),
        ("PMP.GSM", "PMP General San Martín", -33.67895344, -71.19821923),
        ("PMP.DEM", "PMP Padre Demetrio Bravo", -33.67974748, -71.20379917),
        ("PMP.ALC", "PMP Alcalde", -33.67995249, -71.20761396),
        ("PMP.MEL", "PMP Melipilla", None, None),
    ]),
    ("PAN", "PASOS A NIVEL", [
        ("PAN.PAL", "PAN Camino Las Palmeras", -33.62032813, -70.86902404),
        ("PAN.ANA", "PAN Camino Santa Ana", -33.63921847, -70.88521886),
        ("PAN.LUI", "PAN San Luis", -33.64973653, -70.90196031),
        ("PAN.EYZ", "PAN Eyzaguirre", -33.66040542, -70.93283334),
        ("PAN.BEN", "PAN Camino Benavente", -33.67107752, -70.97062416),
        ("PAN.ART", "PAN Arturo Prat", -33.67469949, -70.97901564),
        ("PAN.ZOZ", "PAN Zozimo Errazuriz", -33.6818697, -70.99941016),
        ("PAN.JUA", "PAN Santa Juana", -33.6830108, -71.01161059),
        ("PAN.PAI", "PAN Paico Alto", -33.68611347, -71.04486463),
        ("PAN.QUI", "PAN Chiñihue-Los Quilos", -33.68938277, -71.08008238),
        ("PAN.CRI", "PAN Chiñihue-El Cristo", -33.6869981, -71.09742658),
        ("PAN.ROS", "PAN Chiñihue-Las Rosas", -33.68067601, -71.12587595),
        ("PAN.SUE", "PAN Suecia", None, None),
    ]),
    ("PTE", "PUENTE", [
        ("PTE.BER", "Puente Berlin 27+414", -33.62605142, -70.87394506),
        ("PTE.CAS", "Puente Canal Castillo 28+946", -33.63719672, -70.88352758),
        ("PTE.TAL", "Puente Talagante 34+440", -33.66036247, -70.93434196),
        ("PTE.SMI", "Puente San Miguel 34+950", -33.66020406, -70.93952835),
        ("PTE.PAI", "Puente El Paico 34+990", -33.66015898, -70.94011269),
        ("PTE.EPA", "Puente Estero El Paico 45+944", -33.6867205, -71.0516895),
        ("PTE.SJO", "Puente Menor San Jose 49+470", -33.68889694, -71.08940898),
    ]),
    ("OBA", "OBRAS DE ARTE", [
        ("OBA.TR1", "Obras de Arte Tramo 1", None, None),
        ("OBA.TR2", "Obras de Arte Tramo 2", None, None),
        ("OBA.TR3", "Obras de Arte Tramo 3", None, None),
    ]),
    ("CAN", "CANALES", [
        ("CAN.TR1", "Canales Tramo 1", None, None),
        ("CAN.TR2", "Canales Tramo 2", None, None),
        ("CAN.TR3", "Canales Tramo 3", None, None),
    ]),
]

# ---------- Vías y Catenarias ----------
#
# Las dos ramas lineales del proyecto. A diferencia de las puntuales NO
# se cargan por planilla: son la estructura fija y viven aquí.
#
# El código de cada tipo dice dónde está: TRAMO - SECTOR - DISCIPLINA.
# El sector es G (General), E (Estación) o D (Desvío).
#
#     T2-E-VC  =  Tramo 2 · Estación · Vía Carga
#     T1-G-VP  =  Tramo 1 · General  · Vía Pasajeros
#     T3-D-VP  =  Tramo 3 · Desvío   · Vía Pasajeros
#
# y la obra agrega la partida:  T2-E-VC.PLA  ->  Plataforma.

# Partidas de cada disciplina. Vía Carga no lleva Poliductos y la Vía
# Secundaria solo tiene tres.
_PARTIDAS_VP = [("MOV", "Movimiento de Tierras"), ("CON", "Contencion"),
                ("DRE", "Drenaje"), ("PLA", "Plataforma"),
                ("POL", "Poliductos"), ("SUP", "Superestructura"),
                ("CFN", "Confinamiento")]
_PARTIDAS_VC = [p for p in _PARTIDAS_VP if p[0] != "POL"]
_PARTIDAS_VS = [("MOV", "Movimiento de Tierras"), ("PLA", "Plataforma"),
                ("SUP", "Superestructura")]
_PARTIDAS_CT = [("CIV", "Catenarias Obras Civiles"),
                ("ELE", "Catenarias Obras Electricas")]

#   (TRAMO, [(SECTOR, [(ID TIPO, DISCIPLINA, partidas), ...]), ...])
VIAS = [
    ("Tramo 1", [
        ("Vías General", [
            ("T1-G-VP", "Vía Pasajeros", _PARTIDAS_VP),
            ("T1-G-VC", "Vía Carga", _PARTIDAS_VC),
        ]),
        ("Vías Estación Talagante", [
            ("T1-E-VP", "Vía Pasajeros", _PARTIDAS_VP),
            ("T1-E-VC", "Vía Carga", _PARTIDAS_VC),
            ("T1-E-VS", "Vía Secundaria", _PARTIDAS_VS),
        ]),
    ]),
    ("Tramo 2", [
        ("Vías General", [
            ("T2-G-VP", "Vía Pasajeros", _PARTIDAS_VP),
            ("T2-G-VC", "Vía Carga", _PARTIDAS_VC),
        ]),
        ("Vías Estación El Monte", [
            ("T2-E-VP", "Vía Pasajeros", _PARTIDAS_VP),
            ("T2-E-VC", "Vía Carga", _PARTIDAS_VC),
            ("T2-E-VS", "Vía Secundaria", _PARTIDAS_VS),
        ]),
    ]),
    ("Tramo 3", [
        ("Vías General", [
            ("T3-G-VP", "Vía Pasajeros", _PARTIDAS_VP),
            ("T3-G-VC", "Vía Carga", _PARTIDAS_VC),
        ]),
        # El Tramo 3 no tiene Vía Secundaria en la planilla.
        ("Vías Estación Melipilla", [
            ("T3-E-VP", "Vía Pasajeros", _PARTIDAS_VP),
            ("T3-E-VC", "Vía Carga", _PARTIDAS_VC),
        ]),
        # Único sector de desvío del proyecto: solo el Tramo 3 lo tiene.
        ("Desvío Chiñihue", [
            ("T3-D-VP", "Vía Pasajeros", _PARTIDAS_VP),
            ("T3-D-VC", "Vía Carga", _PARTIDAS_VC),
        ]),
    ]),
]

#   (TRAMO, [(ID TIPO, SECTOR, partidas), ...])
CATENARIAS = [
    ("Tramo 1", [
        ("T1-G-CT", "Catenarias General", _PARTIDAS_CT),
        ("T1-E-CT", "Catenarias Estación Talagante", _PARTIDAS_CT),
    ]),
    ("Tramo 2", [
        ("T2-G-CT", "Catenarias General", _PARTIDAS_CT),
        ("T2-E-CT", "Catenarias Estación El Monte", _PARTIDAS_CT),
    ]),
    ("Tramo 3", [
        ("T3-G-CT", "Catenarias General", _PARTIDAS_CT),
        ("T3-E-CT", "Catenarias Estación Melipilla", _PARTIDAS_CT),
    ]),
]

# El árbol se abre en dos: lo puntual y lo lineal. Vías y Catenarias no
# son raíces propias, cuelgan de Obras Lineales —así el primer nivel
# refleja la división con la que trabaja la obra, sin repetir nodos—.
RAMA_PUNTUALES = "Obras Puntuales"
RAMA_LINEALES = "Obras Lineales"
RAMA_VIAS = "Vías"
RAMA_CATENARIAS = "Catenarias"
RAMAS_RAIZ = (RAMA_PUNTUALES, RAMA_LINEALES)

# Códigos de tipo que pertenecen a las ramas fijas. Se suman los de
# versiones anteriores de la planilla —cuando las vías no se abrían por
# tramo y sector— para que cargar una copia vieja no los reviva como
# obras puntuales.
CODIGOS_LINEALES_ANTIGUOS = {"VIP", "VIC", "CAT", "T1VP", "T1VC", "T2VP",
                             "T2VC", "T3VP", "T3VC", "T1CT", "T2CT", "T3CT"}


def _obras_de(codigo_tipo, partidas):
    """Nodos de obra de un tipo: T2-E-VC + .PLA -> T2-E-VC.PLA."""
    return [{"nombre": nombre, "tipo": "grupo", "codigo": f"{codigo_tipo}.{sufijo}",
             "pregunta": "Seleccione la actividad",
             "hijos": sin_actividades()}
            for sufijo, nombre in partidas]


def rama_vias():
    """Vías -> Tramo -> Sector -> Disciplina -> Partida -> Actividad -> Step.

    Cuelga de Obras Lineales, no del primer nivel. Dos niveles más que
    las puntuales: el tramo y el sector (General o la estación). Como la
    profundidad del árbol es libre, no hay nada que ajustar fuera de
    esta función.
    """
    return {
        "nombre": RAMA_VIAS, "tipo": "grupo", "codigo": "VIA",
        "pregunta": "Seleccione el tramo",
        "hijos": [
            {"nombre": tramo, "tipo": "grupo",
             "pregunta": "Seleccione el sector",
             "hijos": [
                 {"nombre": sector, "tipo": "grupo",
                  "pregunta": "Seleccione la vía",
                  "hijos": [
                      {"nombre": disciplina, "tipo": "grupo", "codigo": codigo_tipo,
                       "pregunta": "Seleccione la partida",
                       "hijos": _obras_de(codigo_tipo, partidas)}
                      for codigo_tipo, disciplina, partidas in tipos
                  ]}
                 for sector, tipos in sectores
             ]}
            for tramo, sectores in VIAS
        ],
    }


def rama_catenarias():
    """Catenarias -> Tramo -> Sector -> Partida -> Actividad -> Step.

    Igual que Vías, cuelga de Obras Lineales.
    """
    return {
        "nombre": RAMA_CATENARIAS, "tipo": "grupo", "codigo": "CAT",
        "pregunta": "Seleccione el tramo",
        "hijos": [
            {"nombre": tramo, "tipo": "grupo",
             "pregunta": "Seleccione el sector",
             "hijos": [
                 {"nombre": sector, "tipo": "grupo", "codigo": codigo_tipo,
                  "pregunta": "Seleccione la partida",
                  "hijos": _obras_de(codigo_tipo, partidas)}
                 for codigo_tipo, sector, partidas in tipos
             ]}
            for tramo, tipos in CATENARIAS
        ],
    }


def rama_obras_lineales():
    """Obras Lineales: agrupa Vías y Catenarias.

    Es la otra mitad del proyecto frente a Obras Puntuales. No agrega
    obras propias: solo pone en un mismo nivel las dos disciplinas
    lineales, para que el primer toque en terreno sea la división real
    con la que trabaja la obra.
    """
    return {
        "nombre": RAMA_LINEALES, "tipo": "obra", "codigo": "OL",
        "pregunta": "Seleccione la disciplina",
        "hijos": [rama_vias(), rama_catenarias()],
    }


def codigos_de_ramas_fijas():
    """Todos los ID TIPO de Vías y Catenarias, más los antiguos."""
    codigos = {c for _, sectores in VIAS for _, tipos in sectores for c, _, _ in tipos}
    codigos |= {c for _, tipos in CATENARIAS for c, _, _ in tipos}
    return codigos | CODIGOS_LINEALES_ANTIGUOS


def rama_obras_puntuales():
    """Rama de Obras Puntuales según la planilla oficial de obras.

        Obras Puntuales -> TIPO OBRA -> OBRA -> Actividad -> Step

    Las actividades que cuelgan de cada obra son el catálogo genérico
    de obra civil: la planilla solo llega hasta el Nivel 2 (la obra).
    Las reales se cargan por obra desde ⚙️ Obras con la BD de
    actividades del diario (ver importar_actividades).
    """
    return {
        "nombre": "Obras Puntuales", "tipo": "obra", "codigo": "OP",
        "pregunta": "Seleccione el tipo de obra",
        "hijos": [
            {"nombre": nombre_tipo, "tipo": "grupo", "codigo": codigo_tipo,
             "pregunta": "Seleccione la obra",
             "hijos": [
                 {"nombre": nombre_obra, "tipo": "grupo", "codigo": codigo_obra,
                  "latitud": lat, "longitud": lon,
                  "pregunta": "Seleccione la actividad",
                  "hijos": sin_actividades()}
                 for codigo_obra, nombre_obra, lat, lon in obras
             ]}
            for codigo_tipo, nombre_tipo, obras in TIPOS_OBRA_PUNTUAL
        ],
    }


def arbol_demo():
    """Semilla del árbol: las dos ramas reales del proyecto.

    Obras Puntuales se reemplaza después con la planilla oficial;
    Obras Lineales —Vías y Catenarias— es fija y se resiembra desde el
    código.
    """
    return [
        rama_obras_puntuales(),
        rama_obras_lineales(),
    ]


# Un solo INSERT para todo el árbol: la semilla, la planilla de obras
# y la BD de actividades escriben exactamente las mismas columnas.
INSERT_NODO = (
    "INSERT INTO nodos (id,nombre,padre_id,tipo,codigo,unidad,pregunta,orden,activo,"
    "latitud,longitud,id_actividad,id_step,hh_lb,cantidad_lb,peso) "
    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)")


def _aplanar(nodos, padre_id, filas, usados):
    """Convierte el árbol anidado en filas (id, nombre, padre_id...)."""
    for orden, n in enumerate(nodos):
        base = f"{padre_id}/{_slug(n['nombre'])}" if padre_id else _slug(n["nombre"])
        nodo_id, extra = base, 2
        while nodo_id in usados:          # el nombre se repite bajo el mismo padre
            nodo_id = f"{base}-{extra}"
            extra += 1
        usados.add(nodo_id)

        filas.append((
            nodo_id, n["nombre"], padre_id, n.get("tipo", "grupo"),
            n.get("codigo", ""), n.get("unidad", ""),
            n.get("pregunta", "Seleccione una opción"), orden, 1,
            n.get("latitud"), n.get("longitud"),
            n.get("id_actividad", ""), n.get("id_step", ""),
            n.get("hh_lb"), n.get("cantidad_lb"), n.get("peso"),
        ))
        if n.get("hijos"):
            _aplanar(n["hijos"], nodo_id, filas, usados)


# ---------- Importación de la planilla oficial de obras ----------

# Nombres de columna aceptados, normalizados (sin tildes, en
# minúsculas). La planilla trae: Nivel · ID TIPO OBRA · TIPO OBRA ·
# ID OBRA · OBRA · Latitud · Longitud.
COLUMNAS_OBRAS = {
    "nivel": "nivel",
    "id tipo obra": "id_tipo",
    "tipo obra": "tipo",
    "id obra": "id_obra",
    "obra": "obra",
    "latitud": "latitud",
    "longitud": "longitud",
}


def _normalizar(texto) -> str:
    if texto is None:
        return ""
    limpio = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", limpio).strip().lower()


def _a_numero(valor):
    if valor is None or str(valor).strip() == "":
        return None
    try:
        return float(str(valor).replace(",", "."))
    except ValueError:
        return None


def leer_planilla_obras(archivo, nombre_archivo=""):
    """Lee la planilla (.xlsx o .csv) y devuelve las filas normalizadas.

    Busca la fila de encabezados en cualquier parte de la hoja: la
    planilla suele traer títulos o filtros arriba de la tabla.
    """
    nombre = (nombre_archivo or getattr(archivo, "name", "")).lower()

    if nombre.endswith(".csv"):
        import csv
        import io
        texto = archivo.getvalue().decode("utf-8-sig", errors="replace")
        dialecto = csv.Sniffer().sniff(texto[:2000], delimiters=";,\t")
        matriz = list(csv.reader(io.StringIO(texto), dialecto))
    else:
        import openpyxl
        wb = openpyxl.load_workbook(archivo, read_only=True, data_only=True)
        matriz = [list(f) for f in wb.active.iter_rows(values_only=True)]
        wb.close()

    # Fila de encabezados: la primera que contiene "tipo obra" y "obra".
    encabezado, inicio = None, 0
    for i, fila in enumerate(matriz):
        celdas = [_normalizar(v) for v in fila]
        if "tipo obra" in celdas and "obra" in celdas:
            encabezado, inicio = celdas, i + 1
            break
    if encabezado is None:
        raise ValueError(
            "No se encontró la fila de encabezados. Se esperan las columnas: "
            "Nivel, ID TIPO OBRA, TIPO OBRA, ID OBRA, OBRA, Latitud, Longitud.")

    indices = {}
    for pos, celda in enumerate(encabezado):
        if celda in COLUMNAS_OBRAS:
            indices.setdefault(COLUMNAS_OBRAS[celda], pos)
    for requerida in ("id_tipo", "tipo", "obra"):
        if requerida not in indices:
            raise ValueError(f"Falta la columna obligatoria '{requerida}' en la planilla.")

    def celda(fila, campo):
        pos = indices.get(campo)
        if pos is None or pos >= len(fila) or fila[pos] is None:
            return ""
        return str(fila[pos]).strip()

    filas = []
    for fila in matriz[inicio:]:
        if not any(v not in (None, "") for v in fila):
            continue
        registro = {
            "nivel": celda(fila, "nivel"),
            "id_tipo": celda(fila, "id_tipo"),
            "tipo": celda(fila, "tipo"),
            "id_obra": celda(fila, "id_obra"),
            "obra": celda(fila, "obra"),
            "latitud": _a_numero(celda(fila, "latitud")),
            "longitud": _a_numero(celda(fila, "longitud")),
        }
        # Las filas de detalle de la planilla oficial vienen con el tipo
        # en blanco —solo lo trae su encabezado de Nivel 1—, así que no
        # se puede exigir tipo para conservarlas.
        if any((registro["id_tipo"], registro["tipo"],
                registro["id_obra"], registro["obra"])):
            filas.append(registro)
    return filas


def _borrar_subarbol(cursor, nodo_id, incluir_raiz=False):
    """Elimina todo lo que cuelga de un nodo (y el nodo si se pide)."""
    semilla = "SELECT id FROM nodos WHERE id = ?" if incluir_raiz \
        else "SELECT id FROM nodos WHERE padre_id = ?"
    cursor.execute(f"""
        WITH RECURSIVE sub(id) AS (
            {semilla}
            UNION ALL
            SELECT n.id FROM nodos n JOIN sub s ON n.padre_id = s.id
        )
        DELETE FROM nodos WHERE id IN (SELECT id FROM sub)
    """, (nodo_id,))


def importar_obras_puntuales(filas):
    """Reconstruye la rama Obras Puntuales desde la planilla.

    Jerarquía resultante:
        Obras Puntuales -> TIPO OBRA -> OBRA -> Especialidad -> Step

    Las filas de Nivel 1 (sin ID OBRA) definen los tipos; las demás
    son obras individuales y se cuelgan de su tipo. Devuelve el
    resumen de lo cargado.
    """
    # La planilla de referencia trae los 17 tipos juntos, puntuales y
    # lineales. Los lineales se ignoran acá: viven en su propia rama y
    # cargarlos también aquí los duplicaría. Se suman los códigos que
    # usaba la planilla antigua, antes de abrirse por tramo, para que
    # cargar una versión vieja no los resucite bajo esta rama.
    codigos_lineales = codigos_de_ramas_fijas()

    # Agrupa por tipo respetando el orden de aparición en la planilla.
    #
    # En la planilla oficial el tipo se escribe UNA vez, en su fila de
    # Nivel 1; las filas de detalle traen ID TIPO OBRA vacío. Por eso el
    # tipo se arrastra hacia abajo en vez de leerse de cada fila.
    tipos, obras_por_tipo, omitidos = [], {}, set()
    tipo_actual = None
    for f in filas:
        if f["id_tipo"]:
            tipo_actual = f["id_tipo"]
            if tipo_actual not in obras_por_tipo and tipo_actual not in codigos_lineales:
                tipos.append((tipo_actual, f["tipo"] or tipo_actual))
                obras_por_tipo[tipo_actual] = []
        if not tipo_actual:
            continue                       # detalle antes de su encabezado
        if tipo_actual in codigos_lineales:
            omitidos.add(tipo_actual)
            continue

        # Fila de detalle: trae ID OBRA propio o un nombre distinto al del tipo.
        es_detalle = bool(f["id_obra"]) and f["id_obra"] != tipo_actual
        if not es_detalle:
            es_detalle = (f["nivel"] not in ("1", "1.0")
                          and _normalizar(f["obra"]) != _normalizar(f["tipo"]))
        if es_detalle and f["obra"]:
            obras_por_tipo[tipo_actual].append(f)

    if not tipos:
        raise ValueError("La planilla no contiene tipos de obra reconocibles.")

    arbol = {
        "nombre": "Obras Puntuales", "tipo": "obra", "codigo": "OP",
        "pregunta": "Seleccione el tipo de obra", "hijos": [],
    }
    total_obras = 0
    for id_tipo, nombre_tipo in tipos:
        detalle = obras_por_tipo[id_tipo]
        if detalle:
            hijos = [{
                "nombre": o["obra"], "tipo": "grupo",
                "codigo": o["id_obra"] or id_tipo,
                "latitud": o["latitud"], "longitud": o["longitud"],
                "pregunta": "Seleccione la especialidad",
                "hijos": sin_actividades(),
            } for o in detalle]
            pregunta = "Seleccione la obra"
            total_obras += len(detalle)
        else:
            # Tipo sin detalle cargado: se mantiene navegable.
            hijos = sin_actividades()
            pregunta = "Seleccione la especialidad"
        arbol["hijos"].append({
            "nombre": nombre_tipo, "tipo": "grupo", "codigo": id_tipo,
            "pregunta": pregunta, "hijos": hijos,
        })

    nodos = _reemplazar_rama(arbol)
    return {"tipos": len(tipos), "obras": total_obras, "nodos": nodos,
            "omitidos_lineales": sorted(omitidos)}


def _reemplazar_rama(arbol):
    """Resiembra una rama raíz completa y devuelve cuántos nodos escribió.

    La raíz conserva su mismo id —se deriva del nombre—, y con él los
    permisos y asignaciones que los supervisores ya tienen sobre ella.
    """
    conn = get_db()
    c = conn.cursor()
    anterior = c.execute(
        "SELECT id FROM nodos WHERE padre_id IS NULL AND nombre = ?",
        (arbol["nombre"],)).fetchone()

    usados = {f["id"] for f in c.execute("SELECT id FROM nodos").fetchall()}
    if anterior:
        # Se liberan los ids de la rama antes de regenerarla: si no, el
        # desambiguador de _aplanar les agregaría un sufijo "-2" y las
        # asignaciones existentes dejarían de calzar.
        _borrar_subarbol(c, anterior["id"], incluir_raiz=True)
        usados = {f["id"] for f in c.execute("SELECT id FROM nodos").fetchall()}

    nuevas = []
    _aplanar([arbol], None, nuevas, usados)
    c.executemany(INSERT_NODO, nuevas)
    conn.commit()
    conn.close()
    return len(nuevas)


def importar_ramas_fijas():
    """Siembra o actualiza Obras Lineales (Vías y Catenarias).

    No recibe archivo: la estructura es constante y vive en VIAS y
    CATENARIAS. Las actividades reales se cargan después, obra por
    obra, con la BD de actividades del diario.
    """
    arbol = rama_obras_lineales()
    nodos = _reemplazar_rama(arbol)

    resumen = {}
    for disciplina in arbol["hijos"]:
        tipos = obras = 0

        # Un "tipo" es el nodo que lleva ID TIPO y cuyos hijos ya son
        # obras; la "obra", cada partida que cuelga de él.
        def contar(nodo):
            nonlocal tipos, obras
            for hijo in nodo["hijos"]:
                if hijo.get("codigo") and hijo["hijos"] and hijo["hijos"][0].get("codigo"):
                    tipos += 1
                    obras += len(hijo["hijos"])
                else:
                    contar(hijo)

        contar(disciplina)
        resumen[disciplina["nombre"]] = {"tipos": tipos, "obras": obras}
    resumen["nodos"] = nodos
    return resumen


# ---------- Importación de la BD de actividades del diario ----------
#
# La planilla del programa trae una fila por Step:
#
#   ID ACT · COD FRENTE · NOMBRE FRENTE · SUBFRENTE · COD WBS ·
#   DETALLE WBS · NOMBRE ACTIVIDAD · COD STEP · STEP · HH LB · UND ·
#   CANTIDAD LB
#
# La app NO se llena con esa base: solo cuelga de la OBRA elegida los
# dos últimos niveles —Actividad y Step— guardando ID ACT y COD STEP
# en el nodo. Así el reporte que emite el supervisor puede devolver
# esos identificadores sin que él tenga que verlos ni escribirlos.

COLUMNAS_ACTIVIDADES = {
    "id act": "id_act",
    "id actividad": "id_act",
    "nombre actividad": "actividad",
    "actividad": "actividad",
    "cod step": "cod_step",
    "codigo step": "cod_step",
    "step": "step",
    "hh lb": "hh_lb",
    "und": "unidad",
    "unidad": "unidad",
    "cantidad lb": "cantidad_lb",
    "cod wbs": "cod_wbs",
    "detalle wbs": "detalle_wbs",
    "subfrente": "subfrente",
    "nombre frente": "frente",
    "cod frente": "cod_frente",
}


def leer_planilla_actividades(archivo, nombre_archivo=""):
    """Lee la BD de actividades y devuelve una fila por Step."""
    nombre = (nombre_archivo or getattr(archivo, "name", "")).lower()

    if nombre.endswith(".csv"):
        import csv
        texto = archivo.getvalue().decode("utf-8-sig", errors="replace")
        dialecto = csv.Sniffer().sniff(texto[:2000], delimiters=";,\t")
        matriz = list(csv.reader(io.StringIO(texto), dialecto))
    else:
        import openpyxl
        wb = openpyxl.load_workbook(archivo, read_only=True, data_only=True)
        matriz = [list(f) for f in wb.active.iter_rows(values_only=True)]
        wb.close()

    encabezado, inicio = None, 0
    for i, fila in enumerate(matriz):
        celdas = [_normalizar(v) for v in fila]
        if "nombre actividad" in celdas and "step" in celdas:
            encabezado, inicio = celdas, i + 1
            break
    if encabezado is None:
        raise ValueError(
            "No se encontró la fila de encabezados. Se esperan al menos las columnas: "
            "ID ACT, NOMBRE ACTIVIDAD, COD STEP, STEP.")

    indices = {}
    for pos, celda in enumerate(encabezado):
        if celda in COLUMNAS_ACTIVIDADES:
            indices.setdefault(COLUMNAS_ACTIVIDADES[celda], pos)
    for requerida in ("actividad", "step"):
        if requerida not in indices:
            raise ValueError(f"Falta la columna obligatoria '{requerida}' en la planilla.")

    def celda(fila, campo):
        pos = indices.get(campo)
        if pos is None or pos >= len(fila) or fila[pos] is None:
            return ""
        return str(fila[pos]).strip()

    filas = []
    for fila in matriz[inicio:]:
        if not any(v not in (None, "") for v in fila):
            continue
        registro = {
            "id_act": celda(fila, "id_act"),
            "actividad": celda(fila, "actividad"),
            "cod_step": celda(fila, "cod_step"),
            "step": celda(fila, "step"),
            "unidad": celda(fila, "unidad") or "und",
            "hh_lb": _a_numero(celda(fila, "hh_lb")),
            "cantidad_lb": _a_numero(celda(fila, "cantidad_lb")),
        }
        if registro["actividad"] and registro["step"]:
            filas.append(registro)
    return filas


def importar_actividades(destino_id, filas, reemplazar=True):
    """Cuelga Actividad -> Step bajo la OBRA `destino_id`.

    Agrupa las filas por (ID ACT, NOMBRE ACTIVIDAD) respetando el orden
    de la planilla. Cada Step conserva su COD STEP, su unidad y su HH de
    línea base para poder contrastarlas con lo reportado.
    """
    conn = get_db()
    c = conn.cursor()

    destino = c.execute("SELECT id, nombre FROM nodos WHERE id=?", (destino_id,)).fetchone()
    if destino is None:
        conn.close()
        raise ValueError("La obra de destino ya no existe en el árbol.")

    actividades, steps_por_act = [], {}
    for f in filas:
        clave = (f["id_act"], f["actividad"])
        if clave not in steps_por_act:
            actividades.append(clave)
            steps_por_act[clave] = []
        steps_por_act[clave].append(f)

    if not actividades:
        conn.close()
        raise ValueError("La planilla no contiene actividades reconocibles.")

    arbol = [{
        "nombre": nombre_act, "tipo": "actividad", "codigo": id_act,
        "id_actividad": id_act,
        "pregunta": "Seleccione la actividad detallada (Step)",
        "hijos": [{
            "nombre": s["step"], "tipo": "step", "codigo": s["cod_step"],
            "unidad": s["unidad"], "id_actividad": id_act, "id_step": s["cod_step"],
            "hh_lb": s["hh_lb"], "cantidad_lb": s["cantidad_lb"],
        } for s in steps_por_act[(id_act, nombre_act)]],
    } for id_act, nombre_act in actividades]

    if reemplazar:
        _borrar_subarbol(c, destino_id)          # solo los hijos: la obra se conserva
    usados = {f["id"] for f in c.execute("SELECT id FROM nodos").fetchall()}

    nuevas = []
    _aplanar(arbol, destino_id, nuevas, usados)
    c.executemany(INSERT_NODO, nuevas)
    c.execute("UPDATE nodos SET pregunta=? WHERE id=?",
              ("Seleccione la actividad", destino_id))
    conn.commit()
    conn.close()

    return {"obra": destino["nombre"], "actividades": len(actividades),
            "steps": sum(len(v) for v in steps_por_act.values()), "nodos": len(nuevas)}


# ---------- Carga masiva desde el export de Primavera P6 ----------
#
# El export trae una fila por actividad, con la jerarquía WBS en la
# sangría de la primera columna y dos columnas propias del proyecto:
#
#   00 OBRA  ->  ID OBRA de la planilla de obras. Es el enlace que
#                permite cargar las 41 obras de una sola pasada, en vez
#                de ir obra por obra.
#   01 FASE  ->  fase del cronograma (hoy siempre "CON").
#
# Planificación todavía no publica los Steps. Hasta que lleguen, y
# según lo confirmado por ellos, **cada actividad tiene un único Step
# que es la propia actividad**: el nodo hoja conserva el Activity ID
# tanto en ID ACT como en COD STEP, de modo que el reporte ya viaja
# con identificadores reales. Cuando existan los Steps de verdad, se
# recarga esa obra con la BD de actividades del diario y este nodo
# pasa a ser el padre de sus Steps, sin tocar código.

# Tipos de obra puntual que el export de P6 no cubre: no cuelgan de
# "Obras Civiles Puntuales" en el cronograma. La carga masiva no los
# toca, así conservan lo que tengan cargado a mano.
# Vías y Catenarias no necesitan aparecer aquí: la carga solo recorre
# la rama de Obras Puntuales.
TIPOS_SIN_CRONOGRAMA = ("OBA", "CAN")

# Errores de codificación del export. La clave es el nombre de un nodo
# WBS; si aparece en la ruta de la actividad, manda sobre "00 OBRA".
# Gana la coincidencia más profunda. Al corregirse en el origen, basta
# con vaciar este diccionario.
CORRECCIONES_P6 = {
    "SER 06 Estación El Monte":    "SER.ELM",   # venía EST.MEL
    "SER 06 Estación Melipilla":   "SER.MEL",   # venía EST.MEL
    "Puente San Miguel pk 34+950": "PTE.SMI",   # venía PTE.SMP (nodo combinado)
    "Puente El Paico pk 34+990":   "PTE.PAI",   # venía PTE.SMP
}

# Profundidad, dentro del WBS del export, del nodo que representa la
# obra. Todo lo que cuelga por debajo es la estructura que se importa.
NIVEL_OBRA_P6 = 5


def leer_export_p6(archivo, nombre_archivo=""):
    """Lee el export de P6 y devuelve una fila por actividad.

    Cada fila trae su ruta WBS completa, reconstruida a partir del
    número de nivel de la primera columna.
    """
    nombre = (nombre_archivo or getattr(archivo, "name", "")).lower()

    if nombre.endswith(".csv"):
        import csv
        texto = archivo.getvalue().decode("utf-8-sig", errors="replace")
        dialecto = csv.Sniffer().sniff(texto[:2000], delimiters=";,\t")
        matriz = list(csv.reader(io.StringIO(texto), dialecto))
    else:
        import openpyxl
        wb = openpyxl.load_workbook(archivo, read_only=True, data_only=True)
        matriz = [list(f) for f in wb.active.iter_rows(values_only=True)]
        wb.close()

    encabezado, inicio = None, 0
    for i, fila in enumerate(matriz):
        celdas = [_normalizar(v) for v in fila]
        if "activity id" in celdas and "activity name" in celdas:
            encabezado, inicio = celdas, i + 1
            break
    if encabezado is None:
        raise ValueError(
            "No se encontró la fila de encabezados. Se esperan al menos las "
            "columnas Activity ID, Activity Name y 00 OBRA.")

    idx = {c: p for p, c in enumerate(encabezado) if c}
    i_id = idx["activity id"]
    i_nombre = idx["activity name"]
    i_obra = idx.get("00 obra")
    i_hh = idx.get("budgeted labor units")
    if i_obra is None:
        raise ValueError("Falta la columna '00 OBRA', que enlaza cada actividad con su obra.")

    # El nivel del WBS va en la primera columna y el nombre del nodo en
    # la anterior a Activity ID; ninguna de las dos tiene encabezado útil.
    i_nivel, i_wbs = 0, max(0, i_id - 1)

    pila, filas = {}, []
    for fila in matriz[inicio:]:
        if not fila or all(v in (None, "") for v in fila):
            continue
        try:
            nivel = int(float(str(fila[i_nivel]).strip()))
        except (TypeError, ValueError, IndexError):
            continue

        act_id = str(fila[i_id] or "").strip() if i_id < len(fila) else ""
        if not act_id:
            # Nodo del WBS: se apila y se descartan los niveles inferiores.
            pila[nivel] = str(fila[i_wbs] or "").strip()
            for k in [x for x in pila if x > nivel]:
                del pila[k]
            continue

        filas.append({
            "id": act_id,
            "nombre": str(fila[i_nombre] or "").strip(),
            "obra": str(fila[i_obra] or "").strip() if i_obra < len(fila) else "",
            "hh": _a_numero(fila[i_hh]) if i_hh is not None and i_hh < len(fila) else None,
            "ruta": [pila[k] for k in sorted(pila)],
        })

    if not filas:
        raise ValueError("El export no contiene actividades.")
    return filas


def _codigo_obra_de(actividad):
    """Código de obra de una actividad, aplicando las correcciones."""
    for nombre in reversed(actividad["ruta"]):
        if nombre in CORRECCIONES_P6:
            return CORRECCIONES_P6[nombre]
    return actividad["obra"]


def _subarbol_p6(actividades):
    """Arma la rama que cuelga de una obra: grupos del WBS y Steps.

    El nodo cuyos hijos son todos Steps se marca como "actividad" para
    que el reporte sepa de qué actividad depende cada Step.
    """
    raiz = {}
    for a in actividades:
        nodo = raiz
        for parte in a["ruta"][NIVEL_OBRA_P6 + 1:]:
            nodo = nodo.setdefault(parte, {})
        nodo.setdefault("@steps", []).append(a)

    def convertir(dic):
        hijos = []
        for nombre, contenido in dic.items():
            if nombre == "@steps":
                continue
            sub = convertir(contenido)
            solo_steps = bool(sub) and all(h["tipo"] == "step" for h in sub)
            hijos.append({
                "nombre": nombre,
                "tipo": "actividad" if solo_steps else "grupo",
                "pregunta": ("Seleccione la actividad detallada (Step)" if solo_steps
                             else "Seleccione la actividad"),
                "hijos": sub,
            })
        for a in dic.get("@steps", []):
            # Un único Step por actividad: la actividad misma.
            hijos.append({
                "nombre": a["nombre"], "tipo": "step", "codigo": a["id"],
                "id_actividad": a["id"], "id_step": a["id"],
                "unidad": "", "hh_lb": a["hh"], "cantidad_lb": None,
            })
        return hijos

    return convertir(raiz)


def importar_actividades_p6(filas, excluir_tipos=TIPOS_SIN_CRONOGRAMA):
    """Carga las actividades del export en todas las obras de una vez.

    Reemplaza lo que cuelga de cada obra alcanzada; las obras que el
    export no menciona quedan intactas.
    """
    conn = get_db()
    c = conn.cursor()

    raiz = c.execute(
        "SELECT id FROM nodos WHERE padre_id IS NULL AND nombre='Obras Puntuales'").fetchone()
    if raiz is None:
        conn.close()
        raise ValueError("Primero cargue la planilla de obras.")

    # Obras del árbol, por código, saltando los tipos excluidos.
    destinos, omitidos = {}, set()
    for tipo in c.execute("SELECT id, codigo FROM nodos WHERE padre_id=?", (raiz["id"],)).fetchall():
        if tipo["codigo"] in excluir_tipos:
            omitidos.add(tipo["codigo"])
            continue
        for obra in c.execute("SELECT id, codigo, nombre FROM nodos WHERE padre_id=?",
                              (tipo["id"],)).fetchall():
            if obra["codigo"]:
                destinos[obra["codigo"]] = dict(obra)

    # Agrupar por obra, aplicando correcciones.
    por_obra, sin_enlace = {}, {}
    for a in filas:
        codigo = _codigo_obra_de(a)
        if codigo in destinos:
            por_obra.setdefault(codigo, []).append(a)
        else:
            nodo = a["ruta"][NIVEL_OBRA_P6] if len(a["ruta"]) > NIVEL_OBRA_P6 else "(sin nodo)"
            sin_enlace.setdefault(nodo, []).append(a["id"])

    if not por_obra:
        conn.close()
        raise ValueError("Ninguna actividad del export coincide con una obra del árbol.")

    total_nodos = 0
    for codigo, actividades in por_obra.items():
        destino = destinos[codigo]
        _borrar_subarbol(c, destino["id"])            # solo los hijos
        usados = {f["id"] for f in c.execute("SELECT id FROM nodos").fetchall()}
        nuevas = []
        _aplanar(_subarbol_p6(actividades), destino["id"], nuevas, usados)
        c.executemany(INSERT_NODO, nuevas)
        c.execute("UPDATE nodos SET pregunta=? WHERE id=?",
                  ("Seleccione la actividad", destino["id"]))
        total_nodos += len(nuevas)

    conn.commit()
    conn.close()

    return {
        "obras": len(por_obra),
        "actividades": sum(len(v) for v in por_obra.values()),
        "nodos": total_nodos,
        "omitidos": sorted(omitidos),
        "sin_enlace": {k: len(v) for k, v in sin_enlace.items()},
    }


# ---------- Carga masiva desde la BD del diario (export R2) ----------
#
# La hoja CON_BD_TAM del libro BD_TAM **no trae fila de encabezados**:
# es un volcado de P6 con columnas fijas. Los nombres de abajo salen de
# la hoja BD_Avance del mismo libro, que sí los trae y usa exactamente
# el mismo orden de columnas.
#
# Lo que hace valioso al export es `00 OBRA`: enlaza cada actividad con
# una obra del árbol —puntual o lineal— y permite cargar todo de una
# pasada en vez de ir obra por obra.
#
# Bajo cada obra quedan los dos niveles que la app necesita:
#
#   WBS Name      -> ACTIVIDAD  ("Terraplén", "Excavación")
#   Activity Name -> STEP       ("Conformación de Terraplén - VC - T1 - Pk 25+500")
#
# Los Steps llegaron en el libro R2_1, pero **no por la columna
# `STEP ID`** —que sigue vacía— sino en veinte columnas al final:
# `P01`, `P01%`, `P02`, `P02%`… hasta `P10`. Cada actividad declara así
# entre uno y diez pasos con su peso, y los pesos suman 1:
#
#   Demolición tablero existente        <- Activity Name (STEP)
#      P01  Excavación            0,25   \
#      P01  Emplantillado         0,25    > sus PASOS, el nivel nuevo
#      P03  Colocación Prefab.    0,50   /
#
# El peso es lo que permitirá calcular avance: reportar un paso terminado
# adelanta su fracción de la actividad, no la actividad entera.
#
# Solo 374 de las ~2.700 actividades traen pasos. Las demás siguen siendo
# su propio Step, con el Activity ID en los dos identificadores.
#
# `04 PLATAFORMA` (001.TRR, 002.SBS…) NO es el COD STEP: es el proceso
# constructivo, y solo lo trae una de cada cinco filas. Se lee para
# poder mostrarlo, pero no se hace pasar por un identificador de Step.

# El libro cambió de nombre de hoja entre revisiones: R2 traía la vista
# ya filtrada «CON_BD_TAM» y R2_1 solo el export completo «BD_Avance».
# Las dos tienen las mismas columnas, así que se acepta cualquiera.
HOJAS_DIARIO = ("CON_BD_TAM", "BD_Avance")

# Pares (nombre del paso, peso) al final de la fila: P01 en la columna
# 27, su peso en la 28, y así de dos en dos hasta P10.
COLUMNA_PRIMER_PASO = 27
PASOS_POR_ACTIVIDAD = 10

# `01 FASE` separa la construcción de todo lo demás que el programa
# también planifica: ingeniería (ING.NEW, ING.ORI), suministros
# (APR.SUM, CSE.CON), permisos (PER.DOM, PER.MOP), hitos (HIT),
# expropiaciones, arqueología… Nada de eso se reporta en terreno, así
# que solo entra `CON`. Es el mismo filtro que aplicaba a mano la hoja
# «CON_BD_TAM» del libro anterior: 2.340 filas de 3.069.
FASES_TERRENO = ("CON",)

COLUMNAS_DIARIO = {
    "nivel": 0,          # Level
    "wbs": 2,            # WBS
    "actividad": 3,      # WBS Name      -> nivel ACTIVIDAD
    "id_act": 4,         # Activity ID
    "estado": 5,         # Activity Status
    "step": 6,           # Activity Name -> nivel STEP
    "id_obra": 7,        # 00 OBRA: el enlace con el árbol
    "fase": 8,           # 01 FASE (hoy siempre CON)
    "tramo": 9,          # 02 TRAMOS
    "via": 10,           # 03 VIA
    "plataforma": 11,    # 04 PLATAFORMA (proceso, no es el Step)
    "cod_step": 12,      # STEP ID (sigue vacía; los pasos vienen en P01..P10)
    "cantidad_lb": 17,   # Num01
    "unidad": 18,        # Text01
    "hh_lb": 19,         # Budgeted Labor Units
}
MIN_COLUMNAS_DIARIO = 20


def leer_bd_diario(archivo, nombre_archivo=""):
    """Lee el export de la BD del diario (sin encabezados, columnas fijas)."""
    nombre = (nombre_archivo or getattr(archivo, "name", "")).lower()

    if nombre.endswith(".csv"):
        import csv
        texto = archivo.getvalue().decode("utf-8-sig", errors="replace")
        matriz = list(csv.reader(io.StringIO(texto), delimiter=";"))
    else:
        import openpyxl
        wb = openpyxl.load_workbook(archivo, read_only=True, data_only=True)
        # El libro trae varias hojas (OBRAS entre ellas): la buena se
        # busca por nombre, no por cuál quedó seleccionada al guardar.
        nombre_hoja = next((h for h in HOJAS_DIARIO if h in wb.sheetnames), None)
        hoja = wb[nombre_hoja] if nombre_hoja else wb.active
        matriz = [list(f) for f in hoja.iter_rows(values_only=True)]
        wb.close()

    def celda(fila, campo):
        pos = COLUMNAS_DIARIO[campo]
        if pos >= len(fila) or fila[pos] is None:
            return ""
        return str(fila[pos]).strip()

    def pasos_de(fila):
        """Los P01..P10 de la fila, con su peso. Vacío si no los trae."""
        pasos = []
        for n in range(PASOS_POR_ACTIVIDAD):
            i = COLUMNA_PRIMER_PASO + n * 2
            if i >= len(fila) or fila[i] is None or str(fila[i]).strip() == "":
                continue
            pasos.append({
                "codigo": f"P{n + 1:02d}",
                "nombre": str(fila[i]).strip(),
                "peso": _a_numero(fila[i + 1] if i + 1 < len(fila) else None),
            })
        return pasos

    filas = []
    for fila in matriz:
        if len(fila) < MIN_COLUMNAS_DIARIO:
            continue
        id_obra, id_act = celda(fila, "id_obra"), celda(fila, "id_act")
        actividad, step = celda(fila, "actividad"), celda(fila, "step")
        # Sin ID OBRA no hay dónde colgarla; sin nombre no hay qué mostrar.
        if not (id_obra and id_act and step):
            continue
        # «CON_BD_TAM» venía sin encabezados, «BD_Avance» los trae: sin
        # esto la fila de títulos entra como una obra llamada «00 OBRA».
        if id_obra == "00 OBRA" or id_act == "Activity ID":
            continue
        # Solo construcción. La fase vacía se deja pasar: un export que
        # no traiga la columna no debe quedar en nada.
        fase = celda(fila, "fase")
        if fase and fase not in FASES_TERRENO:
            continue
        filas.append({
            "id_obra": id_obra,
            "id_act": id_act,
            "actividad": actividad or step,
            "step": step,
            # Mientras STEP ID venga vacía, la actividad es su propio
            # Step y lleva su Activity ID en los dos identificadores.
            "cod_step": celda(fila, "cod_step") or id_act,
            "plataforma": celda(fila, "plataforma"),
            "unidad": celda(fila, "unidad"),
            "hh_lb": _a_numero(celda(fila, "hh_lb")),
            "cantidad_lb": _a_numero(celda(fila, "cantidad_lb")),
            "estado": celda(fila, "estado"),
            "fase": fase,
            "pasos": pasos_de(fila),
        })

    if not filas:
        raise ValueError(
            f"No se reconoció ninguna actividad. Se espera la hoja {HOJA_DIARIO} "
            "del libro BD_TAM: sin encabezados, con 00 OBRA en la columna 8.")
    return filas


def importar_bd_diario(filas):
    """Carga las actividades del diario en todas las obras de una vez.

    Reemplaza lo que cuelga de cada obra alcanzada; las obras que el
    export no menciona quedan intactas. Recorre el árbol completo, así
    que sirve igual para Obras Puntuales y para Obras Lineales.
    """
    conn = get_db()
    c = conn.cursor()

    # Todas las obras del árbol, por su ID OBRA. La obra es el nodo con
    # código con punto (EST.TAL, T2-E-VC.PLA), igual que en nodos_obra().
    destinos = {f["codigo"]: dict(f) for f in c.execute(
        "SELECT id, codigo, nombre FROM nodos "
        "WHERE activo=1 AND codigo LIKE '%.%'").fetchall()}

    por_obra, sin_enlace, corregidas = {}, {}, {}
    for f in filas:
        codigo = f["id_obra"]
        # Mismo arreglo que en el export de P6: algunos ID OBRA vienen
        # combinados (PTE.SMP para dos puentes distintos) y es el nombre
        # de la actividad el que los separa.
        if codigo not in destinos and f["actividad"] in CORRECCIONES_P6:
            corregido = CORRECCIONES_P6[f["actividad"]]
            if corregido in destinos:
                corregidas[f["actividad"]] = corregido
                codigo = corregido
        if codigo in destinos:
            por_obra.setdefault(codigo, []).append(f)
        else:
            sin_enlace.setdefault(f["id_obra"], []).append(f["id_act"])

    if not por_obra:
        conn.close()
        raise ValueError("Ningún ID OBRA del export coincide con una obra del árbol.")

    total_nodos = total_steps = 0
    for codigo, actividades in por_obra.items():
        destino = destinos[codigo]

        # Agrupa Actividad -> Steps respetando el orden del export.
        orden, steps_por_act = [], {}
        for f in actividades:
            if f["actividad"] not in steps_por_act:
                orden.append(f["actividad"])
                steps_por_act[f["actividad"]] = []
            steps_por_act[f["actividad"]].append(f)

        def nodo_step(s):
            """El Step, y debajo sus pasos si el programa los declaró.

            Con pasos el nodo deja de ser hoja: el supervisor elige uno
            más adentro, y ese es el que se reporta. Sin pasos —lo normal
            todavía— el Step se sigue reportando tal cual.
            """
            nodo = {
                "nombre": s["step"], "tipo": "step", "codigo": s["id_act"],
                "unidad": s["unidad"], "id_actividad": s["id_act"],
                "id_step": s["cod_step"], "hh_lb": s["hh_lb"],
                "cantidad_lb": s["cantidad_lb"],
            }
            # Un paso único no es una decisión: sería una pantalla con un
            # solo botón. El Step se queda con su nombre —más informativo
            # que el del paso— y adopta su código y su peso.
            if len(s.get("pasos") or []) == 1:
                unico = s["pasos"][0]
                nodo["id_step"] = unico["codigo"]
                nodo["peso"] = unico["peso"]
                return nodo
            if s.get("pasos"):
                nodo["tipo"] = "actividad"
                nodo["pregunta"] = "Seleccione el paso ejecutado"
                nodo["hijos"] = [{
                    "nombre": p["nombre"], "tipo": "step", "codigo": p["codigo"],
                    "unidad": s["unidad"], "id_actividad": s["id_act"],
                    "id_step": p["codigo"], "peso": p["peso"],
                    # La HH y la cantidad de línea base son de la actividad
                    # completa: al paso le toca su fracción según el peso.
                    "hh_lb": (s["hh_lb"] or 0) * (p["peso"] or 0) or None,
                    "cantidad_lb": (s["cantidad_lb"] or 0) * (p["peso"] or 0) or None,
                } for p in s["pasos"]]
            return nodo

        arbol = [{
            "nombre": nombre_act, "tipo": "actividad",
            "pregunta": "Seleccione la actividad detallada (Step)",
            "hijos": [nodo_step(s) for s in steps_por_act[nombre_act]],
        } for nombre_act in orden]

        _borrar_subarbol(c, destino["id"])           # solo los hijos
        usados = {f["id"] for f in c.execute("SELECT id FROM nodos").fetchall()}
        nuevas = []
        _aplanar(arbol, destino["id"], nuevas, usados)
        c.executemany(INSERT_NODO, nuevas)
        c.execute("UPDATE nodos SET pregunta=? WHERE id=?",
                  ("Seleccione la actividad", destino["id"]))
        total_nodos += len(nuevas)
        total_steps += len(actividades)

    conn.commit()
    conn.close()

    return {
        "obras": len(por_obra),
        "actividades": sum(len({f["actividad"] for f in v}) for v in por_obra.values()),
        "steps": total_steps,
        "nodos": total_nodos,
        "sin_enlace": {k: len(v) for k, v in sin_enlace.items()},
        "corregidas": corregidas,
        # Cuántas filas traían STEP ID propio. Sigue en 0: los pasos
        # llegan por las columnas P01..P10, no por esa columna.
        "con_step_propio": sum(1 for v in por_obra.values() for f in v
                               if f["cod_step"] != f["id_act"]),
        # Los pasos P01..P10: cuántas actividades los declaran y cuántos
        # nodos nuevos significan.
        "con_pasos": sum(1 for v in por_obra.values() for f in v if f.get("pasos")),
        "pasos": sum(len(f.get("pasos") or []) for v in por_obra.values() for f in v),
    }


# El programa, empaquetado junto a la app. En Streamlit Cloud el disco
# es efímero: la base nace vacía en cada reinicio, y lo que se cargue a
# mano desde ⚙️ Obras se pierde. Sin este archivo el supervisor abre la
# app y encuentra las 156 obras sin una sola actividad.
#
# Lo genera `exportar_programa()` desde el libro del diario, y pesa poco
# porque va comprimido: son las mismas filas que devuelve
# leer_bd_diario(), con sus pasos.
ARCHIVO_PROGRAMA = os.path.join(BASE_DIR, "programa_tam.json.gz")


def cargar_programa_empaquetado():
    """Siembra las actividades del programa que viajan con la app.

    Devuelve el resumen del importador, o None si no hay archivo.
    """
    if not os.path.exists(ARCHIVO_PROGRAMA):
        return None
    import gzip
    try:
        with gzip.open(ARCHIVO_PROGRAMA, "rt", encoding="utf-8") as f:
            filas = json.load(f)
        return importar_bd_diario(filas) if filas else None
    except Exception:                                # archivo corrupto
        return None


def seed_db():
    """Carga datos iniciales solo si la base está vacía."""
    conn = get_db()
    c = conn.cursor()

    if c.execute("SELECT COUNT(*) FROM nodos").fetchone()[0] == 0:
        filas = []
        _aplanar(arbol_demo(), None, filas, set())
        c.executemany(INSERT_NODO, filas)
        # El árbol recién sembrado son obras vacías. Las actividades del
        # programa vienen empaquetadas con la app: se cuelgan ahora, sin
        # que nadie tenga que subir el libro a mano después de cada
        # reinicio. `importar_bd_diario` abre su propia conexión, por eso
        # se confirma esta antes.
        conn.commit()
        cargar_programa_empaquetado()

    # La tabla `supervisores` arranca vacía a propósito. Antes se sembraba
    # con siete nombres de ejemplo —Juan Pérez, Carlos Soto…— y salían
    # como opciones en «¿A nombre de qué supervisor registra?»: un reporte
    # real podía quedar a nombre de alguien que no existe. Los supervisores
    # de verdad son los usuarios con perfil Supervisor, más los que dé de
    # alta Oficina Técnica o lleguen de la lista de SharePoint.

    if c.execute("SELECT COUNT(*) FROM maquinaria").fetchone()[0] == 0:
        c.executemany(
            "INSERT OR IGNORE INTO maquinaria (empresa, patente, tipo, activo, creado_en, creado_por) "
            "VALUES (?,?,?,1,?,'semilla')",
            [(e, p, t, datetime.now().strftime("%d-%m-%Y %H:%M"))
             for e, p, t in MAQUINARIA_INICIAL])

    # Usuarios definidos en Secrets: es la forma de tener cuentas reales
    # en Streamlit Cloud sin escribir contraseñas en el repositorio. Si
    # existen, mandan ellos y NO se crean los de demostración.
    usuarios_secretos = secreto("usuarios", defecto={})
    if usuarios_secretos:
        for datos in usuarios_secretos.values():
            usuario = str(datos.get("usuario", "")).strip().lower()
            if not usuario:
                continue
            c.execute("""INSERT INTO usuarios (usuario, clave, nombre, empresa, perfil, activo)
                         VALUES (?,?,?,?,?,1)
                         ON CONFLICT(usuario) DO UPDATE SET
                             clave=excluded.clave, nombre=excluded.nombre,
                             empresa=excluded.empresa, perfil=excluded.perfil, activo=1""",
                      (usuario, hash_clave(str(datos.get("password", ""))),
                       datos.get("nombre", usuario), datos.get("empresa", "Sacyr"),
                       datos.get("perfil", "Supervisor")))
            # Asignaciones de arranque: nombre de obra raíz o id de nodo,
            # separados por ';'. La bitácora de Sheets manda por encima.
            uid = c.execute("SELECT id FROM usuarios WHERE usuario=?", (usuario,)).fetchone()[0]
            for referencia in str(datos.get("obras", "")).split(";"):
                referencia = referencia.strip()
                if not referencia:
                    continue
                nodo = c.execute(
                    "SELECT id FROM nodos WHERE id=? UNION ALL "
                    "SELECT id FROM nodos WHERE padre_id IS NULL AND nombre=? LIMIT 1",
                    (referencia, referencia)).fetchone()
                if nodo:
                    c.execute("INSERT OR IGNORE INTO permisos (usuario_id, nodo_id) VALUES (?,?)",
                              (uid, nodo["id"]))
        conn.commit()
        conn.close()
        return

    if c.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0] == 0:
        # Los dos únicos usuarios de la marcha blanca. Los demás se
        # crean desde 👥 Usuarios, o llegan del flujo de SharePoint.
        usuarios = [
            ("jaime", "mamboscuro", "Jaime",         "Sacyr", "Supervisor"),
            ("admin", "1234",       "Administrador", "Sacyr", "Administrador"),
        ]
        c.executemany(
            "INSERT INTO usuarios (usuario, clave, nombre, empresa, perfil, activo) VALUES (?,?,?,?,?,1)",
            [(u, hash_clave(p), n, e, pf) for u, p, n, e, pf in usuarios])

        # Jaime es supervisor: sin permisos no vería nada, así que parte
        # con las dos ramas completas. Se recorta desde 👥 Asignar.
        uid = c.execute("SELECT id FROM usuarios WHERE usuario='jaime'").fetchone()[0]
        for r in c.execute("SELECT id FROM nodos WHERE padre_id IS NULL").fetchall():
            c.execute("INSERT OR IGNORE INTO permisos (usuario_id, nodo_id) VALUES (?,?)",
                      (uid, r["id"]))

    conn.commit()
    conn.close()


# ---------- Consultas de maquinaria ----------

def subcontratos_con_maquinaria():
    """Empresas que tienen parque cargado, para el selector del Step."""
    conn = get_db()
    filas = conn.execute(
        "SELECT DISTINCT empresa FROM maquinaria WHERE activo=1 ORDER BY empresa").fetchall()
    conn.close()
    return [f["empresa"] for f in filas]


def maquinaria_de(empresa):
    """Máquinas de un subcontrato, en orden de patente."""
    if not empresa:
        return []
    conn = get_db()
    filas = conn.execute(
        "SELECT patente, tipo FROM maquinaria WHERE activo=1 AND empresa=? "
        "ORDER BY tipo, patente", (empresa,)).fetchall()
    conn.close()
    return [dict(f) for f in filas]


def agregar_maquinaria(empresa, patente, tipo, usuario=""):
    """Da de alta una máquina que no estaba en el catálogo.

    Devuelve (ok, mensaje). La patente se normaliza a mayúsculas sin
    espacios: es lo que la vuelve un identificador y no un texto libre.
    """
    empresa = (empresa or "").strip().upper()
    patente = re.sub(r"\s+", "", (patente or "")).upper()
    tipo = (tipo or "").strip().upper()

    if not empresa:
        return False, "Indique el subcontrato dueño de la máquina."
    if not patente:
        return False, "Indique la patente."
    if not tipo:
        return False, "Indique el tipo de máquina (excavadora, camión tolva…)."

    conn = get_db()
    existe = conn.execute(
        "SELECT tipo, activo FROM maquinaria WHERE empresa=? AND patente=?",
        (empresa, patente)).fetchone()
    if existe:
        # Puede estar dada de baja: reactivarla es más sano que duplicarla.
        conn.execute("UPDATE maquinaria SET activo=1, tipo=? WHERE empresa=? AND patente=?",
                     (tipo, empresa, patente))
        conn.commit()
        conn.close()
        return True, f"{patente} ya estaba registrada en {empresa}: quedó disponible."

    conn.execute(
        "INSERT INTO maquinaria (empresa, patente, tipo, activo, creado_en, creado_por) "
        "VALUES (?,?,?,1,?,?)",
        (empresa, patente, tipo, datetime.now().strftime("%d-%m-%Y %H:%M"), usuario))
    conn.commit()
    conn.close()

    # Queda anotada en la bitácora: en Streamlit Cloud la base local se
    # borra en cada reinicio y sin esto el alta duraría una tarde.
    registrar_maquinaria_sheets(empresa, patente, tipo, usuario)
    return True, f"{patente} · {tipo} agregada a {empresa}."


# ============================================================
# 4. Consultas del árbol
# ============================================================

def hijos_de(padre_id):
    conn = get_db()
    filas = conn.execute(
        "SELECT * FROM nodos WHERE activo=1 AND padre_id IS ? ORDER BY orden, nombre",
        (padre_id,)).fetchall()
    conn.close()
    return [dict(f) for f in filas]


# ---------- Asignación de actividades por usuario ----------
#
# El id de cada nodo es su ruta completa dentro del árbol:
#
#   obras-puntuales/estaciones/estacion-talagante/muros-de-.../excavacion
#
# Eso permite resolver la visibilidad con prefijos, sin recorrer el
# árbol: si a alguien se le asigna una OBRA, ve todo lo que cuelga de
# ella; si se le asigna un Step suelto, sigue viendo los niveles de
# arriba —son el camino para llegar— pero nada más de ese ramal.

def nodos_asignados(usuario_id):
    conn = get_db()
    filas = conn.execute(
        "SELECT nodo_id FROM permisos WHERE usuario_id = ?", (usuario_id,)).fetchall()
    conn.close()
    return [f["nodo_id"] for f in filas]


def permisos_de(auth):
    """Nodos asignados al usuario de la sesión, cacheados por rerun.

    None significa «sin límite»: es lo que corresponde a los perfiles
    que ven todo (Administración, Oficina Técnica, Inspector).
    """
    if PERFILES.get(auth["perfil"], {}).get("ve_todo"):
        return None
    cache = st.session_state.get("_permisos_cache")
    if not cache or cache[0] != auth["id"]:
        cache = (auth["id"], nodos_asignados(auth["id"]))
        st.session_state["_permisos_cache"] = cache
    return cache[1]


def limpiar_cache_permisos():
    """Olvida lo cacheado del árbol y de los permisos.

    Se llama tras cada importación: es el único punto donde el árbol
    cambia, así que también es donde el índice del buscador caduca.
    """
    st.session_state.pop("_permisos_cache", None)
    try:
        indice_busqueda.clear()
    except Exception:                    # sin runtime de Streamlit (tests)
        pass


def es_visible(nodo_id, asignados):
    """¿El usuario puede ver este nodo o algo que cuelgue de él?"""
    if asignados is None:
        return True
    for a in asignados:
        if nodo_id == a or nodo_id.startswith(a + "/") or a.startswith(nodo_id + "/"):
            return True
    return False


def filtrar_visibles(nodos, asignados):
    if asignados is None:
        return nodos
    return [n for n in nodos if es_visible(n["id"], asignados)]


def hijos_visibles(padre_id, auth):
    """Hijos de un nodo, ya recortados a lo asignado al usuario."""
    return filtrar_visibles(hijos_de(padre_id), permisos_de(auth))


def raices_visibles(auth):
    """Obras que el usuario tiene permitido ver (Nivel 2)."""
    conn = get_db()
    filas = conn.execute(
        "SELECT * FROM nodos WHERE activo=1 AND padre_id IS NULL ORDER BY orden, nombre").fetchall()
    conn.close()
    return filtrar_visibles([dict(f) for f in filas], permisos_de(auth))


def contar_hijos(ids, asignados=None):
    """Nº de hijos de cada nodo, en una sola consulta.

    Con `asignados` se cuentan solo los hijos que el usuario puede ver:
    de otro modo la pantalla ofrecería «(11)» donde hay 2.
    """
    if not ids:
        return {}
    conn = get_db()
    marcas = ",".join("?" * len(ids))
    if asignados is None:
        filas = conn.execute(
            f"SELECT padre_id, COUNT(*) c FROM nodos WHERE activo=1 AND padre_id IN ({marcas}) "
            "GROUP BY padre_id", ids).fetchall()
        conn.close()
        return {f["padre_id"]: f["c"] for f in filas}

    filas = conn.execute(
        f"SELECT id, padre_id FROM nodos WHERE activo=1 AND padre_id IN ({marcas})", ids).fetchall()
    conn.close()
    conteos = {}
    for f in filas:
        if es_visible(f["id"], asignados):
            conteos[f["padre_id"]] = conteos.get(f["padre_id"], 0) + 1
    return conteos


def contar_hojas(nodo_id):
    """Total de Steps que cuelgan de un nodo (recursivo)."""
    conn = get_db()
    fila = conn.execute("""
        WITH RECURSIVE sub(id) AS (
            SELECT id FROM nodos WHERE id = ?
            UNION ALL
            SELECT n.id FROM nodos n JOIN sub s ON n.padre_id = s.id WHERE n.activo = 1
        )
        SELECT COUNT(*) c FROM sub s
        WHERE NOT EXISTS (SELECT 1 FROM nodos h WHERE h.padre_id = s.id AND h.activo = 1)
    """, (nodo_id,)).fetchone()
    conn.close()
    return fila["c"]


def ruta_de(nodo_id):
    """Ruta completa desde la raíz hasta el nodo (trazabilidad)."""
    conn = get_db()
    filas = conn.execute("""
        WITH RECURSIVE ruta(id, nombre, codigo, tipo, unidad, pregunta,
                            latitud, longitud, id_actividad, id_step,
                            hh_lb, cantidad_lb, padre_id, nivel) AS (
            SELECT id, nombre, codigo, tipo, unidad, pregunta,
                   latitud, longitud, id_actividad, id_step,
                   hh_lb, cantidad_lb, padre_id, 0 FROM nodos WHERE id = ?
            UNION ALL
            SELECT n.id, n.nombre, n.codigo, n.tipo, n.unidad, n.pregunta,
                   n.latitud, n.longitud, n.id_actividad, n.id_step,
                   n.hh_lb, n.cantidad_lb, n.padre_id, r.nivel + 1
            FROM nodos n JOIN ruta r ON n.id = r.padre_id
        )
        SELECT * FROM ruta ORDER BY nivel DESC
    """, (nodo_id,)).fetchall()
    conn.close()
    return [dict(f) for f in filas]


def _sin_tildes(texto):
    """Deja el texto comparable: sin tildes, sin ñ, en minúsculas.

    En terreno se escribe con el pulgar y sin acentos. «excavacion»,
    «Excavación» y «EXCAVACION» tienen que encontrar lo mismo, igual
    que «chinihue» y «Chiñihue».
    """
    import unicodedata
    plano = unicodedata.normalize("NFKD", (texto or "").casefold())
    return "".join(c for c in plano if not unicodedata.combining(c))


def _palabras(texto):
    """Trozos buscables: descarta signos y palabras de una sola letra."""
    import re
    return [p for p in re.split(r"[^0-9a-z]+", _sin_tildes(texto)) if len(p) > 1]


def _erratas(palabra, candidata, tope):
    """Distancia de edición, rendida en cuanto supera `tope`.

    Cuenta letras cambiadas, sobrantes o faltantes. Se corta apenas se
    pasa del presupuesto: no interesa cuán distintas son dos palabras
    que ya se descartaron.
    """
    if abs(len(palabra) - len(candidata)) > tope:
        return tope + 1
    previa = list(range(len(candidata) + 1))
    for i, a in enumerate(palabra, start=1):
        actual = [i]
        for j, b in enumerate(candidata, start=1):
            actual.append(min(previa[j] + 1,            # sobra una letra
                              actual[j - 1] + 1,        # falta una letra
                              previa[j - 1] + (a != b)))  # cambiada
        if min(actual) > tope:
            return tope + 1
        previa = actual
    return previa[-1]


def _parecido(palabra, candidata):
    """¿Son la misma palabra, tolerando dedazos?

    Acepta que la escrita sea el principio de la real —«excav» encuentra
    «excavación»— y perdona errores según el largo: uno hasta seis
    letras, dos de siete en adelante. Con eso «terraplan», «terrapln» y
    hasta «exacvacion» llegan a donde tienen que llegar, sin que
    palabras cortas y parecidas se confundan entre sí.
    """
    if candidata.startswith(palabra):
        return True
    if len(palabra) < 4:
        return False
    tope = 1 if len(palabra) < 7 else 2
    return _erratas(palabra, candidata, tope) <= tope


@st.cache_data(show_spinner=False)
def indice_busqueda():
    """Los nodos ya troceados en palabras, listos para comparar.

    Partir 5.800 nombres con expresiones regulares en cada tecla se
    nota; hacerlo una vez, no. Se rehace solo cuando cambia el árbol,
    desde `limpiar_cache_permisos()`.
    """
    conn = get_db()
    filas = conn.execute(
        "SELECT id, nombre, codigo, id_actividad, id_step, tipo FROM nodos "
        "WHERE activo = 1 AND nombre IS NOT NULL AND nombre <> ''").fetchall()
    conn.close()

    indice, vocabulario = [], set()
    for f in filas:
        # El id del nodo ya es su ruta en texto plano y sin tildes, así
        # que buscar ahí es buscar en el nombre y en toda su ruta a la
        # vez, sin una consulta por nodo.
        ruta_plana = f["id"].replace("/", " ").replace("-", " ")
        del_nodo = set(_palabras(f["nombre"]))
        de_ruta = set(_palabras(ruta_plana))
        vocabulario |= del_nodo | de_ruta
        indice.append((
            dict(f), del_nodo, de_ruta,
            [_sin_tildes(x) for x in (f["codigo"], f["id_actividad"], f["id_step"]) if x],
        ))
    return indice, vocabulario


def buscar_actividad(texto, auth, limite=12):
    """Busca por nombre y devuelve la actividad con su ruta.

    El supervisor busca lo que ve en terreno —«terraplen tramo 2»— no
    un código. Se buscan **todas las palabras**, en cualquier orden, en
    el nombre del nodo y en el de su ruta, así «terraplen melipilla»
    encuentra el terraplén que cuelga de esa estación.

    Es tolerante a propósito: ignora tildes y mayúsculas, acepta
    palabras a medias y aguanta un dedazo por palabra. Escribir mal es
    lo normal cuando se reporta con una mano y guantes.

    Los identificadores siguen sirviendo: quien llegue con un
    `C0008390` o un `T3-D-VC.PLA` en la mano también lo encuentra.

    Respeta lo asignado: cada perfil encuentra solo lo suyo.
    """
    texto = (texto or "").strip()
    if len(texto) < 3:
        return []

    buscadas = _palabras(texto)
    plano = _sin_tildes(texto)
    if not buscadas:
        return []

    indice, vocabulario = indice_busqueda()

    # Cada palabra escrita se resuelve UNA vez contra el vocabulario:
    # así los dedazos se calculan sobre ~3.000 palabras distintas y no
    # sobre las ~90.000 que aparecen repartidas por los nodos.
    aceptables = [{v for v in vocabulario if _parecido(p, v)} for p in buscadas]

    asignados = permisos_de(auth)
    candidatos = []
    for f, del_nodo, de_ruta, ids in indice:
        if not es_visible(f["id"], asignados):
            continue

        aciertos = en_nombre = 0
        for p, validas in zip(buscadas, aceptables):
            if del_nodo & validas:
                aciertos += 1
                en_nombre += 1
            elif (de_ruta & validas) or any(p in i for i in ids):
                aciertos += 1
        # Un identificador escrito entero vale por sí solo.
        if aciertos < len(buscadas) and not any(plano in i or i in plano for i in ids):
            continue

        nombre_plano = _sin_tildes(f["nombre"])
        candidatos.append((
            0 if plano in ids else 1,                     # es exactamente ese código
            0 if plano in nombre_plano else 1,            # el nombre contiene lo escrito
            -en_nombre,                                   # cuánto coincide en el nombre…
            0 if nombre_plano.startswith(plano) else 1,   # …y si además empieza por ahí
            0 if f["tipo"] == "step" else 1,              # el Step es lo que se reporta
            len(f["nombre"]),                             # el más corto es el más preciso
            f["nombre"], dict(f),
        ))

    candidatos.sort(key=lambda x: x[:7])
    resultados = []
    for *_, f in candidatos[:limite]:
        ruta = ruta_de(f["id"])
        resultados.append({
            "id": f["id"], "nombre": f["nombre"], "tipo": f["tipo"],
            "ids": " · ".join(dict.fromkeys(
                x for x in (f["codigo"], f["id_actividad"], f["id_step"]) if x)),
            "ruta": " › ".join(n["nombre"] for n in ruta[:-1]),
        })
    return resultados


def nodos_obra(raiz_id):
    """Las obras de una rama, con su ruta, sin importar la profundidad.

    Una obra es el nodo del que cuelgan las actividades. Se reconoce por
    su código: los de obra llevan punto —EST.TAL, T2-E-VC.PLA— y los de
    tipo no —EST, T2-E-VC—. Sirve igual para Obras Puntuales, que tiene
    dos niveles, y para Vías, que tiene cuatro.
    """
    conn = get_db()
    filas = conn.execute("""
        WITH RECURSIVE sub(id, nombre, codigo, ruta, nivel) AS (
            SELECT id, nombre, codigo, '', 0 FROM nodos WHERE id = ?
            UNION ALL
            SELECT n.id, n.nombre, n.codigo,
                   CASE WHEN s.ruta = '' THEN n.nombre ELSE s.ruta || ' › ' || n.nombre END,
                   s.nivel + 1
            FROM nodos n JOIN sub s ON n.padre_id = s.id
            WHERE n.activo = 1 AND s.nivel < 8
        )
        SELECT id, nombre, codigo, ruta FROM sub
        WHERE codigo LIKE '%.%' ORDER BY ruta
    """, (raiz_id,)).fetchall()
    conn.close()
    return [dict(f) for f in filas]


def es_hoja(nodo_id):
    conn = get_db()
    fila = conn.execute(
        "SELECT COUNT(*) c FROM nodos WHERE activo=1 AND padre_id = ?", (nodo_id,)).fetchone()
    conn.close()
    return fila["c"] == 0


def es_obra(nodo):
    """¿Es el nodo de obra individual (Nivel 3), el que agrupa Actividades?

    Se reconoce por el código, con la misma regla que usan nodos_obra()
    y ubicacion_de_ruta(): los de obra llevan punto —EST.TAL,
    T2-E-VC.PLA— y los de tipo o tramo no. Es el nivel donde vive el
    apartado de "Reportar actividad fuera del programa".
    """
    return "." in (nodo.get("codigo") or "")


# Niveles que la BD de actividades llama «Actividad». La rama demo usa
# «especialidad» y la importada, «actividad»: para el reporte son lo mismo.
TIPOS_ACTIVIDAD = ("actividad", "especialidad")


def ubicacion_de_ruta(ruta):
    """Dónde se ejecutó lo reportado, sacado de la ruta recorrida.

    La obra es el nodo cuyo código lleva punto —`EST.TAL`,
    `T1-G-VC.MOV`— y el tipo es su padre, que lleva el ID TIPO OBRA. Es
    la misma regla que usa `nodos_obra()`, y por eso vale para las dos
    ramas pese a que tienen distinta profundidad:

        Obras Puntuales › ESTACIONES › Estación Talagante › …
                          ^tipo        ^obra
        Obras Lineales › Vías › Tramo 1 › Vías General › Vía Carga ›
                                ^tramo    ^sector        ^tipo
          Movimiento de Tierras › …
          ^obra

    **El tramo y el sector no son adorno**: sin ellos `T1-G-VC.MOV` y
    `T1-E-VC.MOV` salen idénticos en la planilla —mismo tipo, misma
    obra— y no se sabe si el trabajo fue en la vía general o dentro de
    la estación. En Catenarias el sector va en el nombre del propio
    tipo («Catenarias Estación El Monte»), así que ahí queda vacío.

    En Obras Puntuales no hay tramo ni sector: quedan en blanco.
    """
    vacio = {"cod_obra": "", "tipo_obra": "", "nombre_obra": "",
             "tramo": "", "sector": ""}
    if not ruta:
        return vacio

    # El tramo se reconoce por el nombre: es el único nivel que se
    # llama así, tanto en Vías como en Catenarias.
    tramo = next((n["nombre"] for n in ruta
                  if (n.get("nombre") or "").lower().startswith("tramo")), "")

    for i in range(len(ruta) - 1, -1, -1):
        if "." in (ruta[i].get("codigo") or ""):
            padre = ruta[i - 1] if i > 0 else {}
            abuelo = ruta[i - 2] if i > 1 else {}
            sector = abuelo.get("nombre", "")
            # Si sobre el tipo está el tramo o la raíz, no hay sector
            # propio: el nombre del tipo ya lo dice.
            if sector == tramo or not abuelo.get("padre_id", abuelo.get("id")):
                sector = ""
            return {"cod_obra": ruta[i].get("codigo") or "",
                    "tipo_obra": padre.get("nombre", ""),
                    "nombre_obra": ruta[i]["nombre"],
                    "tramo": tramo, "sector": sector}

    # Árbol sin códigos (demo, o ramas aún sin planilla): se cae a las
    # posiciones de siempre para no quedarse sin nada que mostrar.
    return {"cod_obra": "",
            "tipo_obra": ruta[1]["nombre"] if len(ruta) > 1 else "",
            "nombre_obra": ruta[2]["nombre"] if len(ruta) > 2 else "",
            "tramo": tramo, "sector": ""}


def obra_de_ruta(ruta):
    """(TIPO_OBRA, OBRA) — atajo sobre ubicacion_de_ruta()."""
    u = ubicacion_de_ruta(ruta)
    return u["tipo_obra"], u["nombre_obra"]


def resolver_ids(ruta):
    """Identificadores de la BD de actividades para el Step reportado.

    El supervisor nunca ve ni escribe estos códigos: se deducen de la
    rama por la que navegó. Devuelve un diccionario con el nombre de la
    actividad, su ID ACT, el COD STEP y la línea base del Step.
    """
    if not ruta:
        return {"actividad": "", "id_actividad": "", "id_step": "",
                "actividad_nodo_id": "", "hh_lb": None, "cantidad_lb": None}

    step = ruta[-1]
    padres = ruta[:-1]

    # Nodo Actividad: el más cercano al Step por arriba. Si la rama no
    # declara el tipo (árbol demo), sirve el padre directo del Step.
    actividad = next((n for n in reversed(padres) if n.get("tipo") in TIPOS_ACTIVIDAD), None)
    if actividad is None and padres:
        actividad = padres[-1]

    id_actividad = (step.get("id_actividad")
                    or (actividad or {}).get("id_actividad")
                    or (actividad or {}).get("codigo") or "")

    return {
        "actividad": (actividad or {}).get("nombre", ""),
        "id_actividad": id_actividad,
        "id_step": step.get("id_step") or step.get("codigo") or "",
        "actividad_nodo_id": (actividad or {}).get("id", ""),
        "hh_lb": step.get("hh_lb"),
        "cantidad_lb": step.get("cantidad_lb"),
    }


def estado_nodo(nodo_id):
    """Estado en que está una actividad, según lo ya reportado.

    Se deduce del último reporte emitido sobre ese Step. Mientras no
    exista el cierre de actividades, es la única fuente que hay.
    """
    conn = get_db()
    fila = conn.execute(
        "SELECT estado FROM reportes WHERE nodo_id=? ORDER BY rowid DESC LIMIT 1",
        (nodo_id,)).fetchone()
    conn.close()
    return (fila["estado"] if fila and fila["estado"] else ESTADO_SIN_REPORTES)


def hh_calculada(horas, personal):
    """HH del reporte: horas de jornada × personal en el frente.

    Es la magnitud que se contrasta con las HH LB del programa. Con
    personal en 0 se devuelven las horas, para no perder el registro
    de una jornada trabajada sin dotación declarada.
    """
    horas = float(horas or 0)
    personal = int(personal or 0)
    return round(horas * personal, 2) if personal else round(horas, 2)


def hm_calculada(horas, maquinas):
    """HM del reporte: horas de jornada × máquinas en el frente.

    Mismo criterio que las HH, pero con el parque. Sin maquinaria
    declarada son 0: no hubo horas máquina que imputar.
    """
    return round(float(horas or 0) * int(maquinas or 0), 2)


# ============================================================
# 5. Guardado durable: Google Sheets + SharePoint
# ============================================================
#
# En el PC de la oficina la base SQLite basta. En Streamlit Cloud NO:
# el disco del contenedor es efímero y se borra en cada reinicio o
# redespliegue. Por eso el reporte viaja a dos destinos:
#
#   SQLite (local, rápido)  ──┬──> Google Sheets   histórico durable
#                             └──> webhook SP      respaldo
#
# Los dos son opcionales e independientes: si no hay Sheets configurado
# la app guarda igual, y si el webhook falla el reporte queda en la cola
# de reenvío. Lo que nunca ocurre es perder el reporte por no tener red.

HOJA_ENTREGA = "Reporte Diario"     # los reportes: igual que el Excel
HOJA_MAQUINARIA = "Maquinaria"
HOJA_ASIGNACIONES = "Asignaciones"

# Maquinaria y Asignaciones se guardan como bitácora: una fila por alta
# o por cambio, nunca se borra ni se reescribe nada. Agregar filas es la
# operación barata en Sheets, y al arrancar se reconstruye el estado
# reproduciendo la bitácora en orden.
ENCABEZADOS_SHEETS = {
    HOJA_MAQUINARIA:   ["Empresa", "Patente", "Tipo", "Activo", "CreadoEn", "CreadoPor"],
    HOJA_ASIGNACIONES: ["Usuario", "NodoId", "Accion", "Fecha", "Por"],
}


def secreto(*claves, defecto=None):
    """Lee st.secrets sin reventar cuando no hay secrets configurados.

    Fuera de Streamlit Cloud lo normal es que no exista el archivo, y
    `st.secrets[...]` lanza excepción en ese caso.
    """
    try:
        valor = st.secrets
        for clave in claves:
            valor = valor[clave]
        return valor
    except Exception:
        return defecto


def hay_sheets():
    return HAY_GSPREAD and secreto("google_sheets", "sheet_id") is not None


# La misma cuenta de servicio sirve para la hoja y para subir las fotos
# a Drive: por eso se piden los dos permisos de una vez. `drive.file`
# solo da acceso a los archivos que crea esta app, no a todo el Drive.
ALCANCES_GOOGLE = ["https://www.googleapis.com/auth/spreadsheets",
                   "https://www.googleapis.com/auth/drive.file"]


def credenciales_google():
    """Credenciales de la cuenta de servicio, o None si no hay Secrets."""
    datos = secreto("google_sheets", "credentials", defecto=None)
    if not datos:
        return None
    try:
        return CredencialesGoogle.from_service_account_info(
            dict(datos), scopes=ALCANCES_GOOGLE)
    except Exception:                                # credenciales malas
        return None


@st.cache_resource(show_spinner=False)
def cliente_sheets():
    """Abre el Google Sheet de Secrets. None si no está configurado.

    Va en cache_resource porque la conexión se reusa entre rerun: montar
    las credenciales en cada interacción agregaría un segundo largo a
    cada toque de pantalla.
    """
    if not hay_sheets():
        return None
    try:
        credenciales = credenciales_google()
        if credenciales is None:
            return None
        return gspread.authorize(credenciales).open_by_key(
            secreto("google_sheets", "sheet_id"))
    except Exception:                                # credenciales malas, sin red
        return None


# ---------- Fotos en Google Drive ----------
#
# Publicada la app, el disco es efímero: lo que se escriba en
# `fotos_reportes/` desaparece al reiniciar. La evidencia tiene que
# vivir fuera, y la hoja guardar un enlace en vez de un archivo.
#
# Se usa la API REST con urllib en vez de google-api-python-client: es
# una sola petición multipart y no vale la pena arrastrar la librería
# entera solo para esto.

DRIVE_SUBIDA = "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart"
DRIVE_PERMISOS = "https://www.googleapis.com/drive/v3/files/{}/permissions"

# Cómo se ve la foto una vez subida:
#   ENLACE_VER    abre el visor de Drive (respeta permisos, siempre sirve)
#   ENLACE_DIRECTO devuelve el archivo en crudo, que es lo que necesita
#                  =IMAGE() para pintar la miniatura dentro de la hoja
ENLACE_VER = "https://drive.google.com/file/d/{}/view"
ENLACE_DIRECTO = "https://lh3.googleusercontent.com/d/{}"


def diagnostico_sheets():
    """Por qué no se pudo abrir el Sheet. Devuelve (ok, mensaje, pistas).

    `cliente_sheets()` devuelve None ante cualquier problema, y hace
    bien: en terreno la causa da igual, el reporte se guarda local y se
    reintenta. Pero al configurar la app la causa es lo único que
    importa, y las tres posibles —API sin habilitar, id equivocado,
    hoja sin compartir— se arreglan en pantallas distintas de Google.
    """
    if not HAY_GSPREAD:
        return False, "Faltan las librerías `gspread` y `google-auth`.", []

    sheet_id = (secreto("google_sheets", "sheet_id") or "").strip()
    datos = secreto("google_sheets", "credentials")
    if not datos:
        return False, "No hay `[google_sheets.credentials]` en Secrets.", []

    correo = dict(datos).get("client_email", "(sin client_email)")
    proyecto = dict(datos).get("project_id", "")

    if "/" in sheet_id or sheet_id.lower().startswith("http"):
        return False, "El `sheet_id` no es un id: parece la URL completa.", [
            "Copie solo lo que va entre `/d/` y `/edit` en la URL del Sheet."]

    try:
        credenciales = CredencialesGoogle.from_service_account_info(
            dict(datos), scopes=ALCANCES_GOOGLE)
    except Exception as error:
        return False, f"Las credenciales no se pudieron leer: {error}", [
            "Revise que `private_key` vaya entre comillas, en una sola línea y "
            "con los `\\n` tal como vienen en el archivo JSON."]

    try:
        libro = gspread.authorize(credenciales).open_by_key(sheet_id)
        return True, libro.title, []
    except Exception as error:
        texto = str(error)
        if "SERVICE_DISABLED" in texto or "has not been used in project" in texto:
            return False, "La API de Google Sheets no está habilitada en el proyecto.", [
                f"Entre a Google Cloud → proyecto **{proyecto}** → APIs y servicios "
                "→ Biblioteca, y habilite **Google Sheets API** y **Google Drive API**.",
                "Tarda un par de minutos en surtir efecto."]
        if "404" in texto or "not found" in texto.lower():
            return False, "Google dice que no existe ninguna hoja con ese `sheet_id`.", [
                "El id está mal copiado, o es el de otro archivo.",
                f"Confirme que la hoja compartida con `{correo}` es la misma del id."]
        if "403" in texto or "permission" in texto.lower():
            return False, "La cuenta de servicio no tiene permiso sobre esa hoja.", [
                f"Comparta el Sheet como **Editor** con `{correo}`.",
                "Ojo con tener dos cuentas de servicio parecidas: el correo "
                "compartido tiene que ser exactamente ese."]
        return False, f"Google respondió: {texto[:300]}", []


def webhook_fotos():
    """URL del Apps Script que guarda las fotos, si está configurado."""
    return (secreto("google_sheets", "webhook_fotos") or "").strip()


def hay_drive():
    """¿Se pueden subir fotos?

    Dos caminos, y basta con uno:

    - **Apps Script** (`webhook_fotos`): un script en la cuenta del
      dueño del Drive recibe la foto y la guarda él. Es el que funciona
      con una cuenta Gmail normal.
    - **Cuenta de servicio** (`carpeta_fotos`): la API de Drive directa.
      Solo sirve contra una **unidad compartida** de Workspace: desde
      2024 una cuenta de servicio no tiene cuota propia y no puede ser
      dueña de un archivo en «Mi unidad».
    """
    if webhook_fotos():
        return True
    return bool(HAY_GSPREAD and secreto("google_sheets", "credentials")
                and secreto("google_sheets", "carpeta_fotos"))


def subir_foto_por_script(nombre, contenido, publica=False):
    """Manda la foto al Apps Script. Devuelve (ok, id de Drive o error).

    El script corre **como el dueño de la cuenta**, así que el archivo
    es suyo y ocupa su cuota. La app no guarda credenciales de Google:
    la URL del script hace de llave, igual que el webhook de SharePoint.
    """
    import urllib.error
    import urllib.request

    url = webhook_fotos()
    cuerpo = json.dumps({
        "nombre": nombre,
        "contenido": base64.b64encode(contenido).decode("ascii"),
        "tipo": "image/jpeg",
        "publica": bool(publica),
        "token": secreto("google_sheets", "token_fotos") or "",
    }).encode("utf-8")

    peticion = urllib.request.Request(
        url, data=cuerpo, method="POST",
        headers={"Content-Type": "application/json; charset=utf-8"})
    try:
        with urllib.request.urlopen(peticion, timeout=90) as r:
            respuesta = r.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as error:
        return False, f"Script HTTP {error.code}: {error.read().decode('utf-8')[:200]}"
    except Exception as error:
        return False, f"Script de fotos: {error}"

    try:
        datos = json.loads(respuesta)
    except ValueError:
        # Apps Script devuelve HTML cuando el despliegue pide iniciar
        # sesión: es el error de configuración más común.
        if "accounts.google.com" in respuesta or "<html" in respuesta.lower():
            return False, ("El script pide iniciar sesión: despliéguelo con acceso "
                           "«Cualquier persona» y ejecutándose como usted.")
        return False, f"El script no devolvió JSON: {respuesta[:160]}"

    if not datos.get("ok"):
        return False, f"El script rechazó la foto: {datos.get('error', respuesta[:160])}"
    if not datos.get("id"):
        return False, "El script no devolvió el id del archivo."
    return True, datos["id"]


def subir_foto_drive(nombre, contenido, publica=False):
    """Sube una imagen a Drive. Devuelve (ok, id del archivo o error)."""
    import urllib.error
    import urllib.request

    if webhook_fotos():
        return subir_foto_por_script(nombre, contenido, publica)

    credenciales = credenciales_google()
    carpeta = secreto("google_sheets", "carpeta_fotos")
    if credenciales is None or not carpeta:
        return False, "Google Drive no está configurado."

    try:
        from google.auth.transport.requests import Request as PeticionGoogle
        credenciales.refresh(PeticionGoogle())
    except Exception as error:
        return False, f"No se pudo autenticar con Drive: {error}"

    limite = "===============rep-app=="
    metadatos = json.dumps({"name": nombre, "parents": [carpeta]})
    cuerpo = b"".join([
        f"--{limite}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n".encode(),
        metadatos.encode("utf-8"),
        f"\r\n--{limite}\r\nContent-Type: image/jpeg\r\n\r\n".encode(),
        contenido,
        f"\r\n--{limite}--\r\n".encode(),
    ])
    peticion = urllib.request.Request(
        DRIVE_SUBIDA, data=cuerpo, method="POST",
        headers={"Authorization": f"Bearer {credenciales.token}",
                 "Content-Type": f"multipart/related; boundary={limite}"})
    try:
        with urllib.request.urlopen(peticion, timeout=60) as r:
            archivo_id = json.loads(r.read().decode("utf-8")).get("id")
    except urllib.error.HTTPError as error:
        detalle = error.read().decode("utf-8", errors="replace")
        # Desde 2024 una cuenta de servicio NO tiene cuota propia: no
        # puede ser dueña de un archivo en «Mi unidad», por más que la
        # carpeta esté compartida con ella. El mensaje de Google se
        # confunde fácil con un problema de permisos, y no lo es.
        if "storage quota" in detalle or "storageQuotaExceeded" in detalle:
            return False, (
                "SIN_CUOTA: una cuenta de servicio no puede ser dueña de archivos "
                "en Drive. Hay que subir con una cuenta de usuario (ver "
                "PUBLICAR-rep-app.md, «Las fotos»).")
        return False, f"Drive HTTP {error.code}: {detalle[:200]}"
    except Exception as error:
        return False, f"Drive: {error}"

    if not archivo_id:
        return False, "Drive no devolvió el id del archivo."

    # Sin esto la miniatura de =IMAGE() sale en blanco para quien no
    # tenga acceso. Es una decisión del usuario, no un automatismo.
    if publica:
        permiso = json.dumps({"role": "reader", "type": "anyone"}).encode("utf-8")
        try:
            urllib.request.urlopen(urllib.request.Request(
                DRIVE_PERMISOS.format(archivo_id), data=permiso, method="POST",
                headers={"Authorization": f"Bearer {credenciales.token}",
                         "Content-Type": "application/json"}), timeout=30)
        except Exception:
            pass                                     # queda privada, pero subida

    return True, archivo_id


def subir_fotos_a_drive(reporte, config=None):
    """Sube las fotos del reporte. Devuelve (enlaces, mensaje de error).

    Guarda el resultado en el reporte para no volver a subirlas si hay
    que reintentar el envío: en terreno un reintento es lo normal, y
    duplicar la evidencia en Drive ensucia la carpeta.

    El error se devuelve en vez de tragarse: si la carpeta no está
    compartida o falta habilitar la API, la única forma de enterarse
    era que la columna FOTO llegara vacía a la planilla.
    """
    guardados = json.loads(reporte["fotos_url"] or "[]")
    if guardados or not hay_drive():
        return guardados, ""

    rutas = json.loads(reporte["fotos"] or "[]")
    if not rutas:
        return [], ""

    config = config or cargar_config_sync()
    publica = bool(config.get("fotos_publicas"))
    enlaces, fallos = [], []
    for i, ruta in enumerate(rutas, start=1):
        completa = os.path.join(BASE_DIR, ruta)
        if not os.path.exists(completa):
            fallos.append(f"{os.path.basename(ruta)}: ya no está en el disco")
            continue
        with open(completa, "rb") as f:
            ok, resultado = subir_foto_drive(
                f"{reporte['id']}_{i:02d}.jpg", f.read(), publica)
        if ok:
            enlaces.append({"id": resultado,
                            "ver": ENLACE_VER.format(resultado),
                            "directo": ENLACE_DIRECTO.format(resultado)})
        else:
            fallos.append(resultado)

    if enlaces:
        conn = get_db()
        conn.execute("UPDATE reportes SET fotos_url=? WHERE id=?",
                     (json.dumps(enlaces), reporte["id"]))
        conn.commit()
        conn.close()
    return enlaces, " · ".join(fallos)


def enlaces_fotos(reporte, config=None):
    """Los enlaces de Drive del reporte, subiéndolas si hace falta."""
    return subir_fotos_a_drive(reporte, config)[0]


def celda_foto(enlaces, publica):
    """Lo que va en la columna FOTO de la hoja.

    Con las fotos públicas se usa =IMAGE(): la evidencia se ve dentro de
    la planilla, sin abrir nada. Si son privadas, un hipervínculo con
    texto —una imagen que Drive no deja cargar saldría como celda rota—.
    """
    if not enlaces:
        return "Sin fotografía"
    if publica:
        return f'=IMAGE("{enlaces[0]["directo"]}")'
    if len(enlaces) == 1:
        return f'=HYPERLINK("{enlaces[0]["ver"]}";"Ver foto")'
    return " ".join(e["ver"] for e in enlaces)


def hoja_sheets(nombre, encabezados=None):
    """Devuelve la pestaña pedida, creándola con sus encabezados."""
    libro = cliente_sheets()
    if libro is None:
        return None
    encabezados = encabezados or ENCABEZADOS_SHEETS.get(nombre) or []
    try:
        hoja = libro.worksheet(nombre)
    except gspread.WorksheetNotFound:
        hoja = libro.add_worksheet(title=nombre, rows=2000,
                                   cols=max(len(encabezados), 10))
        if encabezados:
            hoja.append_row(encabezados)
        return hoja
    if not encabezados:
        return hoja

    actuales = hoja.row_values(1)
    if not actuales:
        hoja.append_row(encabezados)
        return hoja

    # La hoja ya existía con menos columnas: se completa la fila de
    # títulos. Pasa al agregar una columna al final —`ENVIADO SP`, por
    # ejemplo— sobre una hoja que ya lleva reportes. Sin esto la columna
    # se llenaría con datos pero sin encabezado, y un flujo que lea la
    # hoja por nombre de columna no la vería nunca.
    if len(encabezados) > len(actuales) and encabezados[:len(actuales)] == actuales:
        hoja.update([encabezados], f"A1:{gspread.utils.rowcol_to_a1(1, len(encabezados))}")
    return hoja


def agregar_filas_sheets(nombre, filas, encabezados=None):
    """Agrega filas al final de una pestaña. Devuelve (ok, mensaje)."""
    if not filas:
        return True, "Nada que agregar."
    if not hay_sheets():
        return False, "Google Sheets no está configurado."
    try:
        hoja = hoja_sheets(nombre, encabezados)
        if hoja is None:
            return False, "No se pudo abrir el Google Sheet (revise Secrets)."
        limpias = [[("" if v is None else v) for v in fila] for fila in filas]
        hoja.append_rows(limpias, value_input_option="USER_ENTERED")
        return True, f"{len(filas)} fila(s) en «{nombre}»."
    except Exception as error:
        return False, f"Google Sheets: {error}"


def leer_filas_sheets(nombre):
    """Lee una pestaña completa como lista de diccionarios."""
    if not hay_sheets():
        return []
    try:
        hoja = hoja_sheets(nombre)
        return hoja.get_all_records() if hoja is not None else []
    except Exception:
        return []


def columnas_entrega():
    """Los encabezados de la hoja que se entrega.

    Son EXACTAMENTE los del Excel —se piden a la misma función que lo
    arma— más el folio delante. Derivarlos en vez de repetirlos es lo
    que evita que la planilla y la hoja se separen con el tiempo:
    agregar una columna al Excel la agrega aquí sola.
    """
    columnas, _ = _columnas_excel(MAQUINAS_EXCEL)
    return ["FOLIO"] + columnas


def fila_entrega(reporte, enlaces=None, publica=False):
    """El reporte en el orden de columnas_entrega().

    El .xlsx de un reporte crece si trae más máquinas que
    `MAQUINAS_EXCEL`; la hoja y la lista de SharePoint no pueden crecer,
    porque sus encabezados ya están escritos. Las máquinas que no caben
    se juntan en la última columna: perder el dato sería peor, y dejar
    la fila más larga que los encabezados correría todo lo demás.
    """
    fila = _fila_excel(reporte)
    base = len(COLUMNAS_EXCEL_BASE)
    maquinas, observaciones = fila[base:-1], fila[-1]
    if len(maquinas) > MAQUINAS_EXCEL:
        sobrantes = [m for m in maquinas[MAQUINAS_EXCEL - 1:] if m]
        maquinas = maquinas[:MAQUINAS_EXCEL - 1] + [" / ".join(sobrantes)]

    # _fila_excel llega hasta OBS; la foto es lo último y aquí es un
    # enlace, no una imagen incrustada como en el .xlsx.
    return ([reporte["id"]] + fila[:base] + maquinas + [observaciones]
            + [celda_foto(enlaces or [], publica)])


def fila_prueba_entrega(auth):
    """Una fila reconocible para probar la conexión con la hoja.

    Se arma por nombre de columna, no por posición: así sigue calzando
    si el Excel cambia sus columnas.
    """
    valores = {
        "FOLIO": "PRUEBA-" + str(uuid.uuid4())[:8].upper(),
        "ACTIVIDAD": "Registro de prueba desde rep-app",
        "FECHA": date.today().strftime("%d-%m-%Y"),
        "SUBCONTRATO": auth.get("empresa", ""),
        "SUPERVISOR": auth.get("nombre", auth.get("usuario", "")),
        "OBS": "Fila de prueba: se puede borrar.",
    }
    return [valores.get(columna, "") for columna in columnas_entrega()]


# Columna de control al final de la hoja. No es un dato del reporte
# —por eso no está en el Excel ni en el contrato de SharePoint— sino la
# marca que necesita el flujo de Power Automate que copia las filas a
# SharePoint: lee las que digan NO, crea el elemento y las pasa a SI.
# Es lo que evita que una segunda pasada duplique lo ya enviado.
COLUMNA_CONTROL_SP = "ENVIADO SP"
CONTROL_PENDIENTE = "NO"


def columnas_hoja():
    """Los encabezados reales de la pestaña: la entrega más el control."""
    return columnas_entrega() + [COLUMNA_CONTROL_SP]


def enviar_a_sheets(reporte, config=None):
    """Escribe el reporte en «Reporte Diario», la única pestaña de datos.

    Lo que se guarda es **exactamente lo que exporta el Excel**: mismas
    columnas, mismos valores, en el mismo orden. Antes había además una
    pestaña con los 50 campos del contrato de SharePoint, y sobraba:
    obligaba a mirar dos tablas distintas para el mismo reporte.

    Al final va `ENVIADO SP`, que no es del reporte: la usa el flujo que
    lleva las filas a SharePoint para saber cuáles le faltan.

    Las fotos no se copian a la hoja: se sube el archivo a Drive y la
    celda guarda su enlace.
    """
    config = config or cargar_config_sync()
    enlaces = enlaces_fotos(reporte, config)
    fila = fila_entrega(reporte, enlaces, bool(config.get("fotos_publicas")))
    return agregar_filas_sheets(
        HOJA_ENTREGA, [fila + [CONTROL_PENDIENTE]], columnas_hoja())


# ---------- Bitácora de maquinaria y asignaciones ----------

def registrar_maquinaria_sheets(empresa, patente, tipo, usuario):
    agregar_filas_sheets(HOJA_MAQUINARIA, [[
        empresa, patente, tipo, 1,
        datetime.now().strftime("%d-%m-%Y %H:%M"), usuario or ""]])


def registrar_asignacion_sheets(usuario, nodo_id, accion, por):
    agregar_filas_sheets(HOJA_ASIGNACIONES, [[
        usuario, nodo_id, accion,
        datetime.now().strftime("%d-%m-%Y %H:%M"), por or ""]])


@st.cache_resource(show_spinner="Recuperando configuración…")
def restaurar_desde_sheets():
    """Reconstruye maquinaria y asignaciones desde la bitácora.

    En Streamlit Cloud la base local nace vacía en cada arranque: sin
    esto, las máquinas que dio de alta terreno y las actividades que
    asignó Oficina Técnica se perderían en cada reinicio. Se ejecuta una
    sola vez por contenedor (cache_resource).
    """
    if not hay_sheets():
        return {"maquinaria": 0, "asignaciones": 0}

    conn = get_db()
    c = conn.cursor()

    maquinas = leer_filas_sheets(HOJA_MAQUINARIA)
    for fila in maquinas:
        empresa = str(fila.get("Empresa", "")).strip().upper()
        patente = str(fila.get("Patente", "")).strip().upper()
        if not empresa or not patente:
            continue
        c.execute(
            "INSERT INTO maquinaria (empresa, patente, tipo, activo, creado_en, creado_por) "
            "VALUES (?,?,?,?,?,?) ON CONFLICT(empresa, patente) DO UPDATE SET "
            "tipo=excluded.tipo, activo=excluded.activo",
            (empresa, patente, str(fila.get("Tipo", "")).strip().upper(),
             1 if str(fila.get("Activo", 1)) not in ("0", "False", "") else 0,
             str(fila.get("CreadoEn", "")), str(fila.get("CreadoPor", ""))))

    # La bitácora se reproduce en orden SOBRE lo ya sembrado: son
    # cambios, no un inventario. Así conviven las asignaciones de
    # arranque —las de Secrets o las del código— con las que hizo
    # Oficina Técnica desde la app, y quitar una sembrada también queda
    # anotado y se respeta.
    movimientos = leer_filas_sheets(HOJA_ASIGNACIONES)
    for fila in movimientos:
        usuario = str(fila.get("Usuario", "")).strip()
        nodo_id = str(fila.get("NodoId", "")).strip()
        accion = str(fila.get("Accion", "")).strip().lower()
        if not usuario or not nodo_id:
            continue
        cuenta = c.execute("SELECT id FROM usuarios WHERE usuario=?", (usuario,)).fetchone()
        if cuenta is None:
            continue
        if accion == "quitar":
            _quitar_local(c, cuenta["id"], nodo_id)
        else:
            _asignar_local(c, cuenta["id"], nodo_id)

    conn.commit()
    conn.close()
    return {"maquinaria": len(maquinas), "asignaciones": len(movimientos)}


# ============================================================
# 5b. Envío a SharePoint (webhook)
# ============================================================
#
# Flujo completo:
#
#   rep-app.py  --POST JSON-->  webhook  -->  lista de SharePoint
#                               (Power Automate / Apps Script)
#
# El reporte se guarda SIEMPRE primero en SQLite y recién después se
# intenta enviar. Si el envío falla —sin red, flujo caído, URL mal
# escrita— el reporte no se pierde: queda con sincronizado=0 y se
# reintenta desde la pantalla de configuración. Es el mismo criterio
# offline-first de la app Reporte-diario.


def cargar_config_sync():
    """Configuración del envío, de la más débil a la más fuerte:

        valores de fábrica  <  rep_app_config.json  <  Secrets

    Secrets manda porque en Streamlit Cloud el archivo JSON vive en un
    disco efímero: lo que se guarde ahí desaparece en el próximo
    reinicio. Las URL y el token son credenciales y su lugar es Secrets.
    """
    config = dict(CONFIG_SYNC_POR_DEFECTO)
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                config.update(json.load(f))
        except (OSError, ValueError):
            pass                       # configuración ilegible: se usa la de fábrica

    for clave in ("modo", "url", "url_login", "token", "enviar_fotos", "max_fotos", "fotos_publicas"):
        valor = secreto("sharepoint", clave)
        if valor is not None:
            config[clave] = valor
    return config


def config_en_secrets():
    """Claves del envío que vienen fijadas desde Secrets (no editables)."""
    return {c for c in ("modo", "url", "url_login", "token", "enviar_fotos", "max_fotos", "fotos_publicas")
            if secreto("sharepoint", c) is not None}


def guardar_config_sync(config):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def _iso_fecha(texto):
    """dd-mm-aaaa -> aaaa-mm-dd. Una columna Fecha de SharePoint no
    acepta el formato chileno: hay que mandarle ISO 8601."""
    try:
        return datetime.strptime(texto, "%d-%m-%Y").strftime("%Y-%m-%d")
    except (TypeError, ValueError):
        return texto or ""


def _iso_datetime(texto):
    """dd-mm-aaaa HH:MM -> aaaa-mm-ddTHH:MM:SS."""
    try:
        return datetime.strptime(texto, "%d-%m-%Y %H:%M").strftime("%Y-%m-%dT%H:%M:%S")
    except (TypeError, ValueError):
        return texto or ""


def payload_reporte(reporte, incluir_fotos=True):
    """El reporte como JSON plano para el flujo de Power Automate.

    Las claves son las de `columnas_sp()` y los valores salen de la
    **misma fila que se escribe en la hoja y en el .xlsx**: SharePoint
    no puede recibir algo distinto de lo que ve terreno.

    Dos ajustes que impone SharePoint, y son de formato, no de dato:
    la fecha viaja en ISO —lo que espera una columna de tipo Fecha— y
    `Foto` lleva los enlaces de Drive, porque una lista no guarda la
    imagen incrustada.
    """
    config = cargar_config_sync()
    fotos = json.loads(reporte["fotos"] or "[]")
    # Los que ya están en Drive; no se sube nada desde aquí, de eso se
    # encarga el guardado en la hoja.
    enlaces = json.loads(reporte["fotos_url"] or "[]")

    nombres = [nombre for nombre, _, _ in columnas_sp()]
    datos = dict(zip(nombres, fila_entrega(reporte, enlaces)))
    datos["Fecha"] = _iso_fecha(reporte["fecha"])
    datos["Foto"] = " ".join(e["ver"] for e in enlaces)

    # Las fotos viajan aparte, en base64, para que el flujo las cree
    # como archivos en una biblioteca documental. Van fuera de las
    # columnas de la lista porque no son un campo de texto.
    if incluir_fotos and config.get("enviar_fotos"):
        adjuntos = []
        for ruta in fotos[: config.get("max_fotos", 4)]:
            completa = os.path.join(BASE_DIR, ruta)
            if not os.path.exists(completa):
                continue
            with open(completa, "rb") as f:
                adjuntos.append({
                    "nombre": f"{reporte['id']}_{os.path.basename(ruta)}",
                    "contenido": base64.b64encode(f.read()).decode("ascii"),
                })
        datos["Fotos"] = adjuntos

    return datos


def enviar_a_webhook(payload, config=None):
    """POST del JSON al webhook. Devuelve (ok, mensaje).

    Se usa urllib de la biblioteca estándar a propósito: la app no
    debe depender de paquetes extra para algo tan básico.
    """
    import urllib.error
    import urllib.request

    config = config or cargar_config_sync()
    url = (config.get("url") or "").strip()
    if not url:
        return False, "No hay URL de webhook configurada."

    cuerpo = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    cabeceras = {"Content-Type": "application/json; charset=utf-8"}
    if config.get("token"):
        cabeceras["X-Token"] = config["token"]

    peticion = urllib.request.Request(url, data=cuerpo, headers=cabeceras, method="POST")
    try:
        with urllib.request.urlopen(peticion, timeout=config.get("tiempo_limite", 30)) as r:
            respuesta = r.read().decode("utf-8", errors="replace")[:300]
            return True, f"HTTP {r.status} · {respuesta or 'sin cuerpo'}"
    except urllib.error.HTTPError as error:
        detalle = error.read().decode("utf-8", errors="replace")[:300]
        return False, f"HTTP {error.code} · {detalle}"
    except urllib.error.URLError as error:
        return False, f"Sin conexión con el webhook: {error.reason}"
    except Exception as error:                       # timeout, DNS, TLS...
        return False, f"Error al enviar: {error}"


def destinos_activos(config=None):
    """Qué destinos hay configurados ahora mismo.

    Drive es un destino como los otros dos: las fotos tienen que salir
    del contenedor aunque la hoja no esté configurada.
    """
    config = config or cargar_config_sync()
    return {"sheets": hay_sheets(), "webhook": config.get("modo") == "webhook",
            "drive": hay_drive()}


def sincronizar_reporte(reporte_id):
    """Envía el reporte a Google Sheets y al webhook, y anota el resultado.

    Cada destino se intenta por separado y solo si aún no está logrado:
    reintentar un reporte que ya llegó a Sheets duplicaría la fila.
    `sincronizado` queda en 1 cuando todos los destinos activos están
    cubiertos; mientras tanto el reporte sigue en la cola de reenvío.
    """
    config = cargar_config_sync()
    activos = destinos_activos(config)

    conn = get_db()
    fila = conn.execute("SELECT * FROM reportes WHERE id=?", (reporte_id,)).fetchone()
    if fila is None:
        conn.close()
        return False, "El reporte no existe."

    if not any(activos.values()):
        conn.close()
        return False, "Sin destinos configurados: el reporte queda solo en la base local."

    avisos, errores = [], []

    # Drive va PRIMERO y por su cuenta. Antes las fotos se subían dentro
    # del envío a Sheets, así que con la hoja mal configurada —o caída—
    # la evidencia se quedaba en el disco del contenedor, que se borra.
    # Subida antes, la hoja y SharePoint reciben el enlace ya hecho.
    drive_ok = True
    if activos["drive"]:
        enlaces, fallo = subir_fotos_a_drive(fila, config)
        if fallo:
            drive_ok = False
            errores.append(f"Drive: {fallo}")
        elif enlaces:
            avisos.append(f"Drive: {len(enlaces)} foto(s)")
        fila = conn.execute("SELECT * FROM reportes WHERE id=?", (reporte_id,)).fetchone()

    sheets_ok = bool(fila["sheets_ok"])
    if activos["sheets"] and not sheets_ok:
        sheets_ok, mensaje = enviar_a_sheets(fila, config)
        (avisos if sheets_ok else errores).append(f"Sheets: {mensaje}")

    # Armar el JSON cuesta leer las fotos en base64, así que no se arma
    # si no hay a dónde enviarlo.
    sp_ok = bool(fila["sp_ok"])
    if activos["webhook"] and not sp_ok:
        sp_ok, mensaje = enviar_a_webhook(payload_reporte(fila), config)
        (avisos if sp_ok else errores).append(f"SharePoint: {mensaje}")

    # Sin las fotos arriba el reporte no está completo: queda en la cola
    # y se reintenta, que es justo lo que se quiere si Drive falló.
    completo = (drive_ok
                and (sheets_ok or not activos["sheets"])
                and (sp_ok or not activos["webhook"]))

    conn.execute("""UPDATE reportes
                    SET sincronizado = ?, sheets_ok = ?, sp_ok = ?,
                        sync_fecha = ?, sync_intentos = sync_intentos + 1, sync_error = ?
                    WHERE id = ?""",
                 (1 if completo else 0, 1 if sheets_ok else 0, 1 if sp_ok else 0,
                  datetime.now().strftime("%d-%m-%Y %H:%M"),
                  " · ".join(errores), reporte_id))
    conn.commit()
    conn.close()
    return completo, " · ".join(errores or avisos) or "Sin cambios."


def sincronizar_pendientes(limite=50):
    """Reintenta la cola de reportes no sincronizados."""
    conn = get_db()
    pendientes = [f["id"] for f in conn.execute(
        "SELECT id FROM reportes WHERE sincronizado = 0 ORDER BY rowid LIMIT ?",
        (limite,)).fetchall()]
    conn.close()

    enviados, fallidos, ultimo_error = 0, 0, ""
    for reporte_id in pendientes:
        ok, mensaje = sincronizar_reporte(reporte_id)
        if ok:
            enviados += 1
        else:
            fallidos += 1
            ultimo_error = mensaje
    return {"total": len(pendientes), "enviados": enviados,
            "fallidos": fallidos, "error": ultimo_error}


def contar_pendientes():
    conn = get_db()
    n = conn.execute("SELECT COUNT(*) c FROM reportes WHERE sincronizado = 0").fetchone()["c"]
    conn.close()
    return n


# ---------- Login contra la lista Usuarios App ----------
#
# Mismo principio que el envío: la app no consulta SharePoint, le
# pregunta a un flujo. Le manda {usuario, clave_hash} y el flujo
# responde con el perfil si las credenciales calzan.
#
# La contraseña nunca viaja: se envía su SHA-256, que es exactamente
# lo que está guardado en la columna ClaveHash. El flujo solo compara
# dos textos.
#
# Si el flujo no responde —sin red en terreno, típico— se cae al
# usuario local. Por eso cada login exitoso se cachea: el supervisor
# que entró una vez en la oficina puede entrar después en el cerro.


def validar_usuario_remoto(usuario, clave):
    """Pregunta al flujo. Devuelve (estado, datos):

        "ok"          credenciales válidas, `datos` trae el perfil
        "rechazado"   el flujo respondió que no son válidas
        "sin_flujo"   no hay webhook de login configurado
        "sin_red"     el flujo no respondió; hay que caer a lo local
    """
    import urllib.error
    import urllib.request

    config = cargar_config_sync()
    url = (config.get("url_login") or "").strip()
    if not url:
        return "sin_flujo", None

    cuerpo = json.dumps({"usuario": usuario, "clave_hash": hash_clave(clave)}).encode("utf-8")
    cabeceras = {"Content-Type": "application/json; charset=utf-8"}
    if config.get("token"):
        cabeceras["X-Token"] = config["token"]

    peticion = urllib.request.Request(url, data=cuerpo, headers=cabeceras, method="POST")
    try:
        with urllib.request.urlopen(peticion, timeout=config.get("tiempo_limite", 30)) as r:
            respuesta = json.loads(r.read().decode("utf-8"))
    except (urllib.error.URLError, TimeoutError, OSError):
        return "sin_red", None
    except ValueError:
        return "sin_red", None            # el flujo no devolvió JSON

    if not respuesta.get("ok"):
        return "rechazado", respuesta
    return "ok", respuesta


def guardar_usuario_local(datos, clave):
    """Cachea el usuario validado para poder entrar sin red la próxima vez."""
    # `obras` ausente ≠ `obras` vacío: si el flujo no manda la columna,
    # las asignaciones hechas en la app se conservan. Antes cualquier
    # login las borraba.
    obras = datos.get("obras")
    if isinstance(obras, str):
        obras = [o.strip() for o in obras.split(";") if o.strip()]

    conn = get_db()
    c = conn.cursor()
    c.execute("""INSERT INTO usuarios (usuario, clave, nombre, empresa, perfil, activo)
                 VALUES (?,?,?,?,?,1)
                 ON CONFLICT(usuario) DO UPDATE SET
                     clave=excluded.clave, nombre=excluded.nombre,
                     empresa=excluded.empresa, perfil=excluded.perfil, activo=1""",
              (datos["usuario"], hash_clave(clave), datos.get("nombre", ""),
               datos.get("empresa", ""), datos.get("perfil", "Supervisor")))

    uid = c.execute("SELECT id FROM usuarios WHERE usuario=?", (datos["usuario"],)).fetchone()[0]
    if obras is not None:
        c.execute("DELETE FROM permisos WHERE usuario_id=?", (uid,))
        for referencia in obras:
            # La lista puede traer el nombre de la obra raíz (como
            # siempre) o el id de cualquier nodo del árbol, que es lo
            # que permite asignar hasta el nivel de Step.
            nodo = c.execute(
                "SELECT id FROM nodos WHERE id=? UNION ALL "
                "SELECT id FROM nodos WHERE padre_id IS NULL AND nombre=? LIMIT 1",
                (referencia, referencia)).fetchone()
            if nodo:
                c.execute("INSERT OR IGNORE INTO permisos (usuario_id, nodo_id) VALUES (?,?)",
                          (uid, nodo["id"]))
    conn.commit()
    fila = c.execute("SELECT * FROM usuarios WHERE usuario=?", (datos["usuario"],)).fetchone()
    conn.close()
    return dict(fila)


def validar_usuario_local(usuario, clave):
    conn = get_db()
    fila = conn.execute(
        "SELECT * FROM usuarios WHERE usuario=? AND clave=? AND activo=1",
        (usuario, hash_clave(clave))).fetchone()
    conn.close()
    return dict(fila) if fila else None


def autenticar(usuario, clave):
    """Punto único de login. Devuelve (auth, origen, mensaje).

    `origen` sirve para que la pantalla diga de dónde salió la
    validación: "sharepoint" o "local".
    """
    estado, datos = validar_usuario_remoto(usuario, clave)

    if estado == "ok":
        datos.setdefault("usuario", usuario)
        return guardar_usuario_local(datos, clave), "sharepoint", ""

    if estado == "rechazado":
        # El flujo respondió y dijo que no. No se intenta lo local:
        # dar de baja a alguien en SharePoint debe surtir efecto.
        motivo = (datos or {}).get("motivo") or "Usuario o contraseña incorrectos."
        return None, "sharepoint", motivo

    auth = validar_usuario_local(usuario, clave)
    if auth:
        aviso = ("Sin conexión con SharePoint: se validó con la copia local."
                 if estado == "sin_red" else "")
        return auth, "local", aviso
    return None, "local", "Usuario o contraseña incorrectos."


# ============================================================
# 6. Estado de sesión y navegación
# ============================================================

def init_estado():
    ss = st.session_state
    ss.setdefault("auth", None)          # usuario autenticado (Nivel 0)
    ss.setdefault("supervisor", None)    # supervisor responsable (Nivel 1)
    ss.setdefault("ruta", [])            # nodos recorridos (Niveles 2..5)
    ss.setdefault("paso", "login")       # login | supervisor | arbol | formulario | guardado | reportes
    ss.setdefault("ultimo_reporte", None)
    ss.setdefault("origen_login", "local")   # "sharepoint" o "local"
    ss.setdefault("asig_ruta", [])       # navegación de la pantalla de asignaciones


def supervisor_directo(auth):
    """¿El usuario que entró ES el supervisor del reporte?

    Sí para el perfil Supervisor: se identifica con su usuario y su
    contraseña, así que preguntarle después quién está a cargo es pedir
    dos veces el mismo dato —y abre la puerta a que firme por otro—.
    Solo los perfiles que ven todo (Administración, Oficina Técnica,
    Inspector) siguen eligiendo supervisor, porque registran o revisan
    en nombre de terceros.
    """
    return not PERFILES.get(auth["perfil"], {}).get("ve_todo", False)


def ir(paso):
    st.session_state.paso = paso
    st.rerun()


def es_reportable(nodo):
    """¿Sobre este nodo se puede hacer un reporte?

    Solo los Step. Una obra a la que aún no se le cargó el programa
    también se queda sin hijos, y dejar que se reporte sobre ella
    produciría un reporte sin actividad ni ID contra el cual cruzarlo.
    """
    return (nodo.get("tipo") or "") == "step"


def bajar(nodo):
    """Entra un nivel. Si el nodo es un Step, abre el formulario."""
    st.session_state.ruta.append(nodo)
    st.session_state.paso = ("formulario"
                             if es_hoja(nodo["id"]) and es_reportable(nodo)
                             else "arbol")
    st.rerun()


def nuevo_id_fuera_programa():
    """Código que no coincide con ningún ID ACT ni COD STEP del programa.

    Prefijo FP + sufijo al azar, nunca un correlativo: así no hay forma
    de que choque con un código real del catálogo, ni ahora ni cuando
    Planificación publique una revisión nueva.
    """
    return "FP-" + uuid.uuid4().hex[:6].upper()


def bajar_fuera_programa(nodo):
    """Entra al reporte de una actividad no contemplada en el programa.

    No hay Actividad ni Step que elegir —por eso se salta directo al
    formulario—: se genera un ID ACT y un COD STEP propios, fuera de
    cualquier catálogo, y el supervisor describe ahí qué se ejecutó
    (ver pantalla_formulario). Cuelga del nodo donde se tocó el botón,
    así que la ubicación del reporte es la del punto del árbol en que
    estaba parado: si fue sobre el tipo de obra y no sobre la obra
    individual, el reporte sale sin COD OBRA.
    """
    ss = st.session_state
    id_act = nuevo_id_fuera_programa()
    actividad_fp = {
        "id": f"{nodo['id']}::fp::{id_act}", "nombre": "Fuera de programa",
        "tipo": "actividad", "codigo": id_act, "id_actividad": id_act,
        "padre_id": nodo["id"],
    }
    step_fp = {
        "id": f"{actividad_fp['id']}::step", "nombre": "Actividad fuera de programa",
        "tipo": "step", "codigo": nuevo_id_fuera_programa(),
        "unidad": "und", "padre_id": actividad_fp["id"],
        "fuera_de_programa": True,
    }
    step_fp["id_step"] = step_fp["codigo"]
    ss.ruta.append(actividad_fp)
    ss.ruta.append(step_fp)
    ss.paso = "formulario"
    st.rerun()


def opciones_de_nivel(ruta, auth):
    """Lo que se vería en pantalla estando parado en `ruta`."""
    return hijos_visibles(ruta[-1]["id"], auth) if ruta else raices_visibles(auth)


def es_nivel_de_paso(ruta, auth):
    """¿Ese nivel tiene un único camino que no es un Step?

    Un nivel así no es una decisión: la app lo cruza sola en vez de
    pedir un toque para el único botón que hay.

    La obra es la excepción: aunque tenga una sola Actividad cargada,
    ahí vive el apartado para reportar algo fuera del programa, así
    que ese nivel nunca se salta ni al bajar ni al volver.
    """
    if ruta and es_obra(ruta[-1]):
        return False
    opciones = opciones_de_nivel(ruta, auth)
    return len(opciones) == 1 and not es_hoja(opciones[0]["id"])


def hay_donde_volver(ruta, auth):
    """¿Queda arriba algún nivel donde el usuario pueda elegir algo?

    Con el salto automático, retroceder a un nivel de paso único es
    inútil: la app volvería a bajar por él. Si no hay ninguno con
    opciones reales, la flecha no se dibuja.
    """
    return any(not es_nivel_de_paso(ruta[:i], auth) for i in range(len(ruta)))


def subir():
    """Vuelve al nivel anterior donde hubo algo que elegir.

    No basta con retroceder uno: si ese nivel tiene un solo camino,
    pantalla_arbol volvería a bajar por él y la flecha parecería
    muerta. Por eso sube hasta encontrar un nivel con opciones.
    """
    ss = st.session_state
    auth = ss.auth
    if ss.ruta:
        ss.ruta.pop()
        while ss.ruta and es_nivel_de_paso(ss.ruta, auth):
            ss.ruta.pop()
        # Si al llegar a la raíz tampoco hay elección, no hay pantalla
        # de árbol que mostrar: se sale a la selección de supervisor.
        if not ss.ruta and es_nivel_de_paso([], auth) and not supervisor_directo(auth):
            ss.paso = "supervisor"
        else:
            ss.paso = "arbol"
    elif supervisor_directo(ss.auth):
        ss.paso = "arbol"
    else:
        ss.paso = "supervisor"
    st.rerun()


def saltar_a(nodo_id):
    """Se planta en un nodo cualquiera, sin recorrer el árbol.

    Reconstruye la ruta completa igual que si el usuario hubiera ido
    tocando nivel por nivel, así el reporte sale idéntico.
    """
    ss = st.session_state
    ss.ruta = ruta_de(nodo_id)
    ss.paso = "formulario" if es_hoja(nodo_id) else "arbol"
    st.rerun()


def cerrar_sesion():
    for k in ("auth", "supervisor", "ruta", "paso", "ultimo_reporte",
              "asig_ruta", "_permisos_cache"):
        st.session_state.pop(k, None)
    init_estado()
    st.rerun()


# ============================================================
# 7. Componentes de interfaz
# ============================================================

def app_header(titulo, subtitulo="", con_volver=True):
    lado = "header-right" if con_volver else ""
    html = (f'<div class="app-header {lado}">'
            f'<div><div class="title">{titulo}</div>'
            f'<div class="sub">{subtitulo}</div></div>'
            f'<div class="brand">SACYR</div></div>')
    if con_volver:
        # Sin `help`: el tooltip queda pegado en pantallas táctiles y
        # tapa el propio botón. El significado de la flecha es evidente.
        with st.container(key="header_row"):
            c1, c2 = st.columns([1, 5], vertical_alignment="center")
            with c1:
                if st.button("←", key="btn_back", width="stretch"):
                    subir()
            with c2:
                st.markdown(html, unsafe_allow_html=True)
    else:
        st.markdown(html, unsafe_allow_html=True)


def breadcrumb():
    """Muestra la ruta recorrida: el usuario siempre sabe dónde está."""
    ss = st.session_state
    partes = []
    if ss.supervisor:
        partes.append(ss.supervisor)
    partes += [n["nombre"] for n in ss.ruta]
    if not partes:
        return
    texto = " › ".join(partes[:-1])
    ultimo = f"<b>{partes[-1]}</b>"
    st.markdown(
        f'<div class="breadcrumb">📍 {texto + " › " if texto else ""}{ultimo}</div>',
        unsafe_allow_html=True)


def bottom_nav():
    # La carga de la planilla de obras es tarea de Oficina Técnica /
    # Administración, no del supervisor en terreno.
    admin = PERFILES.get(st.session_state.auth["perfil"], {}).get("ve_todo", False)
    # En modo local todos los reportes están "pendientes" por definición:
    # el contador solo tiene sentido cuando hay un webhook configurado.
    pendientes = (contar_pendientes()
                  if admin and cargar_config_sync().get("modo") == "webhook" else 0)
    with st.container(key="bottom_nav"):
        columnas = st.columns(6 if admin else 3)
        with columnas[0]:
            if st.button("🏗️ Reportar", key="nav_arbol"):
                st.session_state.ruta = []
                ir("arbol")
        with columnas[1]:
            if st.button("🗂️ Reportes", key="nav_reportes"):
                ir("reportes")
        if admin:
            with columnas[2]:
                if st.button("⚙️ Obras", key="nav_obras"):
                    ir("obras")
            with columnas[3]:
                if st.button("👥 Asignar", key="nav_asignar"):
                    st.session_state.asig_ruta = []
                    ir("asignar")
            with columnas[4]:
                aviso = f" ({pendientes})" if pendientes else ""
                if st.button(f"☁️ Envío{aviso}", key="nav_envio"):
                    ir("envio")
        with columnas[-1]:
            if st.button("🚪 Salir", key="btn_salir"):
                cerrar_sesion()


def buscador_por_id(auth):
    """Atajo por nombre, disponible en cualquier nivel y para todo perfil.

    Va plegado para no competir con los botones grandes del nivel: en
    terreno lo normal es navegar, y buscar es el atajo de quien ya sabe
    qué actividad viene a reportar.
    """
    with st.expander("🔎 Buscar actividad"):
        texto = st.text_input(
            "Nombre de la actividad", key="busq_id",
            placeholder="Ej. terraplen tramo 2, excavacion melipilla",
            label_visibility="collapsed").strip()

        if len(texto) < 3:
            st.markdown(
                '<p class="hint">Escriba parte del nombre: <b>no importan las tildes, '
                'las mayúsculas ni un dedazo</b>. Puede poner varias palabras en '
                'cualquier orden —«terraplen tramo 2»— y también sirve un código si '
                'lo tiene a mano.</p>', unsafe_allow_html=True)
            return

        resultados = buscar_actividad(texto, auth)
        if not resultados:
            st.markdown(
                '<p class="hint">Nada coincide entre sus actividades. Pruebe con '
                'menos palabras, o con una sola.</p>', unsafe_allow_html=True)
            return

        st.markdown(f'<p class="hint">{len(resultados)} coincidencia(s):</p>',
                    unsafe_allow_html=True)
        for r in resultados:
            icono = "📄" if r["tipo"] == "step" else "📁"
            # El texto buscado no se limpia: Streamlit prohíbe tocar el
            # session_state de un widget ya instanciado, y además el
            # expander vuelve plegado al saltar, así que no estorba.
            if st.button(f"{icono}  {r['nombre']}", key=clave_widget("busq", r["id"]),
                         width="stretch"):
                saltar_a(r["id"])
            st.markdown(
                f'<p class="busq-ruta">{r["ruta"]}<br><code>{r["ids"]}</code></p>',
                unsafe_allow_html=True)


def ruta_texto(ruta):
    return " › ".join(n["nombre"] for n in ruta)


# ============================================================
# 7b. Jornada, cámara y sello de la evidencia
# ============================================================

def horas_por_fecha(fecha):
    """Horas de la jornada según el día de la semana.

    Es el dato definitivo, no una propuesta: el formulario ya no tiene
    apartado de jornada. Lunes y martes 9 h · miércoles a viernes 8 h ·
    sábado y domingo 6 h.
    """
    return HORAS_POR_DIA.get(fecha.weekday(), 8.0)


def etiqueta_dia(fecha):
    return DIAS_SEMANA[fecha.weekday()]


def fmt_horas(horas):
    """9.0 -> «9,0». La obra lee coma decimal, no punto."""
    return f"{horas:.1f}".replace(".", ",")


def ultimo_sello(usuario):
    """Últimos datos de ubicación que usó este usuario.

    El supervisor trabaja días seguidos en el mismo camino y comuna:
    volver a escribirlos en cada reporte es puro roce. Se rescatan del
    último reporte suyo que los traiga."""
    base = dict(SELLO_POR_DEFECTO)
    try:
        conn = get_db()
        fila = conn.execute(
            """SELECT camino, comuna, region, pais, altitud, pk_desde, pk_hasta
               FROM reportes
               WHERE usuario = ? AND IFNULL(camino,'') <> ''
               ORDER BY rowid DESC LIMIT 1""", (usuario,)).fetchone()
        conn.close()
    except sqlite3.Error:
        fila = None
    if fila:
        base.update({
            "camino": fila["camino"] or base["camino"],
            "comuna": fila["comuna"] or base["comuna"],
            "region": fila["region"] or base["region"],
            "pais": fila["pais"] or base["pais"],
            "altitud": fila["altitud"] if fila["altitud"] is not None else base["altitud"],
        })
    return base


def selector_camara():
    """Permite elegir cámara trasera o frontal.

    `st.camera_input` no expone la cámara a usar: siempre pide la que
    el navegador considere por defecto —en el celular, casi siempre la
    frontal, que es inútil para documentar un frente de obra—. Aquí se
    parcha `getUserMedia` en la página para que toda petición de video
    lleve el `facingMode` elegido. El parche queda instalado en la
    ventana, así que sobrevive a los rerun de Streamlit.
    """
    opciones = {"📷 Trasera": "environment", "🤳 Frontal": "user"}
    elegida = st.radio("Cámara", list(opciones), horizontal=True, key="f_camara",
                       help="En el computador, si hay una sola cámara no cambia nada.")
    facing = opciones[elegida]

    # `st.iframe` con HTML inline usa srcdoc: el marco hereda el origen
    # de la página y puede tocar `window.parent`. Sin eso, el parche no
    # llegaría al widget de cámara, que vive en la página de arriba.
    st.iframe(
        """
        <script>
        (function () {
            var win = window.parent;
            var md = win.navigator.mediaDevices;
            if (!md || !md.getUserMedia) { return; }
            if (!md.__sacyrParche) {
                md.__sacyrParche = true;
                md.__sacyrOriginal = md.getUserMedia.bind(md);
                md.getUserMedia = function (restricciones) {
                    var r = Object.assign({}, restricciones || {});
                    if (r.video) {
                        var v = (typeof r.video === 'object') ? Object.assign({}, r.video) : {};
                        v.facingMode = { ideal: win.__sacyrCamara || 'environment' };
                        r.video = v;
                    }
                    return md.__sacyrOriginal(r);
                };
            }
            win.__sacyrCamara = 'FACING';
        })();
        </script>
        """.replace("FACING", facing), height="content")

    # Quien llama usa este valor en la clave de `st.camera_input`: al
    # cambiar de cámara el widget se remonta, y solo así vuelve a pedir
    # el stream tomando el facingMode nuevo.
    return facing


# ============================================================
#  SELLO DE LAS FOTOS — tamaños y calidad
#  ------------------------------------------------------------
#  ESTE ES EL BLOQUE A EDITAR si el sello ocupa mucho o poco.
#  No hay tamaños sueltos en el resto del código: todo sale de aquí.
#
#  La letra se calcula como  ancho_de_la_foto / SELLO_DIVISOR, para que
#  el sello se vea igual en una foto de 1 Mpx que en una de 12.
#
#     ¿Letras más chicas?  ->  SUBIR  SELLO_DIVISOR  (55 -> 65)
#     ¿Letras más grandes? ->  BAJAR  SELLO_DIVISOR  (55 -> 45)
#     ¿Líneas más juntas?  ->  BAJAR  SELLO_INTERLINEA
# ============================================================

SELLO_DIVISOR = 55           # letra base = ancho / este número
SELLO_TAM_MINIMO = 13        # px: piso para que no desaparezca en fotos chicas
SELLO_TITULO = 1.15          # el título, en veces la letra base
SELLO_INTERLINEA = 1.28      # separación entre líneas, en veces la letra
SELLO_MARGEN = 0.85          # margen al borde, en veces la letra
SELLO_CONTORNO = 0.11        # grosor del contorno negro, en veces la letra
SELLO_TITULO_MAX_LINEAS = 2  # el título largo se parte hasta en estas líneas

# Calidad de la fotografía guardada. Subirlas mejora la evidencia pero
# engorda el JSON del webhook, porque las fotos viajan en base64.
FOTO_LADO_MAX = 2560         # px del lado mayor
FOTO_CALIDAD = 95            # calidad JPEG, 1 a 100


def _fuente(tam):
    """Fuente del sello. Windows en la oficina, DejaVu en el servidor."""
    for ruta in ("C:/Windows/Fonts/arialbd.ttf",
                 "C:/Windows/Fonts/arial.ttf",
                 "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
                 "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"):
        try:
            return ImageFont.truetype(ruta, tam)
        except OSError:
            continue
    try:
        return ImageFont.load_default(size=tam)
    except TypeError:                        # Pillow < 10
        return ImageFont.load_default()


def lineas_sello(datos):
    """Texto que se imprime sobre la foto, en el orden en que se lee."""
    lineas = [datos["fecha_hora"]]

    posicion = []
    if datos.get("coords"):
        posicion.append(datos["coords"])
    if datos.get("altitud"):
        posicion.append(f"{datos['altitud']:.0f} m.s.n.m.")
    if datos.get("orientacion") and datos["orientacion"] != "—":
        posicion.append(f"Vista {datos['orientacion']}")
    if posicion:
        lineas.append(" · ".join(posicion))

    lugar = [t for t in (datos.get("camino"), datos.get("comuna")) if t]
    if lugar:
        lineas.append(" · ".join(lugar))
    territorio = [t for t in (datos.get("region"), datos.get("pais")) if t]
    if territorio:
        lineas.append(" · ".join(territorio))

    if datos.get("pk"):
        lineas.append(datos["pk"])
    if datos.get("pie"):
        lineas.append(datos["pie"])
    return lineas


def _ancho_texto(texto, fuente):
    """Ancho en píxeles de un texto. `getlength` no está en Pillow viejo."""
    try:
        return fuente.getlength(texto)
    except AttributeError:
        return fuente.getsize(texto)[0]


def _recortar(texto, fuente, ancho_max):
    """Corta el texto con puntos suspensivos hasta que quepa."""
    if _ancho_texto(texto, fuente) <= ancho_max:
        return texto
    while texto and _ancho_texto(texto + "…", fuente) > ancho_max:
        texto = texto[:-1]
    return (texto.rstrip() + "…") if texto else ""


def _envolver(texto, fuente, ancho_max, max_lineas):
    """Parte el texto por palabras en como mucho `max_lineas` renglones.

    El título del Step suele ser largo ("PMP TAL - Instalación de faenas,
    movilización de equipos, cierres perimetrales…") y antes se salía de
    la foto por el lado derecho.
    """
    palabras = texto.split()
    if not palabras:
        return []

    lineas, actual = [], ""
    for palabra in palabras:
        prueba = f"{actual} {palabra}".strip()
        if actual and _ancho_texto(prueba, fuente) > ancho_max:
            lineas.append(actual)
            actual = palabra
            if len(lineas) == max_lineas:
                break
        else:
            actual = prueba
    else:
        lineas.append(actual)

    if len(lineas) > max_lineas:
        lineas = lineas[:max_lineas]
    # Lo que no cupo se resume en la última línea con puntos suspensivos.
    sobrante = texto[len(" ".join(lineas)):].strip()
    if sobrante:
        lineas[-1] = _recortar(lineas[-1] + " " + sobrante, fuente, ancho_max)
    return [_recortar(l, fuente, ancho_max) for l in lineas]


def estampar_foto(contenido, titulo, lineas):
    """Devuelve la foto en JPEG con los datos escritos abajo.

    Solo letras: nada de panel de color. El recuadro azul tapaba un
    tercio de la evidencia —justo la base del frente, que es lo que hay
    que mirar—. La legibilidad la da un contorno oscuro sobre cada
    letra, que funciona igual sobre cielo claro o sobre sombra.
    """
    if not HAY_PILLOW:
        return contenido
    try:
        imagen = Image.open(io.BytesIO(contenido))
        imagen = ImageOps.exif_transpose(imagen).convert("RGB")
    except Exception:                        # archivo corrupto o no imagen
        return contenido

    # Un celular actual entrega 12 Mpx: pesa demasiado para viajar en
    # base64 dentro del JSON del webhook. Se reduce, pero con remuestreo
    # LANCZOS y sin recomprimir barato: la foto anterior salía blanda
    # por el lado corto y por la calidad JPEG.
    if max(imagen.size) > FOTO_LADO_MAX:
        imagen.thumbnail((FOTO_LADO_MAX, FOTO_LADO_MAX), Image.LANCZOS)
    ancho, alto = imagen.size

    base = max(SELLO_TAM_MINIMO, int(ancho / SELLO_DIVISOR))
    fuente = _fuente(base)
    fuente_titulo = _fuente(int(base * SELLO_TITULO))
    margen = int(base * SELLO_MARGEN)
    salto = int(base * SELLO_INTERLINEA)
    grosor = max(2, int(base * SELLO_CONTORNO))   # contorno que reemplaza al panel

    lienzo = ImageDraw.Draw(imagen)

    # Nada puede sobrepasar el margen derecho: el título se parte en
    # varias líneas y el resto se recorta.
    util = ancho - margen * 2
    filas = [(t, True) for t in _envolver(titulo, fuente_titulo, util,
                                          SELLO_TITULO_MAX_LINEAS)] if titulo else []
    filas += [(_recortar(str(t), fuente, util), False) for t in lineas if str(t).strip()]

    x = margen
    y = alto - margen - salto * len(filas)
    for texto, es_titulo in filas:
        lienzo.text(
            (x, y), texto,
            font=fuente_titulo if es_titulo else fuente,
            fill=(255, 255, 255) if es_titulo else (238, 243, 249),
            stroke_width=grosor, stroke_fill=(0, 0, 0))
        y += salto

    salida = io.BytesIO()
    imagen.save(salida, format="JPEG", quality=FOTO_CALIDAD,
                subsampling=0, optimize=True)
    return salida.getvalue()


# ---------- Excel del reporte ----------
#
# Lo que se entrega al emitir el reporte: una planilla con los datos que
# se imputan al programa y la evidencia pegada al lado. Es el formato en
# que la obra revisa y firma, y el que se archiva.

# Los tres niveles que identifican lo reportado, de lo general a lo
# concreto: TIPO_OBRA (ESTACIONES, SER, PASOS VEHICULARES DESNIVELADOS…),
# OBRA (Estación El Monte) y ACTIVIDAD (lo realmente ejecutado, que es
# el nombre del Step al que corresponden COD ACTIVIDAD y COD STEPS).
COLUMNAS_EXCEL_BASE = ["COD ACTIVIDAD", "COD STEPS", "COD OBRA",
                       "TRAMO", "SECTOR", "TIPO_OBRA", "OBRA", "ACTIVIDAD",
                       "FECHA", "HH", "HM", "CANT", "UND", "PERSONAL",
                       "SUBCONTRATO", "SUPERVISOR"]
ANCHOS_EXCEL_BASE = [16, 12, 14, 10, 26, 24, 28, 46, 12, 8, 8, 10, 8, 10, 16, 22]

COLUMNAS_EXCEL_FIN = ["OBS", "FOTO"]
ANCHOS_EXCEL_FIN = [44, 34]

# Cada máquina va en su propia columna. Se emiten al menos estas tres
# para que dos reportes se puedan apilar sin que se corran las columnas;
# si el reporte trae más máquinas, se agregan las que falten.
MAQUINAS_EXCEL = 3
ANCHO_MAQUINA = 24
ANCHO_FOTO_PX = 230


def _maquinas_de(reporte):
    """Las máquinas del reporte, una por elemento (PATENTE · TIPO)."""
    return [linea.strip() for linea in (reporte["equipos"] or "").splitlines()
            if linea.strip()]


def _columnas_excel(n_maquinas):
    """Encabezados y anchos para un reporte con `n_maquinas` máquinas."""
    n = max(n_maquinas, MAQUINAS_EXCEL)
    columnas = (COLUMNAS_EXCEL_BASE
                + [f"MAQUINA USADA {i}" for i in range(1, n + 1)]
                + COLUMNAS_EXCEL_FIN)
    anchos = ANCHOS_EXCEL_BASE + [ANCHO_MAQUINA] * n + ANCHOS_EXCEL_FIN
    return columnas, anchos


def _fila_excel(reporte):
    """Los datos del reporte en el orden de _columnas_excel()."""
    maquinas = _maquinas_de(reporte)
    n = max(len(maquinas), MAQUINAS_EXCEL)
    celdas_maquinas = [maquinas[i] if i < len(maquinas) else "" for i in range(n)]
    if not maquinas:
        celdas_maquinas[0] = "—"

    # Los reportes anteriores a estas columnas no las traen: se deducen
    # de la ruta guardada para que la planilla no salga con huecos.
    ubicacion = {"cod_obra": reporte["cod_obra"] if "cod_obra" in reporte.keys() else "",
                 "tramo": reporte["tramo"] if "tramo" in reporte.keys() else "",
                 "sector": reporte["sector"] if "sector" in reporte.keys() else ""}
    if not ubicacion["cod_obra"] and reporte["nodo_id"]:
        ubicacion = ubicacion_de_ruta(ruta_de(reporte["nodo_id"]))

    return [
        reporte["id_actividad"] or "—",
        reporte["id_step"] or "—",
        ubicacion["cod_obra"] or "—",
        ubicacion["tramo"] or "—",
        ubicacion["sector"] or "—",
        reporte["tipo_obra"] or "—",
        reporte["nombre_obra"] or "—",
        reporte["step_nombre"] or "—",     # lo ejecutado: el Step reportado
        reporte["fecha"] or "",
        reporte["hh_calculada"],
        reporte["hm_calculada"],
        reporte["cantidad"],
        reporte["unidad"] or "—",
        reporte["personal"],
        reporte["subcontrato"] or "—",
        reporte["supervisor"] or reporte["usuario"] or "",
        *celdas_maquinas,
        reporte["observaciones"] or "",
    ]


def excel_reporte(reporte):
    """Devuelve el .xlsx del reporte, con las fotos incrustadas.

    Una fila por fotografía, repitiendo los datos: así cada evidencia
    queda en su propia celda y la planilla se puede filtrar y pegar
    debajo de otra sin desarmar nada. Sin fotos, una sola fila.
    """
    import openpyxl
    from openpyxl.drawing.image import Image as ImagenExcel
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.title = "Reporte"

    columnas, anchos = _columnas_excel(len(_maquinas_de(reporte)))
    col_obs = len(columnas) - 1          # penúltima
    col_foto = len(columnas)             # última

    azul = PatternFill("solid", fgColor="10263F")
    for i, (titulo, ancho) in enumerate(zip(columnas, anchos), start=1):
        celda = hoja.cell(row=1, column=i, value=titulo)
        celda.font = Font(bold=True, color="FFFFFF", size=10)
        celda.fill = azul
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        hoja.column_dimensions[get_column_letter(i)].width = ancho
    hoja.row_dimensions[1].height = 28
    hoja.freeze_panes = "A2"

    datos = _fila_excel(reporte)
    rutas = json.loads(reporte["fotos"] or "[]")
    existentes = [os.path.join(BASE_DIR, r) for r in rutas
                  if os.path.exists(os.path.join(BASE_DIR, r))]

    # Las imágenes se sostienen desde aquí: openpyxl las lee recién al
    # guardar, y un BytesIO recolectado antes deja la celda vacía.
    buffers = []

    for indice, fila_n in enumerate(range(2, max(len(existentes), 1) + 2)):
        for columna, valor in enumerate(datos, start=1):
            celda = hoja.cell(row=fila_n, column=columna, value=valor)
            celda.alignment = Alignment(vertical="top", wrap_text=columna == col_obs)
        hoja.row_dimensions[fila_n].height = 22

        if indice >= len(existentes):
            hoja.cell(row=fila_n, column=col_foto, value="Sin fotografía")
            continue

        ruta = existentes[indice]
        try:
            with open(ruta, "rb") as f:
                contenido = f.read()
            buffer = io.BytesIO(contenido)
            imagen = ImagenExcel(buffer)
            # Se escala al ancho de la columna conservando la proporción.
            proporcion = ANCHO_FOTO_PX / float(imagen.width or ANCHO_FOTO_PX)
            imagen.width = ANCHO_FOTO_PX
            imagen.height = int((imagen.height or ANCHO_FOTO_PX) * proporcion)
            hoja.add_image(imagen, f"{get_column_letter(col_foto)}{fila_n}")
            # Alto de fila en puntos (1 px ≈ 0,75 pt) para que la foto quepa.
            hoja.row_dimensions[fila_n].height = imagen.height * 0.75 + 6
            buffers.append(buffer)
        except Exception:                    # imagen corrupta o formato raro
            hoja.cell(row=fila_n, column=col_foto, value=os.path.basename(ruta))

    # Segunda hoja con el detalle completo: el Excel principal se queda
    # en lo pedido, pero la trazabilidad no se pierde.
    detalle = libro.create_sheet("Detalle")
    detalle.column_dimensions["A"].width = 24
    detalle.column_dimensions["B"].width = 70
    for fila_n, (etiqueta, valor) in enumerate([
            ("Folio", reporte["id"]),
            ("Emitido", reporte["creado_en"]),
            ("Ruta WBS", reporte["ruta_texto"]),
            ("Actividad", reporte["actividad"] or "—"),
            ("Step", reporte["step_nombre"]),
            ("Estado", reporte["estado"]),
            ("Jornada (h)", reporte["horas"]),
            ("Personal", reporte["personal"]),
            ("Cantidad ejecutada", f"{reporte['cantidad']} {reporte['unidad'] or ''}".strip()),
            ("HH línea base", reporte["hh_lb"]),
            ("Cantidad línea base", reporte["cantidad_lb"]),
            ("Materiales", reporte["materiales"] or "—"),
            ("Incidencias", reporte["incidencias"] or "Sin incidencias"),
            ("Riesgos", reporte["riesgos"] or "—"),
            ("Tramo", f"PK{reporte['pk_desde']} a PK{reporte['pk_hasta']}"
                      if reporte["pk_desde"] or reporte["pk_hasta"] else "—"),
            ("Ubicación", " · ".join(t for t in (reporte["camino"], reporte["comuna"],
                                                 reporte["region"]) if t) or "—"),
            ("Coordenadas", f"{reporte['latitud']}, {reporte['longitud']}"
                            if reporte["latitud"] else "—"),
            ("Firmado por", reporte["firma_nombre"]),
    ], start=1):
        detalle.cell(row=fila_n, column=1, value=etiqueta).font = Font(bold=True, size=10)
        detalle.cell(row=fila_n, column=2, value=valor).alignment = Alignment(wrap_text=True)

    salida = io.BytesIO()
    libro.save(salida)
    return salida.getvalue()


def boton_excel(reporte, clave):
    """Botón de descarga del Excel. Silencioso si algo falla: nunca debe
    impedir que se vea la confirmación de un reporte ya guardado."""
    try:
        contenido = excel_reporte(reporte)
    except Exception as error:
        st.caption(f"No se pudo generar el Excel: {error}")
        return
    st.download_button(
        "📊 Descargar Excel del reporte", data=contenido,
        file_name=f"Reporte_{reporte['id']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=clave, width="stretch")


def reporte_por_id(reporte_id):
    conn = get_db()
    fila = conn.execute("SELECT * FROM reportes WHERE id=?", (reporte_id,)).fetchone()
    conn.close()
    return fila


def limpiar_formulario():
    """Borra los campos del reporte para que el siguiente parta limpio."""
    for clave in [k for k in st.session_state if k.startswith("f_")]:
        st.session_state.pop(clave, None)


# ============================================================
# 8. Pantallas
# ============================================================

def pantalla_login():
    """Nivel 0 — Identifica al usuario y, con él, empresa y perfil."""
    app_header("Reporte Diario", "Inicio de sesión", con_volver=False)
    st.markdown('<div class="pregunta">Ingrese sus credenciales</div>', unsafe_allow_html=True)

    with st.form("form_login"):
        usuario = st.text_input("Usuario", placeholder="Ej. jaime").strip().lower()
        clave = st.text_input("Contraseña", type="password")
        entrar = st.form_submit_button("Entrar", width="stretch")

    if entrar:
        auth, origen, mensaje = autenticar(usuario, clave)
        if auth:
            st.session_state.auth = auth
            st.session_state.origen_login = origen
            limpiar_cache_permisos()
            # El supervisor entra directo a SUS actividades: es él quien
            # se autenticó, no hay a quién elegir.
            directo = supervisor_directo(auth)
            st.session_state.supervisor = auth["nombre"] if directo else None
            st.session_state.ruta = []
            if mensaje:
                st.warning(mensaje)
            ir("arbol" if directo else "supervisor")
        else:
            st.error(mensaje or "Usuario o contraseña incorrectos.")

    # Bloque de ayuda para pruebas. Desaparece solo en cuanto hay
    # usuarios reales en Secrets: publicar en internet una lista de
    # cuentas con la contraseña al lado sería entregar la app.
    if secreto("usuarios"):
        return

    with st.expander("Usuarios de la marcha blanca"):
        st.markdown(
            "| Usuario | Contraseña | Perfil | Obras |\n"
            "|---|---|---|---|\n"
            "| `jaime` | `mamboscuro` | Supervisor | Lineales + Puntuales |\n"
            "| `admin` | `1234` | Administrador | Todas |")


def pantalla_supervisor():
    """Supervisor responsable — solo para perfiles que registran o
    revisan en nombre de otros. El perfil Supervisor no pasa por aquí:
    entra directo a sus actividades."""
    auth = st.session_state.auth
    app_header("Supervisor responsable", f"{auth['empresa']} · {auth['perfil']}", con_volver=False)

    st.markdown('<div class="pregunta">¿A nombre de qué supervisor registra?</div>',
                unsafe_allow_html=True)
    st.markdown('<p class="hint">La información quedará asociada a este supervisor.</p>',
                unsafe_allow_html=True)

    # Los supervisores con usuario propio son la fuente principal: son
    # los que además tienen actividades asignadas.
    conn = get_db()
    filas = list(conn.execute(
        "SELECT nombre FROM usuarios WHERE activo=1 AND perfil='Supervisor' AND empresa=? "
        "ORDER BY nombre", (auth["empresa"],)).fetchall())
    filas += list(conn.execute(
        "SELECT nombre FROM supervisores WHERE activo=1 AND empresa=? ORDER BY nombre",
        (auth["empresa"],)).fetchall())
    conn.close()

    # Se deduplica por el slug, no por el nombre: «Jaime» y «jaime» son
    # la misma persona escrita distinto, y con la clave del botón salida
    # del slug, dejar las dos hacía reventar la pantalla entera.
    nombres, vistos = [], set()
    for f in filas:
        nombre = (f["nombre"] or "").strip()
        clave = _slug(nombre)
        if not nombre or clave in vistos:
            continue
        vistos.add(clave)
        nombres.append(nombre)

    if st.button(f"👤  {auth['nombre']}  ·  yo mismo", key="sup_yo", width="stretch"):
        st.session_state.supervisor = auth["nombre"]
        st.session_state.ruta = []
        ir("arbol")

    if not nombres:
        st.warning(f"No hay supervisores registrados para {auth['empresa']}.")

    for i, nombre in enumerate(nombres):
        if _slug(nombre) == _slug(auth["nombre"]):
            continue                     # ya está arriba, como «yo mismo»
        if st.button(f"○  {nombre}", key=f"sup_{i}_{_slug(nombre)}", width="stretch"):
            st.session_state.supervisor = nombre
            st.session_state.ruta = []
            ir("arbol")

    st.divider()
    if st.button("🚪 Cerrar sesión", key="btn_salir", width="stretch"):
        cerrar_sesion()


def pantalla_arbol():
    """Niveles 2 a 5 — Un nivel por pantalla, filtrado por lo anterior."""
    ss = st.session_state
    auth = ss.auth
    ruta = ss.ruta

    if not ruta:
        # Nivel 1: obra raíz. Solo las que el usuario tiene asignadas.
        opciones = raices_visibles(auth)
        pregunta = "Seleccione la obra"
        titulo, subtitulo = "Mis actividades", f"{auth['empresa']} · {ss.supervisor or auth['nombre']}"
    else:
        actual = ruta[-1]
        opciones = hijos_visibles(actual["id"], auth)
        pregunta = actual.get("pregunta") or "Seleccione una opción"
        titulo = actual["nombre"]
        subtitulo = ruta_texto(ruta[:-1]) or auth["empresa"]

    # Nivel de paso: un único camino que no es un Step. No se pinta una
    # pantalla para pedir el único botón que hay; se cruza sola. Es lo
    # que hace que un supervisor de un solo frente entre directo a sus
    # actividades.
    if (len(opciones) == 1 and not es_hoja(opciones[0]["id"])
            and not (ruta and es_obra(ruta[-1]))):
        bajar(opciones[0])
        return

    # La flecha solo se dibuja si arriba queda algo que elegir, o si
    # existe la pantalla de supervisor a la cual volver.
    app_header(titulo, subtitulo,
               con_volver=hay_donde_volver(ruta, auth) or not supervisor_directo(auth))
    breadcrumb()
    buscador_por_id(auth)

    # Va arriba, junto al buscador: si queda al final de la lista hay que
    # bajar por seis partidas —o trece Steps— para descubrir que existe.
    # Se ofrece en todo nivel dentro de una obra raíz, no solo en la obra
    # individual: lo que no está en el programa aparece a cualquier altura.
    if ruta:
        apartado_fuera_de_programa(ruta[-1])

    if not opciones:
        if permisos_de(auth) is not None and not ruta:
            st.warning("No tiene actividades asignadas. Pida a Oficina Técnica que le "
                       "asigne obras o actividades desde 👥 Asignar.")
        else:
            st.info("Este nivel no tiene actividades cargadas.")
        bottom_nav()
        return

    # Nada de «NIVEL 7»: ese número es la profundidad del árbol, un dato
    # interno que a quien reporta no le dice nada y le hace creer que
    # existe una jerarquía numerada. Dónde está lo cuenta la miga de pan,
    # y qué tiene que hacer, la pregunta.
    st.markdown(f'<div class="pregunta">{pregunta}</div>', unsafe_allow_html=True)

    # Nunca se muestran cientos de registros: si el nivel es ancho,
    # se filtra antes de pintar.
    if len(opciones) > MAX_OPCIONES_SIN_BUSCADOR:
        filtro = st.text_input("🔍 Buscar en este nivel", key=f"buscar_{len(ruta)}",
                               placeholder="Escriba para filtrar...").strip().lower()
        if filtro:
            opciones = [o for o in opciones
                        if filtro in o["nombre"].lower() or filtro in (o["codigo"] or "").lower()]
        if len(opciones) > MAX_OPCIONES_SIN_BUSCADOR:
            st.markdown(f'<p class="hint">{len(opciones)} opciones · use el buscador para acotar.</p>',
                        unsafe_allow_html=True)
            opciones = opciones[:MAX_OPCIONES_SIN_BUSCADOR]

    conteos = contar_hijos([o["id"] for o in opciones], permisos_de(auth))

    for o in opciones:
        n_hijos = conteos.get(o["id"], 0)
        if n_hijos == 0 and not es_reportable(o):
            # Obra sin el programa cargado. Se muestra para que se sepa
            # que existe, pero no se puede reportar contra ella: no hay
            # actividad ni ID que poner en el reporte.
            codigo = f"[{o['codigo']}]  " if o["codigo"] else ""
            st.button(f"📂  {codigo}{o['nombre']}   (sin actividades)",
                      key=clave_widget("vacio", o["id"]), width="stretch",
                      disabled=True,
                      help="Oficina Técnica todavía no cargó las actividades "
                           "del programa para esta obra.")
            continue
        if n_hijos == 0:
            # Es un Step: al tocarlo se abre el formulario del reporte.
            unidad = f"  ·  {o['unidad']}" if o["unidad"] else ""
            etiqueta = f"📄  {o['nombre']}{unidad}"
            clave = clave_widget("step", o["id"])
        else:
            codigo = f"[{o['codigo']}]  " if o["codigo"] else ""
            etiqueta = f"📁  {codigo}{o['nombre']}   ({n_hijos})"
            clave = clave_widget("nodo", o["id"])
        if st.button(etiqueta, key=clave, width="stretch"):
            bajar(o)

    bottom_nav()


def apartado_fuera_de_programa(nodo):
    """El botón para reportar algo que no está en el programa.

    Va arriba de la lista, con su propio divisor: no es una opción más
    del catálogo —no tiene ID ni Step previstos, se generan al tocarlo
    (ver bajar_fuera_programa())— y al final de la lista no se
    encuentra.
    """
    if st.button("🆓  Reportar actividad fuera del programa",
                 key=clave_widget("fp", nodo["id"]), width="stretch",
                 help="Para lo ejecutado que no figura en el programa de la obra. "
                      "Se genera un ID de actividad y un Step propios, fuera del "
                      "catálogo."):
        bajar_fuera_programa(nodo)
    st.divider()


def pantalla_formulario():
    """Formulario del Reporte Diario — solo datos de ejecución.
    Toda la jerarquía ya viene resuelta desde la navegación.

    La jornada no se pregunta: el reporte es del día en que se hace y
    las horas salen del día de la semana (ver HORAS_POR_DIA).

    No se usa `st.form` a propósito: el formulario tiene que reaccionar
    mientras se llena —la dotación cambia la HH calculada, la incidencia
    abre su detalle, la cámara elegida remonta el visor— y dentro de un
    form nada de eso ocurre hasta apretar el botón.
    """
    ss = st.session_state
    auth = ss.auth
    ruta = ss.ruta
    step = ruta[-1]

    # Al cambiar de Step (o al volver tras guardar) el formulario parte
    # limpio: arrastrar datos de otra actividad es cómo se contaminan
    # los reportes.
    if ss.get("f_step_id") != step["id"]:
        limpiar_formulario()
        ss["f_step_id"] = step["id"]

    # Actividad fuera de programa: el ID ACT y el COD STEP ya se
    # generaron al tocar el botón (bajar_fuera_programa()); no vienen
    # de ningún catálogo, así que el nombre a mostrar se pide más abajo.
    fuera_de_programa = bool(step.get("fuera_de_programa"))

    app_header("Reporte Diario",
               "Actividad fuera de programa" if fuera_de_programa else step["nombre"])

    # Campos que la planilla y SharePoint quieren como columnas propias:
    # la rama raíz, y dónde se ejecutó según los códigos de la ruta.
    obra = ruta[0]["nombre"] if ruta else ""
    ubicacion = ubicacion_de_ruta(ruta)
    tipo_obra, nombre_obra = ubicacion["tipo_obra"], ubicacion["nombre_obra"]
    especialidad = next((n["nombre"] for n in ruta if n["tipo"] == "especialidad"), "")

    # ID ACT y COD STEP de la BD de actividades: se deducen de la rama
    # recorrida y viajan con el reporte. No se piden en pantalla —el
    # supervisor no tiene por qué conocerlos— pero sí se muestran como
    # referencia para que se pueda auditar contra el programa.
    ids = resolver_ids(ruta)

    # Lo único que falta de la actividad fuera de programa: que el
    # supervisor diga qué se ejecutó, porque no viene de ningún
    # catálogo que lo diga por él.
    descripcion_fp = ""
    if fuera_de_programa:
        st.markdown(
            '<p class="hint">⚠️ Actividad fuera del programa: el ID ACT y el COD '
            'STEP se generaron solo para este reporte y no están en el catálogo. '
            'Describa qué se ejecutó.</p>', unsafe_allow_html=True)
        descripcion_fp = st.text_input(
            "Actividad ejecutada", key="f_desc_fp",
            placeholder="Ej. Reparación de cerco perimetral dañado por temporal"
        ).strip()
    nombre_actividad = descripcion_fp or step["nombre"]

    # Coordenadas de la obra más cercana en la ruta (vienen de la planilla).
    # OJO con el nombre: `ubicacion` es el diccionario de arriba, y se usa
    # en el INSERT del final. Esto es solo el texto que se pinta.
    coords = next((n for n in reversed(ruta) if n.get("latitud") and n.get("longitud")), None)
    coords_texto = (f' · 📍 {coords["latitud"]:.5f}, {coords["longitud"]:.5f}'
                    if coords else "")

    # Jornada: no se pregunta. El reporte es del día en que se hace y las
    # horas salen del día de la semana. Se muestran en la tarjeta de
    # contexto para que el supervisor vea con qué se va a guardar.
    fecha = date.today()
    horas = horas_por_fecha(fecha)

    codigos = " · ".join(t for t in (
        f'ACT {ids["id_actividad"]}' if ids["id_actividad"] else "",
        f'STEP {ids["id_step"]}' if ids["id_step"] else "") if t)

    # El estado va junto a los códigos, como referencia de en qué punto
    # está la actividad. No se edita aquí: se deduce de lo ya reportado.
    estado_vigente = estado_nodo(step["id"])
    insignia = (f'<span class="estado" style="background:'
                f'{COLOR_ESTADO.get(estado_vigente, "#6C757D")}">{estado_vigente}</span>')

    st.markdown(
        f'<div class="ctx-card">'
        f'<div class="step">📄 {nombre_actividad}</div>'
        f'<div class="ruta">{ruta_texto(ruta)}</div>'
        f'<div class="meta">👷 {ss.supervisor or auth["nombre"]} · 🏢 {auth["empresa"]}'
        f'{coords_texto}</div>'
        + (f'<div class="codigos">🔗 {codigos}{insignia}</div>'
           if codigos else f'<div class="codigos">{insignia}</div>') +
        f'<div class="jornada">{etiqueta_dia(fecha)} {fecha.strftime("%d-%m-%Y")} · '
        f'jornada de {fmt_horas(horas)} h</div>'
        f'</div>', unsafe_allow_html=True)

    if not PERFILES.get(auth["perfil"], {}).get("puede_reportar", False):
        st.warning(f"El perfil **{auth['perfil']}** tiene acceso de solo consulta: "
                   "puede navegar el árbol pero no registrar reportes.")
        bottom_nav()
        return

    unidad = step.get("unidad") or "und"
    sello = ultimo_sello(auth["usuario"])

    # Lo primero del Step: quién lo ejecutó. De esa respuesta depende
    # la lista de maquinaria que se ofrece más abajo.
    st.markdown("**Ejecución**")
    empresas = subcontratos_con_maquinaria()
    propia = (auth["empresa"] or "").strip().upper()
    if propia and propia not in empresas:
        empresas = [propia] + empresas
    indice = empresas.index(propia) if propia in empresas else 0
    subcontrato = st.selectbox(
        "Subcontrato que ejecutó la actividad", empresas, index=indice, key="f_subcontrato",
        help="Define qué maquinaria aparece disponible en Recursos.")

    cantidad = st.number_input(f"Cantidad ejecutada ({unidad})",
                               min_value=0.0, step=1.0, value=0.0, key="f_cant")
    # El estado se muestra arriba, junto a los códigos. Se deja de pedir
    # aquí: cerrar una actividad será un paso propio, no una casilla que
    # el supervisor marque al pasar.
    if MOSTRAR_ESTADO_EDITABLE:
        estado = st.selectbox("Estado de la actividad", ESTADOS, index=0, key="f_estado")
    else:
        estado = ESTADO_POR_DEFECTO

    st.markdown("**Recursos**")
    personal = st.number_input("Personal en el frente", min_value=0, step=1, value=0,
                               key="f_personal")

    # HH calculada: es el dato que se compara contra las HH LB del
    # programa. Se muestra en pantalla para que el supervisor vea qué
    # se va a informar, pero no se escribe a mano: sale de la jornada
    # y de la dotación que ya declaró.
    hh = hh_calculada(horas, personal)
    if MOSTRAR_LINEA_BASE and ids["hh_lb"]:
        consumo = hh / ids["hh_lb"] * 100 if ids["hh_lb"] else 0
        detalle = (f" · línea base {fmt_horas(ids['hh_lb'])} HH "
                   f"({consumo:.0f}% consumido en esta jornada)")
    else:
        detalle = ""
    st.markdown(
        f'<p class="hint">⏱️ <b>HH calculada: {fmt_horas(hh)}</b> '
        f'({fmt_horas(horas)} h × {personal} persona/s){detalle}</p>',
        unsafe_allow_html=True)

    # La HH LB se sigue guardando en el reporte aunque no se muestre:
    # oficina técnica la necesita para contrastar rendimiento.
    if MOSTRAR_NOMBRES_PERSONAL:
        personal_nombres = st.text_area(
            "Nombres del personal (opcional)", height=68, key="f_nombres",
            placeholder="Un nombre por línea")
    else:
        personal_nombres = ""

    # Maquinaria del subcontrato declarado arriba. La clave del widget
    # incluye la empresa: al cambiar de subcontrato el multiselect se
    # remonta limpio, en vez de arrastrar patentes que ya no existen en
    # la lista nueva.
    maquinas = maquinaria_de(subcontrato)
    tipos_por_patente = {m["patente"]: m["tipo"] for m in maquinas}
    if maquinas:
        patentes = st.multiselect(
            f"Maquinaria de {subcontrato}", list(tipos_por_patente),
            format_func=lambda p: f"{p} · {tipos_por_patente[p]}",
            key=f"f_maq_{_slug(subcontrato)}",
            help="Puede elegir una o varias. Solo aparece la del subcontrato declarado.")
    else:
        patentes = []
        st.markdown(
            f'<p class="hint">{subcontrato} no tiene maquinaria registrada. '
            f'Agréguela aquí abajo.</p>', unsafe_allow_html=True)

    # El alta de una máquina termina en `st.rerun()` —solo así aparece
    # de inmediato en el multiselect, que ya se dibujó más arriba—, y un
    # rerun descarta lo que se haya escrito en pantalla. Por eso el aviso
    # viaja en session_state y se muestra en la pasada siguiente.
    aviso_maquina = ss.pop("f_maq_aviso", None)
    if aviso_maquina:
        st.success(aviso_maquina)

    with st.expander("➕ La máquina no está en la lista"):
        st.markdown(
            '<p class="hint">Queda registrada en el parque del subcontrato y '
            'disponible para los próximos reportes.</p>', unsafe_allow_html=True)
        c5, c6 = st.columns(2)
        with c5:
            nueva_patente = st.text_input("Patente", key="f_maq_patente",
                                          placeholder="Ej. VTTY-67")
        with c6:
            nuevo_tipo = st.text_input("Tipo de máquina", key="f_maq_tipo",
                                       placeholder="Ej. EXCAVADORA")
        # La clave incluye el subcontrato: con una clave fija, Streamlit
        # ignora `value` en los rerun siguientes y el campo se quedaba
        # con la empresa anterior —la máquina terminaba registrada en el
        # parque equivocado—.
        nueva_empresa = st.text_input(
            "Subcontrato dueño", value=subcontrato, key=f"f_maq_empresa_{_slug(subcontrato)}",
            help="Cámbielo solo si la máquina es de una empresa que aún no está en la lista.")
        if st.button("Registrar maquinaria", key="f_maq_agregar", width="stretch"):
            ok, mensaje = agregar_maquinaria(nueva_empresa, nueva_patente, nuevo_tipo,
                                             auth["usuario"])
            if ok:
                if _normalizar(nueva_empresa) != _normalizar(subcontrato):
                    mensaje += (f" Seleccione «{nueva_empresa.strip().upper()}» como "
                                "subcontrato ejecutor para poder usarla.")
                ss["f_maq_aviso"] = mensaje
                # Los campos se vacían para que el alta siguiente parta limpia.
                for clave in ("f_maq_patente", "f_maq_tipo"):
                    ss.pop(clave, None)
                st.rerun()
            else:
                st.error(mensaje)

    # `equipos` sigue siendo el texto legible que ya esperaba SharePoint
    # (una máquina por línea); `equipos_patentes` es la misma selección
    # en forma de identificadores, para cruzarla sin parsear texto.
    equipos = "\n".join(f"{p} · {tipos_por_patente[p]}" for p in patentes)
    equipos_patentes = ";".join(patentes)

    hm = hm_calculada(horas, len(patentes))
    if patentes:
        st.markdown(
            f'<p class="hint">🚜 <b>HM calculada: {fmt_horas(hm)}</b> '
            f'({fmt_horas(horas)} h × {len(patentes)} máquina/s)</p>',
            unsafe_allow_html=True)

    if MOSTRAR_MATERIALES:
        materiales = st.text_area(
            "Materiales", height=68, key="f_materiales",
            placeholder="Un material por línea. Ej. Hormigón G25 · 12 m3")
    else:
        materiales = ""

    if MOSTRAR_SEGURIDAD:
        st.markdown("**Seguridad**")
        hubo_incidencia = st.checkbox("Hubo incidencia / accidente", key="f_incidencia")
        # El detalle solo existe si hubo algo que detallar: un campo vacío
        # permanente en pantalla es una invitación a ignorarlo.
        incidencias = ""
        if hubo_incidencia:
            incidencias = st.text_area(
                "Detalle de la incidencia", height=90, key="f_detalle",
                placeholder="Qué ocurrió, a quién, a qué hora y qué medidas se tomaron")
        riesgos = st.text_area("Riesgos detectados", height=68, key="f_riesgos",
                               placeholder="Condiciones inseguras observadas en el frente")
    else:
        hubo_incidencia, incidencias, riesgos = False, "", ""

    observaciones = st.text_area("Observaciones", height=90, key="f_obs",
                                 placeholder="Novedades, interferencias, motivos de retraso...")

    st.markdown("**Evidencia fotográfica**")
    facing = selector_camara()
    # El visor se abre recién cuando lo piden: así el parche de la
    # cámara ya está puesto cuando se solicita el stream, y a quien solo
    # sube fotos de la galería no se le pide permiso de cámara.
    foto_camara = None
    if ss.get("f_camara_abierta"):
        foto_camara = st.camera_input("Tomar fotografía", key=f"f_cam_{facing}")
    elif st.button("📷 Abrir cámara", key="f_abrir_cam", width="stretch"):
        ss["f_camara_abierta"] = True
        st.rerun()
    fotos_archivo = st.file_uploader("O adjuntar desde la galería",
                                     type=["jpg", "jpeg", "png"],
                                     accept_multiple_files=True, key="f_galeria")

    with st.expander("📍 Datos que se imprimen sobre la foto", expanded=True):
        estampar = st.checkbox("Imprimir los datos sobre la foto", value=True, key="f_sello")
        camino = st.text_input("Ruta / camino", value=sello["camino"], key="f_camino",
                               placeholder="Ej. Ruta G-78")
        c3, c4 = st.columns(2)
        with c3:
            pk_desde = st.text_input("PK desde", key="f_pk1", placeholder="52500")
            comuna = st.text_input("Comuna", value=sello["comuna"], key="f_comuna")
            altitud = st.number_input("Altitud (m.s.n.m.)", min_value=0.0, max_value=7000.0,
                                      step=1.0, value=float(sello["altitud"]), key="f_altitud")
        with c4:
            pk_hasta = st.text_input("PK hasta", key="f_pk2", placeholder="52580")
            region = st.text_input("Región", value=sello["region"], key="f_region")
            orientacion = st.selectbox("Orientación de la foto", ORIENTACIONES, index=0,
                                       key="f_orient",
                                       help="Punto cardinal hacia el que apunta la cámara")
        pais = st.text_input("País", value=sello["pais"], key="f_pais")
        if not HAY_PILLOW:
            st.warning("Pillow no está instalado: las fotos se guardan sin sello.")

    pk_desde, pk_hasta = pk_desde.strip(), pk_hasta.strip()
    texto_pk = ""
    if pk_desde and pk_hasta:
        texto_pk = f"Entre PK{pk_desde} a PK{pk_hasta}"
    elif pk_desde or pk_hasta:
        texto_pk = f"PK{pk_desde or pk_hasta}"

    datos_sello = {
        "fecha_hora": f"{fecha.strftime('%d-%m-%Y')}, {datetime.now().strftime('%H:%M')}",
        "coords": (f"{coords['latitud']:.5f}, {coords['longitud']:.5f}" if coords else ""),
        "altitud": altitud,
        "orientacion": orientacion,
        "camino": camino.strip(),
        "comuna": comuna.strip(),
        "region": region.strip(),
        "pais": pais.strip(),
        "pk": texto_pk,
        "pie": f"{nombre_obra or obra} · {auth['empresa']}",
    }
    fotos_listas = preparar_fotos(foto_camara, fotos_archivo, estampar,
                                  nombre_actividad, lineas_sello(datos_sello))
    if fotos_listas:
        st.image(fotos_listas[0][1], width="stretch",
                 caption=f"Así se guardará la evidencia ({len(fotos_listas)} foto/s)")

    st.markdown("**Firma**")
    firma_nombre = st.text_input("Nombre de quien firma",
                                 value=ss.supervisor or auth["nombre"], key="f_firma")
    firma_ok = st.checkbox("Declaro que la información registrada es correcta", key="f_ok")

    guardar = st.button("✅ Guardar reporte", key="btn_guardar", width="stretch")

    if guardar:
        errores = []
        if not firma_ok:
            errores.append("Debe firmar la declaración para guardar el reporte.")
        if not firma_nombre.strip():
            errores.append("Indique el nombre de quien firma.")
        if hubo_incidencia and not incidencias.strip():
            errores.append("Marcó que hubo incidencia: describa el detalle.")
        if fuera_de_programa and not descripcion_fp:
            errores.append("Describa qué actividad fuera del programa se ejecutó.")
        if errores:
            for e in errores:
                st.error(e)
            return

        reporte_id = str(uuid.uuid4())[:8].upper()
        rutas_foto = guardar_fotos(reporte_id, fotos_listas)

        conn = get_db()
        conn.execute("""INSERT INTO reportes (
            id, creado_en, usuario, empresa, perfil, supervisor,
            nodo_id, step_nombre, ruta_ids, ruta_texto, obra, especialidad,
            actividad, id_actividad, id_step, hh_calculada, hm_calculada, hh_lb, cantidad_lb,
            subcontrato, equipos_patentes,
            fecha, hora_inicio, hora_termino, horas,
            cantidad, unidad, avance_pct, estado,
            personal, personal_nombres, equipos, materiales,
            clima, temperatura, observaciones, incidencias, riesgos,
            firma_nombre, firma_ok, fotos,
            latitud, longitud, tipo_obra, nombre_obra,
            camino, comuna, region, pais, altitud, orientacion, pk_desde, pk_hasta,
            cod_obra, tramo, sector,
            sincronizado
        ) VALUES (?,?,?,?,?,?, ?,?,?,?,?,?, ?,?,?,?,?,?,?, ?,?, ?,?,?,?, ?,?,?,?, ?,?,?,?, ?,?,?,?,?, ?,?,?,
                  ?,?,?,?, ?,?,?,?,?,?,?,?, ?,?,?, 0)""", (
            reporte_id, datetime.now().strftime("%d-%m-%Y %H:%M"),
            auth["usuario"], auth["empresa"], auth["perfil"], ss.supervisor,
            step["id"], nombre_actividad, json.dumps([n["id"] for n in ruta]),
            ruta_texto(ruta), obra, especialidad,
            # Enganche con el programa: la HH calculada se recalcula
            # aquí para que sea la de los valores efectivamente guardados.
            ids["actividad"], ids["id_actividad"], ids["id_step"],
            hh_calculada(horas, personal), hm_calculada(horas, len(patentes)),
            ids["hh_lb"], ids["cantidad_lb"],
            subcontrato, equipos_patentes,
            fecha.strftime("%d-%m-%Y"),
            # El avance acumulado, el clima y las horas de entrada y
            # salida salieron del formulario: las columnas quedan por
            # compatibilidad, sin dato. Las horas las fija el día.
            "", "",
            horas, cantidad, unidad, None, estado,
            personal, personal_nombres.strip(), equipos.strip(), materiales.strip(),
            "", None, observaciones.strip(),
            incidencias.strip() if hubo_incidencia else "", riesgos.strip(),
            firma_nombre.strip(), 1, json.dumps(rutas_foto),
            coords["latitud"] if coords else None,
            coords["longitud"] if coords else None,
            tipo_obra, nombre_obra,
            camino.strip(), comuna.strip(), region.strip(), pais.strip(),
            altitud or None, "" if orientacion == "—" else orientacion,
            pk_desde, pk_hasta,
            ubicacion["cod_obra"], ubicacion["tramo"], ubicacion["sector"],
        ))
        conn.commit()
        conn.close()

        # El reporte ya está a salvo en SQLite. El envío es un extra:
        # si falla, queda en la cola y se reintenta desde ⚙️ Envío.
        ss.ultimo_reporte = reporte_id
        ss.ultimo_sync = sincronizar_reporte(reporte_id)
        # Se marca el formulario como consumido: la limpieza ocurre al
        # entrar de nuevo, antes de crear los widgets (Streamlit no deja
        # tocar la clave de un widget ya instanciado).
        ss["f_step_id"] = None
        ir("guardado")

    bottom_nav()


def preparar_fotos(foto_camara, fotos_archivo, estampar, titulo, lineas):
    """Deja la evidencia lista: [(nombre, bytes)] ya sellados.

    Se prepara antes de guardar para que la vista previa muestre
    exactamente el archivo que se va a almacenar y enviar.
    """
    listas = []
    for archivo in ([foto_camara] if foto_camara else []) + list(fotos_archivo or []):
        contenido = archivo.getvalue()
        nombre = getattr(archivo, "name", "") or "camara.jpg"
        if estampar:
            contenido = estampar_foto(contenido, titulo, lineas)
            nombre = os.path.splitext(nombre)[0] + ".jpg"
        listas.append((nombre, contenido))
    return listas


def guardar_fotos(reporte_id, fotos):
    """Guarda la evidencia en disco y devuelve las rutas relativas."""
    rutas = []
    if not fotos:
        return rutas
    destino = os.path.join(FOTOS_DIR, reporte_id)
    os.makedirs(destino, exist_ok=True)
    for i, (nombre, contenido) in enumerate(fotos, start=1):
        ext = os.path.splitext(nombre)[1].lower() or ".jpg"
        if ext not in (".jpg", ".jpeg", ".png"):
            ext = ".jpg"
        ruta = os.path.join(destino, f"{i:02d}{ext}")
        with open(ruta, "wb") as f:
            f.write(contenido)
        rutas.append(os.path.relpath(ruta, BASE_DIR))
    return rutas


def pantalla_guardado():
    """Confirmación: se muestra la trazabilidad completa del registro."""
    ss = st.session_state
    app_header("Reporte guardado", f"Folio {ss.ultimo_reporte}", con_volver=False)

    st.success(f"Reporte **{ss.ultimo_reporte}** registrado correctamente.")
    st.markdown(f'<div class="breadcrumb">📍 <b>{ruta_texto(ss.ruta)}</b></div>',
                unsafe_allow_html=True)

    # El guardado local ya ocurrió; aquí solo se informa si además llegó
    # a sus destinos o si quedó en la cola de reenvío.
    ok, mensaje = ss.get("ultimo_sync") or (False, "")
    activos = destinos_activos()
    destinos = " y ".join(n for n, v in (("Google Sheets", activos["sheets"]),
                                         ("SharePoint", activos["webhook"]),
                                         ("Drive", activos["drive"])) if v)
    if ok and destinos:
        st.caption(f"☁️ Guardado en {destinos}.")
    elif destinos:
        st.warning(f"Guardado en el equipo, pendiente de enviar a {destinos}. {mensaje}")
    else:
        st.caption("💾 Guardado local (no hay destino configurado).")

    # Las fotos merecen su propia línea: son lo único que no sobrevive al
    # reinicio si se queda en el disco del contenedor.
    reporte_recien = reporte_por_id(ss.ultimo_reporte)
    n_fotos = len(json.loads((reporte_recien["fotos"] if reporte_recien else "") or "[]"))
    if n_fotos:
        subidas = json.loads((reporte_recien["fotos_url"] or "[]"))
        if subidas:
            st.caption(f"📷 {len(subidas)} de {n_fotos} foto(s) en Google Drive.")
        elif activos["drive"]:
            st.warning(f"📷 Las {n_fotos} foto(s) NO subieron a Drive: quedan solo en "
                       f"este equipo. Reintente desde ☁️ Envío.")
        else:
            st.warning(f"📷 {n_fotos} foto(s) guardadas solo en este equipo: falta "
                       f"configurar la carpeta de Google Drive.")

    # La planilla del reporte recién emitido: es el momento en que las
    # fotos están en disco y el supervisor la puede enviar o archivar.
    reporte = reporte_por_id(ss.ultimo_reporte)
    if reporte is not None:
        st.divider()
        boton_excel(reporte, "btn_excel_guardado")
        st.markdown(
            '<p class="hint">COD ACTIVIDAD · COD STEPS · FECHA · HH · HM · SUBCONTRATO · '
            'SUPERVISOR · MAQUINA USADA · OBS · FOTO, con la evidencia incrustada.</p>',
            unsafe_allow_html=True)
        st.divider()

    if st.button("➕ Otro reporte en el mismo Step", key="btn_mismo", width="stretch"):
        ir("formulario")
    if st.button("⬅️ Volver al nivel anterior", key="btn_nivel", width="stretch"):
        if ss.ruta:
            # Fuera de programa son dos nodos generados (Actividad y
            # Step) que no existen en el árbol: subir un solo nivel
            # dejaría parado en la Actividad fantasma, sin nada que
            # mostrar. Se sube directo a la obra real.
            extra = ss.ruta.pop()
            if extra.get("fuera_de_programa") and ss.ruta:
                ss.ruta.pop()
        ir("arbol")
    if st.button("🏗️ Empezar desde la obra", key="btn_inicio", width="stretch"):
        ss.ruta = []
        ir("arbol")
    if st.button("🗂️ Ver mis reportes", key="btn_ver", width="stretch"):
        ir("reportes")


def pantalla_reportes():
    """Consulta de lo registrado, con la ruta completa de cada reporte."""
    auth = st.session_state.auth
    app_header("Mis reportes", auth["nombre"], con_volver=False)

    conn = get_db()
    if PERFILES.get(auth["perfil"], {}).get("ve_todo"):
        filas = conn.execute(
            "SELECT * FROM reportes ORDER BY rowid DESC LIMIT 100").fetchall()
    else:
        filas = conn.execute(
            "SELECT * FROM reportes WHERE usuario=? ORDER BY rowid DESC LIMIT 100",
            (auth["usuario"],)).fetchall()
    conn.close()

    if not filas:
        st.info("Todavía no hay reportes registrados.")
    else:
        st.markdown(f'<p class="hint">{len(filas)} reporte(s).</p>', unsafe_allow_html=True)

    for r in filas:
        color = {"Terminada": "#2FA84F", "Paralizada": "#DC3545",
                 "No ejecutada": "#6C757D"}.get(r["estado"], "#0B3B8A")
        with st.expander(f"{r['fecha']} · {r['step_nombre']} · {r['estado']}"):
            st.markdown(
                f'<span class="pill" style="background:{color};color:#fff">{r["estado"]}</span> '
                f'&nbsp;<span class="rep-meta">Folio {r["id"]}</span>', unsafe_allow_html=True)
            st.markdown(f'<div class="breadcrumb">📍 {r["ruta_texto"]}</div>', unsafe_allow_html=True)

            # Los reportes nuevos no traen hora de entrada ni de salida:
            # la jornada la fija el día. Los antiguos sí, y se muestran.
            if r["hora_inicio"] or r["hora_termino"]:
                linea_jornada = (f"- **Horario:** {r['hora_inicio'] or '—'} a "
                                 f"{r['hora_termino'] or '—'} ({r['horas']} h)  \n")
            else:
                linea_jornada = f"- **Jornada:** {fmt_horas(r['horas'] or 0)} h  \n"

            linea_hh = ""
            if r["hh_calculada"] is not None:
                contra_lb = (f" · línea base {fmt_horas(r['hh_lb'])} HH"
                             if MOSTRAR_LINEA_BASE and r["hh_lb"] else "")
                linea_hh = f"- **HH calculada:** {fmt_horas(r['hh_calculada'])} HH{contra_lb}  \n"
            if r["hm_calculada"]:
                linea_hh += f"- **HM calculada:** {fmt_horas(r['hm_calculada'])} HM  \n"
            linea_ids = ""
            if r["id_actividad"] or r["id_step"]:
                linea_ids = (f"- **Actividad › Step:** `{r['id_actividad'] or '—'}` › "
                             f"`{r['id_step'] or '—'}`  \n")
            linea_pk = ""
            if r["pk_desde"] or r["pk_hasta"]:
                camino = f" · {r['camino']}" if r["camino"] else ""
                linea_pk = f"- **Tramo:** PK{r['pk_desde']} a PK{r['pk_hasta']}{camino}  \n"

            st.markdown(
                f"- **Supervisor:** {r['supervisor'] or '—'}  \n"
                f"- **Empresa:** {r['empresa']}  \n"
                + (f"- **Ejecutó:** {r['subcontrato']}  \n" if r["subcontrato"] else "")
                + f"{linea_jornada}"
                f"- **Ejecutado:** {r['cantidad']} {r['unidad']}  \n"
                f"- **Personal:** {r['personal']}  \n"
                f"{linea_hh}{linea_ids}{linea_pk}"
                f"- **Firmado por:** {r['firma_nombre']}")
            for etiqueta, campo in (("Equipos", "equipos"), ("Materiales", "materiales"),
                                     ("Observaciones", "observaciones"),
                                     ("Incidencias", "incidencias"), ("Riesgos", "riesgos")):
                if r[campo]:
                    st.markdown(f"**{etiqueta}:** {r[campo]}")
            fotos = json.loads(r["fotos"] or "[]")
            if fotos:
                st.markdown(f"**Evidencia:** {len(fotos)} foto(s)")
                for ruta in fotos:
                    completa = os.path.join(BASE_DIR, ruta)
                    if os.path.exists(completa):
                        st.image(completa, width="stretch")
            boton_excel(r, f"btn_excel_{r['id']}")

    st.divider()
    bottom_nav()


def pantalla_obras():
    """Carga de la planilla oficial de obras (Oficina Técnica).

    Reemplaza la rama Obras Puntuales completa: tipos de obra,
    obras individuales y sus coordenadas.
    """
    auth = st.session_state.auth
    app_header("Planilla de obras", "Obras Puntuales", con_volver=False)

    if not PERFILES.get(auth["perfil"], {}).get("ve_todo"):
        st.warning("Solo Oficina Técnica y Administración pueden cargar la planilla.")
        bottom_nav()
        return

    conn = get_db()
    raiz = conn.execute(
        "SELECT id FROM nodos WHERE padre_id IS NULL AND nombre='Obras Puntuales'").fetchone()
    resumen = []
    if raiz:
        resumen = conn.execute("""
            SELECT n.codigo, n.nombre,
                   (SELECT COUNT(*) FROM nodos h WHERE h.padre_id = n.id) hijos
            FROM nodos n WHERE n.padre_id = ? ORDER BY n.orden""", (raiz["id"],)).fetchall()
    conn.close()

    st.markdown('<div class="pregunta">Estructura cargada</div>', unsafe_allow_html=True)
    if resumen:
        st.markdown("\n".join(
            f"- `{r['codigo']}` **{r['nombre']}** — {r['hijos']} ítem(s)" for r in resumen))
    else:
        st.info("Todavía no hay obras puntuales cargadas.")

    st.divider()
    st.markdown('<div class="pregunta">Cargar planilla</div>', unsafe_allow_html=True)
    st.markdown(
        '<p class="hint">Archivo .xlsx o .csv con las columnas '
        '<b>Nivel · ID TIPO OBRA · TIPO OBRA · ID OBRA · OBRA · Latitud · Longitud</b>. '
        'Las filas de Nivel 1 son los tipos; las de detalle, las obras individuales. '
        'Los encabezados pueden estar en cualquier fila.</p>', unsafe_allow_html=True)

    archivo = st.file_uploader("Planilla de obras", type=["xlsx", "xlsm", "csv"])
    st.markdown(
        '<p class="hint">⚠️ La carga reemplaza la rama Obras Puntuales completa. '
        'Los reportes ya emitidos conservan su ruta como texto, pero dejan de '
        'apuntar a un nodo vigente.</p>', unsafe_allow_html=True)
    confirmar = st.checkbox("Confirmo el reemplazo de la estructura actual")

    if archivo and confirmar and st.button("📥 Importar planilla", key="btn_importar",
                                           width="stretch"):
        try:
            filas = leer_planilla_obras(archivo, archivo.name)
            resultado = importar_obras_puntuales(filas)
        except Exception as error:                      # planilla mal formada
            st.error(f"No se pudo importar: {error}")
        else:
            st.success(
                f"Planilla cargada: {resultado['tipos']} tipos de obra, "
                f"{resultado['obras']} obras y {resultado['nodos']} nodos en total.")
            if resultado.get("omitidos_lineales"):
                st.caption("Tipos lineales ignorados (viven en su propia rama): "
                           + ", ".join(resultado["omitidos_lineales"]) + ".")
            st.session_state.ruta = []
            limpiar_cache_permisos()
            st.rerun()

    # ---- Obras Lineales: estructura fija, sin planilla ----
    st.divider()
    st.markdown('<div class="pregunta">Obras Lineales · estructura fija</div>',
                unsafe_allow_html=True)
    st.markdown(
        '<p class="hint">Vías y Catenarias. No se cargan por planilla: son la '
        'estructura fija del proyecto, definida en <code>VIAS</code> y '
        '<code>CATENARIAS</code>. El código de cada tipo dice dónde está —'
        '<code>T2-E-VC</code> es Tramo 2 · Estación · Vía Carga—. Resembrar conserva '
        'las asignaciones de supervisores.</p>', unsafe_allow_html=True)

    conn = get_db()
    raiz_fija = conn.execute(
        "SELECT id FROM nodos WHERE padre_id IS NULL AND nombre=?",
        (RAMA_LINEALES,)).fetchone()
    if raiz_fija is None:
        st.info(f"Todavía no se ha sembrado la rama {RAMA_LINEALES}.")
    else:
        for disciplina in conn.execute(
                "SELECT id, nombre, codigo FROM nodos WHERE padre_id=? ORDER BY orden",
                (raiz_fija["id"],)).fetchall():
            tramos = conn.execute("SELECT COUNT(*) x FROM nodos WHERE padre_id=?",
                                  (disciplina["id"],)).fetchone()["x"]
            st.markdown(f"- **{disciplina['nombre']}** — {tramos} tramos, "
                        f"{len(nodos_obra(disciplina['id']))} obras, "
                        f"{contar_hojas(disciplina['id'])} hojas")
    conn.close()

    if st.button("📥 Sembrar / actualizar Obras Lineales", key="btn_sembrar_fijas",
                 width="stretch"):
        try:
            resultado = importar_ramas_fijas()
        except Exception as error:
            st.error(f"No se pudo sembrar: {error}")
        else:
            st.success("  ·  ".join(
                [f"{nombre}: {d['tipos']} tipos, {d['obras']} obras"
                 for nombre, d in resultado.items() if nombre != "nodos"]
                + [f"{resultado['nodos']} nodos"]))
            st.session_state.ruta = []
            limpiar_cache_permisos()
            st.rerun()

    # ---- El programa que viaja dentro de la app ----
    st.divider()
    st.markdown('<div class="pregunta">Programa empaquetado</div>', unsafe_allow_html=True)
    st.markdown(
        '<p class="hint">Las actividades del diario viajan comprimidas junto a la app '
        '(<code>programa_tam.json.gz</code>) y se cargan solas la primera vez que se '
        'crea la base. Este botón las vuelve a cargar: sirve cuando el contenedor '
        'conservó una base anterior y el árbol quedó desactualizado.</p>',
        unsafe_allow_html=True)

    if not os.path.exists(ARCHIVO_PROGRAMA):
        st.warning(
            "No se encontró `programa_tam.json.gz` junto a la app. Sin él, las obras "
            "quedan vacías tras cada reinicio: hay que subirlo al repositorio.")
    else:
        conn = get_db()
        con_programa = conn.execute(
            "SELECT COUNT(DISTINCT id_actividad) c FROM nodos "
            "WHERE id_actividad IS NOT NULL AND id_actividad <> ''").fetchone()["c"]
        conn.close()
        st.markdown(
            f'<p class="hint">Archivo: {os.path.getsize(ARCHIVO_PROGRAMA) / 1024:.0f} KB · '
            f'en el árbol hay <b>{con_programa}</b> actividades del programa.</p>',
            unsafe_allow_html=True)
        if st.button("📥 Recargar el programa empaquetado", key="btn_recargar_programa",
                     width="stretch"):
            resultado = cargar_programa_empaquetado()
            if resultado is None:
                st.error("No se pudo leer el archivo del programa.")
            else:
                st.success(
                    f"{resultado['steps']} actividades y {resultado['pasos']} pasos "
                    f"en {resultado['obras']} obras.")
                st.session_state.ruta = []
                limpiar_cache_permisos()
                st.rerun()

    # ---- Carga masiva desde el export de P6 ----
    st.divider()
    st.markdown('<div class="pregunta">Actividades desde el cronograma (P6)</div>',
                unsafe_allow_html=True)
    st.markdown(
        '<p class="hint">Export de Primavera P6 con la jerarquía WBS y las columnas '
        '<b>Activity ID · Activity Name · 00 OBRA</b>. Carga todas las obras de una '
        'pasada: <code>00 OBRA</code> enlaza cada actividad con la suya. Mientras no '
        'existan los Steps, cada actividad se carga como un Step único que es la propia '
        'actividad. Solo alcanza a Obras Puntuales: la rama lineal no está en esta '
        'parte del cronograma.</p>', unsafe_allow_html=True)

    archivo_p6 = st.file_uploader("Export de P6", type=["xlsx", "xlsm", "csv"], key="p6_archivo")
    confirmar_p6 = st.checkbox(
        "Confirmo el reemplazo de las actividades de las obras alcanzadas", key="p6_confirmar")

    if archivo_p6 and confirmar_p6 and st.button("📥 Cargar actividades del cronograma",
                                                 key="btn_importar_p6", width="stretch"):
        try:
            filas_p6 = leer_export_p6(archivo_p6, archivo_p6.name)
            resultado = importar_actividades_p6(filas_p6)
        except Exception as error:                      # export mal formado
            st.error(f"No se pudo importar: {error}")
        else:
            st.success(
                f"{resultado['actividades']} actividades cargadas en "
                f"{resultado['obras']} obras ({resultado['nodos']} nodos).")
            if resultado["omitidos"]:
                st.caption(f"Tipos no tocados: {', '.join(resultado['omitidos'])}.")
            if resultado["sin_enlace"]:
                detalle = "\n".join(f"- {k}: {v} actividad(es)"
                                    for k, v in resultado["sin_enlace"].items())
                st.warning("Actividades sin obra reconocible, quedaron fuera:\n\n" + detalle)
            st.session_state.ruta = []
            limpiar_cache_permisos()
            st.rerun()

    # ---- Carga masiva desde la BD del diario ----
    st.divider()
    st.markdown('<div class="pregunta">BD del diario · todas las obras</div>',
                unsafe_allow_html=True)
    st.markdown(
        '<p class="hint">Export de la BD del diario, <b>sin fila de encabezados</b>: '
        'columnas fijas, con <b>ID OBRA</b> en la octava. Ese código enlaza cada '
        'actividad con su obra, así que carga las dos ramas de una pasada —puntuales '
        'y lineales—. Bajo cada obra quedan <b>WBS Name</b> como actividad y '
        '<b>Activity Name</b> como Step, con su COD STEP, unidad, HH y cantidad de '
        'línea base.</p>', unsafe_allow_html=True)

    archivo_bd = st.file_uploader("BD del diario", type=["xlsx", "xlsm", "csv"],
                                  key="bd_archivo")
    confirmar_bd = st.checkbox(
        "Confirmo el reemplazo de las actividades de las obras alcanzadas",
        key="bd_confirmar")

    if archivo_bd and confirmar_bd and st.button(
            "📥 Cargar BD del diario", key="btn_importar_bd", width="stretch"):
        try:
            filas_bd = leer_bd_diario(archivo_bd, archivo_bd.name)
            resultado = importar_bd_diario(filas_bd)
        except Exception as error:                      # export mal formado
            st.error(f"No se pudo importar: {error}")
        else:
            st.success(
                f"{resultado['steps']} Steps en {resultado['actividades']} actividades, "
                f"repartidos en {resultado['obras']} obras ({resultado['nodos']} nodos).")
            if resultado.get("con_step_propio"):
                st.caption(f"{resultado['con_step_propio']} con STEP ID propio del "
                           "programa; el resto es su propia actividad.")
            else:
                st.caption("La columna STEP ID viene vacía: cada actividad queda como "
                           "su propio Step. Al publicarse los Steps, se recarga y listo.")
            if resultado.get("corregidas"):
                st.caption("ID OBRA corregido por el nombre de la actividad: "
                           + ", ".join(f"{k} → {v}"
                                       for k, v in resultado["corregidas"].items()) + ".")
            if resultado["sin_enlace"]:
                detalle = "\n".join(f"- `{k}`: {v} actividad(es)"
                                    for k, v in sorted(resultado["sin_enlace"].items()))
                st.warning("ID OBRA que no existe en el árbol, quedaron fuera:\n\n" + detalle)
            st.session_state.ruta = []
            limpiar_cache_permisos()
            st.rerun()

    # ---- Actividades y Steps de una obra ----
    st.divider()
    st.markdown('<div class="pregunta">Actividades y Steps de una obra</div>',
                unsafe_allow_html=True)
    st.markdown(
        '<p class="hint">Carga los dos últimos niveles del árbol —<b>Actividad</b> y '
        '<b>Step</b>— desde la BD de actividades del diario, para la obra que elija. '
        'Se leen <b>ID ACT · NOMBRE ACTIVIDAD · COD STEP · STEP · HH LB · UND · '
        'CANTIDAD LB</b>; el resto de columnas se ignora. Esos identificadores quedan '
        'guardados en el nodo y viajan solos en cada reporte: el supervisor nunca los '
        'escribe.</p>', unsafe_allow_html=True)

    # La carga es genérica por nodo destino, así que sirve igual para
    # las dos ramas: solo hay que preguntar primero de cuál se trata.
    conn = get_db()
    marcas = ",".join("?" * len(RAMAS_RAIZ))
    ramas = [dict(f) for f in conn.execute(
        f"SELECT id, nombre FROM nodos WHERE padre_id IS NULL "
        f"AND nombre IN ({marcas}) ORDER BY nombre", RAMAS_RAIZ).fetchall()]
    conn.close()

    if not ramas:
        st.info("Primero cargue la planilla de obras o siembre las Obras Lineales.")
        obras = []
    else:
        opciones_rama = {r["nombre"]: r for r in ramas}
        rama = opciones_rama[st.selectbox("Rama", list(opciones_rama), key="act_rama")]
        # Las ramas no tienen la misma profundidad —Vías mete tramo y
        # sector— así que en vez de encadenar selectores por nivel se
        # listan las obras con su ruta. El buscador del selector hace el
        # resto.
        obras = nodos_obra(rama["id"])

    if not obras:
        if ramas:
            st.info("Esa rama no tiene obras cargadas.")
    else:
        opciones_obra = {f"[{o['codigo']}] {o['ruta']}": o for o in obras}
        obra_destino = opciones_obra[st.selectbox(
            f"Obra de destino ({len(obras)})", list(opciones_obra), key="act_obra")]
        archivo_act = st.file_uploader("BD de actividades", type=["xlsx", "xlsm", "csv"],
                                       key="act_archivo")
        st.markdown(
            f'<p class="hint">⚠️ Reemplaza las actividades que hoy cuelgan de '
            f'<b>{obra_destino["ruta"]}</b>.</p>', unsafe_allow_html=True)
        confirmar_act = st.checkbox("Confirmo el reemplazo de sus actividades",
                                    key="act_confirmar")

        if archivo_act and confirmar_act and st.button(
                "📥 Cargar actividades", key="btn_importar_act", width="stretch"):
            try:
                filas_act = leer_planilla_actividades(archivo_act, archivo_act.name)
                resultado = importar_actividades(obra_destino["id"], filas_act)
            except Exception as error:              # planilla mal formada
                st.error(f"No se pudo importar: {error}")
            else:
                st.success(
                    f"{resultado['obra']}: {resultado['actividades']} actividad(es) y "
                    f"{resultado['steps']} Step(s) cargados.")
                st.session_state.ruta = []
                limpiar_cache_permisos()
                st.rerun()

    st.divider()
    bottom_nav()


# ---------- Asignación de actividades a supervisores ----------

def estado_asignacion(nodo_id, asignados):
    """directo | heredado | parcial | no, respecto de lo ya asignado."""
    if nodo_id in asignados:
        return "directo"
    if any(nodo_id.startswith(a + "/") for a in asignados):
        return "heredado"
    if any(a.startswith(nodo_id + "/") for a in asignados):
        return "parcial"
    return "no"


def ancestro_asignado(nodo_id, asignados):
    return next((a for a in asignados if nodo_id.startswith(a + "/")), None)


# Estas dos escriben en la conexión que reciben y no confirman ni
# registran nada: las usan tanto la pantalla —que además anota el
# movimiento en la bitácora— como la restauración desde Sheets, que
# justamente NO debe volver a anotar lo que está reproduciendo.

def _asignar_local(conexion, usuario_id, nodo_id):
    conexion.execute(
        "DELETE FROM permisos WHERE usuario_id=? AND (nodo_id=? OR nodo_id LIKE ?)",
        (usuario_id, nodo_id, nodo_id + "/%"))
    conexion.execute("INSERT OR IGNORE INTO permisos (usuario_id, nodo_id) VALUES (?,?)",
                     (usuario_id, nodo_id))


def _quitar_local(conexion, usuario_id, nodo_id):
    conexion.execute(
        "DELETE FROM permisos WHERE usuario_id=? AND (nodo_id=? OR nodo_id LIKE ?)",
        (usuario_id, nodo_id, nodo_id + "/%"))


def asignar_nodo(usuario_id, nodo_id, usuario="", por=""):
    """Asigna un nodo y todo lo que cuelga de él.

    Los permisos que ya existían más abajo se borran: quedarían
    duplicando lo mismo y ensucian la lista de asignaciones.
    """
    conn = get_db()
    _asignar_local(conn, usuario_id, nodo_id)
    conn.commit()
    conn.close()
    if usuario:
        registrar_asignacion_sheets(usuario, nodo_id, "asignar", por)


def quitar_nodo(usuario_id, nodo_id, usuario="", por=""):
    conn = get_db()
    _quitar_local(conn, usuario_id, nodo_id)
    conn.commit()
    conn.close()
    if usuario:
        registrar_asignacion_sheets(usuario, nodo_id, "quitar", por)


def excluir_nodo(usuario_id, nodo_id, ancestro_id, usuario="", por=""):
    """Quita una rama que se heredaba de un nodo asignado más arriba.

    No se puede simplemente borrar: lo asignado es el ancestro. Se
    reemplaza por sus hijos, bajando nivel a nivel hasta el nodo que se
    quiere excluir, y ese se deja fuera. Es el mismo desglose que haría
    a mano quien asignó de más.
    """
    ruta = [n["id"] for n in ruta_de(nodo_id)]
    if ancestro_id not in ruta:
        return
    tramo = ruta[ruta.index(ancestro_id):]

    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM permisos WHERE usuario_id=? AND nodo_id=?", (usuario_id, ancestro_id))
    hermanos = []
    for padre, siguiente in zip(tramo, tramo[1:]):
        for h in c.execute("SELECT id FROM nodos WHERE activo=1 AND padre_id=?",
                           (padre,)).fetchall():
            if h["id"] != siguiente:
                c.execute("INSERT OR IGNORE INTO permisos (usuario_id, nodo_id) VALUES (?,?)",
                          (usuario_id, h["id"]))
                hermanos.append(h["id"])
    conn.commit()
    conn.close()

    # El desglose se anota como los movimientos que realmente lo componen,
    # para que reproducir la bitácora dé exactamente este mismo estado.
    if usuario:
        registrar_asignacion_sheets(usuario, ancestro_id, "quitar", por)
        for hermano in hermanos:
            registrar_asignacion_sheets(usuario, hermano, "asignar", por)


def pantalla_asignaciones():
    """Qué puede reportar cada supervisor, a cualquier altura del árbol.

        Obras Puntuales › TIPO OBRA › OBRA › Actividad › Step

    Se asigna en el nivel que convenga: la obra completa, un tipo de
    obra, una actividad o un Step suelto. Lo asignado arrastra todo lo
    que cuelga, así que la asignación normal —«este supervisor lleva la
    Estación Talagante»— es un solo toque.
    """
    ss = st.session_state
    auth = ss.auth
    app_header("Asignar actividades", "Quién reporta qué", con_volver=False)

    if not PERFILES.get(auth["perfil"], {}).get("ve_todo"):
        st.warning("Solo Oficina Técnica y Administración pueden asignar actividades.")
        bottom_nav()
        return

    conn = get_db()
    supervisores = [dict(f) for f in conn.execute(
        "SELECT id, usuario, nombre, empresa FROM usuarios "
        "WHERE activo=1 AND perfil='Supervisor' ORDER BY empresa, nombre").fetchall()]
    conn.close()

    if not supervisores:
        st.info("No hay usuarios con perfil Supervisor.")
        bottom_nav()
        return

    etiquetas = {f"{s['nombre']} · {s['empresa']} ({s['usuario']})": s for s in supervisores}
    elegido = etiquetas[st.selectbox("Supervisor", list(etiquetas), key="asig_usuario")]
    asignados = nodos_asignados(elegido["id"])

    # ---- Lo que hoy tiene asignado ----
    st.markdown('<div class="pregunta">Actividades asignadas</div>', unsafe_allow_html=True)
    if not asignados:
        st.markdown('<p class="hint">Sin asignaciones: hoy no ve ninguna obra.</p>',
                    unsafe_allow_html=True)
    else:
        for nodo_id in asignados[:40]:
            ruta = ruta_de(nodo_id)
            if not ruta:                      # nodo borrado por una reimportación
                continue
            c1, c2 = st.columns([6, 1], vertical_alignment="center")
            with c1:
                st.markdown(f'<div class="breadcrumb">📍 {ruta_texto(ruta)}</div>',
                            unsafe_allow_html=True)
            with c2:
                if st.button("✕", key=f"quitar_{clave_widget('a', nodo_id)}",
                             help="Quitar esta asignación"):
                    quitar_nodo(elegido["id"], nodo_id, elegido["usuario"], auth["usuario"])
                    limpiar_cache_permisos()
                    st.rerun()
        if len(asignados) > 40:
            st.markdown(f'<p class="hint">…y {len(asignados) - 40} más.</p>',
                        unsafe_allow_html=True)

    # ---- Navegador del árbol ----
    st.divider()
    st.markdown('<div class="pregunta">Agregar desde el árbol</div>', unsafe_allow_html=True)

    ruta_nav = ss.asig_ruta
    if ruta_nav:
        st.markdown(f'<div class="breadcrumb">📂 {ruta_texto(ruta_nav)}</div>',
                    unsafe_allow_html=True)
        if st.button("⬅️ Subir un nivel", key="asig_subir", width="stretch"):
            ruta_nav.pop()
            st.rerun()
        opciones = hijos_de(ruta_nav[-1]["id"])
    else:
        st.markdown('<p class="hint">Nivel raíz: obras del proyecto.</p>',
                    unsafe_allow_html=True)
        opciones = hijos_de(None)

    if len(opciones) > MAX_OPCIONES_SIN_BUSCADOR:
        filtro = st.text_input("🔍 Buscar en este nivel", key=f"asig_buscar_{len(ruta_nav)}",
                               placeholder="Escriba para filtrar...").strip().lower()
        if filtro:
            opciones = [o for o in opciones
                        if filtro in o["nombre"].lower() or filtro in (o["codigo"] or "").lower()]
        if len(opciones) > MAX_OPCIONES_SIN_BUSCADOR:
            st.markdown(f'<p class="hint">{len(opciones)} opciones · use el buscador para acotar.</p>',
                        unsafe_allow_html=True)
            opciones = opciones[:MAX_OPCIONES_SIN_BUSCADOR]

    if not opciones:
        st.info("Este nivel no tiene nodos.")
        st.divider()
        bottom_nav()
        return

    conteos = contar_hijos([o["id"] for o in opciones])
    marcas = {"directo": "✅", "heredado": "🔒", "parcial": "◐", "no": "➕"}

    for o in opciones:
        estado = estado_asignacion(o["id"], asignados)
        n_hijos = conteos.get(o["id"], 0)
        codigo = f"[{o['codigo']}] " if o["codigo"] else ""
        icono = "📁" if n_hijos else "📄"
        detalle = f" ({n_hijos})" if n_hijos else ""

        c1, c2 = st.columns([5, 1], vertical_alignment="center")
        with c1:
            etiqueta = f"{icono} {codigo}{o['nombre']}{detalle}"
            if n_hijos:
                if st.button(etiqueta, key=clave_widget("asignav", o["id"]), width="stretch"):
                    ruta_nav.append(o)
                    st.rerun()
            else:
                st.markdown(f'<div class="asig-hoja">{etiqueta}</div>', unsafe_allow_html=True)
        with c2:
            ayuda = {"directo": "Asignado aquí · quitar",
                     "heredado": "Viene de un nivel superior · excluir esta rama",
                     "parcial": "Tiene partes asignadas · asignar todo",
                     "no": "Asignar esta rama completa"}[estado]
            if st.button(marcas[estado], key=clave_widget("asigmarca", o["id"]), help=ayuda):
                if estado == "directo":
                    quitar_nodo(elegido["id"], o["id"], elegido["usuario"], auth["usuario"])
                elif estado == "heredado":
                    excluir_nodo(elegido["id"], o["id"],
                                 ancestro_asignado(o["id"], asignados),
                                 elegido["usuario"], auth["usuario"])
                else:
                    asignar_nodo(elegido["id"], o["id"], elegido["usuario"], auth["usuario"])
                limpiar_cache_permisos()
                st.rerun()

    st.markdown(
        '<p class="hint">✅ asignado aquí · 🔒 heredado de un nivel superior · '
        '◐ con partes asignadas · ➕ sin asignar. Lo que se asigna incluye todo '
        'lo que cuelga debajo.</p>', unsafe_allow_html=True)

    st.divider()
    bottom_nav()


def pantalla_envio():
    """Estado y configuración de los dos destinos: Google Sheets
    (histórico durable) y el webhook de SharePoint (respaldo)."""
    auth = st.session_state.auth
    app_header("Guardado de datos", "Sheets y SharePoint", con_volver=False)

    if not PERFILES.get(auth["perfil"], {}).get("ve_todo"):
        st.warning("Solo Oficina Técnica y Administración pueden configurar el envío.")
        bottom_nav()
        return

    config = cargar_config_sync()
    fijos = config_en_secrets()
    activos = destinos_activos(config)
    pendientes = contar_pendientes()

    st.markdown('<div class="pregunta">Estado</div>', unsafe_allow_html=True)
    conn = get_db()
    enviados = conn.execute("SELECT COUNT(*) c FROM reportes WHERE sincronizado=1").fetchone()["c"]
    en_sheets = conn.execute("SELECT COUNT(*) c FROM reportes WHERE sheets_ok=1").fetchone()["c"]
    conn.close()
    c1, c2, c3 = st.columns(3)
    c1.metric("En Google Sheets", en_sheets)
    c2.metric("Completos", enviados)
    c3.metric("Pendientes", pendientes)

    if not any(activos.values()):
        st.warning("Sin destinos configurados: los reportes quedan **solo** en la base "
                   "local. En Streamlit Cloud eso significa perderlos en el próximo "
                   "reinicio. Configure Google Sheets en Secrets.")

    st.divider()
    st.markdown('<div class="pregunta">Google Sheets · histórico</div>', unsafe_allow_html=True)
    if not HAY_GSPREAD:
        st.error("Faltan las librerías `gspread` y `google-auth`. Agréguelas a "
                 "`requirements.txt` y vuelva a desplegar.")
    elif activos["sheets"]:
        libro = cliente_sheets()
        if libro is None:
            _, causa, pistas = diagnostico_sheets()
            st.error(f"No se pudo abrir el Sheet. **{causa}**")
            for pista in pistas:
                st.markdown(f'<p class="hint">· {pista}</p>', unsafe_allow_html=True)
            credenciales = secreto("google_sheets", "credentials") or {}
            st.markdown(
                '<p class="hint">Con lo que hay en Secrets ahora:</p>',
                unsafe_allow_html=True)
            st.code(f'sheet_id     = {secreto("google_sheets", "sheet_id") or "(vacío)"}\n'
                    f'client_email = {dict(credenciales).get("client_email", "(vacío)")}\n'
                    f'project_id   = {dict(credenciales).get("project_id", "(vacío)")}',
                    language=None)
            if st.button("🔄 Reintentar la conexión", key="btn_reintentar_sheets",
                         width="stretch"):
                cliente_sheets.clear()               # el cache guardó el None
                st.rerun()
        else:
            st.success(f"Conectado a «{libro.title}».")
            st.markdown(
                f'<p class="hint">Pestañas: <b>{HOJA_ENTREGA}</b> — los reportes, con '
                f'las mismas {len(columnas_entrega())} columnas que exporta el Excel —, '
                f'<b>{HOJA_MAQUINARIA}</b> y <b>{HOJA_ASIGNACIONES}</b> '
                f'(bitácora que permite reconstruir la configuración tras un '
                f'reinicio).</p>', unsafe_allow_html=True)
            with st.expander(f"Columnas de «{HOJA_ENTREGA}»"):
                st.markdown(" · ".join(f"`{c}`" for c in columnas_entrega()))
                st.markdown(
                    f'<p class="hint">Y al final <code>{COLUMNA_CONTROL_SP}</code>, '
                    f'que no es del reporte: la app la escribe en '
                    f'<b>{CONTROL_PENDIENTE}</b> y el flujo que copia las filas a '
                    f'SharePoint la pasa a SI cuando ya las llevó.</p>',
                    unsafe_allow_html=True)
            if st.button("🧪 Escribir una fila de prueba en Sheets", key="btn_probar_sheets",
                         width="stretch"):
                ok, mensaje = agregar_filas_sheets(
                    HOJA_ENTREGA, [fila_prueba_entrega(auth) + [CONTROL_PENDIENTE]],
                    columnas_hoja())
                (st.success if ok else st.error)(f"{'✅' if ok else '❌'} {mensaje}")
    else:
        st.info("No configurado. En Secrets: `[google_sheets]` con `sheet_id` y "
                "`[google_sheets.credentials]` con la cuenta de servicio. "
                "Ver `PUBLICAR-rep-app.md`.")

    # ---- Fotos en Drive ----
    st.divider()
    st.markdown('<div class="pregunta">Fotografías · Google Drive</div>',
                unsafe_allow_html=True)
    st.markdown(
        '<p class="hint">Publicada la app, el disco se borra en cada reinicio: la '
        'evidencia tiene que salir del servidor. Se sube a una carpeta de Drive y la '
        'columna <b>FOTO</b> guarda su enlace.</p>', unsafe_allow_html=True)

    if hay_drive():
        if webhook_fotos():
            st.success("Fotos por Apps Script: las guarda la cuenta dueña del Drive.")
            st.markdown(
                '<p class="hint">La app no guarda credenciales de Google para esto: '
                'manda la foto al script y él la crea en su propio Drive, con su cuota. '
                'Es lo que funciona con una cuenta Gmail normal.</p>',
                unsafe_allow_html=True)
        else:
            st.success("Carpeta de Drive configurada (cuenta de servicio).")
            st.warning(
                "Este modo solo funciona contra una **unidad compartida** de Google "
                "Workspace. En un Drive personal, Google responde *«Service Accounts "
                "do not have storage quota»* y la foto no sube. Para una cuenta Gmail, "
                "use el Apps Script: `deploy_rep_app/apps-script-fotos.gs`.")
        publicas = st.checkbox(
            "Fotos visibles para cualquiera con el enlace",
            value=bool(config.get("fotos_publicas")),
            disabled="fotos_publicas" in fijos,
            help="Necesario para que la hoja muestre la miniatura con =IMAGE(). "
                 "Sin esto la celda queda como enlace y pide permiso al abrirlo.")
        if publicas != bool(config.get("fotos_publicas")) and "fotos_publicas" not in fijos:
            guardar_config_sync({"fotos_publicas": publicas})
            st.rerun()
        st.markdown(
            '<p class="hint">La celda queda como '
            + ('<code>=IMAGE(…)</code>: la foto se ve dentro de la planilla.</p>'
               if publicas else
               '<code>=HYPERLINK(…)</code>: abre la foto en Drive, con permiso.</p>'),
            unsafe_allow_html=True)

        # «Configurada» solo dice que hay un id en Secrets. Que la subida
        # funcione depende de la API habilitada y de que la carpeta esté
        # compartida, y eso no se sabe hasta intentarlo: mejor ahora que
        # con la primera foto de terreno.
        if st.button("🧪 Subir una imagen de prueba a Drive", key="btn_probar_drive",
                     width="stretch"):
            pixel = base64.b64decode(
                "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8U"
                "HRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/wAALCAABAAEBAREA"
                "/8QAFAABAAAAAAAAAAAAAAAAAAAACf/EABQQAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQEA"
                "AD8AKp//2Q==")
            ok, resultado = subir_foto_drive(
                f"PRUEBA_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg", pixel,
                bool(config.get("fotos_publicas")))
            if ok:
                st.success("✅ Subida correcta. Borre el archivo de prueba de la carpeta.")
                st.markdown(f"[Ver el archivo subido]({ENLACE_VER.format(resultado)})")
            else:
                st.error(f"❌ {resultado}")
                if resultado.startswith("SIN_CUOTA"):
                    st.markdown(
                        '<p class="hint">· No es un problema de permisos: compartir la '
                        'carpeta no lo arregla. Google quitó la cuota de las cuentas de '
                        'servicio, así que el robot no puede ser <b>dueño</b> del archivo. '
                        'Se resuelve subiendo con una cuenta de usuario: ver '
                        '<b>«Las fotos»</b> en <code>PUBLICAR-rep-app.md</code>.</p>',
                        unsafe_allow_html=True)
                elif "SERVICE_DISABLED" in resultado or "accessNotConfigured" in resultado:
                    st.markdown(
                        '<p class="hint">· Falta habilitar <b>Google Drive API</b> en el '
                        'proyecto de Google Cloud.</p>', unsafe_allow_html=True)
                elif "404" in resultado or "403" in resultado:
                    st.markdown(
                        '<p class="hint">· La carpeta no existe con ese id, o no está '
                        'compartida como <b>Editor</b> con la cuenta de servicio.</p>',
                        unsafe_allow_html=True)
    else:
        st.info("No configurado: las fotos quedan solo en este equipo y se pierden en "
                "el próximo reinicio. Siga «Las fotos» en `PUBLICAR-rep-app.md`: "
                "publique `apps-script-fotos.gs` en su cuenta de Google y ponga su URL "
                "en Secrets como `webhook_fotos`.")

    st.divider()
    st.markdown('<div class="pregunta">SharePoint · respaldo</div>', unsafe_allow_html=True)
    st.markdown(
        '<p class="hint">Pegue la URL del disparador HTTP de un flujo de Power Automate '
        '(o el Web App de un Google Apps Script). La app hace POST de un JSON y el flujo '
        'crea el elemento en la lista. Es también el único destino de las fotos.</p>',
        unsafe_allow_html=True)
    if fijos:
        st.markdown(
            f'<p class="hint">🔒 Definido en Secrets (no editable aquí): '
            f'<code>{"</code>, <code>".join(sorted(fijos))}</code>.</p>',
            unsafe_allow_html=True)

    with st.form("form_sync"):
        modo = st.selectbox(
            "Modo", ["local", "webhook"],
            index=["local", "webhook"].index(config.get("modo", "local")),
            disabled="modo" in fijos,
            help="local = no envía a SharePoint · webhook = además envía")
        url = st.text_input("URL del webhook de reportes", value=config.get("url", ""),
                            disabled="url" in fijos,
                            placeholder="https://prod-00.westus.logic.azure.com/workflows/...")
        url_login = st.text_input(
            "URL del webhook de login (opcional)", value=config.get("url_login", ""),
            disabled="url_login" in fijos,
            placeholder="https://prod-00.westus.logic.azure.com/workflows/...",
            help="Valida contra la lista Usuarios App. Vacío = login solo local.")
        token = st.text_input("Token (opcional)", value=config.get("token", ""),
                              type="password", disabled="token" in fijos,
                              help="Si se define, viaja como cabecera X-Token")
        enviar_fotos = st.checkbox("Incluir fotos en el envío (base64)",
                                   value=config.get("enviar_fotos", True),
                                   disabled="enviar_fotos" in fijos)
        guardar = st.form_submit_button("💾 Guardar configuración", width="stretch")

    if guardar:
        # Lo fijado en Secrets no se escribe al archivo: quedaría un valor
        # fantasma que nunca se usa y confunde al siguiente que lo lea.
        nuevo = {**config, "modo": modo, "url": url.strip(),
                 "url_login": url_login.strip(),
                 "token": token.strip(), "enviar_fotos": enviar_fotos}
        guardar_config_sync({k: v for k, v in nuevo.items() if k not in fijos})
        st.success("Configuración guardada.")
        st.rerun()

    st.divider()
    st.markdown('<div class="pregunta">Pruebas</div>', unsafe_allow_html=True)

    if st.button("🧪 Enviar registro de prueba", key="btn_probar", width="stretch"):
        # La misma fila de prueba que se escribe en la hoja, con las
        # claves de la lista: así el JSON de prueba trae exactamente los
        # campos del contrato, y sirve para generar el esquema del flujo.
        nombres = [nombre for nombre, _, _ in columnas_sp()]
        prueba = dict(zip(nombres, fila_prueba_entrega(auth)))
        prueba["Fecha"] = date.today().strftime("%Y-%m-%d")
        ok, mensaje = enviar_a_webhook(prueba)
        (st.success if ok else st.error)(f"{'✅' if ok else '❌'} {mensaje}")

    if pendientes and st.button(f"🔄 Reenviar {pendientes} pendiente(s)",
                                key="btn_reenviar", width="stretch"):
        resultado = sincronizar_pendientes()
        if resultado["fallidos"]:
            st.error(f"{resultado['enviados']} enviados, {resultado['fallidos']} con error. "
                     f"{resultado['error']}")
        else:
            st.success(f"{resultado['enviados']} reporte(s) enviados.")
        st.rerun()

    st.divider()
    st.markdown('<div class="pregunta">Login contra SharePoint</div>', unsafe_allow_html=True)

    if config.get("url_login"):
        st.markdown(
            '<p class="hint">Configurado. Si el flujo no responde, la app valida con la '
            'copia local: quien ya entró una vez puede seguir entrando sin red.</p>',
            unsafe_allow_html=True)
    else:
        st.markdown(
            '<p class="hint">Sin configurar: el login usa solo la base local.</p>',
            unsafe_allow_html=True)

    c4, c5 = st.columns(2)
    with c4:
        usuario_prueba = st.text_input("Usuario", key="login_prueba_usuario")
    with c5:
        clave_prueba = st.text_input("Contraseña", type="password", key="login_prueba_clave")

    if st.button("🔑 Probar login contra el flujo", key="btn_probar_login", width="stretch"):
        if not usuario_prueba:
            st.error("Escriba un usuario.")
        else:
            estado, datos = validar_usuario_remoto(usuario_prueba.strip().lower(), clave_prueba)
            if estado == "ok":
                st.success(f"✅ Validado: {datos.get('nombre')} · {datos.get('empresa')} · "
                           f"{datos.get('perfil')} · obras: {datos.get('obras')}")
            elif estado == "rechazado":
                st.warning(f"El flujo respondió que las credenciales no son válidas. "
                           f"{(datos or {}).get('motivo', '')}")
            elif estado == "sin_flujo":
                st.info("No hay URL de login configurada.")
            else:
                st.error("El flujo no respondió (sin red, URL mala o error en el flujo).")

    with st.expander("🔐 Generar hash de una contraseña"):
        st.markdown(
            '<p class="hint">La lista guarda solo el SHA-256, nunca la contraseña. '
            'Genere el hash aquí y péguelo en la columna <code>ClaveHash</code>.</p>',
            unsafe_allow_html=True)
        texto = st.text_input("Contraseña a convertir", key="gen_hash")
        if texto:
            st.code(hash_clave(texto), language=None)

    st.divider()
    with st.expander("📋 Listas que debe tener SharePoint"):
        st.markdown(
            "**1. Reportes Diarios** — el flujo recibe un JSON con estas claves, que "
            "son **las mismas columnas que exporta el Excel**:\n\n"
            "| Columna | Tipo | Contenido |\n|---|---|---|\n" +
            "\n".join(f"| `{n}` | {t} | {d} |" for n, t, d in columnas_sp()) +
            "\n\nSi está activo el envío de fotos, llega además un arreglo `Fotos` con "
            "objetos `{nombre, contenido}` (imagen en base64) para que el flujo la cree "
            "en la biblioteca de evidencias.\n\n"
            "**2. Usuarios App** — fuente de verdad del login:\n\n"
            "| Columna | Tipo | Contenido |\n|---|---|---|\n" +
            "\n".join(f"| `{n}` | {t} | {d} |" for n, t, d in COLUMNAS_SP_USUARIOS) +
            "\n\n**3. Supervisores** — los del Nivel 1:\n\n"
            "| Columna | Tipo | Contenido |\n|---|---|---|\n" +
            "\n".join(f"| `{n}` | {t} | {d} |" for n, t, d in COLUMNAS_SP_SUPERVISORES) +
            "\n\n**4. Evidencias** — biblioteca de documentos, una carpeta por folio.\n\n"
            "El árbol de obras NO va en SharePoint: son más de 4.700 nodos que cambian "
            "muy poco. Vive en el código y se actualiza con ⚙️ Obras.")

    st.divider()
    bottom_nav()


# ============================================================
# 9. Enrutador
# ============================================================

def main():
    init_db()
    seed_db()
    # Reconstruye maquinaria y asignaciones desde la bitácora de Sheets.
    # Es lo que permite reiniciar el contenedor sin perder lo que
    # configuró Oficina Técnica ni las máquinas que dio de alta terreno.
    restaurar_desde_sheets()
    init_estado()

    ss = st.session_state

    if ss.auth is None:
        pantalla_login()
        return

    if ss.supervisor is None:
        if supervisor_directo(ss.auth):
            ss.supervisor = ss.auth["nombre"]
            if ss.paso in ("login", "supervisor"):
                ss.paso = "arbol"
        elif ss.paso != "supervisor":
            ss.paso = "supervisor"

    pantallas = {
        "supervisor": pantalla_supervisor,
        "arbol": pantalla_arbol,
        "formulario": pantalla_formulario,
        "guardado": pantalla_guardado,
        "reportes": pantalla_reportes,
        "obras": pantalla_obras,
        "asignar": pantalla_asignaciones,
        "envio": pantalla_envio,
    }
    inicio = pantalla_arbol if supervisor_directo(ss.auth) else pantalla_supervisor
    pantallas.get(ss.paso, inicio)()


if __name__ == "__main__":
    main()
