import io, os, re, tempfile, subprocess, shutil, csv, json, hashlib, copy, base64, sqlite3, xml.etree.ElementTree as ET
from difflib import SequenceMatcher
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import openpyxl
import fitz  # PyMuPDF
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, JsCode
from streamlit_sortables import sort_items

APP_TITLE = 'Herramienta de Extracción de Datos de Fichas Técnicas'
MAX_PDFS_PERMITIDOS = 5
DB_PATH = os.path.join(os.path.dirname(__file__), "cartera.db")

# --- GESTIÓN DE BASE DE DATOS SQLITE (CARTERA DE EXÁMENES) ---
def init_db():
    """Inicializa la base de datos SQLite si no existe."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cartera_examenes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre_examen TEXT UNIQUE,
            codigo_examen TEXT,
            indicaciones TEXT,
            tecnologia TEXT,
            entrega_normal TEXT,
            entrega_urgente TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cartera_metadata (
            clave TEXT PRIMARY KEY,
            valor TEXT
        )
    """)
    conn.commit()
    conn.close()

init_db()


def guardar_cartera_db(df):
    """Guarda la Cartera extrayendo ESTRICTAMENTE los nombres desde la Columna 6 (DESCRIPCION)."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    col_nombre_series = None
    
    for col in df.columns:
        if str(col).strip().upper() in ["DESCRIPCION", "DESCRIPCIÓN"]:
            col_nombre_series = df[col]
            break

    if col_nombre_series is None and len(df.columns) >= 6:
        col_nombre_series = df.iloc[:, 5]  # Columna 6 de Excel
    elif col_nombre_series is None and len(df.columns) > 0:
        col_nombre_series = df.iloc[:, 0]

    cols_exactas = {norm(str(c)): c for c in df.columns}

    def buscar_col(alias_list):
        for a in alias_list:
            if a in cols_exactas:
                return cols_exactas[a]
        for a in alias_list:
            for k_norm, k_real in cols_exactas.items():
                if a in k_norm:
                    return k_real
        return None

    col_codigo = buscar_col(["codigo examen", "codigo host", "codigo", "código"])
    col_indicaciones = buscar_col(["indicaciones", "preparacion paciente", "preparación paciente", "preparacion"])
    col_tecnologia = buscar_col(["tecnologia", "tecnología", "equipamiento", "equipo", "instrumento"])
    col_normal = buscar_col(["entrega normal", "plazo normal", "normal", "tat normal"])
    col_urgente = buscar_col(["entrega urgente", "plazo urgente", "urgente", "tat urgente"])

    cursor.execute("DELETE FROM cartera_examenes")

    filas_guardadas = 0
    for idx, nombre_raw in col_nombre_series.items():
        nombre = str(nombre_raw).strip()

        if not nombre or nombre.casefold() in ["nan", "none", "null", "descripcion", "descripción"] or nombre.isdigit():
            continue

        row = df.loc[idx]

        codigo = str(row.get(col_codigo, "")).strip() if col_codigo else ""
        indicaciones = str(row.get(col_indicaciones, "")).strip() if col_indicaciones else ""
        tecnologia = str(row.get(col_tecnologia, "")).strip() if col_tecnologia else ""
        normal = str(row.get(col_normal, "")).strip() if col_normal else ""
        urgente = str(row.get(col_urgente, "")).strip() if col_urgente else ""

        codigo = "" if codigo.casefold() in ["nan", "none"] else codigo
        indicaciones = "" if indicaciones.casefold() in ["nan", "none"] else indicaciones
        tecnologia = "" if tecnologia.casefold() in ["nan", "none"] else tecnologia
        normal = "" if normal.casefold() in ["nan", "none"] else normal
        urgente = "" if urgente.casefold() in ["nan", "none"] else urgente

        cursor.execute("""
            INSERT OR REPLACE INTO cartera_examenes (nombre_examen, codigo_examen, indicaciones, tecnologia, entrega_normal, entrega_urgente)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (nombre, codigo, indicaciones, tecnologia, normal, urgente))
        filas_guardadas += 1

    ZONA_HORARIA = ZoneInfo("America/Santiago")
    fecha_actual = datetime.now(ZONA_HORARIA).isoformat()
    cursor.execute("INSERT OR REPLACE INTO cartera_metadata (clave, valor) VALUES ('ultima_actualizacion', ?)", (fecha_actual,))

    conn.commit()
    conn.close()
    return True, f"Se importaron exitosamente {filas_guardadas} exámenes desde la Columna 6 (DESCRIPCION) a la Cartera."


def obtener_lista_cartera_db():
    """Retorna la lista ordenada de nombres/DESCRIPCION de exámenes almacenados en la Cartera."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT nombre_examen FROM cartera_examenes WHERE nombre_examen NOT GLOB '[0-9]*' ORDER BY nombre_examen ASC")
    rows = cursor.fetchall()
    conn.close()
    return [r[0] for r in rows if r[0]]


def buscar_top_coincidencias_cartera(nombre_buscar, top_n=5):
    """
    Busca los N mejores candidatos por similitud (Fuzzy Matching) en la Cartera para 'nombre_buscar'.
    Retorna una lista de tuplas: [(nombre_db, score_0_a_100), ...]
    """
    if not nombre_buscar:
        return []

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT nombre_examen FROM cartera_examenes")
    rows = cursor.fetchall()
    conn.close()

    if not rows:
        return []

    resultados = []
    norm_buscar = norm(nombre_buscar)

    for row in rows:
        nombre_db = row[0]
        norm_db = norm(nombre_db)

        if norm_buscar == norm_db:
            score = 1.0
        elif norm_buscar in norm_db or norm_db in norm_buscar:
            ratio = SequenceMatcher(None, norm_buscar, norm_db).ratio()
            score = max(0.88, ratio)
        else:
            score = SequenceMatcher(None, norm_buscar, norm_db).ratio()

        resultados.append((nombre_db, round(score * 100, 1)))

    resultados.sort(key=lambda x: x[1], reverse=True)
    return resultados[:top_n]


def obtener_info_cartera_por_nombre(nombre_exacto_cartera):
    """Consulta SQLite obteniendo los datos completos de un examen específico de la Cartera."""
    if not nombre_exacto_cartera:
        return None
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT codigo_examen, indicaciones, tecnologia, entrega_normal, entrega_urgente FROM cartera_examenes WHERE nombre_examen = ?", (nombre_exacto_cartera.strip(),))
    row = cursor.fetchone()
    conn.close()

    if row:
        codigo, indicaciones, tecnologia, normal, urgente = row
        plazo_partes = []
        if normal:
            plazo_partes.append(f"Entrega Normal: {normal}")
        if urgente:
            plazo_partes.append(f"Entrega Urgente: {urgente}")
        plazo_total = " | ".join(plazo_partes)

        return {
            "Código": codigo or "",
            "Preparación del paciente": indicaciones or "",
            "Equipamiento utilizado": tecnologia or "",
            "Plazo entrega resultados": plazo_total or ""
        }
    return None


def fecha_actualizacion_cartera():
    """Consulta la última fecha de actualización y valida los 6 meses (180 días)."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT valor FROM cartera_metadata WHERE clave = 'ultima_actualizacion'")
    row = cursor.fetchone()
    conn.close()

    if not row or not row[0]:
        return None, True, "No se ha subido ninguna Cartera de Exámenes a la base de datos."

    ZONA_HORARIA = ZoneInfo("America/Santiago")
    fecha_dt = datetime.fromisoformat(row[0])
    dias_transcurridos = (datetime.now(ZONA_HORARIA) - fecha_dt).days
    es_antiguo = dias_transcurridos >= 180

    msg_fecha = fecha_dt.strftime("%d/%m/%Y %H:%M")
    return fecha_dt, es_antiguo, msg_fecha


# --- GESTIÓN DE LABORATORIOS ---
LABORATORIOS_PREDETERMINADOS = [
    "Laboratorio CRS Cordillera Oriente",
]
LABORATORIOS_FILE = os.path.join(os.path.dirname(__file__), "laboratorios.json")


def cargar_laboratorios():
    """Carga la lista de laboratorios guardados."""
    try:
        with open(LABORATORIOS_FILE, "r", encoding="utf-8") as archivo:
            guardados = json.load(archivo)
        if isinstance(guardados, dict) and isinstance(guardados.get("laboratorios"), list):
            lista = guardados["laboratorios"]
        elif isinstance(guardados, list):
            lista = LABORATORIOS_PREDETERMINADOS + guardados
        else:
            lista = LABORATORIOS_PREDETERMINADOS
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        lista = LABORATORIOS_PREDETERMINADOS

    laboratorios = []
    vistos = set()
    for nombre in lista:
        nombre = str(nombre).strip()
        if nombre and nombre.casefold() not in vistos:
            laboratorios.append(nombre)
            vistos.add(nombre.casefold())
    return sorted(laboratorios, key=str.casefold)


def guardar_lista_laboratorios(laboratorios):
    """Guarda la lista completa de laboratorios."""
    limpia = []
    vistos = set()
    for nombre in laboratorios:
        nombre = str(nombre).strip()
        if nombre and nombre.casefold() not in vistos:
            limpia.append(nombre)
            vistos.add(nombre.casefold())
    try:
        with open(LABORATORIOS_FILE, "w", encoding="utf-8") as archivo:
            json.dump({"laboratorios": limpia}, archivo, ensure_ascii=False, indent=2)
        return True
    except OSError:
        return False


def guardar_laboratorio(nombre):
    """Añade un laboratorio nuevo."""
    nombre = nombre.strip()
    laboratorios = cargar_laboratorios()
    if not nombre or nombre.casefold() in {x.casefold() for x in laboratorios}:
        return False
    laboratorios.append(nombre)
    return guardar_lista_laboratorios(laboratorios)


# --- GESTIÓN DE ENCARGADOS (ACTUALIZADO POR / AUTORIZADO POR) ---
ENCARGADOS_PREDETERMINADOS = [
    "Ignacio Salgado Fuentes",
    "Lissette Valenzuela Gonzalez",
]
ENCARGADOS_FILE = os.path.join(os.path.dirname(__file__), "encargados.json")


def cargar_encargados(tipo="actualizado"):
    """Carga la lista de encargados para 'actualizado' o 'autorizado'."""
    try:
        with open(ENCARGADOS_FILE, "r", encoding="utf-8") as archivo:
            data = json.load(archivo)
        if isinstance(data, dict) and isinstance(data.get(tipo), list):
            lista = data[tipo]
        else:
            lista = ENCARGADOS_PREDETERMINADOS
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        lista = ENCARGADOS_PREDETERMINADOS

    encargados = []
    vistos = set()
    for nombre in lista:
        nombre = str(nombre).strip()
        if nombre and nombre.casefold() not in vistos:
            encargados.append(nombre)
            vistos.add(nombre.casefold())
    return sorted(encargados, key=str.casefold)


def guardar_encargados(tipo, lista):
    """Guarda la lista de encargados para el tipo especificado."""
    try:
        data = {}
        if os.path.exists(ENCARGADOS_FILE):
            try:
                with open(ENCARGADOS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                data = {}
        if not isinstance(data, dict):
            data = {}

        limpia = []
        vistos = set()
        for nombre in lista:
            nombre = str(nombre).strip()
            if nombre and nombre.casefold() not in vistos:
                limpia.append(nombre)
                vistos.add(nombre.casefold())

        data[tipo] = sorted(limpia, key=str.casefold)
        with open(ENCARGADOS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except OSError:
        return False


AREA_CONFIG = {
    "Microbiología": {"template": "plantilla_microbiologia.xlsx"},
    "Hematología y coagulación": {"template": "plantilla_hematologia.xlsx"},
    "Bioquímica": {"template": "plantilla_bioquimica.xlsx"},
    "Hormonas e inmunidad": {"template": None},
    "Derivados": {"template": None},
}
EXAMENES_FILE = os.path.join(os.path.dirname(__file__), "examenes_por_area.json")


def cargar_examenes_por_area():
    """Carga exámenes guardados; inicialmente usa los nombres de hojas de las plantillas."""
    guardados = None
    try:
        with open(EXAMENES_FILE, "r", encoding="utf-8") as f:
            guardados = json.load(f)
    except (OSError, json.JSONDecodeError):
        guardados = None
    resultado = {area: list(guardados.get(area, [])) for area in AREA_CONFIG} if isinstance(guardados, dict) else {a: [] for a in AREA_CONFIG}
    for area, config in AREA_CONFIG.items():
        ruta = os.path.join(os.path.dirname(__file__), config["template"] or "")
        if config["template"] and os.path.exists(ruta) and not (isinstance(guardados, dict) and area in guardados):
            try:
                wb = load_workbook(ruta, read_only=True, data_only=True)
                resultado[area] = list(wb.sheetnames)
                wb.close()
            except Exception:
                resultado[area] = []
    return {a: sorted(list(dict.fromkeys(str(x).strip() for x in xs if str(x).strip())), key=str.casefold) for a, xs in resultado.items()}


def guardar_examenes_por_area(examenes):
    try:
        with open(EXAMENES_FILE, "w", encoding="utf-8") as f:
            json.dump(examenes, f, ensure_ascii=False, indent=2)
        return True
    except OSError:
        return False


def nombre_examen_valido(nombre):
    nombre = str(nombre or "").strip()
    if not nombre:
        return False, "Escribe un nombre para el examen."
    if len(nombre) > 31:
        return False, "El nombre debe tener 31 caracteres o menos para coincidir exactamente con la hoja de Excel."
    if re.search(r'[\\/*?:\[\]]', nombre):
        return False, "El nombre no puede contener \\ / ? * : [ ]. Excel no permite esos caracteres en los nombres de hoja."
    return True, ""


def ajustar_ancho_columnas(ws):
    """Ajuste aproximado de ancho de columnas."""
    from openpyxl.utils import get_column_letter
    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        for row in range(1, ws.max_row + 1):
            value = ws.cell(row, col_idx).value
            if value is not None:
                max_len = max(max_len, max((len(line) for line in str(value).splitlines()), default=0))
        ws.column_dimensions[get_column_letter(col_idx)].width = max(12, min(max_len + 2, 65))


def convertir_subindices(texto):
    """Convierte cifras pegadas a letras en subíndice (CO2 → CO₂), evitando expresiones matemáticas."""
    tabla = str.maketrans("0123456789", "₀₁₂₃₄₅₆₇₈₉")
    operadores = "=<>≤≥±×÷≈+-*/^"
    resultado = []
    for linea in str(texto or "").splitlines():
        def reemplazar(match):
            inicio, fin = match.span()
            alrededor = linea[max(0, inicio - 3):min(len(linea), fin + 4)]
            if any(op in alrededor for op in operadores):
                return match.group(1)
            return match.group(1).translate(tabla)

        resultado.append(re.sub(r"(?<=[A-Za-z])(\d+)(?=[A-Za-z]|\b)", reemplazar, linea))
    return "\n".join(resultado)


def quitar_citas_superindices(texto):
    """Elimina citas bibliográficas de superíndices sin afectar cifras decimales."""
    if not texto:
        return ""

    s = str(texto)
    s = re.sub(r"\[\d+(?:\s*,\s*\d+)*\]", "", s)
    s = re.sub(r"\(\d+(?:\s*,\s*\d+)*\)", "", s)
    s = re.sub(r"(?<=[a-zA-ZáéíóúñÁÉÍÓÚÑ])\d+(?:\s*,\s*\d+)*(?=\s|[.,;:!?\n]|$)", "", s)
    s = re.sub(r"(?<=[a-zA-ZáéíóúñÁÉÍÓÚÑ][.,;:])\d+(?:\s*,\s*\d+)*(?=\s|[.,;:!?\n]|$)", "", s)
    s = re.sub(r"\s+([.,;:])", r"\1", s)
    s = re.sub(r"^\s*,\s*", "", s, flags=re.MULTILINE)
    s = re.sub(r"[ \t]+", " ", s)

    return s.strip()


def es_linea_basura_o_encabezado(s, examen_activo=None):
    """Detecta y filtra números de página, códigos de versión y encabezados repetidos."""
    s_clean = s.strip()
    if not s_clean:
        return True
    if re.fullmatch(r"(?:p[aá]gina\s*)?\d+\s*(?:de|[/])\s*\d+", s_clean, flags=re.I):
        return True
    if re.fullmatch(r"\d{7,}[vV]?\d*(?:\.\d+)*\.?", s_clean):
        return True
    if re.fullmatch(r"\d{4}-\d{2},\s*v\s*[\d\.]+\s*.*", s_clean, flags=re.I):
        return True
    if s_clean.casefold() in ["cobas®", "roche diagnostics", "español", "content"]:
        return True
    if examen_activo:
        term_exam = str(examen_activo).strip().casefold()
        if s_clean.casefold() == term_exam or norm(s_clean) in norm(term_exam):
            return True
    return False


def limpiar_valor_excel(valor, examen_activo=None):
    """Quita espacios accidentales, líneas de ruido y citas bibliográficas conservando valores válidos."""
    if valor is None:
        return ""

    valor_str = str(valor).strip()
    if not valor_str:
        return ""

    if examen_activo and valor_str.casefold() == str(examen_activo).strip().casefold():
        return valor_str

    valor_sin_citas = quitar_citas_superindices(valor)
    lineas = [re.sub(r"[ \t]+", " ", linea).strip() for linea in valor_sin_citas.replace("\r", "").split("\n")]
    salida = []
    blancos = 0
    for linea in lineas:
        if not linea:
            blancos += 1
            if blancos == 1 and salida:
                salida.append("")
            continue

        if es_linea_basura_o_encabezado(linea, examen_activo):
            if examen_activo and linea.casefold() == str(examen_activo).strip().casefold() and len(lineas) == 1:
                salida.append(linea)
                continue
            blancos += 1
            if blancos == 1 and salida:
                salida.append("")
            continue

        blancos = 0
        if re.fullmatch(r"https?://\S+|www\.\S+", linea, flags=re.I):
            continue
        salida.append(linea)
    return convertir_subindices("\n".join(salida).strip())


# --- ALIAS ESPECÍFICOS POR CADA ÁREA DEL LABORATORIO ---

FIELD_ALIASES_BASE = {
    'Código': [],
    'Nombre LIS': [],
    'Laboratorio de procesamiento': [],
    'Recepción de muestras': ['recepcion de muestras', 'recepción de muestras', 'recepcion'],
    'Procesamiento de muestras': ['procesamiento de muestras', 'procesamiento de muetsras', 'procesamiento'],
    'Plazo entrega resultados': ['tiempo de respuesta', 'plazo entrega resultados', 'plazo de entrega', 'tat', 'tiempo procesamiento'],
    'Preparación del paciente': ['preparacion paciente', 'preparación paciente', 'preparacion del paciente', 'preparación del paciente'],
    'Procedimiento de toma de muestras': ['Procedimiento', 'Procedimiento de toma de muestras', 'Procedimiento de examen', 'Procedimiento del examen'],
    'Muestra requerida y recipiente': ['tipo de muestra', 'muestra requerida', 'muestra', 'tipo y volumen de muestra', 'material biologico', 'material biológico', 'volumen requerido', 'obtencion y preparacion de las muestras', 'obtención y preparación de las muestras'],
    'Rotulación': ['Rotulación', 'Rotulacion'],
    'Estabilidad de la muestra': ['estabilidad muestra', 'estabilidad de la muestra', 'Conservación y estabilidad', 'conservacion y estabilidad'],
    'Condiciones de envío al Laboratorio': ['Condiciones de envío al Laboratorio', 'temperatura de transporte', 'condiciones de envio', 'condiciones de envío', 'envio al laboratorio', 'envío al laboratorio', 'transporte'],
    'Criterios de rechazo': ['Rechazo', 'Criterios de rechazo', 'Criterios rechazo', 'Criterios del rechazo'],
    'Método de procesamiento': ['metodo utilizado', 'método utilizado', 'metodo de procesamiento', 'método de procesamiento', 'metodologia', 'metodología', 'metodo', 'principio del test', 'principio', 'principio de la prueba', 'principio del ensayo'],
    'Test incluidos': ['test incluidos', 'pruebas incluidas', 'examenes incluidos', 'exámenes incluidos'],
    'Valores de referencia': ['valor de referencia', 'valores de referencia', 'valores normales', 'rango de referencia'],
    'Valor crítico': ['valor critico', 'valor crítico', 'valores criticos', 'valores críticos'],
    'Límite de detección': ['limite de deteccion', 'límite de detección', 'limite deteccion', 'límite deteccion', 'límite de cuantificación', 'limite de cuantificacion', 'loq'],
    'Equipamiento utilizado': ['equipamiento utilizado', 'equipo utilizado', 'instrumentacion', 'instrumentación', 'equipo', 'system', 'sistema', 'informacion del sistema', 'información del sistema'],
    'Información clínica': ['caracteristicas', 'características', 'utilidad clinica', 'utilidad clínica', 'informacion clinica', 'información clínica'],
    'Referencias': ['referencias', 'bibliografia', 'bibliografía', 'fuentes'],
    'Actualizado por': [],
    'Fecha actualización': [],
    'Autorizado por': [],
}

FIELD_ALIASES_MICROBIOLOGIA = {
    'Código': [],
    'Nombre LIS': [],
    'Laboratorio de procesamiento': [],
    'Recepción de muestras': ['recepcion de muestras', 'recepción de muestras', 'recepcion'],
    'Procesamiento de muestras': ['procesamiento de muestras', 'procesamiento', 'tiempo siembra'],
    'Plazo entrega resultados': ['tiempo de respuesta', 'plazo entrega resultados', 'plazo de entrega', 'tat', 'tiempo incubacion', 'dias cultivo'],
    'Preparación del paciente': ['preparacion paciente', 'preparación paciente', 'preparacion del paciente'],
    'Procedimiento de toma de muestras': ['Procedimiento', 'Procedimiento de toma de muestras', 'obtencion muestra microbiologica', 'tecnica hisopado', 'toma de muestra'],
    'Muestra requerida y recipiente': ['tipo de muestra', 'muestra requerida', 'muestra', 'medio de transporte', 'frasco estéril', 'hisopo', 'recipiente'],
    'Rotulación': ['Rotulación', 'Rotulacion'],
    'Estabilidad de la muestra': ['estabilidad muestra', 'estabilidad de la muestra', 'Conservación y estabilidad'],
    'Condiciones de envío al Laboratorio': ['Condiciones de envío al Laboratorio', 'temperatura de transporte', 'condiciones de envio', 'transporte muestra'],
    'Criterios de rechazo': ['Rechazo', 'Criterios de rechazo', 'muestra contaminada', 'hisopo seco'],
    'Método de procesamiento': ['metodo utilizado', 'método utilizado', 'metodo de procesamiento', 'metodologia', 'tintura gram', 'siembra en agar', 'cultivo', 'antibiograma', 'pcr microbiologico'],
    'Test incluidos': ['test incluidos', 'pruebas incluidas', 'identificacion bacteriana', 'sensibilidad antibiotica'],
    'Valores de referencia': ['valor de referencia', 'valores de referencia', 'flora normal', 'sin desarrollo bacteriano', 'resultado esperado'],
    'Valor crítico': ['valor critico', 'valor crítico', 'patógeno crítico', 'alerta microbiológica'],
    'Límite de detección': ['limite de deteccion', 'límite de detección', 'ufc/ml min', 'sensibilidad analítica'],
    'Equipamiento utilizado': ['equipamiento utilizado', 'equipo utilizado', 'estufa incubadora', 'autoclave', 'vitek', 'maldi-tof', 'sistema automatizado de hemocultivos'],
    'Información clínica': ['caracteristicas', 'características', 'utilidad clinica', 'informacion clinica', 'asociación infecciosa'],
    'Referencias': ['referencias', 'bibliografia', 'clsi', 'eucast'],
    'Actualizado por': [],
    'Fecha actualización': [],
    'Autorizado por': [],
}

FIELD_ALIASES_HEMATOLOGIA = {
    'Código': [],
    'Nombre LIS': [],
    'Laboratorio de procesamiento': [],
    'Recepción de muestras': ['recepcion de muestras', 'recepción de muestras', 'recepcion'],
    'Procesamiento de muestras': ['procesamiento de muestras', 'procesamiento', 'homogeneizacion'],
    'Plazo entrega resultados': ['tiempo de respuesta', 'plazo entrega resultados', 'plazo de entrega', 'tat', 'tiempo procesamiento'],
    'Preparación del paciente': ['preparacion paciente', 'preparación paciente', 'ayuno hematologia'],
    'Procedimiento de toma de muestras': ['Procedimiento', 'Procedimiento de toma de muestras', 'puncion venosa', 'anticoagulante edta', 'tubo citrato'],
    'Muestra requerida y recipiente': ['tipo de muestra', 'muestra requerida', 'sangre total', 'plasma citratado', 'tubo tapa morada', 'tubo tapa azul'],
    'Rotulación': ['Rotulación', 'Rotulacion'],
    'Estabilidad de la muestra': ['Estabilidad:'],
    'Condiciones de envío al Laboratorio': ['Condiciones de envío al Laboratorio', 'temperatura de transporte', 'envio temperatura ambiente', 'no congelar sangre total'],
    'Criterios de rechazo': ['Rechazo', 'Criterios de rechazo', 'muestra coagulada', 'muestra hemolisada', 'volumen insuficiente', 'relación anticoagulante'],
    'Método de procesamiento': ['Principio del test','metodo utilizado', 'método utilizado', 'metodo de procesamiento', 'metodologia'],
    'Test incluidos': ['test incluidos', 'pruebas incluidas', 'fórmula leucocitaria', 'indices hematimetricos'],
    'Valores de referencia': ['valores teóricos'],
    'Valor crítico': ['valor critico', 'valor crítico', 'plaquetopenia severa', 'inr crítico', 'hemoglobina crítica'],
    'Límite de detección': ['limite de detección:', 'límite de detección', 'rango de linealidad'],
    'Equipamiento utilizado': ['equipamiento utilizado', 'equipo utilizado', 'contador hematologico', 'coagulometro', 'sysmex', 'stago'],
    'Información clínica': ['caracteristicas', 'características'],
    'Referencias': ['referencias', 'bibliografia', 'fuentes'],
    'Actualizado por': [],
    'Fecha actualización': [],
    'Autorizado por': [],
}

FIELD_ALIASES_BIOQUIMICA = {
    'Código': [],
    'Nombre LIS': [],
    'Laboratorio de procesamiento': [],
    'Recepción de muestras': ['recepcion de muestras', 'recepción de muestras', 'recepcion'],
    'Procesamiento de muestras': ['procesamiento de muestras', 'procesamiento', 'centrifugacion', 'tiempo suero'],
    'Plazo entrega resultados': ['tiempo de respuesta', 'plazo entrega resultados', 'plazo de entrega', 'tat'],
    'Preparación del paciente': ['preparacion paciente', 'preparación paciente', 'ayuno de 8 a 12 horas', 'restricción dietética'],
    'Procedimiento de toma de muestras': ['Procedimiento', 'Procedimiento de toma de muestras', 'puncion venosa', 'tubo suero gel'],
    'Muestra requerida y recipiente': ['tipo de muestra', 'muestra requerida', 'suero', 'plasma', 'orina de 24 horas', 'tubo tapa roja', 'tubo tapa amarilla'],
    'Rotulación': ['Rotulación', 'Rotulacion'],
    'Estabilidad de la muestra': ['estabilidad muestra', 'estabilidad de la muestra', 'estabilidad del suero', 'conservacion refrigerada'],
    'Condiciones de envío al Laboratorio': ['Condiciones de envío al Laboratorio', 'temperatura de transporte', 'suero separado', 'transporte refrigerado'],
    'Criterios de rechazo': ['Rechazo', 'Criterios de rechazo', 'suero hemolisado', 'suero icterico', 'suero lipemico'],
    'Método de procesamiento': ['metodo utilizado', 'método utilizado', 'metodo de procesamiento', 'metodologia', 'espectrofotometria', 'enzimatico', 'colorimetrico', 'electrodo ion selectivo', 'ise'],
    'Test incluidos': ['test incluidos', 'pruebas incluidas', 'perfil hepatico', 'perfil lipidico', 'perfil renal'],
    'Valores de referencia': ['valor de referencia', 'valores de referencia', 'rangos de referencia', 'valores normales'],
    'Valor crítico': ['valor critico', 'valor crítico', 'potasio crítico', 'glicemia crítica', 'sodio crítico'],
    'Límite de detección': ['limite de deteccion', 'límite de detección', 'linealidad analítica', 'loq'],
    'Equipamiento utilizado': ['equipamiento utilizado', 'equipo utilizado', 'analizador bioquimico', 'cobas', 'architect'],
    'Información clínica': ['caracteristicas', 'características', 'utilidad clinica', 'evaluación metabólica', 'función renal'],
    'Referencias': ['referencias', 'bibliografia', 'fuentes'],
    'Actualizado por': [],
    'Fecha actualización': [],
    'Autorizado por': [],
}

FIELD_ALIASES_HORMONAS = {
    'Código': [],
    'Nombre LIS': [],
    'Laboratorio de procesamiento': [],
    'Recepción de muestras': ['recepcion de muestras', 'recepción de muestras', 'recepcion'],
    'Procesamiento de muestras': ['procesamiento de muestras', 'procesamiento'],
    'Plazo entrega resultados': ['tiempo de respuesta', 'plazo entrega resultados', 'plazo de entrega', 'tat'],
    'Preparación del paciente': ['preparacion paciente', 'preparación paciente', 'hora de toma de muestra', 'ritmo circadiano', 'ciclo menstrual'],
    'Procedimiento de toma de muestras': ['Procedimiento', 'Procedimiento de toma de muestras', 'puncion venosa'],
    'Muestra requerida y recipiente': ['tipo de muestra', 'muestra requerida', 'suero', 'plasma edta', 'tubo suero con separador'],
    'Rotulación': ['Rotulación', 'Rotulacion'],
    'Estabilidad de la muestra': ['estabilidad muestra', 'estabilidad de la muestra', 'estabilidad hormonas', 'congelación suero'],
    'Condiciones de envío al Laboratorio': ['Condiciones de envío al Laboratorio', 'temperatura de transporte', 'envio congelado', 'cadena de frio'],
    'Criterios de rechazo': ['Rechazo', 'Criterios de rechazo', 'hemolisis', 'descongelamiento previo'],
    'Método de procesamiento': ['metodo utilizado', 'método utilizado', 'metodo de procesamiento', 'metodologia', 'quimioluminiscencia', 'eclia', 'elisa', 'ria', 'electroquimioluminiscencia'],
    'Test incluidos': ['test incluidos', 'pruebas incluidas', 'perfil tiroideo', 'panel hormonal'],
    'Valores de referencia': ['valor de referencia', 'valores de referencia', 'valores por trimestre', 'valores por fase menstrual'],
    'Valor crítico': ['valor critico', 'valor crítico', 'alerta endocrina'],
    'Límite de detección': ['limite de deteccion', 'límite de detección', 'sensibilidad funcional'],
    'Equipamiento utilizado': ['equipamiento utilizado', 'equipo utilizado', 'analizador de inmunoensayo', 'cobas e', 'centaur', 'immulite'],
    'Información clínica': ['caracteristicas', 'características', 'utilidad clinica', 'evaluación endocrina'],
    'Referencias': ['referencias', 'bibliografia', 'fuentes'],
    'Actualizado por': [],
    'Fecha actualización': [],
    'Autorizado por': [],
}

FIELD_ALIASES_DERIVADOS = {
    'Código': [],
    'Nombre LIS': [],
    'Laboratorio de procesamiento': [],
    'Recepción de muestras': ['recepcion de muestras', 'recepción de muestras', 'recepcion'],
    'Procesamiento de muestras': ['procesamiento de muestras', 'procesamiento', 'preparacion envio'],
    'Plazo entrega resultados': ['tiempo de respuesta', 'plazo entrega resultados', 'plazo de entrega', 'tat derivación', 'dias hábiles'],
    'Preparación del paciente': ['preparacion paciente', 'preparación paciente', 'indicaciones laboratorio externo'],
    'Procedimiento de toma de muestras': ['Procedimiento', 'Procedimiento de toma de muestras', 'protocolo de derivación'],
    'Muestra requerida y recipiente': ['tipo de muestra', 'muestra requerida', 'alícuota', 'volumen minimo enviado'],
    'Rotulación': ['Rotulación', 'Rotulacion', 'etiqueta derivación'],
    'Estabilidad de la muestra': ['estabilidad muestra', 'estabilidad de la muestra', 'estabilidad en viaje'],
    'Condiciones de envío al Laboratorio': ['Condiciones de envío al Laboratorio', 'temperatura de transporte', 'transporte refrigerado', 'transporte congelado', 'nieve seca'],
    'Criterios de rechazo': ['Rechazo', 'Criterios de rechazo', 'incumplimiento cadena de frio', 'volumen insuficiente'],
    'Método de procesamiento': ['metodo utilizado', 'método utilizado', 'metodo de procesamiento', 'metodologia', 'espectrometria de masas', 'pcr', 'secuenciacion'],
    'Test incluidos': ['test incluidos', 'pruebas incluidas', 'panel genético', 'screening de derivación'],
    'Valores de referencia': ['valor de referencia', 'valores de referencia', 'rangos laboratorio de referencia'],
    'Valor crítico': ['valor critico', 'valor crítico'],
    'Límite de detección': ['limite de deteccion', 'límite de detección'],
    'Equipamiento utilizado': ['equipamiento utilizado', 'equipo utilizado', 'tecnología laboratorio de origen'],
    'Información clínica': ['caracteristicas', 'características', 'utilidad clinica', 'diagnóstico genético / derivado'],
    'Referencias': ['referencias', 'bibliografia', 'fuentes'],
    'Actualizado por': [],
    'Fecha actualización': [],
    'Autorizado por': [],
}

FIELD_ALIASES_POR_AREA = {
    "Microbiología": FIELD_ALIASES_MICROBIOLOGIA,
    "Hematología": FIELD_ALIASES_HEMATOLOGIA,
    "Hematología y coagulación": FIELD_ALIASES_HEMATOLOGIA,
    "Bioquímica": FIELD_ALIASES_BIOQUIMICA,
    "Hormonas e inmunidad": FIELD_ALIASES_HORMONAS,
    "Derivados": FIELD_ALIASES_DERIVADOS,
}

SECCIONES_DOCUMENTO_MAESTRAS = [
    'uso previsto', 'caracteristicas', 'características',
    'reactivos - soluciones de trabajo', 'reactivos', 'medidas de precaucion y advertencias',
    'medidas de precaucion', 'medidas de prevención', 'preparacion de los reactivos',
    'preparación de los reactivos', 'obtencion y preparacion de las muestras',
    'obtención y preparación de las muestras',
    'material suministrado', 'material requerido adicionalmente', 'material requerido',
    'realizacion del test', 'realización del test', 'calibracion', 'calibración',
    'control de calidad', 'limitaciones del analisis - interferencias', 'limitaciones del analisis',
    'limitaciones del análisis', 'limites e intervalos', 'límites e intervalos', 'valores teoricos',
    'valores teóricos', 'datos especificos de funcionamiento', 'datos específicos de funcionamiento',
    'precision', 'precisión', 'comparacion de metodos', 'comparación de métodos',
    'limite de cuantificacion', 'límite de cuantificación', 'loq',
    'referencias bibliograficas', 'referencias bibliográficas', 'simbolos', 'símbolos'
]


def get_field_aliases(area=None):
    if area and area in FIELD_ALIASES_POR_AREA:
        return FIELD_ALIASES_POR_AREA[area]
    return FIELD_ALIASES_BASE


def norm(s):
    import unicodedata
    s = str(s or '').strip().lower()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    s = re.sub(r'\s+', ' ', s)
    return s


def extract_pdf_text(uploaded_bytes, filename, examen_activo=None):
    try:
        doc = fitz.open(stream=uploaded_bytes, filetype="pdf")
        pages_text = []

        for page in doc:
            page_height = page.rect.height
            page_width = page.rect.width
            raw_blocks = page.get_text("blocks")
            valid_blocks = []

            for b in raw_blocks:
                if b[6] != 0:
                    continue

                x0, y0, x1, y1, text = b[0], b[1], b[2], b[3], b[4].strip()

                es_margen_superior = y0 < 50
                es_margen_inferior = y1 > (page_height - 50)

                if es_margen_superior or es_margen_inferior:
                    lineas_bloque = [l.strip() for l in text.split('\n') if l.strip()]
                    if all(es_linea_basura_o_encabezado(l, examen_activo) for l in lineas_bloque):
                        continue

                if text and not es_linea_basura_o_encabezado(text, examen_activo):
                    valid_blocks.append((x0, y0, x1, y1, text))

            def get_col_key(block):
                bx0, by0, bx1, by1, _ = block
                if (bx1 - bx0) > page_width * 0.6:
                    col = 0
                elif bx0 < page_width * 0.48:
                    col = 1
                else:
                    col = 2
                return (col, by0)

            valid_blocks.sort(key=get_col_key)

            page_lines = [b[4] for b in valid_blocks]
            pages_text.append("\n".join(page_lines))

        text = "\n\n".join(pages_text)

        if len(re.sub(r"\s", "", text)) >= 50:
            return text, "PyMuPDF (Columna por Columna)"
    except Exception:
        pass

    try:
        from rapidocr_onnxruntime import RapidOCR
        import numpy as np
        from PIL import Image

        engine = RapidOCR()
        doc = fitz.open(stream=uploaded_bytes, filetype="pdf")
        parts = []

        for page in doc:
            pix = page.get_pixmap(dpi=200)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            result, _ = engine(np.array(img))
            if result:
                lineas_pagina = [line[1] for line in result]
                parts.append("\n".join(lineas_pagina))

        texto_ocr = "\n\n".join(parts)
        if len(re.sub(r"\s", "", texto_ocr)) < 10:
            raise RuntimeError("El documento es una imagen o escaneo de muy baja resolución/calidad.")

        return texto_ocr, "PyMuPDF (RapidOCR)"

    except ImportError:
        raise RuntimeError("La librería 'rapidocr-onnxruntime' no está instalada.")
    except Exception as err_ocr:
        raise RuntimeError(f"Fallo en el motor OCR: {err_ocr}")


def clean_value(v):
    v = re.sub(r'[ \t]+', ' ', v).strip()
    return v.strip(' :\t')


def parse_labeled_text(text, field_aliases=None, examen_activo=None):
    if field_aliases is None:
        field_aliases = FIELD_ALIASES_BASE

    lines = [x.rstrip() for x in text.replace('\r', '').split('\n')]
    found = {}
    current = None

    label_patterns = []
    for field, aliases in field_aliases.items():
        if field == "Código":
            continue
        for a in aliases:
            label_patterns.append((field, a))
    label_patterns.sort(key=lambda x: len(x[1]), reverse=True)

    norm_maestras = [norm(sec) for sec in SECCIONES_DOCUMENTO_MAESTRAS]

    for raw in lines:
        s = raw.strip()
        if not s or es_linea_basura_o_encabezado(s, examen_activo):
            continue

        ns = norm(s)
        matched_field = None
        remainder = ''

        for field, alias in label_patterns:
            na = norm(alias)
            if ns == na:
                matched_field = field
                remainder = ''
                break
            if ns.startswith(na + ':') or ns.startswith(na + ' -'):
                matched_field = field
                idx = len(alias)
                remainder = s[idx:].lstrip(' :-').strip()
                break

        if matched_field:
            current = matched_field
            if remainder:
                found[current] = (found.get(current, '') + ('\n' if found.get(current) else '') + clean_value(remainder)).strip()
            else:
                found.setdefault(current, '')
            continue

        is_main_doc_section = False
        for sec in norm_maestras:
            if ns == sec or ns.startswith(sec + ':'):
                is_main_doc_section = True
                break

        if is_main_doc_section:
            current = None
            continue

        if current:
            found[current] = (found.get(current, '') + ('\n' if found.get(current) else '') + s).strip()

    for k, v in list(found.items()):
        found[k] = limpiar_valor_excel(v, examen_activo)

    return found


def infer_name_code(text, data):
    equipos_encontrados = re.findall(r'\bcobas\s+(?:[a-zA-Z]\s*)?\d{3,4}\b|\bcobas\s+(?:pure|pro)\b', text, flags=re.I)
    if equipos_encontrados and not data.get('Equipamiento utilizado'):
        equipos_limpios = []
        for eq in equipos_encontrados:
            eq_norm = re.sub(r'\s+', ' ', eq).strip()
            if eq_norm.casefold() not in [x.casefold() for x in equipos_limpios]:
                equipos_limpios.append(eq_norm)
        data['Equipamiento utilizado'] = "\n".join(equipos_limpios)

    if not data.get('Método de procesamiento'):
        patrones_metodo = [
            r'\bm[eé]todo\s+de\s+[a-zA-ZáéíóúñÁÉÍÓÚÑ\s]{3,35}(?=\.|\n|,|$)',
            r'\b(?:coagulom[eé]tric[oa]|turbidim[eé]tric[oa]|fotom[eé]tric[oa]|enzim[aá]tic[oa]|espectrofotom[eé]tric[oa]|quimioluminiscencia|ECLIA|ELISA|RIA|inmunoensayo|nefelom[eé]tric[oa]|cromog[eé]nic[oa]|aglutinaci[oó]n|impedancia|citometr[ií]a\s+de\s+flujo|colorim[eé]tric[oa]|punto\s+final|cin[eé]tic[oa])\b'
        ]
        metodos_encontrados = []
        for pat in patrones_metodo:
            coincidencias = re.findall(pat, text, flags=re.I)
            for c in coincidencias:
                c_clean = re.sub(r'\s+', ' ', c).strip()
                if c_clean.casefold() not in [x.casefold() for x in metodos_encontrados]:
                    metodos_encontrados.append(c_clean)
        if metodos_encontrados:
            data['Método de procesamiento'] = "\n".join(metodos_encontrados[:3])

    return data


def apply_to_sheet(ws, data):
    rows = {}
    for r in range(1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or '').strip()
        if label:
            rows[norm(label)] = r
    updates = []
    for field, value in data.items():
        if not value:
            continue
        target = field
        r = rows.get(norm(target))
        if r is None:
            for existing, rr in rows.items():
                if existing == norm(target) or norm(target) in existing or existing in norm(target):
                    r = rr
                    break
        if r:
            old = ws.cell(r, 2).value
            ws.cell(r, 2).value = value
            ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical='top')
            updates.append((field, old, value, r))
    return updates


def crear_hoja_examen(wb, nombre_examen, field_aliases=None):
    if field_aliases is None:
        field_aliases = FIELD_ALIASES_BASE
    ok, mensaje = nombre_examen_valido(nombre_examen)
    if not ok:
        raise ValueError(mensaje)
    if nombre_examen in wb.sheetnames:
        return wb[nombre_examen]
    ws = wb.create_sheet(title=nombre_examen)
    for fila, campo in enumerate(field_aliases, start=1):
        ws.cell(row=fila, column=1, value=campo)
        ws.cell(row=fila, column=2, value=nombre_examen if norm(campo) == norm("Nombre LIS") else "")
    return ws


def process(uploaded_bytes, filename, template_bytes, centro='', nombre_examen='', field_aliases=None,
            actualizado_por='', autorizado_por='', datos_sobrescribir=None):
    if field_aliases is None:
        field_aliases = FIELD_ALIASES_BASE
    if not nombre_examen:
        raise ValueError('Selecciona un examen desde el menú del área antes de analizar el PDF.')
    ok, mensaje = nombre_examen_valido(nombre_examen)
    if not ok:
        raise ValueError(mensaje)

    text, method = extract_pdf_text(uploaded_bytes, filename, nombre_examen)
    data = infer_name_code(text, parse_labeled_text(text, field_aliases, nombre_examen))

    if datos_sobrescribir:
        for campo, val in datos_sobrescribir.items():
            if val:
                data[campo] = val

    ZONA_HORARIA = ZoneInfo("America/Santiago")
    fecha_actual = datetime.now(ZONA_HORARIA).strftime('%d/%m/%Y %H:%M')

    data['Nombre LIS'] = nombre_examen.strip()
    data['Laboratorio de procesamiento'] = (centro or '').strip()
    data['Actualizado por'] = (actualizado_por or '').strip()
    data['Autorizado por'] = (autorizado_por or '').strip()
    data['Fecha actualización'] = fecha_actual

    wb = load_workbook(io.BytesIO(template_bytes))

    if nombre_examen in wb.sheetnames:
        ws = wb[nombre_examen]
        score, reason = 1.0, 'Examen seleccionado por el usuario'
    elif len(wb.sheetnames) == 1 and norm(wb.sheetnames[0]) in ['sheet', 'sheet1', 'hoja1', 'hoja 1']:
        ws = wb.active
        ws.title = nombre_examen.strip()
        for fila, campo in enumerate(field_aliases, start=1):
            ws.cell(row=fila, column=1, value=campo)
            ws.cell(row=fila, column=2, value=nombre_examen.strip() if norm(campo) == norm("Nombre LIS") else "")
        score, reason = 1.0, 'Ficha renombrada desde el examen seleccionado'
    else:
        ws = crear_hoja_examen(wb, nombre_examen, field_aliases)
        score, reason = 1.0, 'Ficha nueva creada desde el examen seleccionado'

    updates = apply_to_sheet(ws, data)

    filas = {norm(str(ws.cell(r, 1).value or '')): r for r in range(1, ws.max_row + 1)}
    for campo, val in [
        ('Nombre LIS', nombre_examen.strip()),
        ('Laboratorio de procesamiento', (centro or '').strip()),
        ('Actualizado por', (actualizado_por or '').strip()),
        ('Autorizado por', (autorizado_por or '').strip()),
        ('Fecha actualización', fecha_actual),
    ]:
        if norm(campo) in filas:
            ws.cell(filas[norm(campo)], 2).value = val

    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and isinstance(cell.value, str):
                cell.value = limpiar_valor_excel(cell.value, nombre_examen)
            cell.alignment = Alignment(vertical='top', wrap_text=True)
    ajustar_ancho_columnas(ws)

    wb.active = ws

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue(), data, nombre_examen, score, reason, method, text


def aplicar_ediciones_a_excel(excel_bytes, edited_df, nombre_examen):
    wb = load_workbook(io.BytesIO(excel_bytes))

    if nombre_examen in wb.sheetnames:
        ws = wb[nombre_examen]
        filas = {}
        for r in range(1, ws.max_row + 1):
            val = str(ws.cell(r, 1).value or '').strip()
            if val:
                filas[norm(val)] = r

        for _, row_data in edited_df.iterrows():
            campo_norm = norm(row_data["Campo"])
            val_editado = str(row_data["Información extraída"] or '').strip()
            if campo_norm in filas:
                r_idx = filas[campo_norm]
                ws.cell(r_idx, 2).value = val_editado
                ws.cell(r_idx, 2).alignment = Alignment(vertical='top', wrap_text=True)

    out_buf = io.BytesIO()
    wb.save(out_buf)
    return out_buf.getvalue()


def comparar_plantillas_excel(bytes_antiguo, bytes_nuevo, nombre_examen):
    def extraer_datos_fichas(raw_bytes):
        if not raw_bytes:
            return {}, 0, []
        try:
            wb = load_workbook(io.BytesIO(raw_bytes), data_only=True)
            hojas = wb.sheetnames
            datos_examen = {}
            if nombre_examen in wb.sheetnames:
                ws = wb[nombre_examen]
                for r in range(1, ws.max_row + 1):
                    campo = str(ws.cell(r, 1).value or '').strip()
                    valor = str(ws.cell(r, 2).value or '').strip()
                    if campo:
                        datos_examen[campo] = valor
            return datos_examen, len(hojas), hojas
        except Exception:
            return {}, 0, []

    datos_ant, total_hojas_ant, hojas_ant = extraer_datos_fichas(bytes_antiguo)
    datos_nue, total_hojas_nue, hojas_nue = extraer_datos_fichas(bytes_nuevo)

    campos_completados_ant = sum(1 for v in datos_ant.values() if v)
    campos_completados_nue = sum(1 for v in datos_nue.values() if v)

    diferencias = []
    todos_campos = list(dict.fromkeys(list(datos_ant.keys()) + list(datos_nue.keys())))

    for campo in todos_campos:
        val_ant = datos_ant.get(campo, "")
        val_nue = datos_nue.get(campo, "")

        if val_ant != val_nue:
            if not val_ant and val_nue:
                tipo = "➕ Campo nuevo incorporado"
            elif val_ant and not val_nue:
                tipo = "➖ Campo eliminado / borrado"
            else:
                tipo = "✏️ Información modificada"

            diferencias.append({
                "Campo": campo,
                "Valor anterior": val_ant or "(Vacío)",
                "Valor nuevo": val_nue or "(Vacío)",
                "Tipo de cambio": tipo
            })

    es_nueva_hoja = (nombre_examen not in hojas_ant) and (nombre_examen in hojas_nue)

    return {
        "hojas_antes": total_hojas_ant,
        "hojas_despues": total_hojas_nue,
        "campos_completados_antes": campos_completados_ant,
        "campos_completados_despues": campos_completados_nue,
        "es_nueva_hoja": es_nueva_hoja,
        "diferencias": diferencias
    }


# --- ACCESO BÁSICO ---
DEFAULT_USERNAME = "CRSCODUOC"
DEFAULT_PASSWORD = "CRSCODUOC"


def check_login():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if st.session_state["authenticated"]:
        return True

    st.session_state.pop("modal_area_activa", None)
    st.session_state["mostrar_modal_examenes"] = False

    _, col_login, _ = st.columns([1, 1.2, 1])

    with col_login:
        st.markdown("<h1 style='text-align: center;'>Acceso a herramienta</h1>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #6c757d;'>Ingresa tus credenciales para continuar.</p>",
                    unsafe_allow_html=True)
        st.write("")

        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("Usuario")
            password = st.text_input("Contraseña", type="password")
            submitted = st.form_submit_button("Iniciar sesión", type="primary", use_container_width=True)

        if submitted:
            try:
                expected_user = st.secrets.get("APP_USERNAME", DEFAULT_USERNAME)
                expected_password = st.secrets.get("APP_PASSWORD", DEFAULT_PASSWORD)
            except Exception:
                expected_user, expected_password = DEFAULT_USERNAME, DEFAULT_PASSWORD

            if (hashlib.sha256(username.encode("utf-8")).digest() ==
                    hashlib.sha256(expected_user.encode("utf-8")).digest() and
                    hashlib.sha256(password.encode("utf-8")).digest() ==
                    hashlib.sha256(expected_password.encode("utf-8")).digest()):
                st.session_state["authenticated"] = True
                st.session_state.pop("modal_area_activa", None)
                st.session_state["mostrar_modal_examenes"] = False
                st.session_state["area_activa"] = None
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")

        st.info("Acceso de prueba: usuario y contraseña iniciales son CRSCODUOC.")

    return False


def export_imported_workbook(template_bytes, fmt, selected_sheet=None):
    wb = load_workbook(io.BytesIO(template_bytes), data_only=False)
    sheet_names = wb.sheetnames

    def sheet_to_rows(ws):
        return [[cell.value for cell in row] for row in ws.iter_rows()]

    if fmt == "XLSX":
        return template_bytes, "excel_importado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    if fmt == "CSV":
        if selected_sheet not in sheet_names:
            raise ValueError("Selecciona una hoja válida para exportar a CSV.")
        ws = wb[selected_sheet]
        output = io.StringIO(newline="")
        writer = csv.writer(output)
        for row in sheet_to_rows(ws):
            writer.writerow(["" if value is None else value for value in row])
        return output.getvalue().encode("utf-8-sig"), "excel_importado.csv", "text/csv; charset=utf-8"

    if fmt == "XLS":
        try:
            import xlwt
        except ImportError as exc:
            raise RuntimeError("Falta la dependencia xlwt. Instálala con: pip install xlwt") from exc
        out_book = xlwt.Workbook(encoding="utf-8")
        used_names = set()
        for ws in wb.worksheets:
            base_name = re.sub(r'[\\/*?:\[\]]', '_', ws.title)[:31] or "Hoja"
            name = base_name
            suffix = 1
            while name.lower() in used_names:
                ending = f"_{suffix}"
                name = base_name[:31 - len(ending)] + ending
                suffix += 1
            used_names.add(name.lower())
            out_ws = out_book.add_sheet(name)
            for row_idx, row in enumerate(ws.iter_rows(), start=0):
                for col_idx, cell in enumerate(row):
                    value = cell.value
                    if value is None:
                        continue
                    if isinstance(value, (str, int, float, bool)):
                        out_ws.write(row_idx, col_idx, str(value)[:32000] if isinstance(value, str) else value)
                    else:
                        out_ws.write(row_idx, col_idx, str(value)[:32000])
        buffer = io.BytesIO()
        out_book.save(buffer)
        return buffer.getvalue(), "resultado_extraccion.xls", "application/vnd.ms-excel"

    if fmt == "JSON":
        payload = {"archivo": "excel_importado.xlsx", "hojas": []}
        for ws in wb.worksheets:
            payload["hojas"].append({
                "nombre": ws.title,
                "filas": sheet_to_rows(ws),
            })
        return json.dumps(payload, ensure_ascii=False, indent=2, default=str).encode(
            "utf-8"), "excel_importado.json", "application/json"

    if fmt == "XML":
        root = ET.Element("libro_excel")
        for ws in wb.worksheets:
            sheet_el = ET.SubElement(root, "hoja", nombre=ws.title)
            for row_idx, row in enumerate(ws.iter_rows(), start=1):
                row_el = ET.SubElement(sheet_el, "fila", numero=str(row_idx))
                for cell in row:
                    if cell.value is None:
                        continue
                    cell_el = ET.SubElement(row_el, "celda", coordenada=cell.coordinate)
                    cell_el.text = str(cell.value)
        xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        return xml_bytes, "excel_importado.xml", "application/xml"

    raise ValueError("Formato de exportación no compatible.")


def guide_tab():
    st.subheader("Guía de desarrollo seguro")
    st.write("Esta guía reúne prácticas recomendadas para desarrollar, probar y mantener la herramienta.")
    with st.expander("1. Protección de documentos y datos", expanded=False):
        st.markdown(
            "- Utilizar documentos ficticios o anonimizados durante las pruebas.\n- Compartir resultados solo con personas autorizadas.")
    with st.expander("2. Usuarios, contraseñas y configuración"):
        st.markdown("- Las credenciales iniciales `CRSCODUOC` / `CRSCODUOC` son solo para pruebas.")
    with st.expander("3. Manejo seguro de archivos"):
        st.markdown("- Aceptar solo los tipos de archivo necesarios y validar su contenido.")
    with st.expander("4. Calidad y revisión de datos"):
        st.markdown("- Marcar claramente los campos no encontrados y los que requieren revisión.")
    with st.expander("5. GitHub y publicación"):
        st.markdown("- Revisar los cambios antes de hacer commit.")


# --- CONFIGURACIÓN DE PÁGINA Y ESTILOS CSS ---
icon_path = os.path.join(os.path.dirname(__file__), "Pestaña.png")

st.set_page_config(
    page_title=APP_TITLE,
    page_icon=icon_path if os.path.exists(icon_path) else "🧪",
    layout="wide"
)

st.markdown("""
<style>
    .ag-theme-alpine ::-webkit-scrollbar {
        width: 8px;
        height: 8px;
    }
    .ag-theme-alpine ::-webkit-scrollbar-track {
        background: #f1f5f9;
        border-radius: 4px;
    }
    .ag-theme-alpine ::-webkit-scrollbar-thumb {
        background-color: #1E293B;
        border-radius: 4px;
    }
    .ag-theme-alpine ::-webkit-scrollbar-thumb:hover {
        background-color: #0f172a;
    }
    .banner-alerta-expiracion {
        background-color: #fee2e2;
        border: 1px solid #fca5a5;
        border-left: 6px solid #dc2626;
        color: #991b1b;
        padding: 12px 16px;
        border-radius: 6px;
        font-size: 0.92rem;
        margin-bottom: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .banner-cola-trabajo {
        background-color: #f0f9ff;
        border: 1px solid #bae6fd;
        border-left: 6px solid #0284c7;
        color: #0369a1;
        padding: 12px 16px;
        border-radius: 6px;
        font-size: 0.95rem;
        margin-bottom: 20px;
    }

    /* ESTILO DE LA COLUMNA DE NÚMEROS: TEXTO BLANCO NATURAL SIN FONDO NI BORDES */
    .num-slot {
        height: 37px;
        display: flex;
        align-items: center;
        justify-content: center;
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        outline: none !important;
        font-weight: bold;
        font-size: 1rem;
        color: #FFFFFF !important;
        margin-bottom: 5px;
    }

    /* CONTENEDOR DE ELEMENTOS ARRASTRABLES (FORMATO Y COLOR ROJO DEFAULT DE STREAMLIT) */
    iframe[title*="streamlit_sortables"] {
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        width: 100% !important;
        max-width: 800px;
    }
</style>
""", unsafe_allow_html=True)

if not check_login():
    st.stop()

# --- VERIFICACIÓN DE 6 MESES PARA LA CARTERA DE EXÁMENES ---
fecha_cartera, cartera_expirada, msg_cartera = fecha_actualizacion_cartera()

if "mostrar_dialogo_expiracion" not in st.session_state and cartera_expirada:
    st.session_state["mostrar_dialogo_expiracion"] = True

if cartera_expirada:
    st.markdown(f"""
    <div class="banner-alerta-expiracion">
        🚨 <b>ATENCIÓN: LA CARTERA DE EXÁMENES SUPERA LOS 6 MESES DE ANTIGÜEDAD</b><br>
        <span style="font-size: 0.85rem;">
        {msg_cartera if not fecha_cartera else f"La última versión de la Cartera de Exámenes fue registrada el <b>{msg_cartera}</b>. Se recomienda actualizar el Excel en la barra lateral."}
        </span>
    </div>
    """, unsafe_allow_html=True)

st.title(APP_TITLE)

if "examenes_por_area" not in st.session_state:
    st.session_state["examenes_por_area"] = cargar_examenes_por_area()
if "area_activa" not in st.session_state:
    st.session_state["area_activa"] = None
if "examen_activo" not in st.session_state:
    st.session_state["examen_activo"] = None
if "aggrid_key_counter" not in st.session_state:
    st.session_state["aggrid_key_counter"] = 0
if "cola_pdf_indice" not in st.session_state:
    st.session_state["cola_pdf_indice"] = 0
if "mostrar_modal_examenes" not in st.session_state:
    st.session_state["mostrar_modal_examenes"] = False


# --- DEFINICIÓN DE DIÁLOGOS EMERGENTES DE STREAMLIT (@st.dialog) ---

@st.dialog("🔔 Alerta de mantenimiento: Cartera de Exámenes")
def modal_alerta_cartera():
    st.warning("⚠️ Han transcurrido más de 6 meses desde la última actualización de la Cartera de Exámenes.")
    st.write(
        "Para garantizar la concordancia exacta entre la información del laboratorio y las fichas extraídas, te sugerimos actualizar la base de datos subiendo el archivo Excel más reciente desde el panel lateral."
    )
    if st.button("Entendido / Ocultar notificación emergente", type="primary", use_container_width=True):
        st.session_state["mostrar_dialogo_expiracion"] = False
        st.rerun()


@st.dialog("Seleccionar coincidencia en Cartera de Exámenes")
def modal_seleccionar_candidato_cartera(candidatos):
    examen_act = st.session_state.get('examen_activo', '')
    st.markdown(f"Selecciona el examen de la **Cartera** que corresponde a **{examen_act}**:")
    st.write("")

    if not candidatos:
        st.info("No se encontraron coincidencias en la Cartera de Exámenes.")
        if st.button("Cerrar", use_container_width=True):
            st.session_state.pop("candidatos_cartera_pendientes", None)
            st.rerun()
        return

    idx_seleccionado = st.radio(
        "Candidatos con mayor porcentaje de similitud:",
        options=range(len(candidatos)),
        format_func=lambda i: f"{candidatos[i][0]}  —  ({candidatos[i][1]}% de coincidencia)"
    )

    st.write("")
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        if st.button("✓ Vincular examen seleccionado", type="primary", use_container_width=True):
            nombre_elegido = candidatos[idx_seleccionado][0]
            st.session_state["cartera_examen_emparejado"] = nombre_elegido
            st.session_state.pop("candidatos_cartera_pendientes", None)
            st.success(f"¡Vinculado con «{nombre_elegido}»!")
            st.rerun()

    with col_c2:
        if st.button("❌ No vincular ninguno", use_container_width=True):
            st.session_state["cartera_examen_emparejado"] = "Sin vincular"
            st.session_state.pop("candidatos_cartera_pendientes", None)
            st.rerun()


@st.dialog("Comparación y Resolución de Discrepancias (PDF vs Cartera)")
def modal_resolver_conflictos():
    st.markdown(f"Se detectaron discrepancias entre el PDF y la Cartera de Exámenes para el examen **{st.session_state.get('examen_activo', '')}**.")
    st.write("Por favor, revisa cada campo afectado y selecciona qué información deseas conservar:")
    st.write("")

    conflictos = st.session_state.get("conflictos_pendientes", {})
    if not conflictos:
        st.info("No hay discrepancias pendientes por resolver.")
        return

    respuestas = {}

    for campo, datos in conflictos.items():
        st.markdown(f"""
        <div style="background-color: #f1f5f9; padding: 10px 14px; border-radius: 6px; border-left: 5px solid #0284c7; margin-top: 10px; margin-bottom: 12px;">
            <span style="font-size: 1.05rem; font-weight: bold; color: #0f172a;">📌 Campo evaluado: <span style="color: #0369a1;">{campo}</span></span>
        </div>
        """, unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            st.info(f"📄 **Valor extraído del PDF ({campo}):**\n\n{datos['pdf'] or '(Vacío)'}")
        with c2:
            st.success(f"🗄️ **Valor en Cartera SQLite ({campo}):**\n\n{datos['cartera'] or '(Vacío)'}")

        opcion = st.radio(
            f"Selecciona el valor que deseas mantener para el campo «{campo}»:",
            options=["Conservar valor extraído del PDF", "Utilizar registro de la Cartera de Exámenes", "Unir ambas informaciones (PDF + Cartera)"],
            key=f"radio_conflicto_{campo}"
        )
        respuestas[campo] = opcion
        st.divider()

    col_btn1, col_btn2 = st.columns([2, 1])
    with col_btn1:
        if st.button("💾Aplicar y actualizar ficha", type="primary", use_container_width=True):
            data_actualizada = copy.deepcopy(st.session_state["data_extracted"])

            for campo, eleccion in respuestas.items():
                val_pdf = conflictos[campo]["pdf"]
                val_cartera = conflictos[campo]["cartera"]

                if eleccion == "Utilizar registro de la Cartera de Exámenes":
                    data_actualizada[campo] = val_cartera
                elif eleccion == "Unir ambas informaciones (PDF + Cartera)":
                    data_actualizada[campo] = f"{val_pdf}\n\n[ Cartera de Exámenes ]:\n{val_cartera}"
                else:
                    data_actualizada[campo] = val_pdf

            st.session_state["data_extracted"] = data_actualizada

            area_res = st.session_state.get("area_resultado", st.session_state.get("area_activa"))
            aliases_area = get_field_aliases(area_res)
            user_edited = st.session_state.get("user_edited_fields", set())
            rows_nuevas = []
            for field in aliases_area:
                value = data_actualizada.get(field, "")
                if field in user_edited:
                    state = "🖋 Modificado por usuario"
                elif field in ["Nombre LIS", "Laboratorio de procesamiento", "Actualizado por", "Autorizado por"]:
                    state = "✓ Ingresado por usuario" if value else "⚠ Pendiente de ingreso"
                elif field == "Fecha actualización":
                    state = "✓ Asignado por sistema" if value else "⚠ Pendiente de asignación"
                else:
                    state = "✓ Encontrado" if value else "⚠ No encontrado"
                rows_nuevas.append({
                    "Campo": field.strip(),
                    "Información extraída": value,
                    "Estado": state,
                })
            df_nuevo = pd.DataFrame(rows_nuevas)

            st.session_state["result_bytes"] = aplicar_ediciones_a_excel(
                st.session_state["result_bytes"],
                df_nuevo,
                st.session_state.get("examen_activo", "")
            )

            st.session_state.pop("conflictos_pendientes", None)
            st.session_state["modal_conflictos_abierto"] = False
            st.success("¡Discrepancias resueltas exitosamente!")
            st.rerun()

    with col_btn2:
        if st.button("❌Cancelar", use_container_width=True):
            st.session_state["modal_conflictos_abierto"] = False
            st.rerun()


@st.dialog("💾 Confirmar actualización de la planilla del área en la herramienta")
def modal_confirmar_actualizacion_planilla():
    area_activa = st.session_state.get("area_activa")
    examen_activo = st.session_state.get("examen_activo")
    comp = st.session_state.get("comparacion_planilla_detalles", {})

    st.markdown(f"Estás a punto de reemplazar la **planilla guardada en la herramienta** para el área **{area_activa}** con los datos recién procesados del examen **{examen_activo}**.")
    st.write("")

    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        st.metric(
            label="Campos del examen",
            value=f"{comp.get('campos_completados_despues', 0)} completados",
            delta=f"{comp.get('campos_completados_despues', 0) - comp.get('campos_completados_antes', 0)} vs versión anterior"
        )
    with col_m2:
        st.metric(
            label="Exámenes/Hojas en área",
            value=f"{comp.get('hojas_despues', 0)} hojas",
            delta=f"+1 nuevo examen" if comp.get("es_nueva_hoja") else "Misma cantidad"
        )
    with col_m3:
        st.metric(
            label="Campos con cambios",
            value=f"{len(comp.get('diferencias', []))} modificados"
        )

    st.divider()

    diferencias = comp.get("diferencias", [])
    if diferencias:
        st.markdown("##### 🔍 Resumen de diferencias detectadas:")
        df_dif = pd.DataFrame(diferencias)
        st.dataframe(
            df_dif[["Campo", "Tipo de cambio", "Valor anterior", "Valor nuevo"]],
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("ℹ️ No se detectaron diferencias en los campos. La versión actual y la nueva son idénticas.")

    st.write("")
    st.warning("⚠️ Esta acción sobreescribirá el archivo de la plantilla en el servidor para que los próximos análisis utilicen este nuevo contenido.")

    col_save1, col_save2 = st.columns([2, 1])

    with col_save1:
        if st.button("✓ Confirmar y actualizar planilla en la herramienta", type="primary", use_container_width=True):
            try:
                config = AREA_CONFIG.get(area_activa, {})
                plantilla_nombre = config.get("template")

                if not plantilla_nombre:
                    plantilla_nombre = f"plantilla_{norm(area_activa).replace(' ', '_')}.xlsx"
                    AREA_CONFIG[area_activa]["template"] = plantilla_nombre

                ruta_guardado = os.path.join(os.path.dirname(__file__), plantilla_nombre)

                with open(ruta_guardado, "wb") as f:
                    f.write(st.session_state["result_bytes"])

                if comp.get("es_nueva_hoja") and examen_activo:
                    actuales = st.session_state["examenes_por_area"].get(area_activa, [])
                    if examen_activo.casefold() not in {x.casefold() for x in actuales}:
                        actuales.append(examen_activo)
                        st.session_state["examenes_por_area"][area_activa] = sorted(actuales, key=str.casefold)
                        guardar_examenes_por_area(st.session_state["examenes_por_area"])

                st.session_state["modal_confirmar_planilla_abierto"] = False
                st.session_state.pop("comparacion_planilla_detalles", None)
                st.success(f"¡La planilla de «{area_activa}» fue actualizada exitosamente en la herramienta!")
                st.rerun()

            except Exception as err_save:
                st.error(f"Error al guardar la planilla en la herramienta: {err_save}")

    with col_save2:
        if st.button("❌ Cancelar", use_container_width=True):
            st.session_state["modal_confirmar_planilla_abierto"] = False
            st.rerun()


@st.dialog("Seleccionar y administrar exámenes")
def modal_examenes_area(area):
    st.subheader(area)
    examenes = st.session_state["examenes_por_area"].get(area, [])
    actual = st.session_state.get("examen_activo") if st.session_state.get("area_activa") == area else None
    indice = examenes.index(actual) if actual in examenes else None
    seleccionado = st.selectbox(
        "Examen de esta área", examenes, index=indice, placeholder="Selecciona un examen",
        key=f"selector_examen_{area}",
        help="El nombre elegido se utilizará exactamente en el campo Nombre LIS y como nombre de la hoja de Excel."
    )
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Usar examen", key=f"usar_examen_{area}", type="primary", use_container_width=True):
            if seleccionado:
                st.session_state["area_activa"] = area
                st.session_state["examen_activo"] = seleccionado
                st.session_state["cartera_examen_emparejado"] = "Sin vincular"
                st.session_state.pop("candidatos_cartera_pendientes", None)
                st.session_state["mostrar_modal_examenes"] = False
                st.session_state.pop("modal_area_activa", None)
                st.rerun()
            else:
                st.warning("Selecciona un examen o agrega uno nuevo.")
    with c2:
        if st.button("Limpiar selección", key=f"limpiar_examen_{area}", use_container_width=True):
            st.session_state["area_activa"] = area
            st.session_state["examen_activo"] = None
            st.session_state["cartera_examen_emparejado"] = "Sin vincular"
            st.session_state.pop("candidatos_cartera_pendientes", None)
            st.session_state["mostrar_modal_examenes"] = False
            st.session_state.pop("modal_area_activa", None)
            st.rerun()

    st.divider()
    with st.expander("➕ Añadir examen"):
        with st.form(f"form_anadir_examen_{area}", clear_on_submit=True):
            nuevo = st.text_input("Nombre exacto del examen", placeholder="Ej.: Hemograma automatizado")
            agregar = st.form_submit_button("Guardar examen")
        if agregar:
            nuevo = nuevo.strip()
            valido, mensaje = nombre_examen_valido(nuevo)
            actuales = st.session_state["examenes_por_area"].get(area, [])
            if not valido:
                st.warning(mensaje)
            elif nuevo.casefold() in {x.casefold() for x in actuales}:
                st.warning("Ese examen ya existe en esta área. Selecciónalo en el menú.")
            else:
                actuales.append(nuevo)
                st.session_state["examenes_por_area"][area] = sorted(actuales, key=str.casefold)
                if guardar_examenes_por_area(st.session_state["examenes_por_area"]):
                    st.session_state["area_activa"] = area
                    st.session_state["examen_activo"] = nuevo
                    st.session_state["cartera_examen_emparejado"] = "Sin vincular"
                    st.session_state.pop("candidatos_cartera_pendientes", None)
                    st.session_state["mostrar_modal_examenes"] = False
                    st.session_state.pop("modal_area_activa", None)
                    st.success(f"Se agregó «{nuevo}».")
                    st.rerun()
                else:
                    st.error("No se pudo guardar la lista en el almacenamiento local.")
    with st.expander("➖ Eliminar examen de la lista"):
        actuales = st.session_state["examenes_por_area"].get(area, [])
        if actuales:
            quitar = st.selectbox("Examen que deseas quitar", actuales, key=f"quitar_examen_{area}")
            confirmar = st.checkbox("Confirmo que quiero quitarlo del menú", key=f"confirmar_quitar_{area}")
            if st.button("Eliminar examen", key=f"eliminar_examen_{area}", disabled=not confirmar):
                st.session_state["examenes_por_area"][area] = [x for x in actuales if x.casefold() != quitar.casefold()]
                if st.session_state.get("examen_activo") == quitar and st.session_state.get("area_activa") == area:
                    st.session_state["examen_activo"] = None
                    st.session_state["cartera_examen_emparejado"] = "Sin vincular"
                if guardar_examenes_por_area(st.session_state["examenes_por_area"]):
                    st.session_state["mostrar_modal_examenes"] = False
                    st.session_state.pop("modal_area_activa", None)
                    st.success(f"Se quitó «{quitar}» del menú de selección.")
                    st.rerun()
                else:
                    st.error("No se pudo guardar el cambio.")
        else:
            st.info("No hay exámenes para eliminar.")


def seleccionar_area(area):
    """Establece el área activa en la sesión y garantiza que los diálogos emergentes se cierren."""
    st.session_state["area_activa"] = area
    st.session_state["mostrar_modal_examenes"] = False
    st.session_state.pop("modal_area_activa", None)


# --- GESTOR DE DIÁLOGOS EMERGENTES (GARANTIZA EXCLUSIVIDAD MUTUA) ---
def ejecutar_gestor_dialogos():
    """Ejecuta como máximo UN solo diálogo emergente por ciclo para evitar errores de Streamlit."""
    if st.session_state.get("modal_confirmar_planilla_abierto"):
        st.session_state["modal_confirmar_planilla_abierto"] = False
        modal_confirmar_actualizacion_planilla()
    elif st.session_state.get("modal_conflictos_abierto"):
        st.session_state["modal_conflictos_abierto"] = False
        modal_resolver_conflictos()
    elif "candidatos_cartera_pendientes" in st.session_state:
        modal_seleccionar_candidato_cartera(st.session_state["candidatos_cartera_pendientes"])
    elif st.session_state.get("mostrar_modal_examenes") and st.session_state.get("modal_area_activa") in AREA_CONFIG:
        modal_examenes_area(st.session_state["modal_area_activa"])
    elif st.session_state.get("mostrar_dialogo_expiracion"):
        modal_alerta_cartera()


# --- SIDEBAR ---
with st.sidebar:
    logo_path = os.path.join(os.path.dirname(__file__), "CRSCO.png")
    if os.path.exists(logo_path):
        st.image(logo_path, use_container_width=True)
    st.divider()
    st.header("🧪 Áreas del laboratorio")
    st.caption("Selecciona un área para trabajar.")
    for area in AREA_CONFIG:
        st.button(
            area,
            key=f"abrir_area_{area}",
            use_container_width=True,
            type="primary" if st.session_state.get("area_activa") == area else "secondary",
            on_click=seleccionar_area,
            args=(area,),
        )
    st.divider()
    st.markdown(f"**Área activa:** {st.session_state.get('area_activa') or 'Sin seleccionar'}")
    st.markdown(f"**Examen activo:** {st.session_state.get('examen_activo') or 'Sin seleccionar'}")
    st.divider()

    st.header("📃 Fichas técnicas")
    template = st.file_uploader("Excel del área (opcional)", type=["xlsx"], key="template_uploader")
    st.caption("Si no cargas uno, se utilizará la plantilla incluida para el área seleccionada.")
    st.divider()

    st.header("💾 Base de datos (SQLite)")
    with st.expander("📊 Cartera de Exámenes", expanded=cartera_expirada):
        st.caption(f"Última actualización: **{msg_cartera}**")
        cartera_file = st.file_uploader("Actualizar Cartera (.xlsx)", type=["xlsx"], key="cartera_uploader_sidebar")
        if cartera_file is not None:
            if st.button("💾 Guardar en Base de Datos", type="primary", use_container_width=True, key="btn_save_cartera"):
                try:
                    df_cartera = pd.read_excel(cartera_file)
                    ok_db, msg_db = guardar_cartera_db(df_cartera)
                    if ok_db:
                        st.success(msg_db)
                        st.session_state["mostrar_dialogo_expiracion"] = False
                        st.rerun()
                    else:
                        st.error(msg_db)
                except Exception as err_db:
                    st.error(f"Error al procesar la Cartera: {err_db}")

    st.divider()
    if st.button("Cerrar sesión"):
        st.session_state["authenticated"] = False
        st.session_state.pop("result_bytes", None)
        st.session_state.pop("modal_area_activa", None)
        st.session_state["mostrar_modal_examenes"] = False
        st.session_state["area_activa"] = None
        st.session_state["cola_pdf_indice"] = 0
        st.rerun()


# --- VISTA PRINCIPAL ---
if st.session_state.get("area_activa") is None:
    st.write("")
    st.info("### 👋 ¡Bienvenido/a a la Herramienta de Extracción de Datos!")
    
    st.markdown(
        """
        Para comenzar a trabajar, selecciona un **área del laboratorio** en el panel lateral de la izquierda.

        #### 📌 Pasos para trabajar:
        1. 🧪 **Selecciona un Área del laboratorio** en el menú de la izquierda (*Microbiología, Hematología y coagulación, Bioquímica, Hormonas e inmunidad, Derivados*).
        2. 📋 **Administra y selecciona el examen** que deseas procesar usando el botón **📋 Administrar exámenes**.
        3. 📄 **Carga las fichas técnicas en PDF** (lote de hasta 5 documentos simultáneos) y organiza la cola de procesamiento si lo deseas.
        4. ⚙️ **Define los parámetros operacionales** (*Laboratorio de procesamiento, Actualizado por, Autorizado por*) y verifica el emparejamiento con la Cartera en SQLite.
        5. 🔎 **Haz clic en "Analizar documento activo"** para realizar la extracción automática de texto, OCR y cruce con la Base Maestra de Cartera.
        6. ✏️ **Revisa y edita la información** en la tabla interactiva, resuelve discrepancias si existen, y actualiza la planilla del área o descarga el Excel consolidado.
        """
    )
    
    st.divider()
    
    # --- MÓDULO DE ESTADO DE LA CARTERA SQLITE EN LA PÁGINA DE BIENVENIDA ---
    st.markdown("#### 📊 Estado de la Cartera SQLite:")
    lista_cartera_home = obtener_lista_cartera_db()
    total_examenes_cartera = len(lista_cartera_home)
    fecha_cartera_home, cartera_expirada_home, msg_cartera_home = fecha_actualizacion_cartera()

    if fecha_cartera_home is None or total_examenes_cartera == 0:
        st.warning("⚠️ **Cartera desactualizada:** No se ha subido ninguna Cartera de Exámenes a la base de datos. Se sugiere cargar el archivo Excel en el menú lateral.")
    elif cartera_expirada_home:
        st.warning(f"⚠️ **Cartera desactualizada (más de 6 meses de antigüedad):** Registrada el **{msg_cartera_home}**. Se recomienda actualizar la base de datos subiendo el archivo Excel desde el panel lateral.")
    else:
        st.success(f"✅ **Cartera de Exámenes cargada y lista:** Última actualización registrada el **{msg_cartera_home}**.")

    col_stat1, col_stat2 = st.columns(2)
    with col_stat1:
        st.metric(
            label="Exámenes registrados en la Cartera Base",
            value=f"{total_examenes_cartera} exámenes"
        )
    with col_stat2:
        st.metric(
            label="Fecha de última actualización",
            value=msg_cartera_home if fecha_cartera_home else "Sin registro"
        )

else:
    tab_pdf, tab_export, tab_guide = st.tabs([
        "Extraer información desde un PDF",
        "Exportar a otro formato",
        "Guía de desarrollo seguro",
    ])

    with tab_pdf:
        st.subheader("Extraer información desde fichas técnicas (PDF)")
        area_activa = st.session_state.get("area_activa")
        examen_activo = st.session_state.get("examen_activo")

        # CABECERA CON BOTÓN PARA ABRIR LA VENTANA DE ADMINISTRAR EXÁMENES
        col_info_area, col_btn_admin = st.columns([3, 1])
        with col_info_area:
            st.info(f"Área seleccionada: **{area_activa}** · Examen activo: **{examen_activo or 'sin seleccionar'}**")
        with col_btn_admin:
            if st.button("📋 Administrar exámenes", type="primary", use_container_width=True, key="btn_abrir_admin_examenes"):
                st.session_state["modal_area_activa"] = area_activa
                st.session_state["mostrar_modal_examenes"] = True
                st.rerun()

        # CARGA DE ARCHIVOS PDF Y CONTROL DE COLA
        pdfs = st.file_uploader(
            f"Cargar fichas técnicas (Lote de hasta {MAX_PDFS_PERMITIDOS} PDF para procesar en cola)",
            type=["pdf"],
            accept_multiple_files=True,
            help=f"Puedes cargar hasta {MAX_PDFS_PERMITIDOS} archivos PDF para procesarlos de forma individual uno tras otro.",
            key="pdf_uploader"
        )

        exceso_archivos = pdfs is not None and len(pdfs) > MAX_PDFS_PERMITIDOS

        if exceso_archivos:
            st.error(f"⚠️ Has seleccionado **{len(pdfs)} archivos**. El límite máximo simultáneo es de **{MAX_PDFS_PERMITIDOS} documentos PDF** para evitar saturar el sistema.")

        if pdfs and not exceso_archivos:
            mapa_pdfs = {pdf.name: pdf for pdf in pdfs}
            nombres_actuales = [pdf.name for pdf in pdfs]

            # 1. Sincronizar el estado de Python cuando cambian los archivos cargados (altas o bajas)
            if ("orden_nombres_pdf" not in st.session_state or 
                set(st.session_state["orden_nombres_pdf"]) != set(nombres_actuales)):
                
                # Conservar el orden de los existentes y descartar eliminados
                orden_previo = st.session_state.get("orden_nombres_pdf", [])
                nuevo_orden_filtrado = [nombre for nombre in orden_previo if nombre in mapa_pdfs]
                
                # Añadir archivos nuevos que no estuvieran en el estado
                for nombre in nombres_actuales:
                    if nombre not in nuevo_orden_filtrado:
                        nuevo_orden_filtrado.append(nombre)
                
                st.session_state["orden_nombres_pdf"] = nuevo_orden_filtrado
                st.session_state["cola_pdf_indice"] = 0
                st.session_state.pop("data_extracted", None)
                st.session_state.pop("user_edited_fields", None)
                st.session_state.pop("extracted_text", None)

            # 2. Key dinámica única basada en la firma de los archivos presentes.
            hash_archivos = "_".join(sorted(nombres_actuales))
            key_sortable_dinamico = f"sortable_queue_{hash_archivos}"

            if len(nombres_actuales) > 1:
                st.markdown("##### 🔀 Organiza el orden de la cola de procesamiento:")
                
                # ESTRUCTURA DE ESTANTERÍA: Números con punto a la izquierda (SIN FONDO, LETRAS BLANCAS) + Fichas arrastrables
                col_posicion, col_fichas = st.columns([1, 14])

                with col_posicion:
                    st.caption("Orden")
                    for i in range(1, len(st.session_state["orden_nombres_pdf"]) + 1):
                        st.markdown(f'<div class="num-slot">{i}.</div>', unsafe_allow_html=True)

                with col_fichas:
                    st.caption("Documentos (arrastra de arriba a abajo)")
                    nuevo_orden_nombres = sort_items(
                        st.session_state["orden_nombres_pdf"],
                        direction="vertical",
                        key=key_sortable_dinamico
                    )

                if nuevo_orden_nombres and nuevo_orden_nombres != st.session_state["orden_nombres_pdf"]:
                    st.session_state["orden_nombres_pdf"] = nuevo_orden_nombres
                    st.session_state["cola_pdf_indice"] = 0
                    st.session_state.pop("data_extracted", None)
                    st.session_state.pop("user_edited_fields", None)
                    st.session_state.pop("extracted_text", None)
                    st.rerun()

            pdfs_ordenados = [mapa_pdfs[nombre] for nombre in st.session_state["orden_nombres_pdf"] if nombre in mapa_pdfs]

            total_pdfs = len(pdfs_ordenados)
            idx_actual = st.session_state.get("cola_pdf_indice", 0)

            if idx_actual >= total_pdfs:
                st.session_state["cola_pdf_indice"] = 0
                idx_actual = 0

            pdf_activo = pdfs_ordenados[idx_actual]

            # BANNER DE ESTADO Y BARRA DE PROGRESO DE LA COLA
            st.markdown(f"""
            <div class="banner-cola-trabajo">
                📚 <b>COLA DE PROCESAMIENTO ACTIVA:</b> Documento <b>{idx_actual + 1}</b> de <b>{total_pdfs}</b><br>
                <span style="font-size: 0.90rem;">📄 Archivo en revisión actual: <b>{pdf_activo.name}</b></span>
            </div>
            """, unsafe_allow_html=True)

            st.progress((idx_actual + 1) / total_pdfs)

            st.divider()

            # --- PANEL DE EMPAREJAMIENTO CON LA CARTERA DE EXÁMENES ---
            st.markdown("##### 🔗 Emparejamiento institucional (Cartera de Exámenes en SQLite)")
            lista_cartera = obtener_lista_cartera_db()

            if not lista_cartera:
                st.warning("⚠️ No hay datos cargados en la Cartera de Exámenes. Puedes subir el archivo Excel desde el panel lateral.")
            else:
                emparejado_actual = st.session_state.get("cartera_examen_emparejado", "Sin vincular")

                col_emp1, col_emp2 = st.columns([2, 1])

                with col_emp1:
                    opciones_cartera = ["Sin vincular"] + lista_cartera
                    idx_emp = opciones_cartera.index(emparejado_actual) if emparejado_actual in opciones_cartera else 0

                    seleccion_manual_cartera = st.selectbox(
                        "Examen vinculado en la Cartera:",
                        opciones_cartera,
                        index=idx_emp,
                        help="Puedes corregir o cambiar manualmente el examen de la Cartera con el que deseas cruzar la información."
                    )
                    if seleccion_manual_cartera != emparejado_actual:
                        st.session_state["cartera_examen_emparejado"] = seleccion_manual_cartera
                        st.rerun()

                with col_emp2:
                    st.write("")
                    st.write("")
                    if st.button("🔍 **Buscar coincidencia**", key="btn_re_fuzzy", use_container_width=False):
                        if examen_activo:
                            candidatos = buscar_top_coincidencias_cartera(examen_activo, top_n=5)
                            if candidatos:
                                st.session_state["candidatos_cartera_pendientes"] = candidatos
                                st.rerun()
                            else:
                                st.warning("No se encontraron coincidencias en la Cartera para este examen.")
                        else:
                            st.warning("Debes seleccionar primero un examen activo usando el botón 'Administrar exámenes'.")

            st.divider()
            st.markdown("##### ⚙️ Parámetros operacionales de la ficha")

            c_lab, c_act, c_aut = st.columns(3)

            with c_lab:
                laboratorios = cargar_laboratorios()
                actual_lab = st.session_state.get("laboratorio_activo")
                idx_lab = laboratorios.index(actual_lab) if actual_lab in laboratorios else None
                sel_lab = st.selectbox(
                    "Laboratorio de procesamiento", laboratorios, index=idx_lab,
                    placeholder="Selecciona laboratorio", key="sel_lab_main"
                )
                st.session_state["laboratorio_activo"] = sel_lab

            with c_act:
                encargados_act = cargar_encargados("actualizado")
                actual_act = st.session_state.get("actualizado_activo")
                idx_act = encargados_act.index(actual_act) if actual_act in encargados_act else None
                sel_act = st.selectbox(
                    "Actualizado por", encargados_act, index=idx_act,
                    placeholder="Selecciona quién actualiza", key="sel_act_main"
                )
                st.session_state["actualizado_activo"] = sel_act

            with c_aut:
                encargados_aut = cargar_encargados("autorizado")
                actual_aut = st.session_state.get("autorizado_activo")
                idx_aut = encargados_aut.index(actual_aut) if actual_aut in encargados_aut else None
                sel_aut = st.selectbox(
                    "Autorizado por", encargados_aut, index=idx_aut,
                    placeholder="Selecciona quién autoriza", key="sel_aut_main"
                )
                st.session_state["autorizado_activo"] = sel_aut

            with st.expander("⚙️ Administrar opciones (Añadir / Eliminar de las listas arriba)"):
                tab_m_lab, tab_m_act, tab_m_aut = st.tabs(["Laboratorios", "Actualizado por", "Autorizado por"])

                with tab_m_lab:
                    col_a, col_b = st.columns(2)
                    with col_a:
                        with st.form("form_add_lab_inline", clear_on_submit=True):
                            new_lab = st.text_input("Añadir nuevo laboratorio")
                            if st.form_submit_button("Guardar laboratorio"):
                                if guardar_laboratorio(new_lab):
                                    st.session_state["laboratorio_activo"] = new_lab.strip()
                                    st.success(f"Guardado: {new_lab}")
                                    st.rerun()
                                else:
                                    st.warning("Nombre no válido o ya existente.")
                    with col_b:
                        labs_del = cargar_laboratorios()
                        if labs_del:
                            with st.form("form_del_lab_inline"):
                                target_del = st.selectbox("Eliminar laboratorio", labs_del)
                                if st.form_submit_button("Eliminar seleccionado"):
                                    updated = [x for x in labs_del if x.casefold() != target_del.casefold()]
                                    guardar_lista_laboratorios(updated)
                                    if st.session_state.get("laboratorio_activo") == target_del:
                                        st.session_state["laboratorio_activo"] = None
                                    st.success("Eliminado correctamente.")
                                    st.rerun()

                with tab_m_act:
                    col_a, col_b = st.columns(2)
                    with col_a:
                        with st.form("form_add_act_inline", clear_on_submit=True):
                            new_act = st.text_input("Añadir encargado (Actualizado por)")
                            if st.form_submit_button("Guardar encargado"):
                                curr = cargar_encargados("actualizado")
                                if new_act.strip() and new_act.strip().casefold() not in {x.casefold() for x in curr}:
                                    curr.append(new_act.strip())
                                    guardar_encargados("actualizado", curr)
                                    st.session_state["actualizado_activo"] = new_act.strip()
                                    st.success(f"Guardado: {new_act}")
                                    st.rerun()
                                else:
                                    st.warning("Nombre no válido o ya existente.")
                    with col_b:
                        acts_del = cargar_encargados("actualizado")
                        if acts_del:
                            with st.form("form_del_act_inline"):
                                target_del = st.selectbox("Eliminar de la lista", acts_del)
                                if st.form_submit_button("Eliminar seleccionado"):
                                    updated = [x for x in acts_del if x.casefold() != target_del.casefold()]
                                    guardar_encargados("actualizado", updated)
                                    if st.session_state.get("actualizado_activo") == target_del:
                                        st.session_state["actualizado_activo"] = None
                                    st.success("Eliminado correctamente.")
                                    st.rerun()

                with tab_m_aut:
                    col_a, col_b = st.columns(2)
                    with col_a:
                        with st.form("form_add_aut_inline", clear_on_submit=True):
                            new_aut = st.text_input("Añadir encargado (Autorizado por)")
                            if st.form_submit_button("Guardar encargado"):
                                curr = cargar_encargados("autorizado")
                                if new_aut.strip() and new_aut.strip().casefold() not in {x.casefold() for x in curr}:
                                    curr.append(new_aut.strip())
                                    guardar_encargados("autorizado", curr)
                                    st.session_state["autorizado_activo"] = new_aut.strip()
                                    st.success(f"Guardado: {new_aut}")
                                    st.rerun()
                                else:
                                    st.warning("Nombre no válido o ya existente.")
                    with col_b:
                        auts_del = cargar_encargados("autorizado")
                        if auts_del:
                            with st.form("form_del_aut_inline"):
                                target_del = st.selectbox("Eliminar de la lista", auts_del)
                                if st.form_submit_button("Eliminar seleccionado"):
                                    updated = [x for x in auts_del if x.casefold() != target_del.casefold()]
                                    guardar_encargados("autorizado", updated)
                                    if st.session_state.get("autorizado_activo") == target_del:
                                        st.session_state["autorizado_activo"] = None
                                    st.success("Eliminado correctamente.")
                                    st.rerun()

            st.divider()

            # VISTA PREVIA INTERACTIVA CON LUPA (MOSTRANDO DIRECTAMENTE EL DOCUMENTO ACTIVO EN COLA)
            with st.expander("📄 Vista previa interactiva del documento PDF en revisión", expanded=True):
                try:
                    doc_preview = fitz.open(stream=pdf_activo.getvalue(), filetype="pdf")
                    num_paginas = len(doc_preview)
                    st.caption(f"Visualizando documento activo en cola: **{pdf_activo.name}** | Páginas: **{num_paginas}**")

                    zoom_level = 120

                    col_prev1, col_prev2 = st.columns([2, 1])

                    with col_prev1:
                        pag_seleccionada = st.slider("Página a visualizar", 1, num_paginas, 1, key=f"pdf_page_slider_{idx_actual}") if num_paginas > 1 else 1

                    with col_prev2:
                        magnify_scale = st.slider("🔎 Potencia Lupa (x)", min_value=1.0, max_value=3.5, value=1.0, step=0.5, key=f"pdf_magnify_slider_{idx_actual}")

                    page = doc_preview[pag_seleccionada - 1]
                    pix = page.get_pixmap(dpi=150)
                    img_b64 = base64.b64encode(pix.tobytes("png")).decode("utf-8")

                    ancho_visual = int(750 * (zoom_level / 100))
                    alto_real_imagen = int(ancho_visual * (pix.height / pix.width))
                    alto_componente = alto_real_imagen + 45  # Espacio exacto para la imagen + el texto inferior

                    html_lupa = f"""
                    <style>
                        body {{
                            margin: 0;
                            padding: 0;
                            overflow: hidden;
                        }}
                    </style>
                    <div style="display: flex; flex-direction: column; align-items: center; width: 100%;">
                        <div class="lupa-container" style="
                            position: relative;
                            overflow: hidden;
                            border: 1px solid #ddd;
                            border-radius: 8px;
                            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
                            display: inline-block;
                            cursor: crosshair;
                            max-width: 100%;
                        ">
                            <img id="pdf-page-img" src="data:image/png;base64,{img_b64}" style="
                                width: {ancho_visual}px;
                                height: auto;
                                display: block;
                                transition: transform 0.15s ease-out;
                                transform-origin: center center;
                            " onmousemove="aplicarLupa(event, this)" onmouseleave="restaurarLupa(this)"/>
                        </div>
                        <p style="font-size: 0.8rem; color: #6c757d; margin-top: 6px;">💡 <b>Consejo:</b> Pasa el cursor sobre la página para activar la lupa interactiva.</p>
                    </div>

                    <script>
                        function aplicarLupa(e, img) {{
                            const rect = img.getBoundingClientRect();
                            const x = ((e.clientX - rect.left) / rect.width) * 100;
                            const y = ((e.clientY - rect.top) / rect.height) * 100;
                            img.style.transformOrigin = `${{x}}% ${{y}}%`;
                            img.style.transform = "scale({magnify_scale})";
                        }}
                        function restaurarLupa(img) {{
                            img.style.transform = "scale(1)";
                            img.style.transformOrigin = "center center";
                        }}
                    </script>
                    """

                    # 3. Renderizado con la altura calculada dinámicamente y scrolling desactivado
                    st.components.v1.html(html_lupa, height=alto_componente, scrolling=False)

                except Exception as err_preview:
                    st.error(f"No se pudo generar la vista previa del PDF: {err_preview}")

            # BOTÓN DE ANÁLISIS
            st.write("")
            col_btn_analisis, col_sp1 = st.columns([1.5, 2])
            with col_btn_analisis:
                iniciar_analisis = st.button("🔎 Analizar documento activo", type="primary", key="analyze_pdf", disabled=exceso_archivos, use_container_width=False)

            if iniciar_analisis:
                if not examen_activo:
                    st.warning("Selecciona un examen antes de continuar")
                elif not st.session_state.get("laboratorio_activo"):
                    st.warning("Selecciona un laboratorio de procesamiento antes de analizar.")
                else:
                    try:
                        st.session_state.pop("data_extracted", None)
                        st.session_state.pop("user_edited_fields", None)
                        st.session_state.pop("extracted_text", None)
                        st.session_state.pop("conflictos_pendientes", None)

                        st.session_state["aggrid_key_counter"] += 1

                        config = AREA_CONFIG[area_activa]
                        aliases_area = get_field_aliases(area_activa)
                        ruta_plantilla = os.path.join(os.path.dirname(__file__), config["template"] or "")

                        if "result_bytes" in st.session_state and idx_actual > 0:
                            template_bytes_init = st.session_state["result_bytes"]
                        elif template is not None:
                            template_bytes_init = template.getvalue()
                        elif config["template"] and os.path.exists(ruta_plantilla):
                            with open(ruta_plantilla, "rb") as f:
                                template_bytes_init = f.read()
                        else:
                            wb_vacio = openpyxl.Workbook()
                            ws_vacio = wb_vacio.active
                            ws_vacio.title = examen_activo
                            for fila, campo in enumerate(aliases_area, start=1):
                                ws_vacio.cell(row=fila, column=1, value=campo)
                                ws_vacio.cell(row=fila, column=2, value=examen_activo if norm(campo) == norm("Nombre LIS") else "")
                            buf = io.BytesIO()
                            wb_vacio.save(buf)
                            template_bytes_init = buf.getvalue()

                        st.session_state["bytes_plantilla_inicial"] = template_bytes_init

                        nombre_vinc = st.session_state.get("cartera_examen_emparejado")
                        info_cartera = obtener_info_cartera_por_nombre(nombre_vinc) if (nombre_vinc and nombre_vinc != "Sin vincular") else None

                        datos_sobrescribir = {}
                        conflictos = {}

                        if info_cartera and info_cartera.get("Código"):
                            datos_sobrescribir["Código"] = info_cartera["Código"]

                        campos_comparables = ["Preparación del paciente", "Equipamiento utilizado", "Plazo entrega resultados"]

                        raw_text, _ = extract_pdf_text(pdf_activo.getvalue(), pdf_activo.name, examen_activo)
                        parsed_pdf_data = infer_name_code(raw_text, parse_labeled_text(raw_text, aliases_area, examen_activo))

                        if info_cartera:
                            for campo in campos_comparables:
                                val_pdf = str(parsed_pdf_data.get(campo, "")).strip()
                                val_cartera = str(info_cartera.get(campo, "")).strip()

                                if val_pdf and val_cartera and val_pdf.casefold() != val_cartera.casefold():
                                    conflictos[campo] = {"pdf": val_pdf, "cartera": val_cartera}
                                elif not val_pdf and val_cartera:
                                    datos_sobrescribir[campo] = val_cartera

                        out_bytes, data, target, score, reason, method, text = process(
                            pdf_activo.getvalue(), pdf_activo.name, template_bytes_init,
                            st.session_state.get("laboratorio_activo", ""),
                            examen_activo, aliases_area,
                            st.session_state.get("actualizado_activo", ""),
                            st.session_state.get("autorizado_activo", ""),
                            datos_sobrescribir=datos_sobrescribir
                        )

                        st.session_state["result_bytes"] = out_bytes
                        st.session_state["data_extracted"] = copy.deepcopy(data)
                        st.session_state["user_edited_fields"] = set()
                        st.session_state["extracted_text"] = text
                        st.session_state["area_resultado"] = area_activa

                        if conflictos:
                            st.session_state["conflictos_pendientes"] = conflictos
                            st.session_state["modal_conflictos_abierto"] = True

                        st.success(f"¡Se procesó correctamente el archivo «{pdf_activo.name}»!")
                        st.rerun()

                    except Exception as e:
                        st.error(f"No se pudo analizar el documento: {e}")
        elif not pdfs:
            st.session_state.pop("result_bytes", None)
            st.session_state.pop("data_extracted", None)
            st.session_state.pop("orden_nombres_pdf", None)
            st.info("Carga uno o más PDF para comenzar la extracción en cola.")

        if "data_extracted" in st.session_state and "result_bytes" in st.session_state and pdfs:
            data = st.session_state.get("data_extracted", {})
            user_edited = st.session_state.get("user_edited_fields", set())
            area_res = st.session_state.get("area_resultado", area_activa)
            aliases_area = get_field_aliases(area_res)

            pdfs_ordenados_rev = [mapa_pdfs[nombre] for nombre in st.session_state.get("orden_nombres_pdf", []) if nombre in mapa_pdfs] if 'mapa_pdfs' in locals() else pdfs
            pdf_actual_nombre = pdfs_ordenados_rev[st.session_state.get('cola_pdf_indice', 0)].name if pdfs_ordenados_rev else "documento activo"

            st.subheader(f"Revisión e información extraída de «{pdf_actual_nombre}»")
            st.metric("Campos encontrados", sum(bool(v) for v in data.values()))

            if "conflictos_pendientes" in st.session_state:
                st.warning("⚠️ **Existen discrepancias detectadas entre la información extraída del PDF y la Cartera de Exámenes.**")
                col_dis1, col_dis2 = st.columns([1, 2])
                with col_dis1:
                    if st.button("⚖️ Resolver discrepancias", type="primary", use_container_width=True):
                        st.session_state["modal_conflictos_abierto"] = True
                        st.rerun()
                with col_dis2:
                    if st.button("🗑️ Omitir y conservar valores actuales del PDF", use_container_width=True):
                        st.session_state.pop("conflictos_pendientes", None)
                        st.rerun()
                st.write("")

            rows = []
            for field in aliases_area:
                value = data.get(field, "")
                if field in user_edited:
                    state = "🖋 Modificado por usuario"
                elif field in ["Nombre LIS", "Laboratorio de procesamiento", "Actualizado por", "Autorizado por"]:
                    state = "✓ Ingresado por usuario" if value else "⚠ Pendiente de ingreso"
                elif field == "Fecha actualización":
                    state = "✓ Asignado por sistema" if value else "⚠ Pendiente de asignación"
                else:
                    state = "✓ Encontrado" if value else "⚠ No encontrado"
                rows.append({
                    "Campo": field.strip(),
                    "Información extraída": value,
                    "Estado": state,
                })
            df_base = pd.DataFrame(rows)

            st.markdown(
                """
                <div style="font-size: 0.85rem; color: #6c757d; line-height: 1.6; margin-bottom: 12px;">
                    <span style="background-color: #B7F7C1; color: #124D20; padding: 2px 8px; border-radius: 4px; font-weight: bold;">Verde:</span> Campo completado correctamente.
                    &nbsp;&nbsp;
                    <span style="background-color: #BAE6FD; color: #0369A1; padding: 2px 8px; border-radius: 4px; font-weight: bold;">Azul:</span> Modificado por usuario.
                    &nbsp;&nbsp;
                    <span style="background-color: #FFF0B3; color: #6B4D00; padding: 2px 8px; border-radius: 4px; font-weight: bold;">Amarillo:</span> Campo pendiente de revisión / no encontrado.
                    <br><br>
                    💡 <b>Haz doble clic sobre cualquier celda de la columna "Información extraída" para editar su contenido directamente en la tabla.</b>
                </div>
                """,
                unsafe_allow_html=True
            )

            get_row_style = JsCode("""
            function(params) {
                var val = params.data['Información extraída'];
                var estado = params.data['Estado'];
                
                if (val && val.toString().trim() !== '' && estado === '🖋 Modificado por usuario') {
                    return {
                        'backgroundColor': '#BAE6FD',
                        'color': '#0369A1',
                        'fontWeight': '600'
                    };
                } else if (val && val.toString().trim() !== '') {
                    return {
                        'backgroundColor': '#B7F7C1',
                        'color': '#124D20',
                        'fontWeight': '600'
                    };
                } else {
                    return {
                        'backgroundColor': '#FFF0B3',
                        'color': '#6B4D00',
                        'fontWeight': '500'
                    };
                }
            }
            """)

            dark_header_style = {
                'backgroundColor': '#1E293B',
                'color': '#FFFFFF',
                'fontWeight': 'bold',
                'fontSize': '0.95rem'
            }

            gb = GridOptionsBuilder.from_dataframe(df_base)
            gb.configure_default_column(
                headerStyle=dark_header_style,
                resizable=False,       # Evita que el usuario ajuste el ancho manualmente
                suppressMovable=True   # Evita que el usuario arrastre y reordene las columnas
            )

            gb.configure_column(
                "Campo",
                editable=False,
                flex=2,
                cellStyle={'fontWeight': 'bold', 'fontSize': '0.9rem'}
            )
            gb.configure_column(
                "Información extraída",
                editable=True,
                flex=5,
                cellEditor='agLargeTextCellEditor',
                cellEditorPopup=True,
                autoHeight=True
            )
            gb.configure_column(
                "Estado",
                editable=False,
                flex=2,
                cellStyle={'textAlign': 'center', 'fontSize': '0.85rem'}
            )

            grid_options = gb.build()
            grid_options['getRowStyle'] = get_row_style
            grid_options['domLayout'] = 'autoHeight'
            grid_options['suppressMovableColumns'] = True  # 👈 Refuerzo global para congelar la posición de las columnas
            
            grid_options['onCellValueChanged'] = JsCode("""
            function(params) {
                if (params.column && params.column.colDef && params.column.colDef.field === 'Información extraída') {
                    var newVal = params.newValue != null ? params.newValue.toString().trim() : '';
                    
                    if (newVal === '') {
                        var campo = params.data['Campo'];
                        if (['Nombre LIS', 'Laboratorio de procesamiento', 'Actualizado por', 'Autorizado por'].includes(campo)) {
                            params.node.setDataValue('Estado', '⚠ Pendiente de ingreso');
                        } else if (campo === 'Fecha actualización') {
                            params.node.setDataValue('Estado', '⚠ Pendiente de asignación');
                        } else {
                            params.node.setDataValue('Estado', '⚠ No encontrado');
                        }
                    } else {
                        params.node.setDataValue('Estado', '🖋 Modificado por usuario');
                    }
                    params.api.refreshCells({rowNodes: [params.node], force: true});
                }
            }
            """)

            grid_response = AgGrid(
                df_base,
                gridOptions=grid_options,
                allow_unsafe_jscode=True,
                update_mode=GridUpdateMode.MODEL_CHANGED,
                fit_columns_on_grid_load=True,
                theme="alpine",
                height=None,
                key=f"aggrid_data_table_{st.session_state['aggrid_key_counter']}"
            )

            edited_df = pd.DataFrame(grid_response['data'])

            if "data_extracted" in st.session_state:
                for idx, row in edited_df.iterrows():
                    field = row["Campo"]
                    val_editado = str(row["Información extraída"] or '').strip()
                    st.session_state["data_extracted"][field] = val_editado
                    
                    if row["Estado"] == "🖋 Modificado por usuario" and val_editado:
                        st.session_state.setdefault("user_edited_fields", set()).add(field)
                    else:
                        st.session_state.get("user_edited_fields", set()).discard(field)

            updated_excel_bytes = aplicar_ediciones_a_excel(
                st.session_state["result_bytes"],
                edited_df,
                examen_activo or ""
            )
            st.session_state["result_bytes"] = updated_excel_bytes

            st.write("")
            st.subheader("Opciones de guardado y avance de cola")

            col_down, col_update, col_next = st.columns(3)

            with col_down:
                st.download_button(
                    "📥 Descargar Excel consolidado (.xlsx)", type="secondary",
                    data=st.session_state["result_bytes"],
                    file_name=f"Ficha_Acumulada_{area_activa or 'lote'}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="download_updated_xlsx",
                    help="Descarga el libro Excel con todas las hojas de los exámenes procesados hasta ahora en este lote."
                )

            with col_update:
                if st.button("💾 Actualizar en herramienta", type="primary", use_container_width=True, key="btn_trigger_save_template"):
                    bytes_iniciales = st.session_state.get("bytes_plantilla_inicial")
                    st.session_state["comparacion_planilla_detalles"] = comparar_plantillas_excel(
                        bytes_iniciales,
                        st.session_state["result_bytes"],
                        examen_activo or ""
                    )
                    st.session_state["modal_confirmar_planilla_abierto"] = True
                    st.rerun()

            with col_next:
                idx_actual = st.session_state.get("cola_pdf_indice", 0)
                total_pdfs = len(pdfs_ordenados_rev) if 'pdfs_ordenados_rev' in locals() and pdfs_ordenados_rev else 0
                
                if idx_actual < total_pdfs - 1:
                    if st.button(f"➡️ Finalizar y pasar al examen ({idx_actual + 2}/{total_pdfs})", type="secondary", use_container_width=True):
                        st.session_state["cola_pdf_indice"] = idx_actual + 1

                        st.session_state["examen_activo"] = None
                        st.session_state["cartera_examen_emparejado"] = "Sin vincular"
                        st.session_state.pop("candidatos_cartera_pendientes", None)

                        st.session_state.pop("data_extracted", None)
                        st.session_state.pop("user_edited_fields", None)
                        st.session_state.pop("extracted_text", None)

                        st.session_state["mostrar_modal_examenes"] = False
                        st.session_state.pop("modal_area_activa", None)

                        st.rerun()
                else:
                    st.success("🎉 ¡Has completado la cola de exámenes! Puedes descargar el Excel consolidado con el botón de la izquierda.")

            with st.expander("Ver texto bruto extraído"):
                st.text_area("Texto", st.session_state.get("extracted_text", ""), height=350, key="raw_extracted_text")

    with tab_export:
        st.subheader("Exportar a otro formato")
        st.write("Esta pestaña permite exportar el libro Excel que cargues en el panel lateral.")
        if template:
            try:
                imported_wb = load_workbook(io.BytesIO(template.getvalue()), read_only=True, data_only=False)
                imported_sheet_names = imported_wb.sheetnames
                imported_wb.close()
                import_col1, import_col2 = st.columns([1, 2])
                with import_col1:
                    imported_format = st.selectbox(
                        "Formato de exportación",
                        ["XLSX", "CSV", "XLS", "JSON", "XML"],
                        key="imported_excel_format",
                    )
                with import_col2:
                    selected_import_sheet = None
                    if imported_format == "CSV":
                        selected_import_sheet = st.selectbox(
                            "Hoja que se exportará a CSV",
                            imported_sheet_names,
                            key="imported_excel_sheet",
                        )
                    else:
                        st.caption("Se exportarán todas las hojas cuando el formato lo permita.")
                imported_bytes, imported_name, imported_mime = export_imported_workbook(
                    template.getvalue(), imported_format, selected_import_sheet
                )
                st.download_button(
                    f"Descargar Excel importado en {imported_format}", type="primary",
                    data=imported_bytes,
                    file_name=imported_name,
                    mime=imported_mime,
                    key="download_imported_excel",
                )
            except Exception as e:
                st.error(f"No se pudo preparar la exportación del Excel importado: {e}")
        else:
            st.info("Carga primero un Excel de fichas técnicas en el menú lateral para habilitar esta exportación.")

    with tab_guide:
        guide_tab()

# --- EJECUCIÓN FINAL DEL GESTOR DE DIÁLOGOS EMERGENTES ---
ejecutar_gestor_dialogos()
