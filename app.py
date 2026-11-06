import io, os, re, html, tempfile, subprocess, shutil, csv, json, hashlib, copy, base64, sqlite3, xml.etree.ElementTree as ET
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
import fuentes_documentales as fd

APP_TITLE = 'Herramienta de Extracción de Datos de Fichas Técnicas'
MAX_PDFS_PERMITIDOS = 5
DB_PATH = os.path.join(os.path.dirname(__file__), "base_de_datos.db")

AREA_CONFIG = {
    "Microbiología": {"template": "plantilla_microbiologia.xlsx"},
    "Hematología y coagulación": {"template": "plantilla_hematologia.xlsx"},
    "Bioquímica": {"template": "plantilla_bioquimica.xlsx"},
    "Hormonas e inmunidad": {"template": None},
    "Derivados": {"template": None},
}


def formatear_fecha_fuente(fecha_iso):
    """Presenta la fecha ISO guardada en SQLite como DD/MM/AAAA HH:MM.

    No cambia el valor almacenado (ISO se mantiene para ordenar/consultar).
    """
    if not fecha_iso:
        return "Sin registro"
    try:
        fecha = datetime.fromisoformat(fecha_iso)
        if fecha.tzinfo is not None:
            fecha = fecha.astimezone(ZoneInfo("America/Santiago"))
        return fecha.strftime("%d/%m/%Y %H:%M")
    except (TypeError, ValueError):
        return "Fecha no disponible"


# --- GESTIÓN DE BASE DE DATOS SQLITE (CARTERA Y ÁREAS) ---
def init_db():
    """Inicializa la base de datos SQLite con tablas para Cartera, Áreas y Metadata."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Cartera de Exámenes
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
    
    # Plantillas/Fichas por Área del Laboratorio
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS plantillas_area (
            area TEXT PRIMARY KEY,
            excel_bytes BLOB,
            fecha_actualizacion TEXT
        )
    """)
    
    # Exámenes pertenecientes a cada área
    # Compatibilidad con bases antiguas: la columna se añade sin perder las plantillas.
    columnas_plantillas = {r[1] for r in cursor.execute("PRAGMA table_info(plantillas_area)").fetchall()}
    if "nombre_archivo" not in columnas_plantillas:
        cursor.execute("ALTER TABLE plantillas_area ADD COLUMN nombre_archivo TEXT")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS examenes_area (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            area TEXT,
            nombre_examen TEXT,
            UNIQUE(area, nombre_examen)
        )
    """)
    
    # Catálogos de selección persistentes (sin dependencia de archivos JSON).
    cursor.execute("""CREATE TABLE IF NOT EXISTS laboratorios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL COLLATE NOCASE UNIQUE,
        activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1))
    )""")
    cursor.execute("""CREATE TABLE IF NOT EXISTS encargados (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL COLLATE NOCASE UNIQUE,
        activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1))
    )""")
    conn.commit()

    # Cargar plantillas predeterminadas desde disco a la BD si aún no existen
    ZONA_HORARIA = ZoneInfo("America/Santiago")
    fecha_init = datetime.now(ZONA_HORARIA).isoformat()

    for area, config in AREA_CONFIG.items():
        cursor.execute("SELECT area FROM plantillas_area WHERE area = ?", (area,))
        if not cursor.fetchone():
            plantilla_nombre = config.get("template")
            if plantilla_nombre:
                ruta = os.path.join(os.path.dirname(__file__), plantilla_nombre)
                if os.path.exists(ruta):
                    try:
                        with open(ruta, "rb") as f:
                            b_data = f.read()
                        cursor.execute("""
                            INSERT OR REPLACE INTO plantillas_area
                            (area, excel_bytes, fecha_actualizacion, nombre_archivo)
                            VALUES (?, ?, ?, ?)
                        """, (area, b_data, fecha_init, plantilla_nombre))
                    except Exception:
                        pass
    # Identificar los archivos originales en instalaciones anteriores solo cuando
    # el contenido SQLite coincide exactamente con el XLSX disponible en disco.
    for area, config in AREA_CONFIG.items():
        nombre = config.get("template")
        if not nombre:
            continue
        ruta = os.path.join(os.path.dirname(__file__), nombre)
        if not os.path.isfile(ruta):
            continue
        cursor.execute("SELECT excel_bytes, nombre_archivo FROM plantillas_area WHERE area=?", (area,))
        fila = cursor.fetchone()
        if fila and fila[0] and not fila[1]:
            with open(ruta, "rb") as f:
                if hashlib.sha256(fila[0]).digest() == hashlib.sha256(f.read()).digest():
                    cursor.execute("UPDATE plantillas_area SET nombre_archivo=? WHERE area=?", (nombre, area))
    conn.commit()
    conn.close()

init_db()
fd.inicializar_fuentes(DB_PATH)


def obtener_plantilla_area_db(area):
    """Obtiene los bytes del archivo Excel y la última fecha de actualización para un área."""
    if not area:
        return None, "Sin registro"
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT excel_bytes, fecha_actualizacion FROM plantillas_area WHERE area = ?", (area,))
    row = cursor.fetchone()
    conn.close()

    if row:
        b_data, f_act = row
        if f_act:
            try:
                f_dt = datetime.fromisoformat(f_act)
                msg_f = f_dt.strftime("%d/%m/%Y %H:%M")
            except Exception:
                msg_f = f_act
        else:
            msg_f = "Sin registro"
        return b_data, msg_f
    return None, "Sin registro"


def obtener_nombre_plantilla_area_db(area):
    """Devuelve el nombre del Excel almacenado o None si la versión antigua no lo registró."""
    with sqlite3.connect(DB_PATH) as conn:
        fila = conn.execute("SELECT nombre_archivo FROM plantillas_area WHERE area=?", (area,)).fetchone()
    return fila[0] if fila and fila[0] else None


def guardar_plantilla_area_db(area, excel_bytes, nombre_archivo=None):
    """Guarda el XLSX y su nombre real; si es generado, se registra ese nuevo nombre."""
    if not area or not excel_bytes:
        return False, "Datos no válidos."
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    ZONA_HORARIA = ZoneInfo("America/Santiago")
    fecha_actual = datetime.now(ZONA_HORARIA).isoformat()

    cursor.execute("""
        INSERT OR REPLACE INTO plantillas_area
        (area, excel_bytes, fecha_actualizacion, nombre_archivo)
        VALUES (?, ?, ?, ?)
    """, (area, excel_bytes, fecha_actual, os.path.basename(nombre_archivo) if nombre_archivo else None))
    conn.commit()
    conn.close()

    # Actualizar lista de exámenes del área según las hojas del Excel
    try:
        wb = load_workbook(io.BytesIO(excel_bytes), read_only=True, data_only=True)
        hojas = list(wb.sheetnames)
        wb.close()
        if hojas:
            examenes_dict = cargar_examenes_por_area()
            examenes_dict[area] = sorted(list(dict.fromkeys(hojas)), key=str.casefold)
            guardar_examenes_por_area(examenes_dict)
    except Exception:
        pass

    return True, f"Se actualizó la base de datos de «{area}» exitosamente."


def cargar_examenes_por_area():
    """Carga los exámenes de cada área desde la BD; de estar vacía, consulta las plantillas cargadas."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    resultado = {area: [] for area in AREA_CONFIG}
    
    cursor.execute("SELECT area, nombre_examen FROM examenes_area ORDER BY nombre_examen ASC")
    rows = cursor.fetchall()
    conn.close()

    for area, nombre in rows:
        if area in resultado:
            resultado[area].append(nombre)

    # Si alguna área no tiene exámenes registrados en la tabla, poblar desde la plantilla SQLite
    for area in AREA_CONFIG:
        if not resultado[area]:
            b_bytes, _ = obtener_plantilla_area_db(area)
            if b_bytes:
                try:
                    wb = load_workbook(io.BytesIO(b_bytes), read_only=True, data_only=True)
                    resultado[area] = list(wb.sheetnames)
                    wb.close()
                except Exception:
                    pass

    return {a: sorted(list(dict.fromkeys(str(x).strip() for x in xs if str(x).strip())), key=str.casefold) for a, xs in resultado.items()}


def guardar_examenes_por_area(examenes_dict):
    """Guarda los exámenes asociados a cada área en SQLite."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM examenes_area")

    for area, lista in examenes_dict.items():
        vistos = set()
        for nombre in lista:
            nombre = str(nombre).strip()
            if nombre and nombre.casefold() not in vistos:
                cursor.execute("INSERT OR IGNORE INTO examenes_area (area, nombre_examen) VALUES (?, ?)", (area, nombre))
                vistos.add(nombre.casefold())

    conn.commit()
    conn.close()
    return True


def guardar_cartera_db(df, nombre_archivo=None, archivo_original=None):
    """Guarda la Cartera y, si se proporciona, conserva el XLSX original y su nombre."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    col_nombre_series = None
    
    for col in df.columns:
        if str(col).strip().upper() in ["DESCRIPCION", "DESCRIPCIÓN"]:
            col_nombre_series = df[col]
            break

    if col_nombre_series is None and len(df.columns) >= 6:
        col_nombre_series = df.iloc[:, 5]
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

        codigo = "" if codigo.casefold() in ["nan", "none", "null"] else codigo
        indicaciones = "" if indicaciones.casefold() in ["nan", "none", "null"] else indicaciones
        tecnologia = "" if tecnologia.casefold() in ["nan", "none", "null"] else tecnologia
        normal = "" if normal.casefold() in ["nan", "none", "null"] else normal
        urgente = "" if urgente.casefold() in ["nan", "none", "null"] else urgente

        cursor.execute("""
            INSERT OR REPLACE INTO cartera_examenes (nombre_examen, codigo_examen, indicaciones, tecnologia, entrega_normal, entrega_urgente)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (nombre, codigo, indicaciones, tecnologia, normal, urgente))
        filas_guardadas += 1

    if filas_guardadas == 0:
        conn.rollback()
        conn.close()
        return False, "No se encontraron prestaciones válidas en la Cartera. No se modificó la base de datos."

    ZONA_HORARIA = ZoneInfo("America/Santiago")
    fecha_actual = datetime.now(ZONA_HORARIA).isoformat()
    cursor.execute("INSERT OR REPLACE INTO cartera_metadata (clave, valor) VALUES ('ultima_actualizacion', ?)", (fecha_actual,))
    if nombre_archivo and archivo_original:
        nombre_seguro = os.path.basename(nombre_archivo)
        cursor.execute("INSERT OR REPLACE INTO cartera_metadata (clave, valor) VALUES ('nombre_archivo', ?)", (nombre_seguro,))
        # El original se almacena igual que los otros documentos de referencia.
        cursor.execute("""INSERT OR REPLACE INTO documentos_fuente
            (fuente_id,nombre_archivo,tipo_mime,archivo,sha256,fecha_actualizacion,registros)
            VALUES (?,?,?,?,?,?,?)""",
            ("cartera_examenes", nombre_seguro,
             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
             sqlite3.Binary(archivo_original),hashlib.sha256(archivo_original).hexdigest(),
             fecha_actual, cursor.execute("SELECT COUNT(*) FROM cartera_examenes").fetchone()[0]))

    conn.commit()
    conn.close()
    return True, f"Se importaron exitosamente {filas_guardadas} exámenes desde la Columna 6 (DESCRIPCION) a la Cartera."


def obtener_lista_cartera_db():
    """Retorna la lista ordenada de nombres/DESCRIPCION de exámenes almacenados en la Cartera."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    # No filtrar nombres que empiecen por un dígito: son exámenes legítimos,
    # como 17-HIDROXIPROGESTERONA y 17-BETA-ESTRADIOL.
    cursor.execute("SELECT nombre_examen FROM cartera_examenes ORDER BY nombre_examen COLLATE NOCASE ASC")
    rows = cursor.fetchall()
    conn.close()
    return [r[0] for r in rows if r[0]]


def buscar_top_coincidencias_cartera(nombre_buscar, top_n=5):
    """Busca los N mejores candidatos por similitud (Fuzzy Matching) en la Cartera."""
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
    """Consulta SQLite obteniendo los datos completos de un examen de la Cartera."""
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

def obtener_cartera_excel_bytes():
    """Genera y retorna el archivo Excel (.xlsx) con todos los exámenes registrados en la Cartera SQLite."""
    conn = sqlite3.connect(DB_PATH)
    query = """
        SELECT 
            nombre_examen AS "DESCRIPCION",
            codigo_examen AS "Código Examen",
            indicaciones AS "Indicaciones / Preparación del paciente",
            tecnologia AS "Equipamiento / Tecnología",
            entrega_normal AS "Entrega Normal",
            entrega_urgente AS "Entrega Urgente"
        FROM cartera_examenes
        ORDER BY nombre_examen ASC
    """
    df = pd.read_sql_query(query, conn)
    conn.close()

    if df.empty:
        return None

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Cartera de Exámenes")
    return output.getvalue()

# --- CATÁLOGOS PERSISTENTES EN SQLITE ---
# Los dos selectbox de responsables comparten la misma tabla de encargados.
def _normalizar_nombres(nombres):
    resultado = {}
    for nombre in nombres:
        nombre = " ".join(str(nombre).split())
        if nombre:
            resultado.setdefault(nombre.casefold(), nombre)
    return resultado


def _listar_catalogo(tabla):
    if tabla not in ("laboratorios", "encargados"):
        raise ValueError("Catálogo no permitido")
    with sqlite3.connect(DB_PATH, timeout=15) as conn:
        filas = conn.execute(f"SELECT nombre FROM {tabla} WHERE activo = 1").fetchall()
    return sorted((fila[0] for fila in filas), key=str.casefold)


def _guardar_catalogo(tabla, nombres):
    """Sincroniza opciones activas; desactiva sin borrar los registros históricos."""
    if tabla not in ("laboratorios", "encargados"):
        raise ValueError("Catálogo no permitido")
    nombres = _normalizar_nombres(nombres)
    try:
        with sqlite3.connect(DB_PATH, timeout=15) as conn:
            existentes = {fila[0].casefold(): fila[0] for fila in
                          conn.execute(f"SELECT nombre FROM {tabla}")}
            for clave, nombre in nombres.items():
                if clave in existentes:
                    conn.execute(f"UPDATE {tabla} SET activo = 1 WHERE nombre = ?", (existentes[clave],))
                else:
                    conn.execute(f"INSERT INTO {tabla} (nombre, activo) VALUES (?, 1)", (nombre,))
            for clave, nombre in existentes.items():
                if clave not in nombres:
                    conn.execute(f"UPDATE {tabla} SET activo = 0 WHERE nombre = ?", (nombre,))
        return True
    except sqlite3.Error as exc:
        st.error(f"No se pudieron guardar los cambios en SQLite: {exc}")
        return False


def cargar_laboratorios():
    return _listar_catalogo("laboratorios")


def guardar_lista_laboratorios(laboratorios):
    return _guardar_catalogo("laboratorios", laboratorios)


def guardar_laboratorio(nombre):
    nombre = " ".join(str(nombre).split())
    actuales = cargar_laboratorios()
    if not nombre or nombre.casefold() in {n.casefold() for n in actuales}:
        return False
    return guardar_lista_laboratorios(actuales + [nombre])


def cargar_encargados(tipo="actualizado"):
    """Ambos tipos comparten las mismas opciones almacenadas en SQLite."""
    return _listar_catalogo("encargados")


def guardar_encargados(tipo, lista):
    """Una modificación desde cualquier pestaña afecta a ambos selectbox."""
    return _guardar_catalogo("encargados", lista)


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


def extraer_solo_valor_referencia(texto):
    """Extrae únicamente el valor o rango numérico/cualitativo principal de los Valores de Referencia, omitiendo explicaciones."""
    if not texto:
        return ""
    t = str(texto).strip()
    if not t or t.casefold() in ["nan", "none", "null"]:
        return ""

    # Cortar párrafos explicativos o notas metodológicas comunes
    patrones_corte = [
        r"\n?\s*Estos valores corresponden.*",
        r"\n?\s*Cada laboratorio deber[ií]a.*",
        r"\n?\s*Datos espec[ií]ficos de funcionamiento.*",
        r"\n?\s*A continuaci[oó]n,\s*se indican.*",
        r"\n?\s*Los resultados de cada laboratorio.*",
        r"\n?\s*Obtenidos a partir de.*",
        r"\n?\s*Consulte la secci[oó]n.*",
        r"\n?\s*Intervalo de referencia.*",
    ]
    for pat in patrones_corte:
        t = re.split(pat, t, flags=re.IGNORECASE | re.DOTALL)[0]

    lineas = [l.strip() for l in t.splitlines() if l.strip()]
    if not lineas:
        return ""

    # Retorna únicamente la primera línea que contiene la cifra/rango principal de referencia
    return lineas[0]


def limpiar_valor_excel(valor, examen_activo=None):
    """Quita espacios accidentales, líneas de ruido y citas bibliográficas conservando valores válidos."""
    if valor is None:
        return ""

    valor_str = str(valor).strip()
    if not valor_str or valor_str.casefold() in ["nan", "none", "null"]:
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
    """Extrae cada página independientemente; usa OCR solo donde falta texto.

    Mantiene el contrato anterior: (texto, descripción_del_método).
    Nunca inventa información cuando no es posible leer una página.
    """
    try:
        doc = fitz.open(stream=uploaded_bytes, filetype="pdf")
    except Exception as exc:
        raise RuntimeError(f"No se pudo abrir el PDF '{filename}': {exc}") from exc

    if not doc.page_count:
        doc.close()
        raise RuntimeError("El PDF no contiene páginas.")

    pages_text = []
    paginas_ocr = []
    engine = None
    try:
        for numero, page in enumerate(doc, start=1):
            page_height = page.rect.height
            page_width = page.rect.width
            valid_blocks = []
            for b in page.get_text("blocks"):
                if b[6] != 0:
                    continue
                x0, y0, x1, y1, content = b[:5]
                content = content.strip()
                if not content:
                    continue
                if y0 < 50 or y1 > page_height - 50:
                    lines = [line.strip() for line in content.splitlines() if line.strip()]
                    if lines and all(es_linea_basura_o_encabezado(line, examen_activo) for line in lines):
                        continue
                if not es_linea_basura_o_encabezado(content, examen_activo):
                    valid_blocks.append((x0, y0, x1, y1, content))

            def sort_key(block):
                x0, y0, x1, _, _ = block
                column = 0 if x1 - x0 > page_width * 0.6 else (1 if x0 < page_width * 0.48 else 2)
                return (column, y0, x0)

            valid_blocks.sort(key=sort_key)
            page_text = "\n".join(b[4] for b in valid_blocks)

            # Un PDF mixto puede contener páginas con texto y páginas escaneadas.
            if len(re.sub(r"\s", "", page_text)) < 50:
                try:
                    if engine is None:
                        from rapidocr_onnxruntime import RapidOCR
                        import numpy as np
                        from PIL import Image
                        engine = RapidOCR()
                    pix = page.get_pixmap(dpi=200)
                    image = Image.open(io.BytesIO(pix.tobytes("png")))
                    result, _ = engine(np.array(image))
                    ocr_text = "\n".join(str(line[1]) for line in (result or []) if len(line) > 1)
                    if len(re.sub(r"\s", "", ocr_text)) > len(re.sub(r"\s", "", page_text)):
                        page_text = ocr_text
                        paginas_ocr.append(numero)
                except ImportError:
                    # Mantener texto digital disponible; si todo es imagen, informar.
                    pass
                except Exception as exc:
                    # Una página defectuosa no debe descartar las demás.
                    import logging
                    logging.getLogger(__name__).warning("OCR falló en página %s: %s", numero, exc)
            pages_text.append(page_text)
    finally:
        doc.close()

    text = "\n\n".join(pages_text).strip()
    if len(re.sub(r"\s", "", text)) < 10:
        raise RuntimeError("No se detectó texto legible. Comprueba la calidad del PDF y que RapidOCR esté instalado.")
    method = "PyMuPDF + RapidOCR (páginas: " + ", ".join(map(str, paginas_ocr)) + ")" if paginas_ocr else "PyMuPDF (Columna por Columna)"
    return text, method


def clean_value(v):
    v = re.sub(r'[ \t]+', ' ', v).strip()
    return v.strip(' :\t')


def extraer_valores_teoricos_hematologia(texto):
    """Extrae exclusivamente cifras/rangos en mg/dL de la sección Valores teóricos.

    No incluye títulos ni explicaciones; exige ambos límites de sección.
    """
    if not texto:
        return ""
    # Normalizar tildes sin alterar las posiciones de los caracteres.
    tabla = str.maketrans("áéíóúÁÉÍÓÚ", "aeiouAEIOU")
    normalizado = str(texto).translate(tabla)
    inicio = re.search(r"\bvalores\s+teoricos\b\s*[:：]?", normalizado, re.I)
    if not inicio:
        return ""
    resto = normalizado[inicio.end():]
    fin = re.search(r"\bdatos\s+especificos\s+de\s+funcionamiento\s+del\s+test\b", resto, re.I)
    if not fin:
        return ""
    seccion = str(texto)[inicio.end():inicio.end() + fin.start()]
    numero = r"\d+(?:[.,]\d+)?"
    # Permite cifras aisladas, intervalos (5-10, 5 a 10) y comparadores.
    patron = re.compile(
        rf"(?<![\w.,])(?:[<>≤≥]\s*)?{numero}"
        rf"(?:\s*(?:[-‐‑‒–—―−﹣－]|\ba\b|\bhasta\b)\s*{numero})?"
        r"\s*mg\s*/\s*d[lL]\b", re.I
    )
    valores = []
    vistos = set()
    for coincidencia in patron.finditer(seccion):
        valor = re.sub(r"\s+", " ", coincidencia.group()).strip()
        clave = valor.casefold()
        if clave not in vistos:
            vistos.add(clave)
            valores.append(valor)
    return " | ".join(valores)


def es_alias_hematologia(field_aliases):
    return field_aliases is FIELD_ALIASES_HEMATOLOGIA or (
        isinstance(field_aliases, dict)
        and field_aliases.get("Valores de referencia") == FIELD_ALIASES_HEMATOLOGIA["Valores de referencia"]
    )


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
            # Buscar separador después de la etiqueta, sin depender de
            # la longitud original de la variante acentuada.
            match = re.match(r"^(.+?)\s*[:：–—-]\s*(.*)$", s)
            if match and norm(match.group(1)) == na:
                matched_field = field
                remainder = match.group(2).strip()
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
        val_limpio = limpiar_valor_excel(v, examen_activo)
        if k == 'Valores de referencia' and not es_alias_hematologia(field_aliases):
            val_limpio = extraer_solo_valor_referencia(val_limpio)
        found[k] = val_limpio

    if es_alias_hematologia(field_aliases):
        seccion = extraer_valores_teoricos_hematologia(text)
        if seccion:
            found['Valores de referencia'] = seccion
        else:
            # No usar el texto parcial de un apartado sin límite final verificado.
            found.pop('Valores de referencia', None)

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
        if not value or str(value).strip().casefold() in ["nan", "none", "null"]:
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


def consolidar_documentos_pdf(documentos, field_aliases, nombre_examen):
    """Extrae cada PDF por separado y completa únicamente los campos vacíos.

    Las diferencias entre valores no se resuelven automáticamente: se devuelven
    para revisión del personal de laboratorio.
    """
    if not documentos:
        raise ValueError("Debes seleccionar al menos un documento PDF.")
    datos = {}
    fuentes = {}
    discrepancias = {}
    textos = []
    metodos = []
    for nombre, contenido in documentos:
        texto, metodo = extract_pdf_text(contenido, nombre, nombre_examen)
        encontrados = infer_name_code(texto, parse_labeled_text(texto, field_aliases, nombre_examen))
        encontrados = {campo: val for campo, val in encontrados.items() if campo not in fd.CAMPOS_CONTROLADOS}
        textos.append(f"===== {nombre} =====\n{texto}")
        metodos.append(f"{nombre}: {metodo}")
        for campo, valor in encontrados.items():
            valor = str(valor or "").strip()
            if not valor or valor.casefold() in ("nan", "none", "null"):
                continue
            if not str(datos.get(campo, "")).strip():
                datos[campo] = valor
                fuentes[campo] = nombre
            elif norm(str(datos[campo])) != norm(valor) and valor != str(datos[campo]):
                opciones = discrepancias.setdefault(campo, [(fuentes[campo], str(datos[campo]))])
                if (nombre, valor) not in opciones:
                    opciones.append((nombre, valor))
    return datos, fuentes, discrepancias, "\n\n".join(textos), " | ".join(metodos)


def process(uploaded_bytes, filename, template_bytes, centro='', nombre_examen='', field_aliases=None,
            actualizado_por='', autorizado_por='', datos_sobrescribir=None, documentos_adicionales=None):
    if field_aliases is None:
        field_aliases = FIELD_ALIASES_BASE
    if not nombre_examen:
        raise ValueError('Selecciona un examen desde el menú del área antes de analizar el PDF.')
    ok, mensaje = nombre_examen_valido(nombre_examen)
    if not ok:
        raise ValueError(mensaje)

    if documentos_adicionales:
        data, _, _, text, method = consolidar_documentos_pdf(
            [(filename, uploaded_bytes)] + documentos_adicionales, field_aliases, nombre_examen
        )
    else:
        text, method = extract_pdf_text(uploaded_bytes, filename, nombre_examen)
        data = infer_name_code(text, parse_labeled_text(text, field_aliases, nombre_examen))

    # Los ocho campos documentales proceden EXCLUSIVAMENTE de SQLite.
    data = {campo: val for campo, val in data.items() if campo not in fd.CAMPOS_CONTROLADOS}
    if datos_sobrescribir:
        for campo, val in datos_sobrescribir.items():
            if val and str(val).strip().casefold() not in ["nan", "none", "null"]:
                data[campo] = val

    ZONA_HORARIA = ZoneInfo("America/Santiago")
    fecha_actual = datetime.now(ZONA_HORARIA).strftime('%d/%m/%Y %H:%M')

    data['Nombre LIS'] = nombre_examen.strip()
    data['Laboratorio de procesamiento'] = (centro or '').strip()
    data['Actualizado por'] = (actualizado_por or '').strip()
    data['Autorizado por'] = (autorizado_por or '').strip()
    data['Fecha actualización'] = fecha_actual

    if ('Valores de referencia' in data and data['Valores de referencia']
            and not es_alias_hematologia(field_aliases)):
        data['Valores de referencia'] = extraer_solo_valor_referencia(data['Valores de referencia'])

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

    # Evitar datos obsoletos si una ficha anterior ya contenía alguno de estos campos.
    for fila in range(1, ws.max_row + 1):
        if str(ws.cell(fila, 1).value or "").strip() in fd.CAMPOS_CONTROLADOS:
            ws.cell(fila, 2).value = None
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


def actualizar_campo_excel(excel_bytes, nombre_examen, campo, valor):
    """Actualiza un único campo de la ficha sin modificar las demás hojas."""
    wb = load_workbook(io.BytesIO(excel_bytes))
    if nombre_examen not in wb.sheetnames:
        raise ValueError("No se encontró la hoja del examen seleccionado.")
    ws = wb[nombre_examen]
    for fila in range(1, ws.max_row + 1):
        if norm(str(ws.cell(fila, 1).value or "")) == norm(campo):
            ws.cell(fila, 2).value = limpiar_valor_excel(valor, nombre_examen)
            break
    else:
        # Las plantillas antiguas pueden no incluir todos los campos del área.
        # Solo se permite añadir un campo reconocido en la configuración actual.
        area_actual = st.session_state.get("area_resultado", st.session_state.get("area_activa"))
        campos_validos = get_field_aliases(area_actual)
        if not any(norm(c) == norm(campo) for c in campos_validos):
            raise ValueError(f"El campo '{campo}' no está configurado para el área seleccionada.")
        fila = ws.max_row + 1
        ws.cell(fila, 1).value = campo
        ws.cell(fila, 2).value = limpiar_valor_excel(valor, nombre_examen)
        ws.cell(fila, 2).alignment = Alignment(vertical='top', wrap_text=True)
        ajustar_ancho_columnas(ws)
    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue()


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
            val_raw = row_data["Información extraída"]
            if pd.isna(val_raw) or val_raw is None:
                val_editado = ""
            else:
                val_editado = str(val_raw).strip()
                if val_editado.casefold() in ["nan", "none", "null"]:
                    val_editado = ""
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
                        datos_examen[campo] = "" if valor.casefold() in ["nan", "none", "null"] else valor
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
                tipo = "Campo nuevo incorporado"
            elif val_ant and not val_nue:
                tipo = "Campo eliminado / borrado"
            else:
                tipo = "Información modificada"

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
        ATENCIÓN: LA CARTERA DE EXÁMENES SUPERA LOS 6 MESES DE ANTIGÜEDAD<br>
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

@st.dialog("Alerta de mantenimiento: Cartera de Exámenes")
def modal_alerta_cartera():
    st.warning("Han transcurrido más de 6 meses desde la última actualización de la Cartera de Exámenes.")
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
        if st.button("Vincular examen seleccionado", type="primary", use_container_width=True):
            nombre_elegido = candidatos[idx_seleccionado][0]
            st.session_state["cartera_examen_emparejado"] = nombre_elegido
            st.session_state["_sel_cartera_pendiente"] = nombre_elegido
            st.session_state.pop("candidatos_cartera_pendientes", None)
            st.success(f"Vinculado con «{nombre_elegido}»!")
            st.rerun()

    with col_c2:
        if st.button("No vincular ninguno", use_container_width=True):
            st.session_state["cartera_examen_emparejado"] = "Sin vincular"
            st.session_state["acreditacion_manual_id"] = None
            st.session_state["_acreditacion_reset_pendiente"] = True
            st.session_state["_sel_cartera_pendiente"] = None
            st.session_state.pop("candidatos_cartera_pendientes", None)
            st.rerun()


def _conflictos_cartera_actualizados(datos):
    """Compara contra la cartera DESPUÉS de decidir los valores entre PDF."""
    vinculo = st.session_state.get("cartera_examen_emparejado")
    info = obtener_info_cartera_por_nombre(vinculo) if vinculo and vinculo != "Sin vincular" else None
    conflictos = {}
    if info:
        for campo in ("Equipamiento utilizado", "Plazo entrega resultados"):
            pdf = str(datos.get(campo, "") or "").strip()
            cartera = str(info.get(campo, "") or "").strip()
            if pdf and cartera and pdf.casefold() != cartera.casefold():
                conflictos[campo] = {"pdf": pdf, "cartera": cartera}
    return conflictos


def _guardar_resolucion_en_ficha(datos, campos_resueltos):
    """La resolución de discrepancias es extracción validada, no edición manual."""
    area = st.session_state.get("area_resultado", st.session_state.get("area_activa"))
    campos = get_field_aliases(area)
    anteriores = set(st.session_state.get("user_edited_fields", set()))
    # Una selección de fuentes PDF/Cartera es validación de extracción.
    # No debe colorearse como una edición directa del usuario.
    anteriores.difference_update(campos_resueltos)
    filas = []
    for campo in campos:
        valor = datos.get(campo, "")
        if valor is None or str(valor).strip().casefold() in ("nan", "none", "null"):
            valor = ""
        filas.append({
            "Campo": campo.strip(),
            "Información extraída": valor,
            "Estado": ("Modificado por usuario" if campo in anteriores
                       else "Encontrado" if valor else "No encontrado"),
        })
    nuevo_excel = aplicar_ediciones_a_excel(
        st.session_state["result_bytes"], pd.DataFrame(filas),
        st.session_state.get("examen_activo", ""))
    st.session_state["result_bytes"] = nuevo_excel
    st.session_state["data_extracted"] = datos
    st.session_state["user_edited_fields"] = anteriores


def _mensaje_discrepancias(cantidad, origen):
    """Mensaje correcto para uno o varios campos."""
    if cantidad == 1:
        return f"Se detectó una discrepancia en 1 campo entre {origen}."
    return f"Se detectaron discrepancias en {cantidad} campos entre {origen}."


def _tarjeta_valor_discrepancia(titulo, valor, color):
    """Tarjeta visual: encabezado coloreado y valor blanco en negrita."""
    paleta = {
        "azul": ("#193149", "#49a8ff"),
        "verde": ("#163a2b", "#5de398"),
    }
    fondo, encabezado = paleta[color]
    texto = html.escape(str(valor or "(Vacío)")) .replace("\n", "<br>")
    titulo_seguro = html.escape(str(titulo))
    st.markdown(
        f'<div style="background:{fondo};padding:16px;border-radius:8px;'
        f'min-height:155px;overflow-wrap:anywhere;">'
        f'<div style="color:{encabezado};font-size:0.96rem;line-height:1.5;">{titulo_seguro}</div>'
        f'<div style="color:#ffffff;font-weight:700;margin-top:14px;'
        f'line-height:1.55;white-space:normal;">{texto}</div></div>',
        unsafe_allow_html=True,
    )


@st.dialog("Comparación y Resolución de Discrepancias")
def modal_discrepancias_unificadas():
    """Una ventana y un botón: primero PDF vs PDF, después PDF vs Cartera."""
    pdf_diferencias = st.session_state.get("discrepancias_pdf", {})
    cartera_diferencias = st.session_state.get("conflictos_pendientes", {})
    etapa = "pdf" if pdf_diferencias else "cartera"
    if etapa == "pdf":
        st.markdown(_mensaje_discrepancias(len(pdf_diferencias), "los documentos PDF"))
        st.write("Revisa cada campo y selecciona la información que deseas conservar. **Etapa 1: PDF vs PDF**.")
        diferencias = pdf_diferencias
    else:
        st.markdown(_mensaje_discrepancias(len(cartera_diferencias), "el PDF y la Cartera de Exámenes"))
        st.write("Por favor, revisa cada campo afectado y selecciona qué información deseas conservar. **Etapa 2: PDF vs Cartera**.")
        diferencias = cartera_diferencias
    if not diferencias:
        st.info("No hay discrepancias pendientes por resolver.")
        if st.button("Cerrar", key="cerrar_dialogo_unificado"):
            st.session_state["modal_discrepancias_unificado"] = False
            st.rerun()
        return
    st.write("")
    respuestas = {}
    for campo, detalle in diferencias.items():
        st.markdown(f"""<div style="background-color: #f1f5f9; padding: 10px 14px; border-radius: 6px; border-left: 5px solid #0284c7; margin-top: 10px; margin-bottom: 12px;"><span style="font-size: 1.05rem; font-weight: bold; color: #0f172a;">Campo evaluado: <span style="color: #0369a1;">{html.escape(campo)}</span></span></div>""", unsafe_allow_html=True)
        # Ambos tipos de discrepancia comparten dos tarjetas y dos selectores.
        # Cada selector se renderiza DENTRO de la columna de su tarjeta.
        # Los callbacks actualizan el estado sin llamar manualmente a st.rerun().
        if etapa == "pdf":
            opciones = list(detalle)
        else:
            opciones = [
                ("PDF extraído", detalle["pdf"]),
                ("Cartera SQLite", detalle["cartera"]),
            ]

        clave = ("unificado_pdf_" if etapa == "pdf" else "unificado_cartera_") + hashlib.sha256(campo.encode()).hexdigest()[:16]
        if st.session_state.get(clave) not in range(len(opciones)):
            st.session_state[clave] = 0

        def _seleccionar_discrepancia(k, indice):
            st.session_state[k] = indice

        columnas = st.columns(len(opciones), gap="medium")
        for indice, (nombre, valor) in enumerate(opciones):
            with columnas[indice]:
                if etapa == "pdf":
                    titulo = f"Valor del PDF ({nombre}):"
                elif indice == 0:
                    titulo = f"Valor extraído del PDF ({campo}):"
                else:
                    titulo = f"Valor en Cartera SQLite ({campo}):"
                _tarjeta_valor_discrepancia(
                    titulo, valor, "azul" if indice % 2 == 0 else "verde"
                )
                st.markdown('<div style="height:10px;"></div>', unsafe_allow_html=True)
                st.button(
                    "🔴" if st.session_state[clave] == indice else "⚪",
                    key=f"elegir_{etapa}_{hashlib.sha256(campo.encode()).hexdigest()[:12]}_{indice}",
                    use_container_width=True,
                    on_click=_seleccionar_discrepancia,
                    args=(clave, indice),
                )
        respuestas[campo] = st.session_state[clave]
        st.divider()
    col1, col2 = st.columns([2, 1])
    with col1:
        etiqueta = "Aplicar y continuar" if etapa == "pdf" else "Aplicar y actualizar ficha"
        if st.button(etiqueta, type="primary", use_container_width=True, key="aplicar_unificado"):
            datos = copy.deepcopy(st.session_state["data_extracted"])
            for campo, eleccion in respuestas.items():
                if etapa == "pdf":
                    datos[campo] = diferencias[campo][eleccion][1]
                else:
                    pdf, cartera = diferencias[campo]["pdf"], diferencias[campo]["cartera"]
                    datos[campo] = pdf if eleccion == 0 else cartera
            try:
                _guardar_resolucion_en_ficha(datos, set(respuestas))
            except Exception as exc:
                st.error(f"No se pudieron aplicar las selecciones: {exc}")
                return
            if etapa == "pdf":
                st.session_state.pop("discrepancias_pdf", None)
                # No reutilizar conflictos calculados con valores anteriores.
                nuevos = _conflictos_cartera_actualizados(datos)
                if nuevos:
                    st.session_state["conflictos_pendientes"] = nuevos
                    st.session_state["modal_discrepancias_unificado"] = True
                else:
                    st.session_state.pop("conflictos_pendientes", None)
                    st.session_state["modal_discrepancias_unificado"] = False
            else:
                st.session_state.pop("conflictos_pendientes", None)
                st.session_state["modal_discrepancias_unificado"] = False
            st.rerun()
    with col2:
        if st.button("Cancelar", use_container_width=True, key="cancelar_unificado"):
            st.session_state["modal_discrepancias_unificado"] = False
            st.rerun()


@st.dialog("Confirmar actualización de la planilla del área en la herramienta")
def modal_confirmar_actualizacion_planilla():
    area_activa = st.session_state.get("area_activa")
    examen_activo = st.session_state.get("examen_activo")
    comp = st.session_state.get("comparacion_planilla_detalles", {})

    st.markdown(f"Estás a punto de reemplazar la **base de datos guardada en SQLite** para el área **{area_activa}** con los datos recién procesados del examen **{examen_activo}**.")
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
        st.markdown("##### Resumen de diferencias detectadas:")
        df_dif = pd.DataFrame(diferencias)
        st.table(df_dif[["Campo", "Tipo de cambio", "Valor anterior", "Valor nuevo"]])
    else:
        st.info("No se detectaron diferencias en los campos. La versión actual y la nueva son idénticas.")

    st.write("")
    st.warning("Esta acción sobreescribirá la Ficha Técnica / Base de Datos del área en SQLite para que los próximos análisis utilicen este nuevo contenido.")

    col_save1, col_save2 = st.columns([2, 1])

    with col_save1:
        if st.button("Confirmar y actualizar base de datos del área", type="primary", use_container_width=True):
            try:
                nombre_generado = (f"Ficha_Tecnica_Actualizada_{area_activa.replace(' ', '_')}_"
                                   f"{datetime.now(ZoneInfo('America/Santiago')).strftime('%Y-%m-%d')}.xlsx")
                guardar_plantilla_area_db(area_activa, st.session_state["result_bytes"], nombre_generado)

                if comp.get("es_nueva_hoja") and examen_activo:
                    actuales = st.session_state["examenes_por_area"].get(area_activa, [])
                    if examen_activo.casefold() not in {x.casefold() for x in actuales}:
                        actuales.append(examen_activo)
                        st.session_state["examenes_por_area"][area_activa] = sorted(actuales, key=str.casefold)
                        guardar_examenes_por_area(st.session_state["examenes_por_area"])

                st.session_state["modal_confirmar_planilla_abierto"] = False
                st.session_state.pop("comparacion_planilla_detalles", None)
                st.success(f"La base de datos de «{area_activa}» fue actualizada exitosamente en SQLite!")
                st.rerun()

            except Exception as err_save:
                st.error(f"Error al guardar la planilla en la herramienta: {err_save}")

    with col_save2:
        if st.button("Cancelar", use_container_width=True):
            st.session_state["modal_confirmar_planilla_abierto"] = False
            st.rerun()


@st.dialog("Seleccionar y administrar exámenes")
def modal_examenes_area(area):
    st.subheader(area)
    examenes = st.session_state["examenes_por_area"].get(area, [])
    actual = st.session_state.get("examen_activo") if st.session_state.get("area_activa") == area else None
    
    key_sel = f"selector_examen_{area}"
    if key_sel not in st.session_state and actual in examenes:
        st.session_state[key_sel] = actual

    seleccionado = st.selectbox(
        "Examen de esta área", 
        examenes, 
        index=None, 
        placeholder="Selecciona un examen",
        key=key_sel,
        help="El nombre elegido se utilizará exactamente en el campo Nombre LIS y como nombre de la hoja de Excel."
    )
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Usar examen", key=f"usar_examen_{area}", type="primary", use_container_width=True):
            if seleccionado:
                st.session_state["area_activa"] = area
                st.session_state["examen_activo"] = seleccionado
                st.session_state["cartera_examen_emparejado"] = "Sin vincular"
                st.session_state["acreditacion_manual_id"] = None
                st.session_state["_acreditacion_reset_pendiente"] = True
                st.session_state["_sel_cartera_pendiente"] = None
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
            st.session_state[key_sel] = None
            st.session_state["cartera_examen_emparejado"] = "Sin vincular"
            st.session_state["acreditacion_manual_id"] = None
            st.session_state["_acreditacion_reset_pendiente"] = True
            st.session_state["_sel_cartera_pendiente"] = None
            st.session_state.pop("candidatos_cartera_pendientes", None)
            st.session_state["mostrar_modal_examenes"] = False
            st.session_state.pop("modal_area_activa", None)
            st.rerun()

    st.divider()
    with st.expander("Añadir examen"):
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
                    st.session_state[key_sel] = nuevo
                    st.session_state["cartera_examen_emparejado"] = "Sin vincular"
                    st.session_state["acreditacion_manual_id"] = None
                    st.session_state["_acreditacion_reset_pendiente"] = True
                    st.session_state["_sel_cartera_pendiente"] = None
                    st.session_state.pop("candidatos_cartera_pendientes", None)
                    st.session_state["mostrar_modal_examenes"] = False
                    st.session_state.pop("modal_area_activa", None)
                    st.success(f"Se agregó «{nuevo}».")
                    st.rerun()
                else:
                    st.error("No se pudo guardar la lista en la base de datos.")
    with st.expander("Eliminar examen de la lista"):
        actuales = st.session_state["examenes_por_area"].get(area, [])
        if actuales:
            quitar = st.selectbox(
                "Examen que deseas quitar", 
                actuales, 
                index=None,
                placeholder="Selecciona un examen para eliminar",
                key=f"quitar_examen_{area}"
            )
            confirmar = st.checkbox("Confirmo que quiero quitarlo del menú", key=f"confirmar_quitar_{area}")
            if st.button("Eliminar examen", key=f"eliminar_examen_{area}", disabled=not confirmar or not quitar):
                st.session_state["examenes_por_area"][area] = [x for x in actuales if x.casefold() != quitar.casefold()]
                if st.session_state.get("examen_activo") == quitar and st.session_state.get("area_activa") == area:
                    st.session_state["examen_activo"] = None
                    st.session_state[key_sel] = None
                    st.session_state["cartera_examen_emparejado"] = "Sin vincular"
                    st.session_state["acreditacion_manual_id"] = None
                    st.session_state["_acreditacion_reset_pendiente"] = True
                    st.session_state["_sel_cartera_pendiente"] = None
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
    st.session_state["acreditacion_manual_id"] = None
    st.session_state["_acreditacion_reset_pendiente"] = True
    st.session_state["mostrar_modal_examenes"] = False
    st.session_state.pop("modal_area_activa", None)


# --- GESTOR DE DIÁLOGOS EMERGENTES ---
def ejecutar_gestor_dialogos():
    """Ejecuta como máximo UN solo diálogo emergente por ciclo para evitar errores de Streamlit."""
    if st.session_state.get("modal_confirmar_planilla_abierto"):
        st.session_state["modal_confirmar_planilla_abierto"] = False
        modal_confirmar_actualizacion_planilla()
    elif st.session_state.get("modal_discrepancias_unificado"):
        modal_discrepancias_unificadas()
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
    st.header("Áreas del laboratorio")
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

    st.header("Fichas técnicas")
    template = st.file_uploader("Cargar Excel del área (opcional)", type=["xlsx"], key="template_uploader")
    st.caption("Si no cargas uno, se utilizará la plantilla almacenada en la base de datos para el área.")
    st.divider()

    # --- SECCIÓN BASE DE DATOS (SQLITE) SIN EMOJIS ---
    st.header("Base de datos (SQLite)")

    # Áreas del laboratorio almacenadas en SQLite (mantener el orden de AREA_CONFIG)
    for area_item in AREA_CONFIG:
        b_bytes_area, fecha_act_area = obtener_plantilla_area_db(area_item)
        with st.expander(f"Área: {area_item}", expanded=False):
            st.caption(f"Última actualización: **{fecha_act_area}**")
            nombre_plantilla = obtener_nombre_plantilla_area_db(area_item)
            if nombre_plantilla:
                st.caption(f"Documento almacenado: **{nombre_plantilla}**")
            elif b_bytes_area:
                st.caption("Documento de origen: no registrado en esta versión anterior de la base de datos.")

            if b_bytes_area:
                ZONA_HORARIA = ZoneInfo("America/Santiago")
                fecha_sb = datetime.now(ZONA_HORARIA).strftime("%d-%m-%Y")
                area_clean_sb = area_item.replace(" ", "_")
                st.download_button(
                    f"Descargar ficha técnica ({area_item})",
                    data=b_bytes_area,
                    file_name=nombre_plantilla or f"Ficha_Técnica_Actualizada_{area_clean_sb}_{fecha_sb}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"dl_area_btn_{area_item}"
                )
            else:
                st.info("No hay Ficha Técnica/Plantilla registrada en la base de datos para esta área.")

            area_file = st.file_uploader(f"Cargar nueva plantilla de {area_item} (.xlsx)", type=["xlsx"], key=f"uploader_area_{area_item}")
            if area_file is not None:
                if st.button(f"Guardar {area_item} en Base de Datos", type="primary", use_container_width=True, key=f"btn_save_area_{area_item}"):
                    ok_area, msg_area = guardar_plantilla_area_db(area_item, area_file.getvalue(), area_file.name)
                    if ok_area:
                        st.session_state["examenes_por_area"] = cargar_examenes_por_area()
                        st.success(msg_area)
                        st.rerun()
                    else:
                        st.error(msg_area)

    # Fuentes documentales al final del menú de la base de datos.
    # Cada origen conserva su documento, fecha, descarga y actualización en SQLite.
    with st.expander("Cartera de exámenes y acreditación", expanded=cartera_expirada):
        tab_acreditacion, tab_examenes = st.tabs(["Acreditación", "Exámenes"])

        with tab_acreditacion:
            meta_acr = fd.documento_fuente(DB_PATH, fd.ACREDITACION_ID, incluir_archivo=True)
            if meta_acr:
                st.caption(f"Última actualización: **{formatear_fecha_fuente(meta_acr['fecha_actualizacion'])}**")
                st.caption(f"Documento almacenado: **{meta_acr['nombre_archivo']}** · {meta_acr['registros']} prestaciones")
                st.download_button(
                    "Descargar cartera de acreditación original",
                    data=meta_acr["archivo"],file_name=meta_acr["nombre_archivo"],
                    mime=meta_acr["tipo_mime"],use_container_width=True,
                    key="descargar_acreditacion_original")
            else:
                st.warning("No está cargada la Cartera de Acreditación.")
            nuevo_acr = st.file_uploader(
                "Cargar nueva Cartera de Acreditación (.xlsx)",type=["xlsx"],
                key="actualizar_acreditacion_archivo")
            if nuevo_acr is not None and st.button(
                    "Guardar nueva Cartera de Acreditación",key="guardar_acreditacion_nueva",
                    type="primary",use_container_width=True):
                try:
                    total = fd.guardar_cartera_acreditacion(DB_PATH,nuevo_acr.getvalue(),nuevo_acr.name)
                    st.session_state["acreditacion_manual_id"] = None
                    st.session_state["_acreditacion_reset_pendiente"] = True
                    st.success(f"Cartera actualizada: {total} prestaciones indexadas.")
                    st.rerun()
                except Exception as ex:
                    st.error(f"No se reemplazó la Cartera: {ex}")


        with tab_examenes:
            st.caption(f"Última actualización: **{msg_cartera}**")
            meta_cartera_original = fd.documento_fuente(DB_PATH, "cartera_examenes", incluir_archivo=True)
            if meta_cartera_original:
                st.caption(f"Documento almacenado: **{meta_cartera_original['nombre_archivo']}**")
            else:
                st.caption("Documento de origen: no registrado en esta versión anterior de la base de datos.")

            b_bytes_cartera = (meta_cartera_original["archivo"] if meta_cartera_original
                               else obtener_cartera_excel_bytes())
            if b_bytes_cartera:
                ZONA_HORARIA = ZoneInfo("America/Santiago")
                fecha_cartera_dl = datetime.now(ZONA_HORARIA).strftime("%d-%m-%Y")
                st.download_button(
                    "Descargar cartera de exámenes",
                    data=b_bytes_cartera,
                    file_name=(meta_cartera_original['nombre_archivo'] if meta_cartera_original
                               else f"Cartera_de_Exámenes_Actualizada_{fecha_cartera_dl}.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="dl_cartera_btn"
                )
            else:
                st.info("No hay exámenes registrados en la Cartera de la base de datos.")

            cartera_file = st.file_uploader("Cargar nueva Cartera (.xlsx)", type=["xlsx"], key="cartera_uploader_sidebar")
            if cartera_file is not None:
                if st.button("Guardar en Base de Datos", type="primary", use_container_width=True, key="btn_save_cartera"):
                    try:
                        df_cartera = pd.read_excel(cartera_file)
                        ok_db, msg_db = guardar_cartera_db(df_cartera, cartera_file.name, cartera_file.getvalue())
                        if ok_db:
                            st.success(msg_db)
                            st.session_state["mostrar_dialogo_expiracion"] = False
                            st.rerun()
                        else:
                            st.error(msg_db)
                    except Exception as err_db:
                        st.error(f"Error al procesar la Cartera: {err_db}")


    # Manual y Normas generales: fuentes independientes bajo un solo apartado.
    # Las pestañas mantienen el orden alfabético (Manual, Normas generales).
    with st.expander("Manual de toma de muestras", expanded=False):
        tab_manual, tab_normas = st.tabs(["Manual", "Normas generales"])

        with tab_manual:
            meta_manual = fd.documento_fuente(DB_PATH, fd.MANUAL_ID, incluir_archivo=True)
            if meta_manual:
                st.caption(f"Última actualización: **{formatear_fecha_fuente(meta_manual['fecha_actualizacion'])}**")
                st.caption(f"Documento almacenado: **{meta_manual['nombre_archivo']}** · 2 campos predeterminados")
                st.download_button(
                    "Descargar manual de toma de muestras original",
                    data=meta_manual["archivo"], file_name=meta_manual["nombre_archivo"],
                    mime=meta_manual["tipo_mime"], use_container_width=True,
                    key="descargar_manual_toma_original")
            else:
                st.warning("No está cargado el Manual de toma de muestras.")
            nuevo_manual = st.file_uploader(
                "Cargar nuevo Manual de toma de muestras (.pdf)", type=["pdf"],
                key="actualizar_manual_toma_archivo")
            if nuevo_manual is not None and st.button(
                    "Guardar nuevo Manual en SQLite", key="guardar_manual_toma_nuevo",
                    type="primary", use_container_width=True):
                try:
                    total = fd.guardar_manual_toma(DB_PATH, nuevo_manual.getvalue(), nuevo_manual.name)
                    st.success(f"Manual actualizado: {total} apartados reconocidos y depurados.")
                    st.rerun()
                except Exception as ex:
                    st.error(f"No se reemplazó el Manual: {ex}")

        with tab_normas:
            st.caption("Normas generales de toma de muestras")
            meta_normas = fd.documento_fuente(DB_PATH, fd.NORMAS_ID, incluir_archivo=True)
            if meta_normas:
                st.caption(f"Última actualización: **{formatear_fecha_fuente(meta_normas['fecha_actualizacion'])}**")
                st.caption(f"Documento almacenado: **{meta_normas['nombre_archivo']}** · 4 campos predeterminados")
                st.download_button(
                    "Descargar documento original de normas",
                    data=meta_normas["archivo"],
                    file_name=meta_normas["nombre_archivo"],
                    mime=meta_normas["tipo_mime"],
                    use_container_width=True, key="descargar_normas_original")
            else:
                st.warning("No están cargadas las normas generales.")
            nuevo_normas = st.file_uploader(
                "Cargar nuevas Normas Generales (.docx o .pdf)", type=["docx", "pdf"],
                key="actualizar_normas_archivo")
            if nuevo_normas is not None and st.button(
                    "Guardar nuevas Normas en SQLite", key="guardar_normas_nuevas",
                    type="primary", use_container_width=True):
                try:
                    total = fd.guardar_normas(DB_PATH, nuevo_normas.getvalue(), nuevo_normas.name)
                    st.success(f"Normas actualizadas: {total} apartados reconocidos.")
                    st.rerun()
                except Exception as ex:
                    st.error(f"No se reemplazó el documento: {ex}")

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
    st.info("### ¡Bienvenido/a a la Herramienta de Extracción de Datos!")
    
    st.markdown(
        """
        Para comenzar a trabajar, selecciona un **área del laboratorio** en el panel lateral de la izquierda.

        #### Pasos para trabajar:
        1. **Selecciona un Área del laboratorio** en el menú de la izquierda (*Microbiología, Hematología y coagulación, Bioquímica, Hormonas e inmunidad, Derivados*).
        2. **Administra y selecciona el examen** que deseas procesar usando el botón **Administrar exámenes**.
        3. **Carga las fichas técnicas en PDF** (lote de hasta 5 documentos simultáneos) y organiza la cola de procesamiento si lo deseas.
        4. **Define los parámetros operacionales** (*Laboratorio de procesamiento, Actualizado por, Autorizado por*) y verifica el emparejamiento con la Cartera en SQLite.
        5. **Haz clic en "Analizar documento activo"** para realizar la extracción automática de texto, OCR y cruce con la Base Maestra de Cartera.
        6. **Revisa y edita la información** en la tabla interactiva, resuelve discrepancias si existen, y actualiza la planilla del área o descarga el Excel consolidado.
        """
    )
    
    st.divider()
    
    # --- MÓDULO DE ESTADO DE LA CARTERA SQLITE EN LA PÁGINA DE BIENVENIDA ---
    st.markdown("#### Estado de la Cartera SQLite:")
    lista_cartera_home = obtener_lista_cartera_db()
    total_examenes_cartera = len(lista_cartera_home)
    fecha_cartera_home, cartera_expirada_home, msg_cartera_home = fecha_actualizacion_cartera()

    if fecha_cartera_home is None or total_examenes_cartera == 0:
        st.warning("Cartera desactualizada: No se ha subido ninguna Cartera de Exámenes a la base de datos. Se sugiere cargar el archivo Excel en el menú lateral.")
    elif cartera_expirada_home:
        st.warning(f"Cartera desactualizada (más de 6 meses de antigüedad): Registrada el **{msg_cartera_home}**. Se recomienda actualizar la base de datos subiendo el archivo Excel desde el panel lateral.")
    else:
        st.success(f"Cartera de Exámenes cargada y lista: Última actualización registrada el **{msg_cartera_home}**.")

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
            if st.button("Administrar exámenes", type="primary", use_container_width=True, key="btn_abrir_admin_examenes"):
                st.session_state["modal_area_activa"] = area_activa
                st.session_state["mostrar_modal_examenes"] = True
                st.rerun()

        modo_procesamiento = st.radio(
            "Modalidad de análisis",
            ["Un PDF por examen (cola actual)", "Varios PDF para un mismo examen"],
            horizontal=True,
            key="modo_procesamiento_pdf",
            help="En el segundo modo todos los PDF completan una sola ficha del examen seleccionado."
        )
        modo_conjunto = modo_procesamiento == "Varios PDF para un mismo examen"
        if st.session_state.get("modo_pdf_anterior") != modo_procesamiento:
            st.session_state["modo_pdf_anterior"] = modo_procesamiento
            st.session_state["cola_pdf_indice"] = 0
            for clave in ("data_extracted", "result_bytes", "extracted_text", "user_edited_fields", "discrepancias_pdf", "fuentes_pdf", "modal_pdf_pdf_abierto", "modal_discrepancias_unificado"):
                st.session_state.pop(clave, None)

        # CARGA DE ARCHIVOS PDF Y CONTROL DE COLA
        pdfs = st.file_uploader(
            (f"Cargar hasta {MAX_PDFS_PERMITIDOS} PDF del mismo examen" if modo_conjunto
             else f"Cargar fichas técnicas (Lote de hasta {MAX_PDFS_PERMITIDOS} PDF para procesar en cola)"),
            type=["pdf"],
            accept_multiple_files=True,
            help=("Todos los archivos se analizan juntos para completar una sola ficha." if modo_conjunto
                  else f"Puedes cargar hasta {MAX_PDFS_PERMITIDOS} archivos PDF para procesarlos de forma individual uno tras otro."),
            key="pdf_uploader"
        )

        exceso_archivos = pdfs is not None and len(pdfs) > MAX_PDFS_PERMITIDOS

        if exceso_archivos:
            st.error(f"Has seleccionado **{len(pdfs)} archivos**. El límite máximo simultáneo es de **{MAX_PDFS_PERMITIDOS} documentos PDF** para evitar saturar el sistema.")

        if pdfs and not exceso_archivos:
            mapa_pdfs = {pdf.name: pdf for pdf in pdfs}
            nombres_actuales = [pdf.name for pdf in pdfs]

            if ("orden_nombres_pdf" not in st.session_state or 
                set(st.session_state["orden_nombres_pdf"]) != set(nombres_actuales)):
                
                orden_previo = st.session_state.get("orden_nombres_pdf", [])
                nuevo_orden_filtrado = [nombre for nombre in orden_previo if nombre in mapa_pdfs]
                
                for nombre in nombres_actuales:
                    if nombre not in nuevo_orden_filtrado:
                        nuevo_orden_filtrado.append(nombre)
                
                st.session_state["orden_nombres_pdf"] = nuevo_orden_filtrado
                st.session_state["cola_pdf_indice"] = 0
                st.session_state.pop("data_extracted", None)
                st.session_state.pop("user_edited_fields", None)
                st.session_state.pop("extracted_text", None)
                st.session_state.pop("result_bytes", None)
                st.session_state.pop("discrepancias_pdf", None)

            hash_archivos = "_".join(sorted(nombres_actuales))
            key_sortable_dinamico = f"sortable_queue_{hash_archivos}"

            if len(nombres_actuales) > 1 and not modo_conjunto:
                st.markdown("##### Organiza el orden de la cola de procesamiento:")
                
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
            idx_actual = 0 if modo_conjunto else st.session_state.get("cola_pdf_indice", 0)

            if idx_actual >= total_pdfs:
                st.session_state["cola_pdf_indice"] = 0
                idx_actual = 0

            pdf_activo = pdfs_ordenados[idx_actual]

            if modo_conjunto:
                st.info(f"**Análisis conjunto:** {total_pdfs} documento(s) para el examen **{examen_activo or 'por seleccionar'}**. Se generará una única ficha. Los valores contradictorios requieren revisión.")
            else:
                st.markdown(f"""
                <div class="banner-cola-trabajo">
                    <b>COLA DE PROCESAMIENTO ACTIVA:</b> Documento <b>{idx_actual + 1}</b> de <b>{total_pdfs}</b><br>
                    <span style="font-size: 0.90rem;">Archivo en revisión actual: <b>{pdf_activo.name}</b></span>
                </div>
                """, unsafe_allow_html=True)
    
                st.progress((idx_actual + 1) / total_pdfs)
    
            st.divider()

            # --- PANEL DE EMPAREJAMIENTO CON LA CARTERA DE EXÁMENES ---
            st.markdown("##### Emparejamiento institucional (Cartera de Exámenes en SQLite)")
            lista_cartera = obtener_lista_cartera_db()

            if not lista_cartera:
                st.warning("No hay datos cargados en la Cartera de Exámenes. Puedes subir el archivo Excel desde el panel lateral.")
            else:
                emparejado_actual = st.session_state.get("cartera_examen_emparejado")
                val_cartera_init = emparejado_actual if (emparejado_actual and emparejado_actual != "Sin vincular" and emparejado_actual in lista_cartera) else None
                
                key_cartera = "sel_cartera_main"
                # Sincronizar antes de instanciar el selectbox; nunca escribir en su key después.
                if "_sel_cartera_pendiente" in st.session_state:
                    st.session_state[key_cartera] = st.session_state.pop("_sel_cartera_pendiente")
                if key_cartera not in st.session_state and val_cartera_init:
                    st.session_state[key_cartera] = val_cartera_init

                col_emp1, col_emp2 = st.columns([2, 1])

                with col_emp1:
                    seleccion_manual_cartera = st.selectbox(
                        "Examen vinculado en la Cartera:",
                        lista_cartera,
                        index=None,
                        placeholder="Selecciona un examen de la Cartera",
                        key=key_cartera,
                        help="Puedes corregir o cambiar manualmente el examen de la Cartera con el que deseas cruzar la información."
                    )
                    st.session_state["cartera_examen_emparejado"] = seleccion_manual_cartera if seleccion_manual_cartera else "Sin vincular"

                with col_emp2:
                    st.write("")
                    st.write("")
                    if st.button("Buscar coincidencia", key="btn_re_fuzzy", use_container_width=False):
                        if examen_activo:
                            candidatos = buscar_top_coincidencias_cartera(examen_activo, top_n=5)
                            if candidatos:
                                st.session_state["candidatos_cartera_pendientes"] = candidatos
                                st.rerun()
                            else:
                                st.warning("No se encontraron coincidencias en la Cartera para este examen.")
                        else:
                            st.warning("Debes seleccionar primero un examen activo usando el botón 'Administrar exámenes'.")

            # Vinculación independiente a «Prestación» (no se confunde con la cartera antigua).
            st.markdown("##### Cartera de Acreditación — coincidencia por Prestación")
            prestaciones_ui = fd.listar_prestaciones_acreditacion(DB_PATH)
            if prestaciones_ui:
                etiqueta_prestaciones = {r[0]: f"{r[1]} · {r[2]}" for r in prestaciones_ui}
                opciones_prestaciones = [r[0] for r in prestaciones_ui]
                # No modificar la key del widget después de haberlo instanciado.
                # Los cambios realizados en diálogos solicitan el reinicio para el próximo rerun.
                if st.session_state.pop("_acreditacion_reset_pendiente", False):
                    st.session_state["acreditacion_manual_selector"] = None
                elif "acreditacion_manual_selector" not in st.session_state:
                    st.session_state["acreditacion_manual_selector"] = None
                elif st.session_state["acreditacion_manual_selector"] not in opciones_prestaciones:
                    st.session_state["acreditacion_manual_selector"] = None
                identificador_manual = st.selectbox(
                    "Prestación vinculada en la Cartera de Acreditación",
                    options=opciones_prestaciones,
                    index=None,
                    placeholder="Selecciona una prestación (opcional)",
                    format_func=lambda id_p: etiqueta_prestaciones.get(id_p,"Desconocida"),
                    key="acreditacion_manual_selector",
                    help="Se utiliza «Prestación» como nombre del examen. Si no coincide exactamente, selecciona la prestación correcta manualmente."
                )
                st.session_state["acreditacion_manual_id"] = identificador_manual
                _, _, seleccion_previa, detalle_match = fd.obtener_valores_fuentes(
                    DB_PATH,examen_activo,area_activa,identificador_manual)
                if seleccion_previa:
                    st.caption(f"Vinculada: **{seleccion_previa['prestacion']}** ({seleccion_previa['area']}) · {detalle_match}")
                elif examen_activo:
                    st.warning("No hay coincidencia exacta. Selecciona la prestación adecuada para evitar datos de otro examen.")
                    sugeridas = fd.candidatos_acreditacion(DB_PATH,examen_activo,area_activa,3)
                    if sugeridas:
                        st.caption("Posibles coincidencias: " + "; ".join(f"{n} ({a})" for _,n,a,_ in sugeridas))
            else:
                st.warning("Aún no hay prestaciones cargadas en la Cartera de Acreditación.")

            st.divider()
            st.markdown("##### Parámetros operacionales de la ficha")

            c_lab, c_act, c_aut = st.columns(3)

            with c_lab:
                laboratorios = cargar_laboratorios()
                actual_lab = st.session_state.get("laboratorio_activo")
                key_lab = "sel_lab_main"
                if key_lab not in st.session_state and actual_lab in laboratorios:
                    st.session_state[key_lab] = actual_lab

                sel_lab = st.selectbox(
                    "Laboratorio de procesamiento", 
                    laboratorios, 
                    index=None,
                    placeholder="Selecciona laboratorio", 
                    key=key_lab
                )
                st.session_state["laboratorio_activo"] = sel_lab

            with c_act:
                encargados_act = cargar_encargados("actualizado")
                actual_act = st.session_state.get("actualizado_activo")
                key_act = "sel_act_main"
                if key_act not in st.session_state and actual_act in encargados_act:
                    st.session_state[key_act] = actual_act

                sel_act = st.selectbox(
                    "Actualizado por", 
                    encargados_act, 
                    index=None,
                    placeholder="Selecciona quién actualiza", 
                    key=key_act
                )
                st.session_state["actualizado_activo"] = sel_act

            with c_aut:
                encargados_aut = cargar_encargados("autorizado")
                actual_aut = st.session_state.get("autorizado_activo")
                key_aut = "sel_aut_main"
                if key_aut not in st.session_state and actual_aut in encargados_aut:
                    st.session_state[key_aut] = actual_aut

                sel_aut = st.selectbox(
                    "Autorizado por", 
                    encargados_aut, 
                    index=None,
                    placeholder="Selecciona quién autoriza", 
                    key=key_aut
                )
                st.session_state["autorizado_activo"] = sel_aut

            with st.expander("Administrar opciones (Añadir / Eliminar de las listas arriba)"):
                tab_m_lab, tab_m_act, tab_m_aut = st.tabs(["Laboratorios", "Actualizado por", "Autorizado por"])

                with tab_m_lab:
                    col_a, col_b = st.columns(2)
                    with col_a:
                        with st.form("form_add_lab_inline", clear_on_submit=True):
                            new_lab = st.text_input("Añadir nuevo laboratorio")
                            if st.form_submit_button("Guardar laboratorio"):
                                if guardar_laboratorio(new_lab):
                                    st.session_state["laboratorio_activo"] = new_lab.strip()
                                    st.session_state["sel_lab_main"] = new_lab.strip()
                                    st.success(f"Guardado: {new_lab}")
                                    st.rerun()
                                else:
                                    st.warning("Nombre no válido o ya existente.")
                    with col_b:
                        labs_del = cargar_laboratorios()
                        if labs_del:
                            with st.form("form_del_lab_inline"):
                                target_del = st.selectbox(
                                    "Eliminar laboratorio", 
                                    labs_del, 
                                    index=None, 
                                    placeholder="Selecciona laboratorio a eliminar",
                                    key="del_lab_select"
                                )
                                if st.form_submit_button("Eliminar seleccionado"):
                                    if target_del:
                                        updated = [x for x in labs_del if x.casefold() != target_del.casefold()]
                                        if not guardar_lista_laboratorios(updated):
                                            st.stop()
                                        if st.session_state.get("laboratorio_activo") == target_del:
                                            st.session_state["laboratorio_activo"] = None
                                            st.session_state["sel_lab_main"] = None
                                        st.session_state.pop("del_lab_select", None)
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
                                    if not guardar_encargados("actualizado", curr):
                                        st.stop()
                                    st.session_state["actualizado_activo"] = new_act.strip()
                                    st.session_state["sel_act_main"] = new_act.strip()
                                    st.success(f"Guardado: {new_act}")
                                    st.rerun()
                                else:
                                    st.warning("Nombre no válido o ya existente.")
                    with col_b:
                        acts_del = cargar_encargados("actualizado")
                        if acts_del:
                            with st.form("form_del_act_inline"):
                                target_del = st.selectbox(
                                    "Eliminar de la lista", 
                                    acts_del, 
                                    index=None, 
                                    placeholder="Selecciona encargado a eliminar",
                                    key="del_act_select"
                                )
                                if st.form_submit_button("Eliminar seleccionado"):
                                    if target_del:
                                        updated = [x for x in acts_del if x.casefold() != target_del.casefold()]
                                        if not guardar_encargados("actualizado", updated):
                                            st.stop()
                                        for estado, widget in (("actualizado_activo", "sel_act_main"),
                                                               ("autorizado_activo", "sel_aut_main")):
                                            if st.session_state.get(estado) == target_del:
                                                st.session_state[estado] = None
                                                st.session_state[widget] = None
                                        if st.session_state.get("actualizado_activo") == target_del:
                                            st.session_state["actualizado_activo"] = None
                                            st.session_state["sel_act_main"] = None
                                        st.session_state.pop("del_act_select", None)
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
                                    if not guardar_encargados("autorizado", curr):
                                        st.stop()
                                    st.session_state["autorizado_activo"] = new_aut.strip()
                                    st.session_state["sel_aut_main"] = new_aut.strip()
                                    st.success(f"Guardado: {new_aut}")
                                    st.rerun()
                                else:
                                    st.warning("Nombre no válido o ya existente.")
                    with col_b:
                        auts_del = cargar_encargados("autorizado")
                        if auts_del:
                            with st.form("form_del_aut_inline"):
                                target_del = st.selectbox(
                                    "Eliminar de la lista", 
                                    auts_del, 
                                    index=None, 
                                    placeholder="Selecciona encargado a eliminar",
                                    key="del_aut_select"
                                )
                                if st.form_submit_button("Eliminar seleccionado"):
                                    if target_del:
                                        updated = [x for x in auts_del if x.casefold() != target_del.casefold()]
                                        if not guardar_encargados("autorizado", updated):
                                            st.stop()
                                        for estado, widget in (("actualizado_activo", "sel_act_main"),
                                                               ("autorizado_activo", "sel_aut_main")):
                                            if st.session_state.get(estado) == target_del:
                                                st.session_state[estado] = None
                                                st.session_state[widget] = None
                                        if st.session_state.get("autorizado_activo") == target_del:
                                            st.session_state["autorizado_activo"] = None
                                            st.session_state["sel_aut_main"] = None
                                        st.session_state.pop("del_aut_select", None)
                                        st.success("Eliminado correctamente.")
                                        st.rerun()

            st.divider()

            # VISTA PREVIA INTERACTIVA CON LUPA
            with st.expander("Vista previa interactiva del documento PDF en revisión", expanded=True):
                try:
                    # La selección de vista previa es independiente del PDF activo
                    # y nunca modifica los documentos que se analizan en conjunto.
                    pdf_vista = pdf_activo
                    if modo_conjunto and total_pdfs > 1:
                        opciones_vista = list(range(total_pdfs))
                        indice_vista = st.selectbox(
                            "Documento PDF a visualizar",
                            options=opciones_vista,
                            index=None,
                            placeholder="Selecciona un PDF para la vista previa",
                            format_func=lambda i: f"{i + 1}. {pdfs_ordenados[i].name}",
                            key="pdf_preview_documento_conjunto",
                            help="Selecciona el archivo que deseas revisar. El análisis seguirá incluyendo todos los PDF cargados."
                        )
                        if indice_vista is not None:
                            pdf_vista = pdfs_ordenados[indice_vista]

                    doc_preview = fitz.open(stream=pdf_vista.getvalue(), filetype="pdf")
                    num_paginas = len(doc_preview)
                    st.caption(f"Visualizando: **{pdf_vista.name}** | Páginas: **{num_paginas}**")

                    zoom_level = 120

                    col_prev1, col_prev2 = st.columns([2, 1])

                    # Cada documento conserva su propia página seleccionada.
                    # El índice está ligado a la posición en la lista de archivos.
                    indice_pagina = (indice_vista if modo_conjunto and total_pdfs > 1 and indice_vista is not None
                                     else idx_actual)
                    with col_prev1:
                        pag_seleccionada = st.slider("Página a visualizar", 1, num_paginas, 1, key=f"pdf_page_slider_{'conjunto' if modo_conjunto else 'cola'}_{indice_pagina}") if num_paginas > 1 else 1

                    with col_prev2:
                        magnify_scale = st.slider("Potencia Lupa (x)", min_value=1.0, max_value=3.5, value=1.0, step=0.5, key=f"pdf_magnify_slider_{'conjunto' if modo_conjunto else 'cola'}_{indice_pagina}")

                    page = doc_preview[pag_seleccionada - 1]
                    pix = page.get_pixmap(dpi=150)
                    img_b64 = base64.b64encode(pix.tobytes("png")).decode("utf-8")

                    ancho_visual = int(750 * (zoom_level / 100))
                    alto_real_imagen = int(ancho_visual * (pix.height / pix.width))
                    alto_componente = alto_real_imagen + 45

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
                        <p style="font-size: 0.8rem; color: #6c757d; margin-top: 6px;"><b>Consejo:</b> Pasa el cursor sobre la página para activar la lupa interactiva.</p>
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

                    st.components.v1.html(html_lupa, height=alto_componente, scrolling=False)

                except Exception as err_preview:
                    st.error(f"No se pudo generar la vista previa del PDF: {err_preview}")

            # BOTÓN DE ANÁLISIS
            st.write("")
            col_btn_analisis, col_sp1 = st.columns([1.5, 2])
            with col_btn_analisis:
                iniciar_analisis = st.button("Analizar todos los PDF del examen" if modo_conjunto else "Analizar documento activo", type="primary", key="analyze_pdf", disabled=exceso_archivos, use_container_width=False)

            if iniciar_analisis:
                if not examen_activo:
                    st.warning("Selecciona un examen antes de continuar")
                elif not st.session_state.get("laboratorio_activo"):
                    st.warning("Selecciona un laboratorio de procesamiento antes de analizar.")
                else:
                    try:
                        # Conservar ediciones confirmadas del mismo examen al volver a analizar.
                        ediciones_previas = {}
                        if st.session_state.get("examen_resultado") == examen_activo:
                            for campo in st.session_state.get("user_edited_fields", set()):
                                valor = st.session_state.get("data_extracted", {}).get(campo, "")
                                if str(valor or "").strip():
                                    ediciones_previas[campo] = valor
                        st.session_state.pop("data_extracted", None)
                        st.session_state.pop("user_edited_fields", None)
                        st.session_state.pop("extracted_text", None)
                        st.session_state.pop("conflictos_pendientes", None)

                        st.session_state["aggrid_key_counter"] += 1

                        aliases_area = get_field_aliases(area_activa)
                        plantilla_bytes_db, _ = obtener_plantilla_area_db(area_activa)

                        if "result_bytes" in st.session_state and idx_actual > 0 and not modo_conjunto:
                            template_bytes_init = st.session_state["result_bytes"]
                        elif template is not None:
                            template_bytes_init = template.getvalue()
                        elif plantilla_bytes_db:
                            template_bytes_init = plantilla_bytes_db
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

                        # La Cartera tradicional sigue suministrando información no documental.
                        # La preparación se obtiene solo de Normas/Cartera de Acreditación.
                        campos_comparables = ["Equipamiento utilizado", "Plazo entrega resultados"]

                        adicionales = [(pdf.name, pdf.getvalue()) for pdf in pdfs_ordenados[1:]] if modo_conjunto else []
                        if modo_conjunto:
                            parsed_pdf_data, fuentes, diferencias, raw_text, _ = consolidar_documentos_pdf(
                                [(pdf.name, pdf.getvalue()) for pdf in pdfs_ordenados], aliases_area, examen_activo
                            )
                            st.session_state["fuentes_pdf"] = fuentes
                            st.session_state["discrepancias_pdf"] = diferencias
                        else:
                            raw_text, _ = extract_pdf_text(pdf_activo.getvalue(), pdf_activo.name, examen_activo)
                            parsed_pdf_data = infer_name_code(raw_text, parse_labeled_text(raw_text, aliases_area, examen_activo))
                            parsed_pdf_data = {campo: val for campo, val in parsed_pdf_data.items()
                                               if campo not in fd.CAMPOS_CONTROLADOS}
                            st.session_state.pop("fuentes_pdf", None)
                            st.session_state.pop("discrepancias_pdf", None)

                        if info_cartera:
                            for campo in campos_comparables:
                                val_pdf = str(parsed_pdf_data.get(campo, "")).strip()
                                val_cartera = str(info_cartera.get(campo, "")).strip()

                                if val_pdf and val_cartera and val_pdf.casefold() != val_cartera.casefold():
                                    conflictos[campo] = {"pdf": val_pdf, "cartera": val_cartera}
                                elif not val_pdf and val_cartera:
                                    datos_sobrescribir[campo] = val_cartera

                        # Ocho campos desde SQLite: seis con Cartera > Normas y
                        # Recepción/Procesamiento exclusivamente desde el Manual.
                        # Ni los PDF analizados ni la Cartera antigua los sobrescriben.
                        valores_doc, origen_doc, prestacion_doc, estado_vinculo = fd.obtener_valores_fuentes(
                            DB_PATH, examen_activo, area_activa,
                            st.session_state.get("acreditacion_manual_id")
                        )
                        datos_sobrescribir.update(valores_doc)
                        st.session_state["origen_campos_documentales"] = origen_doc
                        st.session_state["prestacion_acreditacion_vinculada"] = prestacion_doc
                        st.session_state["estado_vinculacion_acreditacion"] = estado_vinculo

                        out_bytes, data, target, score, reason, method, text = process(
                            pdf_activo.getvalue(), pdf_activo.name, template_bytes_init,
                            st.session_state.get("laboratorio_activo", ""),
                            examen_activo, aliases_area,
                            st.session_state.get("actualizado_activo", ""),
                            st.session_state.get("autorizado_activo", ""),
                            datos_sobrescribir=datos_sobrescribir,
                            documentos_adicionales=adicionales
                        )

                        # Las ediciones manuales confirmadas conservan la prioridad máxima.
                        if ediciones_previas:
                            data.update(ediciones_previas)
                            for campo, valor in ediciones_previas.items():
                                out_bytes = actualizar_campo_excel(out_bytes, examen_activo, campo, valor)
                        st.session_state["result_bytes"] = out_bytes
                        st.session_state["data_extracted"] = copy.deepcopy(data)
                        st.session_state["user_edited_fields"] = set(ediciones_previas)
                        st.session_state["examen_resultado"] = examen_activo
                        st.session_state["extracted_text"] = text
                        st.session_state["area_resultado"] = area_activa

                        st.session_state.pop("conflictos_pendientes", None)
                        st.session_state["modal_conflictos_abierto"] = False
                        st.session_state["modal_discrepancias_unificado"] = False
                        if conflictos:
                            st.session_state["conflictos_pendientes"] = conflictos
                            # PDF vs Cartera se abre solo por su propio botón,
                            # nunca simultáneamente con PDF vs PDF.

                        st.success(f"Se analizaron {total_pdfs} PDF para el examen seleccionado." if modo_conjunto else f"¡Se procesó correctamente el archivo «{pdf_activo.name}»!")
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

            st.subheader("Revisión de la ficha consolidada (todos los PDF)" if modo_conjunto else f"Revisión e información extraída de «{pdf_actual_nombre}»")
            diferencias_pdf = st.session_state.get("discrepancias_pdf", {}) if modo_conjunto else {}
            diferencias_cartera = st.session_state.get("conflictos_pendientes", {})
            if diferencias_pdf or diferencias_cartera:
                if diferencias_pdf:
                    st.warning(_mensaje_discrepancias(len(diferencias_pdf), "los documentos PDF"))
                if diferencias_cartera:
                    st.warning(_mensaje_discrepancias(len(diferencias_cartera), "el PDF y la Cartera de Exámenes"))
                if st.button("Comparar y resolver discrepancias", type="primary", key="abrir_discrepancias_unificadas"):
                    st.session_state["modal_discrepancias_unificado"] = True
                    st.rerun()
            elif modo_conjunto:
                st.success("No hay discrepancias pendientes de revisión.")

            with st.expander("Procedencia de los ocho campos documentales", expanded=False):
                st.caption("No cambia los tres estados ni sus colores. Los PDF no alimentan estos campos.")
                origen_registrado = st.session_state.get("origen_campos_documentales", {})
                for campo in fd.CAMPOS_CONTROLADOS:
                    st.markdown(f"**{campo}:** {origen_registrado.get(campo, 'Sin información válida en SQLite')}")

            st.metric("Campos encontrados", sum(bool(v) for v in data.values()))

            rows = []
            for field in aliases_area:
                value = data.get(field, "")
                if pd.isna(value) or str(value).strip().casefold() in ["nan", "none", "null"]:
                    value = ""
                if field in user_edited:
                    state = "Modificado por usuario"
                elif field in ["Nombre LIS", "Laboratorio de procesamiento", "Actualizado por", "Autorizado por"]:
                    state = "Ingresado por usuario" if value else "Pendiente de ingreso"
                elif field == "Fecha actualización":
                    state = "Asignado por sistema" if value else "Pendiente de asignación"
                else:
                    state = "Encontrado" if value else "No encontrado"
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
                    <span style="background-color: #BAE6FD; color: #0369A1; padding: 2px 8px; border-radius: 4px; font-weight: bold;">Azul:</span> Campo modificado por usuario.
                    &nbsp;&nbsp;
                    <span style="background-color: #FFF0B3; color: #6B4D00; padding: 2px 8px; border-radius: 4px; font-weight: bold;">Amarillo:</span> Campo pendiente de revisión / no encontrado.
                    <br><br>
                    <b>Haz doble clic sobre cualquier celda de la columna "Información extraída" para editar su contenido directamente en la tabla.</b>
                </div>
                """,
                unsafe_allow_html=True
            )

            get_row_style = JsCode("""
            function(params) {
                var val = params.data['Información extraída'];
                var estado = params.data['Estado'];
                var valStr = val != null ? val.toString().trim() : '';
                if (valStr.toLowerCase() === 'nan' || valStr.toLowerCase() === 'none' || valStr.toLowerCase() === 'null') {
                    valStr = '';
                }
                
                if (valStr !== '' && estado === 'Modificado por usuario') {
                    return {
                        'backgroundColor': '#BAE6FD',
                        'color': '#0369A1',
                        'fontWeight': '600'
                    };
                } else if (valStr !== '') {
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
                resizable=False,
                suppressMovable=True
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
            grid_options['suppressMovableColumns'] = True
            
            grid_options['onCellValueChanged'] = JsCode("""
            function(params) {
                if (params.column && params.column.colDef && params.column.colDef.field === 'Información extraída') {
                    var newVal = params.newValue != null ? params.newValue.toString().trim() : '';
                    if (newVal.toLowerCase() === 'nan' || newVal.toLowerCase() === 'none' || newVal.toLowerCase() === 'null') {
                        newVal = '';
                    }
                    
                    if (newVal === '') {
                        var campo = params.data['Campo'];
                        if (['Nombre LIS', 'Laboratorio de procesamiento', 'Actualizado por', 'Autorizado por'].includes(campo)) {
                            params.node.setDataValue('Estado', 'Pendiente de ingreso');
                        } else if (campo === 'Fecha actualización') {
                            params.node.setDataValue('Estado', 'Pendiente de asignación');
                        } else {
                            params.node.setDataValue('Estado', 'No encontrado');
                        }
                    } else {
                        params.node.setDataValue('Estado', 'Modificado por usuario');
                    }
                    params.api.refreshCells({rowNodes: [params.node], force: true});
                }
            }
            """)

            filas_cant = len(df_base)
            altura_tabla = filas_cant * 42 + 55

            grid_response = AgGrid(
                df_base,
                gridOptions=grid_options,
                allow_unsafe_jscode=True,
                update_mode=GridUpdateMode.MODEL_CHANGED,
                fit_columns_on_grid_load=True,
                theme="alpine",
                height=altura_tabla,
                key=f"aggrid_data_table_{st.session_state['aggrid_key_counter']}"
            )

            edited_df = pd.DataFrame(grid_response['data'])

            if "data_extracted" in st.session_state:
                for idx, row in edited_df.iterrows():
                    field = row["Campo"]
                    val_raw = row["Información extraída"]
                    if pd.isna(val_raw) or val_raw is None:
                        val_editado = ""
                    else:
                        val_editado = str(val_raw).strip()
                        if val_editado.casefold() in ["nan", "none", "null"]:
                            val_editado = ""
                    st.session_state["data_extracted"][field] = val_editado
                    
                    if row["Estado"] == "Modificado por usuario" and val_editado:
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
                ZONA_HORARIA = ZoneInfo("America/Santiago")
                fecha_str = datetime.now(ZONA_HORARIA).strftime("%d-%m-%Y")
                area_clean = (area_activa or "General").replace(" ", "_")
                nombre_archivo_excel = f"Ficha_Técnica_Actualizada_{area_clean}_{fecha_str}.xlsx"

                st.download_button(
                    "Descargar Excel consolidado (.xlsx)", type="secondary",
                    data=st.session_state["result_bytes"],
                    file_name=nombre_archivo_excel,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="download_updated_xlsx",
                    help="Descarga el libro Excel con todas las hojas de los exámenes procesados hasta ahora en este lote."
                )

            with col_update:
                if st.button("Actualizar en herramienta", type="primary", use_container_width=True, key="btn_trigger_save_template",
                             disabled=bool(st.session_state.get("discrepancias_pdf") or st.session_state.get("conflictos_pendientes")),
                             help="Resuelve primero las discrepancias pendientes."):
                    bytes_iniciales = st.session_state.get("bytes_plantilla_inicial")
                    st.session_state["comparacion_planilla_detalles"] = comparar_plantillas_excel(
                        bytes_iniciales,
                        st.session_state["result_bytes"],
                        examen_activo or ""
                    )
                    st.session_state["modal_confirmar_planilla_abierto"] = True
                    st.rerun()

            with col_next:
                if modo_conjunto:
                    st.info("Análisis conjunto finalizado. Descarga o actualiza la ficha única.")
                idx_actual = st.session_state.get("cola_pdf_indice", 0)
                total_pdfs = len(pdfs_ordenados_rev) if 'pdfs_ordenados_rev' in locals() and pdfs_ordenados_rev else 0
                
                if not modo_conjunto and idx_actual < total_pdfs - 1:
                    if st.button(f"Finalizar y pasar al examen ({idx_actual + 2}/{total_pdfs})", type="secondary", use_container_width=True):
                        st.session_state["cola_pdf_indice"] = idx_actual + 1

                        st.session_state["examen_activo"] = None
                        st.session_state["acreditacion_manual_id"] = None
                        st.session_state["_acreditacion_reset_pendiente"] = True
                        st.session_state["cartera_examen_emparejado"] = "Sin vincular"
                        st.session_state["_sel_cartera_pendiente"] = None
                        st.session_state.pop("candidatos_cartera_pendientes", None)

                        st.session_state.pop("data_extracted", None)
                        st.session_state.pop("user_edited_fields", None)
                        st.session_state.pop("extracted_text", None)

                        st.session_state["mostrar_modal_examenes"] = False
                        st.session_state.pop("modal_area_activa", None)

                        st.rerun()
                elif not modo_conjunto:
                    st.success("¡Has completado la cola de exámenes! Puedes descargar el Excel consolidado con el botón de la izquierda.")

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
                        index=None,
                        placeholder="Selecciona formato",
                        key="imported_excel_format",
                    )
                with import_col2:
                    selected_import_sheet = None
                    if imported_format == "CSV":
                        selected_import_sheet = st.selectbox(
                            "Hoja que se exportará a CSV",
                            imported_sheet_names,
                            index=None,
                            placeholder="Selecciona una hoja",
                            key="imported_excel_sheet",
                        )
                    else:
                        st.caption("Se exportarán todas las hojas cuando el formato lo permita.")
                
                if imported_format:
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