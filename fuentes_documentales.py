"""Fuentes documentales controladas de Testeo.

- Normas Generales de Toma de Muestras: cuatro valores por defecto.
- Cartera de Acreditación: seis valores por prestación (clave «Prestación»).
- Manual de toma de muestras: dos fragmentos institucionales delimitados.

Los PDF analizados por el usuario no completan estos ocho campos.
Los reemplazos de archivos y sus índices se realizan en una sola transacción.
"""
from __future__ import annotations

import io
import re
import sqlite3
import hashlib
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from difflib import SequenceMatcher
from zoneinfo import ZoneInfo

import fitz
from openpyxl import load_workbook

NORMAS_ID = "normas_generales"
ACREDITACION_ID = "cartera_acreditacion"
MANUAL_ID = "manual_toma_muestras"
CAMPOS_MANUAL = ("Recepción de muestras", "Procesamiento de muestras")
CAMPOS_NORMAS = (
    "Preparación del paciente",
    "Procedimiento de toma de muestras",
    "Rotulación",
    "Criterios de rechazo",
)
MAPA_ACREDITACION = {
    "Preparación del paciente": "Preparación del paciente · texto",
    "Muestra requerida y recipiente": "Muestra y recipiente · texto",
    "Procedimiento de toma de muestras": "Procedimiento de toma · texto",
    "Rotulación": "Rotulación · texto",
    "Condiciones de envío al Laboratorio": "Traslado · texto",
    "Criterios de rechazo": "Criterios de rechazo · texto",
}
CAMPOS_CONTROLADOS = tuple(MAPA_ACREDITACION) + CAMPOS_MANUAL
ENCABEZADOS_NORMAS = {
    "Preparación del paciente": "Preparación del paciente",
    "Procedimiento de toma de muestras": "Procedimiento de toma de muestras",
    "Rotulación": "Rotulación",
    "Criterios de rechazo": "Criterios de rechazo",
}


def normalizar(texto):
    base = unicodedata.normalize("NFKD", str(texto or "").casefold())
    base = "".join(c for c in base if not unicodedata.combining(c))
    base = re.sub(r"[^\w]+", " ", base, flags=re.UNICODE)
    return " ".join(base.split())


def area_equivalente(nombre):
    a = normalizar(nombre)
    if "bioquim" in a:
        return "bioquimica"
    if "hematolog" in a or "coagul" in a:
        return "hematologia"
    if "microbiolog" in a:
        return "microbiologia"
    if "hormon" in a or "inmun" in a:
        return "hormonas"
    if "derivad" in a or "convenio" in a:
        return "derivados"
    return a


def texto_valido(valor, estado=None):
    """'No aplica' es un dato documentado válido; 'Pendiente' no lo es."""
    t = str(valor or "").strip()
    if not t or normalizar(t) in ("nan", "none", "null", "pendiente", "sin informacion", "sin documentar"):
        return False
    if estado is not None:
        e = normalizar(estado)
        if e in ("pendiente", "sin documentar", "no encontrado", "incompleto", ""):
            return False
        # El archivo usa Específico/Heredado como categorías documentadas.
        if e not in ("especifico", "heredado"):
            return False
    return not bool(re.match(r"^(pendiente\b|por definir\b|falta incorporar\b|sin documentar\b)", normalizar(t)))


def _fecha():
    return datetime.now(ZoneInfo("America/Santiago")).isoformat(timespec="seconds")


def inicializar_fuentes(db_path):
    with sqlite3.connect(db_path) as con:
        con.execute("""CREATE TABLE IF NOT EXISTS documentos_fuente (
            fuente_id TEXT PRIMARY KEY,
            nombre_archivo TEXT NOT NULL,
            tipo_mime TEXT NOT NULL,
            archivo BLOB NOT NULL,
            sha256 TEXT NOT NULL,
            fecha_actualizacion TEXT NOT NULL,
            registros INTEGER NOT NULL DEFAULT 0
        )""")
        con.execute("""CREATE TABLE IF NOT EXISTS normas_generales_campos (
            campo TEXT PRIMARY KEY,
            contenido TEXT NOT NULL
        )""")
        con.execute("""CREATE TABLE IF NOT EXISTS prestaciones_acreditacion (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fila_excel INTEGER NOT NULL UNIQUE,
            identificador TEXT,
            codigo TEXT,
            prestacion TEXT NOT NULL,
            prestacion_norm TEXT NOT NULL,
            area TEXT,
            area_norm TEXT,
            datos_json TEXT NOT NULL,
            estados_json TEXT NOT NULL
        )""")
        con.execute("CREATE INDEX IF NOT EXISTS idx_prestacion_norm ON prestaciones_acreditacion(prestacion_norm)")
        con.execute("""CREATE TABLE IF NOT EXISTS manual_toma_campos (
            campo TEXT PRIMARY KEY,
            contenido TEXT NOT NULL
        )""")


def _pdf_como_texto(documento):
    with fitz.open(stream=documento, filetype="pdf") as pdf:
        return "\n".join(p.get_text(sort=True) for p in pdf)



def _texto_manual_sin_margenes(documento):
    """Recupera solo bloques del cuerpo (sin cabeceras, pies o folios).

    Se utiliza la posición física del texto y después frases delimitadoras,
    evitando copiar páginas completas o números incrustados en el documento.
    """
    if not documento:
        raise ValueError("El PDF del manual está vacío.")
    paginas = []
    try:
        with fitz.open(stream=documento, filetype="pdf") as pdf:
            if not pdf.page_count:
                raise ValueError("El PDF del manual no contiene páginas.")
            for pagina in pdf:
                alto = pagina.rect.height
                # En el manual original el texto útil ocupa aproximadamente
                # y=131..728 pt; los membretes terminan cerca de y=115 y el
                # folio empieza en y=735. Las proporciones toleran reescalados.
                minimo, maximo = alto * (120 / 792), alto * (730 / 792)
                partes = []
                for bloque in pagina.get_text("blocks", sort=True):
                    x0, y0, x1, y1, contenido = bloque[:5]
                    if not isinstance(contenido, str):
                        continue
                    if y0 < minimo or y1 > maximo:
                        continue
                    contenido = contenido.strip()
                    if contenido:
                        partes.append(contenido)
                if partes:
                    paginas.append(" ".join(partes))
    except (ValueError, RuntimeError) as ex:
        raise ValueError(f"No se pudo leer el PDF del manual: {ex}") from ex
    texto = re.sub(r"\s+", " ", " ".join(paginas)).strip()
    if len(texto) < 200:
        raise ValueError("El PDF no tiene texto digital suficiente; no se sobrescribirá el manual almacenado.")
    return texto


def extraer_manual_toma(documento, nombre_archivo):
    """Obtiene exactamente los dos pasajes solicitados, sin títulos o márgenes.

    Si cambia la estructura del manual y los límites dejan de existir, rechaza
    la actualización antes de escribir en SQLite (sin extracciones parciales).
    """
    if not str(nombre_archivo).lower().endswith(".pdf"):
        raise ValueError("El Manual de toma de muestras debe ser un archivo PDF.")
    texto = _texto_manual_sin_margenes(documento)

    def ubicar(patron, desde=0, identificador="fragmento"):
        match = re.search(patron, texto[desde:], flags=re.IGNORECASE)
        if not match:
            raise ValueError(f"No se reconoció {identificador} en el manual. Se conservó la versión anterior.")
        return desde + match.start(), desde + match.end()

    _, fin_titulo_recepcion = ubicar(
        r"\bX\s*[.]\s*RECEPCI[OÓ]N\s+Y\s+ROTULACI[OÓ]N\s+DE\s+MUESTRAS\b",
        identificador="el título X. RECEPCIÓN Y ROTULACIÓN DE MUESTRAS")
    inicio_recepcion, _ = ubicar(
        r"Las\s+muestras\s+deben\s+ser\s+recepcionadas\s+por\s+un\s+t[eé]cnico\s+param[eé]dico\s+o\s+profesional\s+de\s+la\s+salud[.]",
        desde=fin_titulo_recepcion, identificador="el inicio de Recepción de muestras")
    _, fin_recepcion = ubicar(
        r"cumplen\s+con\s+lo\s+estipulado\s+por\s+el\s+Laboratorio\s+del\s+CRSCO[.]",
        desde=inicio_recepcion, identificador="el final de Recepción de muestras")
    _, fin_titulo_procesamiento = ubicar(
        r"\b2\s*[.]\s*MUESTRAS\s+PROCEDENTES\s+DE\s+LA\s+TOMA\s+DE\s+MUESTRAS\s+DEL\s+LABORATORIO\s+DEL\s+CRSCO\b",
        desde=fin_recepcion, identificador="el título 2. MUESTRAS PROCEDENTES...")
    inicio_procesamiento, _ = ubicar(
        r"Las\s+muestras\s+son\s+rotuladas\s+en\s+la\s+Toma\s+de\s+Muestras\s+con\s+una\s+etiqueta",
        desde=fin_titulo_procesamiento, identificador="el inicio de Procesamiento de muestras")
    _, fin_procesamiento = ubicar(
        r"se\s+ingresan\s+en\s+forma\s+digital\s+en\s+el\s+proceso\s+de\s+agendamiento\s+del\s+paciente\s+en\s+la\s+secretar[ií]a\s+de\s+laboratorio[.]",
        desde=inicio_procesamiento, identificador="el final de Procesamiento de muestras")

    recepcion = texto[inicio_recepcion:fin_recepcion].strip()
    procesamiento = texto[inicio_procesamiento:fin_procesamiento].strip()
    # Validación de integridad y seguridad: jamás aceptar cabeceras o folios.
    for nombre, contenido in (("Recepción de muestras", recepcion),
                              ("Procesamiento de muestras", procesamiento)):
        if len(contenido) < 300:
            raise ValueError(f"El apartado {nombre} parece incompleto; se conservará la versión anterior.")
        if re.search(r"(?:Cod[.]\s*Caract[.]|C[oó]digo\s*:|P[aá]gina\s*:|Vigencia\s*:)", contenido, re.I):
            raise ValueError(f"Se detectó texto de encabezado o pie en {nombre}; se conservará la versión anterior.")
    return {"Recepción de muestras": recepcion,
            "Procesamiento de muestras": procesamiento}


def guardar_manual_toma(db_path, archivo, nombre_archivo):
    """Actualiza en una sola transacción el original y sus dos valores limpios."""
    campos = extraer_manual_toma(archivo, nombre_archivo)
    with sqlite3.connect(db_path) as con:
        con.execute("DELETE FROM manual_toma_campos")
        con.executemany("INSERT INTO manual_toma_campos(campo,contenido) VALUES (?,?)", list(campos.items()))
        con.execute("""INSERT OR REPLACE INTO documentos_fuente
            (fuente_id,nombre_archivo,tipo_mime,archivo,sha256,fecha_actualizacion,registros)
            VALUES (?,?,?,?,?,?,?)""",
            (MANUAL_ID, nombre_archivo, "application/pdf", sqlite3.Binary(archivo),
             hashlib.sha256(archivo).hexdigest(), _fecha(), len(campos)))
    return len(campos)


def _docx_como_texto(documento):
    # Extrae párrafos y tablas en el orden físico del documento Word.
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with zipfile.ZipFile(io.BytesIO(documento)) as z:
        raw = z.read("word/document.xml")
    root = ET.fromstring(raw)
    body = root.find("w:body", ns)
    if body is None:
        raise ValueError("El DOCX no contiene cuerpo de documento.")
    partes = []
    for nodo in body.iter():
        if nodo.tag == "{%s}p" % ns["w"]:
            txt = "".join((t.text or "") for t in nodo.findall(".//w:t", ns)).strip()
            if txt:
                partes.append(txt)
    return "\n".join(partes)


def extraer_normas(documento, nombre_archivo):
    archivo = str(nombre_archivo).lower()
    if archivo.endswith(".pdf"):
        texto = _pdf_como_texto(documento)
    elif archivo.endswith(".docx"):
        texto = _docx_como_texto(documento)
    else:
        raise ValueError("Las Normas Generales solo admiten PDF o DOCX.")
    texto = texto.replace("\u200b", "").replace("\ufeff", "")
    # Encontrar un encabezado de apartado, no coincidencias dentro de párrafos.
    # En DOCX el número de una lista puede ser estilo Word, no texto del párrafo.
    patron = re.compile(
        r"(?im)^\s*(?:[1-4]\s*[.)]\s*)?"
        r"(Preparaci[oó]n del paciente|Procedimiento de toma de muestras|Rotulaci[oó]n|Criterios de rechazo)\s*$"
    )
    matches = list(patron.finditer(texto))
    secciones = {}
    titulos_normalizados = {normalizar(titulo): campo for campo, titulo in ENCABEZADOS_NORMAS.items()}
    for i, m in enumerate(matches):
        nombre = titulos_normalizados.get(normalizar(m.group(1)))
        if not nombre:
            continue
        final = matches[i + 1].start() if i + 1 < len(matches) else len(texto)
        contenido = texto[m.end():final].strip()
        contenido = re.sub(r"\n{3,}", "\n\n", contenido)
        if contenido:
            secciones[nombre] = contenido
    faltan = [c for c in CAMPOS_NORMAS if c not in secciones]
    if faltan:
        raise ValueError("No se reconocen los cuatro apartados esperados: " + ", ".join(faltan))
    return secciones


def _leer_cartera(documento):
    """Lee cabeceras por su nombre real; NO depende de posiciones fijas."""
    wb = load_workbook(io.BytesIO(documento), read_only=True, data_only=True)
    try:
        if "Cartera_Completa" not in wb.sheetnames:
            raise ValueError("El Excel debe contener la hoja 'Cartera_Completa'.")
        ws = wb["Cartera_Completa"]
        encabezados = None
        fila_header = None
        for n, fila in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
            mapa = {normalizar(c): ix for ix, c in enumerate(fila) if c is not None}
            if "prestacion" in mapa and "area" in mapa:
                encabezados, fila_header = mapa, n
                break
        if not encabezados:
            raise ValueError("No se encontraron los encabezados «Prestación» y «Área».")
        necesarios = []
        for origen in MAPA_ACREDITACION.values():
            k = normalizar(origen)
            necesarios += [k, k.replace(" texto", " estado")]
        ausentes = [x for x in necesarios if x not in encabezados]
        if ausentes:
            raise ValueError("Faltan columnas obligatorias: " + ", ".join(ausentes))

        def celda(fila, columna):
            pos = encabezados.get(normalizar(columna))
            return fila[pos] if pos is not None and pos < len(fila) else None

        registros = []
        for n, row in enumerate(ws.iter_rows(min_row=fila_header + 1, values_only=True), start=fila_header + 1):
            prestacion = str(celda(row, "Prestación") or "").strip()
            if not prestacion or normalizar(prestacion) in ("prestacion", "nan", "none"):
                continue
            area = str(celda(row, "Área") or "").strip()
            datos = {}
            estados = {}
            for campo, titulo in MAPA_ACREDITACION.items():
                datos[campo] = str(celda(row, titulo) or "").strip()
                estados[campo] = str(celda(row, titulo.replace("· texto", "· estado")) or "").strip()
            registros.append((n, str(celda(row,"ID") or "").strip(),
                              str(celda(row,"Código") or "").strip(), prestacion,
                              normalizar(prestacion), area, area_equivalente(area),
                              datos, estados))
        if not registros:
            raise ValueError("La Cartera no contiene prestaciones válidas.")
        return registros
    finally:
        wb.close()


def guardar_normas(db_path, archivo, nombre_archivo):
    campos = extraer_normas(archivo, nombre_archivo)  # validar antes de modificar SQLite
    tipo = ("application/pdf" if nombre_archivo.lower().endswith(".pdf")
            else "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    with sqlite3.connect(db_path) as con:
        con.execute("DELETE FROM normas_generales_campos")
        con.executemany("INSERT INTO normas_generales_campos(campo,contenido) VALUES (?,?)", list(campos.items()))
        con.execute("""INSERT OR REPLACE INTO documentos_fuente
            (fuente_id,nombre_archivo,tipo_mime,archivo,sha256,fecha_actualizacion,registros)
            VALUES (?,?,?,?,?,?,?)""",
            (NORMAS_ID,nombre_archivo,tipo,sqlite3.Binary(archivo),hashlib.sha256(archivo).hexdigest(),_fecha(),4))
    return len(campos)


def guardar_cartera_acreditacion(db_path, archivo, nombre_archivo):
    import json
    registros = _leer_cartera(archivo)  # validar antes de modificar SQLite
    filas = [(numero,ident,codigo,prestacion,prest_norm,area,area_norm,
              json.dumps(datos,ensure_ascii=False),json.dumps(estados,ensure_ascii=False))
             for numero,ident,codigo,prestacion,prest_norm,area,area_norm,datos,estados in registros]
    with sqlite3.connect(db_path) as con:
        con.execute("DELETE FROM prestaciones_acreditacion")
        con.executemany("""INSERT INTO prestaciones_acreditacion
            (fila_excel,identificador,codigo,prestacion,prestacion_norm,area,area_norm,datos_json,estados_json)
            VALUES (?,?,?,?,?,?,?,?,?)""", filas)
        con.execute("""INSERT OR REPLACE INTO documentos_fuente
            (fuente_id,nombre_archivo,tipo_mime,archivo,sha256,fecha_actualizacion,registros)
            VALUES (?,?,?,?,?,?,?)""", (ACREDITACION_ID,nombre_archivo,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            sqlite3.Binary(archivo),hashlib.sha256(archivo).hexdigest(),_fecha(),len(filas)))
    return len(filas)


def documento_fuente(db_path, fuente_id, incluir_archivo=False):
    campos = ("fuente_id,nombre_archivo,tipo_mime,sha256,fecha_actualizacion,registros,archivo"
              if incluir_archivo else "fuente_id,nombre_archivo,tipo_mime,sha256,fecha_actualizacion,registros")
    with sqlite3.connect(db_path) as con:
        con.row_factory = sqlite3.Row
        fila = con.execute(f"SELECT {campos} FROM documentos_fuente WHERE fuente_id=?",(fuente_id,)).fetchone()
        return dict(fila) if fila else None


def listar_prestaciones_acreditacion(db_path):
    with sqlite3.connect(db_path) as con:
        return con.execute("SELECT id, prestacion, area FROM prestaciones_acreditacion ORDER BY prestacion COLLATE NOCASE, area").fetchall()


def candidatos_acreditacion(db_path, nombre_examen, area, limite=8):
    objetivo = normalizar(nombre_examen)
    area_norm = area_equivalente(area)
    with sqlite3.connect(db_path) as con:
        filas = con.execute("SELECT id,prestacion,area,prestacion_norm,area_norm FROM prestaciones_acreditacion").fetchall()
    scored = []
    for identificador,nombre,nombre_area,nombre_norm,area_dest in filas:
        ratio = SequenceMatcher(None,objetivo,nombre_norm).ratio() if objetivo else 0
        if nombre_norm == objetivo:
            ratio = 1.0
        score = ratio + (0.02 if area_norm and area_norm==area_dest else 0)
        scored.append((identificador,nombre,nombre_area,score))
    scored.sort(key=lambda row:(-row[3],row[1]))
    return scored[:limite]


def _buscar_match(con, nombre_examen, area, seleccion_manual=None):
    # Un emparejamiento difuso nunca se aplica automáticamente.
    if seleccion_manual is not None:
        return con.execute("SELECT id,prestacion,area,datos_json,estados_json FROM prestaciones_acreditacion WHERE id=?",(int(seleccion_manual),)).fetchone(), "manual"
    normal = normalizar(nombre_examen)
    if not normal:
        return None, "sin examen"
    filas = con.execute("""SELECT id,prestacion,area,datos_json,estados_json
                          FROM prestaciones_acreditacion WHERE prestacion_norm=?""",(normal,)).fetchall()
    if len(filas)==1:
        if area and area_equivalente(filas[0][2]) != area_equivalente(area):
            return None, "coincidencia en otra área; requiere confirmación"
        return filas[0], "exacta"
    if len(filas)>1:
        area_norm = area_equivalente(area)
        filtradas = [fila for fila in filas if area_equivalente(fila[2]) == area_norm]
        if len(filtradas)==1:
            return filtradas[0], "exacta por área"
        return None, "ambigua"
    return None, "sin coincidencia exacta"


def obtener_valores_fuentes(db_path, nombre_examen, area=None, seleccion_manual=None):
    """Regresa (valores, orígenes, prestación vinculada, mensaje).

    Ningún PDF analizado interviene en los ocho campos controlados.
    Los seis campos de Acreditación priorizan Cartera > Normas.
    Los dos campos adicionales se obtienen exclusivamente del Manual.
    """
    import json
    valores, origen = {}, {}
    with sqlite3.connect(db_path) as con:
        for campo, contenido in con.execute("SELECT campo, contenido FROM normas_generales_campos"):
            if campo in CAMPOS_NORMAS and texto_valido(contenido):
                valores[campo] = contenido
                origen[campo] = "Normas Generales de Toma de Muestras"
        for campo, contenido in con.execute("SELECT campo, contenido FROM manual_toma_campos"):
            if campo in CAMPOS_MANUAL and texto_valido(contenido):
                valores[campo] = contenido
                origen[campo] = "Manual de toma de muestras"
        fila, mensaje = _buscar_match(con,nombre_examen,area,seleccion_manual)
        seleccion = None
        if fila:
            identificador, prestacion, area_prestacion, datos_raw, estados_raw = fila
            seleccion = {"id": identificador, "prestacion": prestacion, "area": area_prestacion}
            datos, estados = json.loads(datos_raw), json.loads(estados_raw)
            for campo in CAMPOS_CONTROLADOS:
                texto = datos.get(campo,"")
                if texto_valido(texto, estados.get(campo)):
                    valores[campo] = texto
                    origen[campo] = "Cartera de Acreditación · " + str(estados.get(campo))
    return valores, origen, seleccion, mensaje
