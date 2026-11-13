"""Fuentes documentales controladas de la Herramienta de extracción de datos para laboratorio.

- Normas Generales de Toma de Muestras: cuatro valores por defecto.
- Cartera de Acreditación: seis valores por prestación (clave «Prestación»).
- Recepción y Procesamiento de muestras: campos presentes en la ficha,
  deliberadamente sin fuente de extracción automática.

Los PDF analizados por el usuario no completan estos ocho campos.
Los reemplazos de archivos y sus índices se realizan en una sola transacción.
"""
from __future__ import annotations

import io
import re
import sqlite3
import hashlib
from contextlib import closing, nullcontext
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
MANUAL_ID = "manual_toma_muestras"  # Identificador legado: solo para retirar datos antiguos.
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

# Orden institucional de consulta para cada campo. Se usa también en el panel
# de procedencia, pero NO altera los valores ni los estados de la extracción.
PRIORIDAD_FUENTES_POR_CAMPO = {
    "Preparación del paciente": (
        "Cartera de Exámenes del CRS",
        "Cartera de acreditación",
        "Normas Generales de Toma de Muestras",
    ),
    "Muestra requerida y recipiente": ("Cartera de acreditación",),
    "Procedimiento de toma de muestras": (
        "Normas Generales de Toma de Muestras", "Cartera de acreditación",
    ),
    "Rotulación": ("Normas Generales de Toma de Muestras", "Cartera de acreditación"),
    "Condiciones de envío al Laboratorio": ("Cartera de acreditación",),
    "Criterios de rechazo": ("Normas Generales de Toma de Muestras", "Cartera de acreditación"),
    "Recepción de muestras": (),
    "Procesamiento de muestras": (),
}


# Solo estos tres campos conservan Normas cuando ambas fuentes son válidas.
CAMPOS_NORMAS_PRIORITARIAS = frozenset((
    "Procedimiento de toma de muestras", "Rotulación", "Criterios de rechazo",
))


def identificar_fuente_procedencia(texto_origen):
    """Convierte el origen real (incluido estado de Acreditación) en su fuente."""
    origen = str(texto_origen or "").strip()
    if origen.startswith("Cartera de Acreditación"):
        return "Cartera de acreditación"
    if origen.startswith("Cartera de Exámenes del CRS"):
        return "Cartera de Exámenes del CRS"
    if origen == "Normas Generales de Toma de Muestras":
        return origen
    return ""
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


# Aviso genérico de Acreditación: no constituye preparación específica.
# Se ignora únicamente en «Preparación del paciente» cuando proviene de
# Acreditación; no se filtran por esta regla las Indicaciones del CRS ni Normas.
AVISO_PREPARACION_GENERAL_ACREDITACION = (
    "Aplicar requisitos generales de preparación del paciente del CRS; "
    "verificar además si la prestación tiene una indicación específica."
)


def preparacion_valida(valor, estado=None, *, fuente_acreditacion=False):
    """Valida preparación, descartando avisos genéricos de Acreditación.

    «No aplica» y sus variantes son válidos SOLO si el texto viene de la
    Cartera de Acreditación y su estado documental está validado. En Normas
    continúan descartándose; Indicaciones CRS usa su propia validación.
    Las comparaciones están normalizadas para respetar el resto de filtros.
    """
    if not texto_valido(valor, estado):
        return False
    limpio = normalizar(valor)
    if limpio in ("no aplica", "no aplicable", "n a", "na") and not fuente_acreditacion:
        return False
    if (limpio.startswith("pendiente incorporar")
            or "pendiente incorporar la instruccion especifica" in limpio
            or "consultar los manuales instructivos del laboratorio externo" in limpio):
        return False
    if fuente_acreditacion and limpio == normalizar(AVISO_PREPARACION_GENERAL_ACREDITACION):
        return False
    return True


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
        # Migración de retiro: eliminar SOLO el Manual heredado.
        # No tocar Normas, carteras, plantillas, cuentas ni auditorías.
        if con.execute("SELECT 1 FROM documentos_fuente WHERE fuente_id=?", (MANUAL_ID,)).fetchone():
            con.execute("DELETE FROM documentos_fuente WHERE fuente_id=?", (MANUAL_ID,))
        if con.execute("""SELECT 1 FROM sqlite_master
                          WHERE type='table' AND name='manual_toma_campos'""").fetchone():
            con.execute("DROP TABLE manual_toma_campos")


def hay_manual_heredado(db_path):
    """Permite respaldar la base SOLO si contiene el Manual que se va a retirar."""
    with closing(sqlite3.connect(db_path)) as con:
        tabla = con.execute("""SELECT 1 FROM sqlite_master
                              WHERE type='table' AND name='manual_toma_campos'""").fetchone()
        docs = con.execute("""SELECT 1 FROM sqlite_master
                             WHERE type='table' AND name='documentos_fuente'""").fetchone()
        archivo = (docs and con.execute(
            "SELECT 1 FROM documentos_fuente WHERE fuente_id=?", (MANUAL_ID,)).fetchone())
        return bool(tabla or archivo)


def _pdf_como_texto(documento):
    with fitz.open(stream=documento, filetype="pdf") as pdf:
        return "\n".join(p.get_text(sort=True) for p in pdf)


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
        contenido = limpiar_formato_normas(nombre, contenido)
        if contenido:
            secciones[nombre] = contenido
    faltan = [c for c in CAMPOS_NORMAS if c not in secciones]
    if faltan:
        raise ValueError("No se reconocen los cuatro apartados esperados: " + ", ".join(faltan))
    return secciones


# Subtítulos documentales comprobados en las Normas generales del CRS.
# Solo se reconoce como título una línea completa, para no poner en negrita
# fragmentos del texto clínico ni inventar una estructura.
SUBTITULOS_NORMAS = {
    "Preparación del paciente": (
        "Indicaciones generales para pacientes ambulatorios",
        "Regla institucional", "Antes de la toma", "Documentación requerida",
    ),
    "Procedimiento de toma de muestras": (
        "Sangre venosa por sistema al vacío", "Preparación", "Material",
        "Obtención", "Después de la toma",
    ),
    "Rotulación": (
        "Rotulación institucional general", "Datos mínimos",
        "Datos operativos del sistema institucional", "Ubicación de la etiqueta",
    ),
    "Criterios de rechazo": (
        "Criterios generales institucionales", "Identificación y solicitud",
        "Muestra y contenedor", "Transporte y conservación",
        "Condiciones fisicoquímicas",
    ),
}


def es_subtitulo_normas(campo, linea):
    return normalizar(linea) in {
        normalizar(t) for t in SUBTITULOS_NORMAS.get(campo, ())
    }


def limpiar_formato_normas(campo, contenido):
    """Reconstruye párrafos, viñetas y subtítulos de Normas PDF o DOCX.

    Un salto visual de línea en el PDF no equivale a un nuevo párrafo: las
    continuaciones se unen hasta la siguiente viñeta o subtítulo reconocido.
    No se reescriben ni se modifican las indicaciones clínicas.
    """
    texto = str(contenido or "").replace("\r\n", "\n").replace("\r", "\n")
    for caracter in ("\u200b", "\ufeff", "\u00ad", "\u2060"):
        texto = texto.replace(caracter, "")
    texto = texto.replace("\u00a0", " ")
    texto = re.sub(r"[\t ]+", " ", texto)

    # El primer PDF institucional incluía una narración que repite todas las
    # instrucciones de las 11 viñetas. Se elimina SOLO cuando se comprueba que
    # están presentes las viñetas equivalentes, no por palabras aisladas.
    if campo == "Procedimiento de toma de muestras":
        repeticion = re.search(
            r"(?s)(?<=Sangre venosa por sistema al vacío\n)(.*?)\n\s*Preparación\s*\n",
            texto,
        )
        if repeticion:
            parrafo = normalizar(repeticion.group(1))
            resto = normalizar(texto[repeticion.end():])
            evidencias = (
                "higiene de manos entre pacientes",
                "verificar que etiquetas y orden correspondan",
                "seleccionar una vena y un sitio indemne",
                "mezclar por inversion suave",
                "eliminar material cortopunzante sin recapsular",
                "comprimir colocar parche y registrar al responsable",
            )
            if all(frase in parrafo and frase in resto for frase in evidencias):
                texto = texto[:repeticion.start()] + "Preparación\n" + texto[repeticion.end():]

    titulos = {normalizar(t) for t in SUBTITULOS_NORMAS.get(campo, ())}
    piezas = []  # (clase, texto)
    for linea in texto.split("\n"):
        linea = linea.strip()
        if not linea:
            continue
        # El PDF anterior tenía este subtítulo sin contenido; nunca fue una
        # instrucción. Se descarta únicamente en esa posición y ese apartado.
        if (campo == "Rotulación"
                and normalizar(linea) == "procedencia y trazabilidad avanzada"
                and not any(x.strip() for x in texto.split(linea, 1)[-1].splitlines())):
            continue
        # Folios o cabeceras claramente identificables, no números clínicos.
        if (re.fullmatch(r"(?:p[aá]gina\s*:?\s*)?\d+\s*(?:de|/)\s*\d+", linea, re.I)
                or re.fullmatch(r"p[aá]gina\s*:?\s*\d+", linea, re.I)):
            continue
        if normalizar(linea) in titulos:
            piezas.append(["titulo", linea])
            continue
        vi = re.match(r"^(?:[●•▪◦∙]\s*|[-–]\s+)(.*)$", linea)
        if vi:
            piezas.append(["vineta", re.sub(r"\s+", " ", vi.group(1)).strip()])
            continue
        # Continuación de una viñeta o párrafo cortado por el ancho del PDF.
        if piezas and piezas[-1][0] in ("vineta", "parrafo"):
            piezas[-1][1] += " " + linea
        else:
            piezas.append(["parrafo", linea])

    lineas = []
    for clase, frase in piezas:
        frase = re.sub(r"\s+", " ", frase).strip()
        if not frase:
            continue
        if clase == "titulo":
            if lineas and lineas[-1] != "":
                lineas.append("")
            lineas.append(frase)
        elif clase == "vineta":
            lineas.append("  • " + frase)
        else:
            lineas.append(frase)
    return "\n".join(lineas).strip()


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


def guardar_normas(db_path, archivo, nombre_archivo, conexion=None):
    campos = extraer_normas(archivo, nombre_archivo)  # validar antes de modificar SQLite
    tipo = ("application/pdf" if nombre_archivo.lower().endswith(".pdf")
            else "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    with (nullcontext(conexion) if conexion is not None else sqlite3.connect(db_path)) as con:
        con.execute("DELETE FROM normas_generales_campos")
        con.executemany("INSERT INTO normas_generales_campos(campo,contenido) VALUES (?,?)", list(campos.items()))
        con.execute("""INSERT OR REPLACE INTO documentos_fuente
            (fuente_id,nombre_archivo,tipo_mime,archivo,sha256,fecha_actualizacion,registros)
            VALUES (?,?,?,?,?,?,?)""",
            (NORMAS_ID,nombre_archivo,tipo,sqlite3.Binary(archivo),hashlib.sha256(archivo).hexdigest(),_fecha(),4))
    return len(campos)


def guardar_cartera_acreditacion(db_path, archivo, nombre_archivo, conexion=None):
    import json
    registros = _leer_cartera(archivo)  # validar antes de modificar SQLite
    filas = [(numero,ident,codigo,prestacion,prest_norm,area,area_norm,
              json.dumps(datos,ensure_ascii=False),json.dumps(estados,ensure_ascii=False))
             for numero,ident,codigo,prestacion,prest_norm,area,area_norm,datos,estados in registros]
    with (nullcontext(conexion) if conexion is not None else sqlite3.connect(db_path)) as con:
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


# Términos descriptivos o administrativos que no identifican por sí solos
# una prestación. No se expande «AC.» a un término clínico supuesto.
_TERMINOS_GENERICOS_ACREDITACION = frozenset({
    "ac", "o", "u", "y", "e", "de", "del", "la", "el", "los", "las",
    "en", "por", "para", "con", "sin", "nuevo", "nueva", "taco",
    "test", "prueba", "examen", "determinacion", "analisis", "perfil",
})


def _terminos_clinicos(nombre_normalizado):
    """Palabras completas distintivas; no deduce sinónimos clínicos."""
    return {t for t in nombre_normalizado.split()
            if len(t) >= 4 and t not in _TERMINOS_GENERICOS_ACREDITACION}


def similitud_sugerencia_clinica(nombre_examen, nombre_prestacion):
    """Devuelve (prioridad, similitud TEXTUAL) con el mismo criterio en ambas carteras.

    Comparar solo caracteres perjudica abreviaturas como «AC. LUPICO» frente
    a «AC. CIRCULANTES O AC. LUPICO NUEVO»: el nombre íntegro está incluido,
    pero el texto adicional reduce artificialmente la semejanza global.

    Se compara la frase completa por palabras independientes y se valora su
    cobertura cuando el nombre consultado tiene varios términos. Los parecidos
    superficiales sin palabras clínicas compartidas se atenúan solo en estas
    consultas compuestas; una sola palabra sigue admitiendo errores tipográficos
    y variantes difusas como PROTROMBINA / PROGESTERONA.

    No traduce siglas, no infiere equivalencias clínicas y no realiza vínculos.
    """
    objetivo = normalizar(nombre_examen)
    candidato = normalizar(nombre_prestacion)
    if not objetivo or not candidato:
        return 0, 0.0
    if objetivo == candidato:
        return 3, 1.0

    literal = SequenceMatcher(None, objetivo, candidato).ratio()
    terminos_obj = _terminos_clinicos(objetivo)
    terminos_cand = _terminos_clinicos(candidato)
    compartidos = terminos_obj & terminos_cand
    frase_completa = bool(re.search(
        r"(?<!\w)" + re.escape(objetivo) + r"(?!\w)", candidato))

    if frase_completa:
        # En una consulta de varias palabras se conserva el orden y la
        # totalidad de la frase. La cobertura de esa frase (1.0) tiene más
        # peso que la longitud de los calificadores administrativos.
        if len(objetivo.split()) > 1 and terminos_obj:
            puntuacion = 0.25 * literal + 0.75
        else:
            # Preservar el criterio ya validado para nombres simples:
            # PROTROMBINA (TP) 91 %, PROTROMBINA TACO NUEVO 75 %.
            puntuacion = 0.75 * literal + 0.25
        return 2, min(1.0, puntuacion)

    if compartidos:
        cobertura_obj = len(compartidos) / len(terminos_obj)
        cobertura_cand = len(compartidos) / len(terminos_cand)
        puntuacion = max(literal, 0.55 + 0.20 * cobertura_obj + 0.10 * cobertura_cand)
        return 1, min(1.0, puntuacion)

    # «AC. LUPICO» y «ACIDO URICO» se asemejan en caracteres, pero no
    # comparten «lupico», la palabra distintiva de la consulta. Su semejanza
    # literal no debería superar una frase que sí conserva el nombre completo.
    if len(objetivo.split()) > 1 and terminos_obj:
        literal *= 0.75
    return 0, literal


def _similitud_sugerencia_acreditacion(nombre_examen, nombre_prestacion):
    """Compatibilidad: Acreditación y CRS utilizan ahora el mismo puntaje."""
    return similitud_sugerencia_clinica(nombre_examen, nombre_prestacion)


def candidatos_acreditacion(
    db_path, nombre_examen, area, limite=8, similitud_minima=0.0,
    nombre_referencia=None,
):
    """Sugiere prestaciones sin vincularlas automáticamente.

    - Filtra ANTES de limitar a cinco opciones visibles.
    - Prioriza nombres clínicos compartidos sobre parecidos de caracteres.
    - Considera el nombre CRS confirmado por el usuario como segunda consulta;
      no afecta la coincidencia exacta, que solo usa el examen activo.
    - El área solo desempata: nunca aumenta el porcentaje visible.
    """
    consultas = list(dict.fromkeys(filter(None, (
        normalizar(nombre_examen), normalizar(nombre_referencia),
    ))))
    if not consultas or limite <= 0:
        return []
    area_norm = area_equivalente(area)
    with sqlite3.connect(db_path) as con:
        filas = con.execute(
            "SELECT id,prestacion,area,prestacion_norm,area_norm "
            "FROM prestaciones_acreditacion"
        ).fetchall()
    candidatas = []
    for identificador, nombre, nombre_area, nombre_norm, area_dest in filas:
        # La consulta de la cartera se normaliza con la misma regla que los
        # nombres almacenados; usar el visible si el índice antiguo está vacío.
        candidato = normalizar(nombre_norm or nombre)
        prioridad, ratio = max(
            (_similitud_sugerencia_acreditacion(consulta, candidato)
             for consulta in consultas),
            key=lambda par: (par[0], par[1]),
        )
        if ratio < similitud_minima:
            continue
        misma_area = bool(area_norm and area_norm == area_dest)
        candidatas.append((prioridad, ratio, misma_area,
                           identificador, nombre, nombre_area))
    candidatas.sort(key=lambda r: (-r[0], -r[1], -r[2], normalizar(r[4]), r[3]))
    return [(identificador, nombre, nombre_area, ratio)
            for _, ratio, _, identificador, nombre, nombre_area in candidatas[:limite]]


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


def coincidencia_exacta_acreditacion(db_path, nombre_examen, area):
    """Retorna (ID, tipo de coincidencia), o (None, None) si es ambigua.

    Reutiliza EXACTAMENTE las reglas de selección institucional de _buscar_match;
    nunca promociona un candidato de similitud a coincidencia exacta.
    """
    with closing(sqlite3.connect(db_path)) as con:
        fila, motivo = _buscar_match(con, nombre_examen, area)
    return (fila[0], motivo) if fila else (None, None)


def sincronizar_vinculo_acreditacion_en_sesion(
    estado, db_path, examen, area, opciones, revision=0
):
    """Sincroniza el selectbox y el vínculo real antes de crear el widget.

    Respeta la selección o eliminación deliberada del usuario. Si cambia el
    examen o el área, vuelve a evaluar el vínculo exacto. Si solo cambia la
    revisión de la Cartera, mantiene una elección manual que todavía exista.
    ``estado`` admite st.session_state y diccionarios de pruebas.
    No escribe en SQLite.
    """
    disponibles = set(opciones)
    contexto = (str(examen or ""), str(area or ""), int(revision or 0))
    anterior = estado.get("contexto_emparejamiento_acreditacion")
    actual = estado.get("acreditacion_manual_selector")
    manual = bool(estado.get("acreditacion_seleccion_confirmada_manualmente", False))
    exacta, detalle = coincidencia_exacta_acreditacion(db_path, examen, area)
    if exacta not in disponibles:
        exacta, detalle = None, None

    if anterior != contexto:
        mismo_examen_area = (isinstance(anterior, tuple) and anterior[:2] == contexto[:2])
        conserva_manual = (
            manual and mismo_examen_area
            and (actual is None or actual in disponibles)
        )
        # Un vínculo no vacío en sesiones de versiones anteriores fue elegido
        # por el usuario: no reemplazarlo por un nombre exacto recién detectado.
        vinculo_heredado = (
            anterior is None and not manual and actual is not None
            and actual in disponibles
        )
        if conserva_manual or vinculo_heredado:
            estado["acreditacion_seleccion_confirmada_manualmente"] = True
        else:
            estado["acreditacion_manual_selector"] = exacta
            estado["acreditacion_seleccion_confirmada_manualmente"] = False
        estado["contexto_emparejamiento_acreditacion"] = contexto
    elif actual not in disponibles and actual is not None:
        estado["acreditacion_manual_selector"] = exacta
        estado["acreditacion_seleccion_confirmada_manualmente"] = False
    elif not manual and actual != exacta:
        # Recuperar widgets antiguos que quedaron vacíos, sin tocar el valor
        # de una selección manual explícita.
        estado["acreditacion_manual_selector"] = exacta

    elegido = estado.get("acreditacion_manual_selector")
    estado["acreditacion_manual_id"] = elegido
    estado["acreditacion_coincidencia_automatica"] = (
        detalle if elegido is not None
        and elegido == exacta
        and not estado.get("acreditacion_seleccion_confirmada_manualmente", False)
        else None
    )
    return elegido


def obtener_valores_fuentes(db_path, nombre_examen, area=None, seleccion_manual=None, indicaciones_cartera=""):
    """Regresa (valores, orígenes, prestación vinculada, mensaje).

    Preparación: Indicaciones CRS > Acreditación válida > Normas.
    Procedimiento/Rotulación/Rechazo: Normas válidas > Acreditación válida.
    Para Muestra/Condiciones: solo Acreditación válida.
    Recepción/Procesamiento: sin fuente; quedan vacíos salvo edición manual.
    """
    import json
    valores, origen = {}, {}
    with sqlite3.connect(db_path) as con:
        for campo, contenido in con.execute("SELECT campo, contenido FROM normas_generales_campos"):
            if campo in CAMPOS_NORMAS and (preparacion_valida(contenido) if campo == "Preparación del paciente" else texto_valido(contenido)):
                valores[campo] = limpiar_formato_normas(campo, contenido)
                origen[campo] = "Normas Generales de Toma de Muestras"
        fila, mensaje = _buscar_match(con,nombre_examen,area,seleccion_manual)
        seleccion = None
        if fila:
            identificador, prestacion, area_prestacion, datos_raw, estados_raw = fila
            seleccion = {"id": identificador, "prestacion": prestacion, "area": area_prestacion}
            datos, estados = json.loads(datos_raw), json.loads(estados_raw)
            for campo in CAMPOS_CONTROLADOS:
                texto = datos.get(campo,"")
                valido = (preparacion_valida(texto, estados.get(campo), fuente_acreditacion=True)
                          if campo == "Preparación del paciente"
                          else texto_valido(texto, estados.get(campo)))
                if valido:
                    # No sustituir un valor válido de Normas en los tres
                    # campos cuya fuente institucional principal es Normas.
                    if (campo in CAMPOS_NORMAS_PRIORITARIAS
                            and origen.get(campo) == "Normas Generales de Toma de Muestras"):
                        continue
                    valores[campo] = texto
                    origen[campo] = "Cartera de Acreditación · " + str(estados.get(campo))
        # «Indicaciones» solo se recibe del examen CRS efectivamente emparejado.
        # «No aplica» documentado en Acreditación ahora es válido y conserva
        # prioridad sobre Normas. En el CRS prevalece cualquier indicación
        # no vacía y válida, incluido «No aplica».
        if texto_valido(indicaciones_cartera):
            valores["Preparación del paciente"] = indicaciones_cartera
            origen["Preparación del paciente"] = "Cartera de Exámenes del CRS · Indicaciones"
    return valores, origen, seleccion, mensaje
