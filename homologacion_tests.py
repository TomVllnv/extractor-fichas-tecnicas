"""Fuente institucional «Homologació de test» para el campo Test incluidos.

El documento se carga como Excel y se conserva íntegro en documentos_fuente.
Solo se asignan valores cuando la identidad del examen es inequívoca; nunca
se aplican coincidencias difusas automáticamente.
"""
from __future__ import annotations

import hashlib
import io
import re
import sqlite3
import unicodedata
from contextlib import nullcontext
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import load_workbook

FUENTE_ID = "homologacion_tests"
ETIQUETA = "Homologació de test"
ENCABEZADO_RESULTADOS = "¿QUÉ EXAMEN O RESULTADOS INCLUYE?"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Archivo institucional aportado para el primer alta. Se distribuye intacto:
# no es una base de datos ni se importa sin autorización del encargado.
ARCHIVO_INICIAL_HEMATOLOGIA = Path(__file__).resolve().parent / "recursos_iniciales" / "Homologacion_de_test.xlsx"


def archivo_inicial_hematologia():
    """Devuelve el Excel original para una importación confirmada, si existe."""
    if not ARCHIVO_INICIAL_HEMATOLOGIA.is_file():
        return None
    return ARCHIVO_INICIAL_HEMATOLOGIA.read_bytes()


# La fuente y su índice se guardan por área. El nombre del archivo no se
# utiliza para inferir su área; el responsable la selecciona en la interfaz.
AREAS = (
    "Microbiología",
    "Hematología y coagulación",
    "Bioquímica",
    "Hormonas e inmunidad",
    "Derivados",
)
IDS_POR_AREA = {
    "Microbiología": "homologacion_tests:microbiologia",
    "Hematología y coagulación": "homologacion_tests:hematologia_y_coagulacion",
    "Bioquímica": "homologacion_tests:bioquimica",
    "Hormonas e inmunidad": "homologacion_tests:hormonas_e_inmunidad",
    "Derivados": "homologacion_tests:derivados",
}


def id_fuente_area(area):
    """Identificador estable de documento, aislado por área."""
    if area not in IDS_POR_AREA:
        raise ValueError("Selecciona un área válida para Homologación de test.")
    return IDS_POR_AREA[area]


def recurso_area(area):
    return "fuente:" + id_fuente_area(area)


def _area_celda_compatible(valor, area):
    """Compara áreas de origen, sin permitir asignaciones entre laboratorios."""
    nombre = _norm(valor)
    if not nombre:
        return True
    equivalencias = {
        "Hematología y coagulación": {
            _norm("Hematología"), _norm("Coagulación"),
            _norm("Hematología y coagulación"),
            # Grupos SUBORDINADOS de esta planilla real, no áreas independientes.
            _norm("Microscopía"), _norm("Perfil compuesto"),
        },
        "Microbiología": {_norm("Microbiología")},
        "Bioquímica": {_norm("Bioquímica")},
        "Hormonas e inmunidad": {_norm("Hormonas"), _norm("Inmunidad"), _norm("Hormonas e inmunidad")},
        "Derivados": {_norm("Derivados")},
    }
    return nombre in equivalencias[area]



def _norm(valor):
    s = unicodedata.normalize("NFKD", str(valor or "").casefold())
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(re.findall(r"[a-z0-9]+", s))


_NOMBRES_EXAMEN = frozenset(map(_norm, (
    "Examen", "Nombre examen", "Nombre del examen", "Nombre de examen",
    "Examen LIS", "Nombre LIS", "Test", "Prueba", "Prestación", "Nombre prestación",
    "Determinación", "Determinaciones", "Nombre del test",
    "Prestación / Nombre", "Prestación / Nombre del examen",
)))
_CODIGOS_EXAMEN = frozenset(map(_norm, (
    "Código LIS", "Código del examen LIS", "Código examen LIS",
)))
_AREAS_EXAMEN = frozenset(map(_norm, (
    "Área", "Area", "Área laboratorio", "Sección", "Unidad", "Grupo",
)))


def _texto_celda(valor):
    if valor is None or isinstance(valor, bool):
        return ""
    # Evita convertir un código entero en '1234.0' y preserva ceros iniciales
    # cuando el archivo almacena la celda como texto.
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()


def _identificar_encabezados(fila):
    indices = {}
    for indice, celda in enumerate(fila):
        clave = _norm(celda)
        if not clave:
            continue
        if clave == _norm(ENCABEZADO_RESULTADOS):
            indices.setdefault("resultados", []).append(indice)
        elif clave in _NOMBRES_EXAMEN:
            indices.setdefault("examen", []).append(indice)
        elif clave in _CODIGOS_EXAMEN:
            indices.setdefault("codigo", []).append(indice)
        elif clave in _AREAS_EXAMEN:
            indices.setdefault("area", []).append(indice)
    return indices


def inspeccionar_encabezados(archivo):
    """Obtiene las hojas con la columna de resultados para poder elegir la clave.

    No infiere columnas arbitrarias: el usuario puede indicar cuál identifica
    el examen antes de confirmar el reemplazo en la base de datos.
    """
    if len(archivo) > 30 * 1024 * 1024:
        raise ValueError("El Excel supera el límite de 30 MB.")
    try:
        libro = load_workbook(io.BytesIO(archivo), read_only=True, data_only=True)
    except Exception as exc:
        raise ValueError("El archivo no es un Excel .xlsx válido.") from exc
    tablas = []
    try:
        for nombre_hoja in libro.sheetnames:
            for numero, fila in enumerate(libro[nombre_hoja].iter_rows(
                    min_row=1, max_row=20, values_only=True), 1):
                if 'resultados' in _identificar_encabezados(fila):
                    tablas.append((nombre_hoja, numero,
                                   tuple(_texto_celda(v) for v in fila)))
    finally:
        libro.close()
    return tablas


def leer_excel(archivo, *, hoja=None, columna_examen=None, columna_codigo=None):
    """Valida Excel y devuelve filas aptas para indexar, sin alterar SQLite.

    Se rechazan las hojas sin columna inequívoca de examen/código, las columnas
    de resultados duplicadas y la ambigüedad entre hojas. Si no hay título de
    examen reconocido, la UI debe pedir explícitamente qué columna lo contiene.
    """
    if not isinstance(archivo, (bytes, bytearray)) or not archivo:
        raise ValueError("Selecciona un archivo Excel válido.")
    if len(archivo) > 30 * 1024 * 1024:
        raise ValueError("El Excel supera el límite de 30 MB.")
    try:
        libro = load_workbook(io.BytesIO(archivo), read_only=True, data_only=True)
    except Exception as exc:
        raise ValueError("El archivo no es un Excel .xlsx válido.") from exc
    try:
        candidatas = []
        for nombre_hoja in libro.sheetnames:
            if hoja is not None and nombre_hoja != hoja:
                continue
            ws = libro[nombre_hoja]
            for numero, fila in enumerate(ws.iter_rows(min_row=1, max_row=20, values_only=True), 1):
                columnas = _identificar_encabezados(fila)
                if "resultados" in columnas:
                    candidatas.append((nombre_hoja, numero, tuple(_texto_celda(v) for v in fila), columnas))
        if not candidatas:
            raise ValueError('No se encontró el encabezado «¿QUÉ EXAMEN O RESULTADOS INCLUYE?» en las primeras 20 filas.')
        if len(candidatas) != 1:
            raise ValueError("Se encontraron varias tablas de resultados. Selecciona una hoja y comprueba que tenga una sola tabla.")
        hoja_elegida, fila_encabezado, encabezados, columnas = candidatas[0]
        if len(columnas['resultados']) != 1:
            raise ValueError("La columna de resultados aparece más de una vez en la misma tabla.")
        if columna_examen is None:
            indices_examen = columnas.get('examen', [])
            if len(indices_examen) > 1:
                raise ValueError("Hay varias columnas que podrían identificar el examen; selecciona una explícitamente.")
            columna_examen = indices_examen[0] if indices_examen else None
        if columna_codigo is None:
            indices_codigo = columnas.get('codigo', [])
            columna_codigo = indices_codigo[0] if len(indices_codigo) == 1 else None
        if columna_examen is None and columna_codigo is None:
            raise ValueError("No se encontró una columna para identificar cada examen. Selecciona la columna de examen antes de importar.")
        if columna_examen is not None and (columna_examen < 0 or columna_examen >= len(encabezados)):
            raise ValueError("La columna de examen seleccionada no existe.")
        if columna_codigo is not None and (columna_codigo < 0 or columna_codigo >= len(encabezados)):
            raise ValueError("La columna de código seleccionada no existe.")
        idx_resultados = columnas['resultados'][0]
        if columna_codigo is not None and any(
            palabra in _norm(encabezados[columna_codigo]).split()
            for palabra in ("fonasa", "relaciona", "relacionados", "externo")
        ):
            raise ValueError("No utilices códigos FONASA ni columnas de códigos relacionados "
                             "como si fueran Código LIS.")
        if idx_resultados in (columna_examen, columna_codigo):
            raise ValueError("La columna de resultados no puede ser el identificador del examen.")
        idx_area = columnas.get('area', [None])[0]
        registros = []
        ws = libro[hoja_elegida]
        for num_fila, fila in enumerate(ws.iter_rows(min_row=fila_encabezado + 1, values_only=True), fila_encabezado + 1):
            if num_fila > 100020:
                raise ValueError("El Excel contiene demasiadas filas para procesar de manera segura.")
            def celda(indice):
                return _texto_celda(fila[indice]) if indice is not None and indice < len(fila) else ""
            resultados = celda(idx_resultados)
            if not resultados:
                continue
            examen = celda(columna_examen)
            codigo = celda(columna_codigo)
            area = celda(idx_area)
            if not examen and not codigo:
                # No atribuir un resultado sin identificador al examen de la fila anterior.
                continue
            registros.append((hoja_elegida, num_fila, examen, _norm(examen), codigo,
                              _norm(codigo), area, _norm(area), resultados))
        if not registros:
            raise ValueError("El Excel no contiene filas con resultados e identificador de examen.")
        return registros, {
            'hoja': hoja_elegida,
            'fila_encabezado': fila_encabezado,
            'encabezados': encabezados,
            'columna_examen': columna_examen,
            'columna_codigo': columna_codigo,
            'registros': len(registros),
        }
    finally:
        libro.close()


def inicializar(db_path, conexion=None):
    with (nullcontext(conexion) if conexion is not None else sqlite3.connect(db_path)) as con:
        con.execute('''CREATE TABLE IF NOT EXISTS homologacion_tests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            hoja TEXT NOT NULL,
            fila_excel INTEGER NOT NULL,
            examen TEXT,
            examen_norm TEXT,
            codigo TEXT,
            codigo_norm TEXT,
            area TEXT,
            area_norm TEXT,
            resultados TEXT NOT NULL,
            area_fuente TEXT NOT NULL DEFAULT ''
        )''')
        columnas = {fila[1] for fila in con.execute('PRAGMA table_info(homologacion_tests)')}
        if 'area_fuente' not in columnas:
            con.execute("ALTER TABLE homologacion_tests ADD COLUMN area_fuente TEXT NOT NULL DEFAULT ''")
        con.execute('CREATE INDEX IF NOT EXISTS idx_homologacion_fuente ON homologacion_tests(area_fuente)')
        # Migración conservadora de instalaciones previas: solo documentos
        # inequívocamente hematológicos. No etiquetar archivos de otras áreas.
        existe_documentos = con.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='documentos_fuente'").fetchone()
        if existe_documentos:
            original = con.execute('SELECT fuente_id, nombre_archivo, tipo_mime, archivo, sha256, fecha_actualizacion, registros FROM documentos_fuente WHERE fuente_id=?', (FUENTE_ID,)).fetchone()
            destino = id_fuente_area('Hematología y coagulación')
            ya_migrado = con.execute('SELECT 1 FROM documentos_fuente WHERE fuente_id=?', (destino,)).fetchone()
            if original and not ya_migrado:
                ambiguas = con.execute("SELECT area FROM homologacion_tests WHERE area_fuente='' AND area_norm <> ''").fetchall()
                if all(_area_celda_compatible(fila[0], 'Hematología y coagulación') for fila in ambiguas):
                    con.execute("UPDATE homologacion_tests SET area_fuente=? WHERE area_fuente=''", ('Hematología y coagulación',))
                    con.execute('INSERT INTO documentos_fuente (fuente_id,nombre_archivo,tipo_mime,archivo,sha256,fecha_actualizacion,registros) VALUES (?,?,?,?,?,?,?)', (destino, *original[1:]))
        con.execute('CREATE INDEX IF NOT EXISTS idx_homologacion_examen ON homologacion_tests(examen_norm, area_norm)')
        con.execute('CREATE INDEX IF NOT EXISTS idx_homologacion_codigo ON homologacion_tests(codigo_norm)')


def guardar(db_path, archivo, nombre_archivo, conexion=None, *, hoja=None, columna_examen=None, columna_codigo=None, area_fuente=None):
    registros, metadatos = leer_excel(archivo, hoja=hoja, columna_examen=columna_examen,
                                     columna_codigo=columna_codigo)
    if not str(nombre_archivo or '').lower().endswith('.xlsx'):
        raise ValueError("El documento debe tener formato .xlsx.")
    # Las nuevas importaciones de la interfaz exigen un área. El modo legado
    # sin área se conserva únicamente para pruebas/migraciones antiguas.
    if area_fuente is not None:
        id_doc = id_fuente_area(area_fuente)
        incompatibles = [fila[6] for fila in registros
                         if fila[6] and not _area_celda_compatible(fila[6], area_fuente)]
        if incompatibles:
            raise ValueError("El archivo contiene exámenes identificados con otra área (por ejemplo: "
                             + incompatibles[0] + "). Revisa la pestaña elegida antes de guardar.")
    else:
        id_doc = FUENTE_ID
    with (nullcontext(conexion) if conexion is not None else sqlite3.connect(db_path)) as con:
        inicializar(db_path, conexion=con)
        con.execute('DELETE FROM homologacion_tests WHERE area_fuente=?', (area_fuente or '',))
        con.executemany('''INSERT INTO homologacion_tests
            (hoja, fila_excel, examen, examen_norm, codigo, codigo_norm, area, area_norm, resultados, area_fuente)
            VALUES (?,?,?,?,?,?,?,?,?,?)''', [(*fila, area_fuente or '') for fila in registros])
        con.execute('''INSERT OR REPLACE INTO documentos_fuente
            (fuente_id,nombre_archivo,tipo_mime,archivo,sha256,fecha_actualizacion,registros)
            VALUES (?,?,?,?,?,?,?)''',
            (id_doc, nombre_archivo, MIME_XLSX, sqlite3.Binary(archivo),
             hashlib.sha256(archivo).hexdigest(),
             datetime.now(ZoneInfo('America/Santiago')).isoformat(timespec='seconds'),
             len(registros)))
    return metadatos


def obtener_tests(db_path, *, examen, area=None, codigo=None):
    """Retorna (valor, detalle). Solo coincidencia inequívoca de código o nombre.

    En caso de ambigüedad NO devuelve valor: se debe resolver manualmente.
    Una coincidencia aproximada jamás autoriza el llenado automático.
    """
    examen_norm, area_norm, codigo_norm = _norm(examen), _norm(area), _norm(codigo)
    # Sin área explícita solo es consultable el índice legado, nunca registros
    # de pestañas nuevas. Evita mezclar estudios homónimos de distintas áreas.
    indice = area if area in AREAS else ''
    if area and area not in AREAS:
        return '', {'estado': 'area_desconocida'}
    with sqlite3.connect(db_path) as con:
        try:
            filas = con.execute('''SELECT hoja, fila_excel, examen, codigo, area,
                                        resultados, examen_norm, codigo_norm, area_norm
                                 FROM homologacion_tests WHERE area_fuente=?''', (indice,)).fetchall()
        except sqlite3.OperationalError as exc:
            if 'no such table' in str(exc).lower():
                return '', {'estado': 'sin_documento'}
            raise
    if not filas:
        return '', {'estado': 'sin_documento_area'}
    por_codigo = [f for f in filas if codigo_norm and f[7] == codigo_norm]
    if area in AREAS and por_codigo and any(f[4] and not _area_celda_compatible(f[4], area) for f in por_codigo):
        return '', {'estado': 'requiere_revision',
                    'motivo': 'El código se encuentra asociado a otra área.'}
    # Código LIS EXACTO y único tiene precedencia incluso si la denominación
    # institucional utiliza una abreviatura diferente. Nunca usar códigos
    # "relacionados" como identidad: solo columna CÓDIGO LIS.
    # Si el nombre corresponde exactamente a otra prestación con otro código,
    # hay un conflicto real y se exige revisión.
    por_nombre = [f for f in filas if examen_norm and f[6] == examen_norm]
    if por_codigo and por_nombre and any(f[7] and f[7] != codigo_norm for f in por_nombre):
        return '', {'estado': 'requiere_revision', 'motivo': 'El nombre y el Código LIS corresponden a prestaciones diferentes.'}
    if por_codigo and len({(f[6], f[7]) for f in por_codigo}) > 1:
        return '', {'estado': 'requiere_revision', 'motivo': 'El mismo Código LIS aparece en más de una prestación.'}
    if por_nombre and area in AREAS:
        por_nombre = [f for f in por_nombre if _area_celda_compatible(f[4], area)]
    candidatas = por_codigo or por_nombre
    if not candidatas:
        return '', {'estado': 'sin_coincidencia'}
    # No aceptar un código diferente conocido para el mismo nombre sin revisión.
    if codigo_norm and por_nombre and any(f[7] and f[7] != codigo_norm for f in por_nombre):
        return '', {'estado': 'requiere_revision', 'motivo': 'Existen códigos distintos para ese nombre.'}
    valores = {_norm(f[5]) for f in candidatas}
    if len(valores) != 1:
        return '', {'estado': 'requiere_revision', 'motivo': 'Hay resultados diferentes para el mismo examen.'}
    # No confundir mismo nombre en áreas diferentes si la fuente especifica área.
    if por_nombre and not por_codigo and not area_norm and len({f[8] for f in por_nombre if f[8]}) > 1:
        return '', {'estado': 'requiere_revision', 'motivo': 'El examen se repite en varias áreas.'}
    return candidatas[0][5], {'estado': 'encontrado', 'hoja': candidatas[0][0],
                              'fila_excel': candidatas[0][1], 'examen': candidatas[0][2],
                              'codigo': candidatas[0][3],
                              'coincidencia': 'codigo_lis' if por_codigo else 'nombre_exacto'}
