"""Lectura conservadora de fichas heterogéneas de laboratorios derivados.

Cada etiqueta/valor es una unidad. Un renglón desconocido pone fin al valor
anterior; nunca se rellena un campo a partir de texto sin etiqueta verificable.
La fuente seleccionada y las discrepancias se resuelven en app.py, no aquí.
"""
from __future__ import annotations

import re
import unicodedata
from collections import defaultdict


def normalizar(texto):
    t = unicodedata.normalize('NFKD', str(texto or '').casefold())
    t = ''.join(ch for ch in t if not unicodedata.combining(ch))
    return ' '.join(re.sub(r'[^a-z0-9]+', ' ', t).split())


# Sin marcas comerciales: equivalencias conceptuales de laboratorios derivados.
# No asociar códigos internos ni FONASA al Código de la ficha (LIS/CRS).
_ETIQUETAS = {
    'plazo': ('tiempo de respuesta', 'plazo de entrega', 'plazo entrega resultados',
              'tiempo de entrega', 'tiempo de respuesta del examen', 'tiempo de informe'),
    'metodo': ('principio del metodo', 'principio del ensayo', 'metodo', 'metodologia',
               'metodo analitico', 'tecnica analitica', 'tecnica de analisis'),
    'clinica': ('proposito del examen', 'proposito de la prueba', 'utilidad clinica',
                'informacion clinica', 'indicacion clinica', 'finalidad del examen',
                'significado clinico', 'caracteristicas del examen'),
    'muestra': ('tipo de muestra requerida', 'muestra requerida', 'tipo de muestra',
                'material biologico', 'tipo de especimen', 'especimen requerido'),
    'muestra_primaria': ('tipo de muestra primaria', 'muestra primaria'),
    'volumen': ('volumen minimo', 'volumen requerido', 'cantidad minima de muestra'),
    'preparacion': ('preparacion del paciente', 'indicaciones al paciente',
                    'preparacion de paciente', 'requisitos del paciente'),
    'toma': ('indicaciones toma de muestra', 'indicaciones para la toma de muestra',
             'procedimiento de toma de muestra', 'obtencion de la muestra',
             'recoleccion de la muestra'),
    'estabilidad': ('estabilidad de la muestra', 'estabilidad', 'conservacion de la muestra',
                    'tiempo de estabilidad', 'almacenamiento de muestra'),
    'transporte': ('condiciones de transporte', 'temperatura de transporte',
                   'condiciones de envio', 'transporte de muestra', 'transporte de muestras',
                   'condiciones de traslado'),
    'rechazo': ('criterios de rechazo', 'criterios de no aceptacion',
                'causales de rechazo', 'muestras no conformes'),
    'referencia': ('valores de referencia', 'intervalos de referencia',
                   'rangos de referencia'),
    'critico': ('valor critico', 'valores criticos', 'valores de alerta'),
    'equipamiento': ('equipo utilizado', 'equipamiento utilizado', 'analizador',
                     'instrumento utilizado', 'plataforma analitica'),
}
_LABEL_TO_KEY = {normalizar(alias): key for key, aliases in _ETIQUETAS.items() for alias in aliases}
# Etiquetas frecuentes que SI deben frenar un campo, pero no son fuente de ese campo.
_ETIQUETAS_SIN_DESTINO = frozenset(map(normalizar, (
    'codigo bk', 'codigo fonasa', 'codigo interno', 'codigo de prestacion',
    'unidad de medida', 'unidad de medicion', 'unidades', 'dias y hora de proceso',
    'dias de proceso', 'medio de transporte', 'observaciones adicionales',
    'componentes/alergeno', 'componentes alergeno', 'informacion general',
    'muestra y transporte', 'observaciones', 'requisitos administrativos',
    'tabla de temperaturas', 'tabla de temperatura', 'rangos de temperaturas',
)))
_ALIASES_ORDENADOS = sorted(_LABEL_TO_KEY, key=len, reverse=True)


def _valor_valido(valor):
    t = str(valor or '').strip()
    n = normalizar(t)
    return bool(t and len(t) <= 2600 and n not in ('nan', 'null', 'none', 'pendiente', 'sin informacion', '-')
                and not re.fullmatch(r'[\W\d_]+', t))


def _separar_prefijo(linea):
    """Acepta solo etiquetas completas al inicio, jamás coincidencias dentro de prosa."""
    s = str(linea or '').strip()
    if not s:
        return None
    n = normalizar(s)
    if n in _LABEL_TO_KEY or n in _ETIQUETAS_SIN_DESTINO:
        return n, ''
    for etiqueta in _ALIASES_ORDENADOS:
        # En texto continuo se exige separador real; espacios simples no bastan.
        patron = re.compile(r'^\s*' + r'\s+'.join(map(re.escape, etiqueta.split())) +
                            r'\s*(?:[:：|]|\s{2,}|\s+[-–—]\s+)\s*(.+)$', re.I)
        m = patron.match(normalizar_acentos_sin_espacios(s))
        if m:
            # Obtener el valor del texto original mediante posición compatible.
            # No se utiliza en la ruta geométrica que ya tiene ambas columnas.
            original = re.match(r'(?is)^\s*(.{1,' + str(len(etiqueta) + 14) +
                                r'}?)(?:\s*[:：|]\s*|\s{2,}|\s+[-–—]\s+)(.+)$', s)
            if original and normalizar(original.group(1)) == etiqueta:
                return etiqueta, original.group(2).strip()
    # Respaldo para OCR sin separación de columnas: una etiqueta completa de
    # al menos tres palabras puede ir seguida del valor con un solo espacio.
    # No aplicar a títulos genéricos como "Método" o "Estabilidad".
    tokens = s.split()
    for etiqueta in _ALIASES_ORDENADOS:
        n = len(etiqueta.split())
        if n < 3 or len(tokens) <= n:
            continue
        if normalizar(' '.join(tokens[:n])) == etiqueta:
            resto = ' '.join(tokens[n:]).strip(' :\t')
            if resto and len(resto) <= 260:
                return etiqueta, resto
    return None


def normalizar_acentos_sin_espacios(texto):
    t = unicodedata.normalize('NFKD', str(texto))
    return ''.join(c for c in t if not unicodedata.combining(c)).casefold()


def _filas_textuales(texto):
    filas = []
    etiqueta = None
    valor = []

    def cerrar():
        nonlocal etiqueta, valor
        if etiqueta is not None:
            filas.append((etiqueta, ' '.join(valor).strip()))
        etiqueta, valor = None, []

    lineas = str(texto or '').splitlines()
    for linea in lineas:
        s = linea.strip()
        if not s:
            continue
        par = _separar_prefijo(s)
        if par:
            cerrar()
            etiqueta, primer = par
            valor = [primer] if primer else []
        elif etiqueta is not None:
            # Continuación de valor: no consumir títulos ni texto largo de otra sección.
            n = normalizar(s)
            encabezado = (n in _ETIQUETAS_SIN_DESTINO or n in _LABEL_TO_KEY or
                          n.startswith(('tabla de temperaturas', 'rangos de temperaturas')) or
                          (':' in s and 2 <= s.index(':') <= 55
                           and len(s[:s.index(':')].split()) <= 7) or
                          (len(s) <= 60 and s.endswith(':')) or
                          (len(s) <= 54 and len(s.split()) <= 6 and s.isupper()))
            if encabezado:
                cerrar()
            elif len(' '.join(valor)) < 2200 and len(s) <= 300:
                valor.append(s)
            else:
                cerrar()
    cerrar()
    return filas


def _filas_geometricas(pdf_bytes):
    """Agrupa palabras por su altura; detecta separación entre columnas y fin de fila.

    El candidato a etiqueta puede ser desconocido: aun así corta la celda previa.
    """
    try:
        import fitz
        documento = fitz.open(stream=pdf_bytes, filetype='pdf')
    except Exception:
        return []
    resultado = []
    try:
        for pagina in documento:
            palabras = [w for w in pagina.get_text('words') if str(w[4]).strip()]
            if not palabras:
                continue
            palabras.sort(key=lambda w: ((w[1] + w[3])/2, w[0]))
            lineas = []
            for w in palabras:
                centro = (w[1] + w[3]) / 2
                if lineas and abs(centro - lineas[-1][0]) < max(2.7, (w[3]-w[1])*0.25):
                    lineas[-1][1].append(w)
                else:
                    lineas.append([centro, [w]])
            pivot = float(pagina.rect.width) * 0.365
            etiqueta = None
            partes = []

            def cerrar():
                nonlocal etiqueta, partes
                if etiqueta is not None:
                    resultado.append((etiqueta, ' '.join(partes).strip()))
                etiqueta, partes = None, []

            for _, palabras_linea in lineas:
                palabras_linea.sort(key=lambda w: w[0])
                # Normalmente la columna derecha empieza después del 36,5 %.
                izq = [w for w in palabras_linea if w[0] < pivot]
                der = [w for w in palabras_linea if w[0] >= pivot]
                t_izq = ' '.join(w[4] for w in izq).strip()
                t_der = ' '.join(w[4] for w in der).strip()
                # Si el límite atraviesa una etiqueta, tomar el mayor espacio
                # horizontal como delimitador solo cuando parezca una tabla.
                if t_izq and not t_der and len(palabras_linea) > 2:
                    huecos = [(palabras_linea[i+1][0]-palabras_linea[i][2], i)
                               for i in range(len(palabras_linea)-1)]
                    hueco, indice = max(huecos)
                    if hueco > 26 and 0 < indice < len(palabras_linea)-1:
                        izq_alt = ' '.join(w[4] for w in palabras_linea[:indice+1])
                        der_alt = ' '.join(w[4] for w in palabras_linea[indice+1:])
                        if normalizar(izq_alt) in _LABEL_TO_KEY or normalizar(izq_alt) in _ETIQUETAS_SIN_DESTINO:
                            t_izq, t_der = izq_alt, der_alt
                clave = normalizar(t_izq)
                if clave in _LABEL_TO_KEY or clave in _ETIQUETAS_SIN_DESTINO:
                    cerrar()
                    etiqueta, partes = clave, ([t_der] if t_der else [])
                elif t_izq and t_der:
                    # Fila no conocida, aun así constituye un límite real.
                    cerrar()
                elif not t_izq and t_der and etiqueta is not None:
                    # La columna derecha puede abarcar múltiples renglones.
                    if len(' '.join(partes)) < 2300:
                        partes.append(t_der)
                elif t_izq:
                    # Encabezado de tabla, otra celda o párrafo separado.
                    # Nunca continuar arrastrando el valor anterior.
                    cerrar()
            cerrar()
    finally:
        documento.close()
    return resultado



# «Tabla de temperaturas» es una leyenda REFERENCIAL, no un nuevo plazo de
# estabilidad. Solo se adjunta el rango de la condición mencionada para ese
# examen, sin inferir que el tiempo corresponda necesariamente al transporte.
_TEMP_RE = re.compile(
    r'(?<![\w\d])(?P<desde>[−-]?\d{1,3})\s*(?:[–—−-]|\ba\b|\bhasta\b)\s*'
    r'(?P<hasta>[−-]?\d{1,3})\s*[°º]?\s*C\b'
    r'|(?<![\w\d])(?P<unico>[−-]?\d{1,3})\s*[°º]?\s*C\b', re.I)
_TEMP_ETIQUETA_RE = re.compile(r'\b(ambiente|refrigerad[ao]s?|congelad[ao]s?)\b', re.I)
_TEMP_TITULO_RE = re.compile(r'\b(?:tabla\s+de\s+temperaturas|rangos?\s+de\s+temperaturas)\b', re.I)
_TEMP_NOMBRES = {'ambiente': 'ambiente', 'refrigerada': 'refrigerada',
                 'congelada': 'congelada'}


def _tipo_temperatura(etiqueta):
    n = normalizar(etiqueta)
    if 'refrigerad' in n:
        return 'refrigerada'
    if 'congelad' in n:
        return 'congelada'
    if n == 'ambiente':
        return 'ambiente'
    return None


def _temperaturas_en_fragmento(fragmento):
    """Devuelve los rangos explícitos (posición, rango) sin interpolar valores."""
    encontrados = []
    for m in _TEMP_RE.finditer(fragmento):
        def numero(n):
            return int(n.replace('−', '-'))
        try:
            lo = numero(m.group('desde') or m.group('unico'))
            hi = numero(m.group('hasta')) if m.group('hasta') is not None else lo
        except (TypeError, ValueError):
            continue
        if not (-80 <= lo <= hi <= 80):
            continue
        fmt = lambda x: str(x).replace('-', '−')
        valor = (f'{fmt(lo)}–{fmt(hi)} °C' if lo != hi else f'{fmt(lo)} °C')
        encontrados.append((m.start(), valor))
    return encontrados


def _tabla_temperaturas_en_texto(texto):
    """Reconoce tablas de tres columnas y pares verticales etiqueta/temperatura.

    Requiere un encabezado de tabla explícito. No asigna temperaturas encontradas
    por azar en otras secciones de la ficha técnica.
    """
    texto = str(texto or '')
    sin_tildes = normalizar_acentos_sin_espacios(texto)
    for titulo in _TEMP_TITULO_RE.finditer(sin_tildes):
        # Limitar el análisis a la tabla y su nota, no a páginas posteriores.
        region = sin_tildes[titulo.end():titulo.end() + 480]
        etiquetas = [(m.start(), _tipo_temperatura(m.group()))
                     for m in _TEMP_ETIQUETA_RE.finditer(region)]
        etiquetas = [(pos, nom) for pos, nom in etiquetas if nom]
        temperaturas = _temperaturas_en_fragmento(region)
        if not etiquetas or not temperaturas:
            continue
        # Tabla horizontal: primero todos los nombres, luego todos los rangos.
        if (len(etiquetas) >= 2 and len(temperaturas) >= len(etiquetas)
                and etiquetas[-1][0] < temperaturas[0][0]):
            asignadas = {}
            for (_, nom), (_, rango) in zip(etiquetas[:3], temperaturas[:3]):
                asignadas.setdefault(nom, rango)
            if asignadas:
                return asignadas
        # Tabla vertical: ambiente -> valor, refrigerada -> valor, etc.
        asignadas = {}
        for indice, (pos, nom) in enumerate(etiquetas):
            limite = etiquetas[indice + 1][0] if indice + 1 < len(etiquetas) else len(region)
            opciones = [rango for posicion, rango in temperaturas if pos < posicion < limite]
            if len(opciones) == 1:
                asignadas.setdefault(nom, opciones[0])
        if asignadas:
            return asignadas
    return {}


def _tabla_temperaturas_en_pdf(pdf_bytes):
    """Lee rangos de una tabla digital por columna, aun si el orden del texto
    interno del PDF no coincide con el orden visual."""
    if not pdf_bytes:
        return {}
    try:
        import fitz
        doc = fitz.open(stream=pdf_bytes, filetype='pdf')
    except Exception:
        return {}
    try:
        for page in doc:
            if not _TEMP_TITULO_RE.search(
                    normalizar_acentos_sin_espacios(page.get_text('text'))):
                continue
            palabras = page.get_text('words')
            # Palabras de títulos, altura y x son fiables en PDFs digitales.
            etiquetas = []
            for w in palabras:
                key = _tipo_temperatura(w[4])
                if key:
                    etiquetas.append((key, (w[0] + w[2])/2, w[3]))
            if len(etiquetas) < 2:
                continue
            # Agrupar los encabezados de una misma fila horizontal, sin
            # confundir menciones posteriores en la nota explicativa.
            candidatas = []
            for etiq in etiquetas:
                grupo = [e for e in etiquetas if abs(e[2]-etiq[2]) < 8]
                if len({g[0] for g in grupo}) >= 2:
                    candidatas = sorted(grupo, key=lambda g: g[1])
                    break
            if not candidatas:
                continue
            asignadas = {}
            for i, (key, x, y) in enumerate(candidatas):
                izq = (x + candidatas[i-1][1])/2 if i else max(0, x-100)
                der = (x + candidatas[i+1][1])/2 if i+1 < len(candidatas) else min(page.rect.width, x+150)
                rect = fitz.Rect(izq, y+1, der, min(page.rect.height, y+75))
                rangos = _temperaturas_en_fragmento(page.get_text('text', clip=rect))
                if len(rangos) == 1:
                    asignadas[key] = rangos[0][1]
            if asignadas:
                return asignadas
    finally:
        doc.close()
    return {}


def _complementar_estabilidad_con_tabla(datos, tabla):
    """Añade una referencia explícita, sin convertir el transporte en estabilidad.

    Si la ficha dice «48 horas» y el transporte es «ambiente», el rango viene
    de la tabla pero NO se afirma «48 horas a 18–28 °C» (no documentado).
    """
    estabilidad = str(datos.get('Estabilidad de la muestra', '') or '').strip()
    if not estabilidad or not tabla or _temperaturas_en_fragmento(estabilidad):
        return
    condiciones = str(datos.get('Condiciones de envío al Laboratorio', '') or '')

    def tipos_indicados(texto):
        encontrados = []
        for m in _TEMP_ETIQUETA_RE.finditer(normalizar(texto)):
            tipo = _tipo_temperatura(m.group())
            if tipo in tabla and tipo not in encontrados:
                encontrados.append(tipo)
        return encontrados

    tipos = tipos_indicados(estabilidad)
    en_estabilidad = bool(tipos)
    if not tipos:
        tipos = tipos_indicados(condiciones)
    if not tipos:
        return
    equivalencias = '; '.join(f'{tipo}: {tabla[tipo]}' for tipo in tipos)
    if en_estabilidad:
        nota = f'Referencia de temperatura según tabla del documento: {equivalencias}.'
    else:
        nota = (f'Referencia de temperatura para las condiciones de transporte '
                f'indicadas (tabla del documento): {equivalencias}.')
    datos['Estabilidad de la muestra'] = estabilidad + '\n' + nota

def _convertir_campos(filas):
    crudos = defaultdict(list)
    for etiqueta, valor in filas:
        if etiqueta not in _LABEL_TO_KEY or not _valor_valido(valor):
            continue
        key = _LABEL_TO_KEY[etiqueta]
        if valor not in crudos[key]:
            crudos[key].append(valor)
    if not crudos:
        return {}
    datos = {}
    simples = {
        'plazo':'Plazo entrega resultados', 'metodo':'Método de procesamiento',
        'clinica':'Información clínica', 'preparacion':'Preparación del paciente',
        'toma':'Procedimiento de toma de muestras',
        'estabilidad':'Estabilidad de la muestra',
        'transporte':'Condiciones de envío al Laboratorio',
        'rechazo':'Criterios de rechazo', 'referencia':'Valores de referencia',
        'equipamiento':'Equipamiento utilizado',
    }
    for clave, campo in simples.items():
        if crudos.get(clave):
            datos[campo] = crudos[clave][0]
    muestras = crudos.get('muestra') or crudos.get('muestra_primaria')
    if muestras:
        piezas = [muestras[0]]
        if crudos.get('volumen'):
            piezas.append('Volumen mínimo: ' + crudos['volumen'][0])
        datos['Muestra requerida y recipiente'] = '\n'.join(piezas)
    return datos


def extraer_ficha_derivados(texto, pdf_bytes=None):
    """Valores estructurados verificables; no decide entre PDF y Cartera.

    Preferir geometría en PDF digital. El texto es el respaldo para OCR cuando
    el PDF no tiene palabras seleccionables. Sin pares fiables, no inferir.
    """
    filas = _filas_geometricas(pdf_bytes) if pdf_bytes else []
    resultados = _convertir_campos(filas)
    # No mezclar filas provenientes de dos representaciones desordenadas.
    # La ruta textual solo se utiliza si la geométrica no detectó valores.
    if not resultados:
        resultados = _convertir_campos(_filas_textuales(texto))
    # La tabla general solo aporta la equivalencia del rango señalado para el
    # examen; no reemplaza la duración ni las fuentes institucionales.
    # En PDFs digitales las coordenadas conservan la asociación de cada
    # encabezado con su columna, incluso si el texto plano se reordena.
    # Para documentos escaneados (sin capa textual), leer la tabla del texto
    # obtenido por lectura óptica.
    tabla = _tabla_temperaturas_en_pdf(pdf_bytes) if pdf_bytes else {}
    if not tabla:
        tabla = _tabla_temperaturas_en_texto(texto)
    _complementar_estabilidad_con_tabla(resultados, tabla)
    return resultados


def son_valores_distintos(campo, uno, otro):
    if not _valor_valido(uno) or not _valor_valido(otro):
        return False
    if campo == 'Equipamiento utilizado':
        import equipamientos as eq
        return not eq.equipamientos_equivalentes(uno, otro)
    return normalizar(uno) != normalizar(otro)
