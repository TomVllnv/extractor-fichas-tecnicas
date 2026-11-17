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
        'critico':'Valor crítico', 'equipamiento':'Equipamiento utilizado',
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
    return resultados


def son_valores_distintos(campo, uno, otro):
    if not _valor_valido(uno) or not _valor_valido(otro):
        return False
    if campo == 'Equipamiento utilizado':
        import equipamientos as eq
        return not eq.equipamientos_equivalentes(uno, otro)
    return normalizar(uno) != normalizar(otro)
