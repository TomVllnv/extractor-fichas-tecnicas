"""Lectura conservadora de apartados explícitos de manuales de procedimiento.

El reconocimiento de documentos puede alterar encabezados y palabras. Se
buscan apartados explícitos y, cuando faltan, fragmentos documentales con
marcadores suficientemente específicos. No se completa mediante inferencias
clínicas ni se sustituyen decisiones del usuario.
"""
import re
import unicodedata
from difflib import SequenceMatcher


def _normalizar(texto):
    valor = unicodedata.normalize('NFKD', str(texto or '').casefold())
    valor = ''.join(c for c in valor if not unicodedata.combining(c))
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9 ]+', ' ', valor)).strip()


_SECCIONES = {
    'Recepción de muestras': (
        'recepcion de muestras', 'recepcion de las muestras',
        'recepcion de la muestra', 'recepcion de muestra',
    ),
    'Procesamiento de muestras': (
        'preparacion de las muestras', 'preparacion de la muestra',
        'preparacion de muestras', 'preparacion de muestra',
        'procesamiento de las muestras', 'procesamiento de muestras',
        'tratamiento de las muestras', 'pretratamiento de las muestras',
    ),
    'Método de procesamiento': (
        'fundamento de la tecnica', 'fundamento del metodo',
        'principio del metodo', 'principio de la tecnica',
        'principio del ensayo', 'metodologia analitica',
        'metodologia', 'metodo analitico', 'tecnica analitica', 'principio analitico',
        'fundamento', 'principio de la prueba', 'principio del test',
    ),
    'Test incluidos': ('test incluidos', 'pruebas incluidas', 'examenes incluidos'),
    'Preparación del paciente': ('preparacion del paciente', 'requisitos del paciente',
                                 'indicaciones al paciente'),
    'Procedimiento de toma de muestras': ('procedimiento de toma de muestras',
                                          'toma de muestra', 'toma de muestras'),
    'Muestra requerida y recipiente': ('muestra requerida y recipiente',
                                       'tipo de muestra', 'tipo de especimen', 'material biologico',
                                       'especimen requerido', 'muestra y recipiente',
                                       'muestras y recipientes'),
    'Rotulación': ('rotulacion', 'identificacion de muestras'),
    'Estabilidad de la muestra': ('estabilidad de la muestra', 'estabilidad',
                                 'estabilidad de muestras', 'conservacion de muestras',
                                 'conservacion de la muestra', 'almacenamiento de muestras'),
    'Condiciones de envío al Laboratorio': ('condiciones de envio al laboratorio',
                                            'transporte de muestras'),
    'Criterios de rechazo': ('criterios de rechazo', 'rechazo de muestras',
                             'criterios de aceptacion y rechazo'),
    'Valores de referencia': ('valores de referencia', 'intervalos de referencia',
                              'rango de referencia'),
    'Valor crítico': ('valor critico', 'valores criticos'),
    'Límite de detección': ('limite de deteccion', 'limites de deteccion',
                           'sensibilidad analitica'),
    'Equipamiento utilizado': ('equipamiento utilizado', 'equipos utilizados',
                                'equipo utilizado', 'equipamiento'),
    'Información clínica': ('informacion clinica', 'utilidad clinica',
                            'interpretacion clinica', 'significado clinico',
                            'aplicaciones clinicas', 'indicaciones clinicas',
                            'caracteristicas', 'caracteristicas del examen',
                            'caracteristicas de la prueba', 'descripcion del examen',
                            'descripcion de la prueba', 'descripcion clinica'),
    'Plazo entrega resultados': ('plazo entrega resultados', 'plazo de entrega de resultados'),
}
_TITULOS_NORMALIZADOS = {
    _normalizar(titulo): campo
    for campo, titulos in _SECCIONES.items()
    for titulo in titulos
}
# Reconocer 6.1 / 6,1 / 6.1, / 6-1, incluso cuando el OCR pierde el punto.
_INDICE = re.compile(r'^\s*\d{1,2}(?:\s*[.,-]\s*\d{1,2}){0,4}\s*[.):,;\-|]?\s*')
_NUMERO_SOLO = re.compile(r'^\s*\d{1,2}(?:\s*[.,-]\s*\d{1,2}){1,4}\s*[.):,;\-|]?\s*$')
_PIE = re.compile(
    r'^(?:p[aá]gina\s*:?\s*\d+(?:\s+de\s+\d+)?|'
    r'c[oó]digo(?:\s+caract\.)?\s*:|versi[oó]n\s*:|'
    r'fecha\s*:|vigencia\s*:|laboratorio\s+cl[ií]nico\s+crsco)', re.I)
_TITULOS_LIMITE = frozenset((
    'preparacion de reactivos', 'procedimiento', 'procedimiento analitico',
    'control de calidad', 'control de la tecnica', 'fluograma',
    'interpretacion de resultados', 'calibracion', 'reporte de resultados',
    'resultados', 'objetivo', 'alcance', 'responsabilidades',
    'materiales y reactivos', 'referencias bibliograficas',
    'anexos', 'anexo', 'registros', 'informe de resultados',
))
_OTROS_TITULOS = (
    'recepcion', 'preparacion', 'procesamiento', 'procedimiento', 'tecnica',
    'fundamento', 'principio', 'equipamiento', 'muestra', 'resultados',
    'control', 'desarrollo', 'objetivo', 'material', 'anexo', 'definiciones',
    'estabilidad', 'plazo', 'rotulacion', 'valores', 'limite', 'limites',
    'criterios', 'rechazo', 'informacion', 'toma', 'requisitos', 'condiciones',
    'flujograma', 'reporte', 'interpretacion', 'bibliografia', 'reactivos',
    'informe', 'calibracion', 'responsabilidades', 'alcance', 'referencias',
    'documentos', 'registros', 'insumos', 'interferencias',
    'introduccion', 'antecedentes', 'generalidades', 'finalidad',
    'descripcion', 'caracteristicas',
)


def _campo_de_titulo(titulo):
    """Acepta errores OCR pequeños SOLO en encabezados con aspecto de título."""
    texto = _normalizar(titulo)
    if not texto:
        return None
    exacto = _TITULOS_NORMALIZADOS.get(texto)
    if exacto:
        return exacto
    # Variante de un título de sección, no un párrafo libre. Las cadenas
    # cortas (p. ej. "fundamento") requieren coincidencia exacta.
    if len(texto) < 14 or len(texto) > 57:
        return None
    for patron, campo in _TITULOS_NORMALIZADOS.items():
        if len(patron) < 14 or abs(len(texto) - len(patron)) > 7:
            continue
        if SequenceMatcher(None, texto, patron).ratio() >= 0.86:
            return campo
    return None


def _es_encabezado(linea):
    """Reconoce un límite de sección numerada o un título OCR aislado.

    Devuelve el título normalizado, incluso si no es un campo extraíble,
    porque esos títulos son fundamentales para NO mezclar secciones.
    """
    valor = str(linea or '').strip().strip(':')
    if not valor or len(valor) > 120:
        return None
    m = _INDICE.match(valor)
    numerado = bool(m)
    nombre = (valor[m.end():] if m else valor).strip(' -.:,;|')
    if not nombre or len(nombre) > 95:
        return None
    texto = _normalizar(nombre)
    if len(texto.split()) > 11:
        return None
    # Un error OCR puede confundir I/l u O/0 dentro de títulos mayúsculos.
    # Exigir mayoría clara de mayúsculas evita confundir texto clínico normal.
    letras = [car for car in nombre if car.isalpha()]
    parece_mayusculas = bool(letras) and sum(car.isupper() for car in letras) / len(letras) >= 0.78
    if not numerado:
        # «Introducción» y «Descripción general» suelen aparecer en letras
        # normales como título aislado; no exigir mayúsculas en esos casos.
        titulos_de_contexto = {
            'introduccion', 'generalidades', 'antecedentes',
            'antecedentes generales', 'descripcion general', 'finalidad',
        }
        if (not parece_mayusculas and texto not in titulos_de_contexto
                and texto not in _TITULOS_LIMITE
                and not _campo_de_titulo(texto)):
            return None
        if not (_campo_de_titulo(texto) or texto in titulos_de_contexto
                or texto in _TITULOS_LIMITE
                or texto.split(' ')[0] in _OTROS_TITULOS):
            return None
    elif not (parece_mayusculas or texto.split(' ')[0] in _OTROS_TITULOS
              or _campo_de_titulo(texto)):
        return None
    return texto


def _reconstruir_lineas_ocr(texto):
    """Une números y títulos partidos; no concatena contenido clínico."""
    originales = str(texto or '').replace('\r', '').split('\n')
    reconstruidas = []
    i = 0
    while i < len(originales):
        linea = originales[i].strip()
        # OCR: "6.1." / "RECEPCION DE" / "LAS MUESTRAS".
        # OCR: "6.1. RECEPCION DE" / "LAS MUESTRAS".
        if _NUMERO_SOLO.fullmatch(linea) or (_INDICE.match(linea) and
                not _campo_de_titulo(_es_encabezado(linea))):
            mejor = None
            for cuantos in (1, 2, 3):
                siguientes = originales[i + 1:i + 1 + cuantos]
                if len(siguientes) != cuantos or any(not v.strip() for v in siguientes):
                    break
                if any(len(v.strip()) > 72 or _NUMERO_SOLO.fullmatch(v.strip())
                       for v in siguientes):
                    break
                candidato = linea + ' ' + ' '.join(v.strip() for v in siguientes)
                titulo = _es_encabezado(candidato)
                if titulo and _campo_de_titulo(titulo):
                    mejor = (candidato, cuantos)
                    break
            if mejor:
                reconstruidas.append(mejor[0])
                i += mejor[1] + 1
                continue
        reconstruidas.append(originales[i])
        i += 1
    return reconstruidas


def _limpiar_fragmento(lineas):
    """Une renglones de un mismo párrafo, sin destruir los párrafos o listas.

    No intenta adivinar separaciones dentro de palabras ya pegadas por OCR.
    El motor especializado es responsable de ese reconocimiento.
    """
    limpias = []
    for original in lineas:
        s = str(original or '').strip()
        if not s:
            if limpias and limpias[-1] != "":
                limpias.append("")
            continue
        if _PIE.match(s):
            continue
        if (len(s) < 115 and s.isupper() and
                ('PROCEDIMIENTO PARA' in s or 'LABORATORIO CLINICO' in s)):
            continue
        if re.match(r'^={2,}\s*P[AÁ]GINA', s, re.I):
            continue
        limpias.append(re.sub(r'[ \t]+', ' ', s))
    parrafos = []
    actual = ""
    for linea in limpias:
        if not linea:
            if actual:
                parrafos.append(actual)
                actual = ""
            continue
        es_vineta = bool(re.match(r'^(?:[•●▪◦∙]|[-–]\s|\d+[.)]\s)', linea))
        if es_vineta:
            if actual:
                parrafos.append(actual)
            actual = linea
            continue
        if actual and (actual.endswith(('.', ':', '!', '?')) and
                       len(actual) > 30):
            parrafos.append(actual)
            actual = linea
        elif actual:
            actual += " " + linea
        else:
            actual = linea
    if actual:
        parrafos.append(actual)
    return "\n\n".join(parrafos).strip()


def _titulo_y_texto_en_misma_linea(linea):
    """OCR puede juntar «6.1 RECEPCION ...» y la primera frase.

    Solo se separan encabezados NUMERADOS cuyos primeros términos coinciden
    con una sección configurada; no se buscan etiquetas dentro de prosa.
    """
    valor = str(linea or '').strip()
    m = _INDICE.match(valor)
    if not m:
        # En manuales digitales también se ve «Introducción: frase clínica».
        # Exigir etiqueta conocida y separador explícito, sin buscar títulos
        # dentro de una oración arbitraria.
        etiqueta = re.match(r'^([^:]{4,55}):\s*(.{18,})$', valor)
        if etiqueta:
            titulo = _normalizar(etiqueta.group(1))
            campo = _campo_de_titulo(titulo)
            if campo or titulo in ('introduccion', 'generalidades',
                                    'descripcion general', 'antecedentes',
                                    'finalidad'):
                return titulo, campo, etiqueta.group(2).strip()
        return None
    cuerpo = valor[m.end():].strip()
    palabras = cuerpo.split()
    if len(palabras) < 4:
        return None
    for cantidad in range(1, min(len(palabras) - 3, 7) + 1):
        titulo = _normalizar(' '.join(palabras[:cantidad]))
        campo = _campo_de_titulo(titulo)
        restante = ' '.join(palabras[cantidad:]).strip(' :-')
        if len(restante) >= 18 and (campo or titulo in _TITULOS_LIMITE):
            return titulo, campo, restante
    return None


# Los patrones textuales son EVIDENCIA literal, no clasificación clínica
# probabilística. No copiar «tiempo de reacción» como plazo de entrega ni
# «rango de medición» como intervalo de referencia.
_PATRONES_CONTEXTUALES = (
    ('Estabilidad de la muestra', re.compile(
        r'\b(?:estabilid(?:ad|ades)|(?:muestras?|plasma|suero)\s+(?:se\s+)?(?:conserva(?:n)?|mantiene(?:n)?)\s+(?:estable[s]?\s+)?(?:hasta|durante|por)\s+\d+)', re.I)),
    ('Valores de referencia', re.compile(r'\b(?:valores?|intervalos?|rangos?)\s+(?:de\s+)?referencia\b', re.I)),
    ('Valor crítico', re.compile(r'\b(?:valores?\s+criticos?|alerta\s+critica)\b', re.I)),
    ('Límite de detección', re.compile(r'\b(?:limites?\s+de\s+deteccion|\bLOD\b)', re.I)),
    ('Preparación del paciente', re.compile(r'\b(?:ayuno|requisitos?\s+(?:para\s+)?(?:el\s+)?paciente|anticoagulantes?\s+orales?)\b', re.I)),
    ('Procesamiento de muestras', re.compile(r'\b(?:centrifugacion|centrifugar|centrifugad[oa]s?|\d+\s*RPM|separacion\s+de\s+(?:plasma|suero))\b', re.I)),
    ('Muestra requerida y recipiente', re.compile(r'\b(?:tubos?\s+(?:celestes?|azules?|morados?|con\s+citrato)|citrato\s+de\s+sodio|plasma\s+(?:citratado|pobre)|sangre\s+total)\b', re.I)),
    ('Criterios de rechazo', re.compile(r'\b(?:se\s+rechaz|rechazo\s+de\s+muestras?|lipemia|hemolisis|coagulos?\s+visibles?|volumen\s+insuficiente)\b', re.I)),
    ('Recepción de muestras', re.compile(r'\b(?:recepcion\s+de\s+(?:las?\s+)?muestras?|registr[oa]\s+(?:en\s+)?(?:el\s+)?LIS|coteja(?:ra)?\s+(?:que\s+)?la\s+orden)\b', re.I)),
    ('Condiciones de envío al Laboratorio', re.compile(r'\b(?:transporte\s+(?:de\s+)?muestras?|envio\s+al\s+laboratorio)\b', re.I)),
    ('Rotulación', re.compile(r'\b(?:rotulacion|rotular\s+(?:el\s+)?tubo|etiquetado\s+de\s+muestras?)\b', re.I)),
    ('Método de procesamiento', re.compile(r'\b(?:metodo\s+(?:de\s+)?(?:procesamiento|deteccion)|\bDR?VVT\b|coagulometria|principio\s+del\s+(?:ensayo|metodo))\b', re.I)),
    ('Test incluidos', re.compile(r'\b(?:pruebas?\s+incluidas?|test\s+incluidos?)\b', re.I)),
    ('Información clínica', re.compile(r'\b(?:utilidad\s+clinica|indicaciones\s+clinicas|interpretacion\s+clinica)\b', re.I)),
    ('Plazo entrega resultados', re.compile(r'\b(?:plazo\s+(?:de\s+)?entrega\s+(?:de\s+)?resultados?\s*[:=-]\s*\d+|resultados?\s+(?:en|dentro\s+de)\s+\d+\s*(?:horas?|dias?))\b', re.I)),
)


def _segmentos_contextuales(contenido):
    """Separar frases preservando porcentajes, decimales y saltos OCR."""
    contenido = str(contenido or '').replace('\r', '')
    fragmentos = []
    for parrafo in re.split(r'\n\s*\n+', contenido):
        partes = re.split(r'(?<=[.!?])\s+(?=[A-ZÁÉÍÓÚÑ])', parrafo.strip())
        for parte in partes:
            parte = ' '.join(parte.split())
            if len(parte) >= 22 or re.match(r'^[•●▪◦∙]\s+', parte):
                fragmentos.append(parte)
    return fragmentos


def _clasificar_fragmento(fragmento):
    normalizado = _normalizar(fragmento)
    for campo, patron in _PATRONES_CONTEXTUALES:
        if patron.search(normalizado):
            return campo
    return None


def _repartir_contexto(texto, hallazgos, titulos):
    """Encabezado propio > frase en otro apartado > frase sin encabezado.

    Una frase sobre centrifugación no contamina «Recepción»; solo se propone
    para «Procesamiento» si falta su apartado dedicado. Se evita duplicación.
    """
    dedicados = {campo for _, campo in titulos
                if campo and (campo != 'Información clínica' or campo in hallazgos)}
    # Separar contenido ya clasificado por encabezado, quitando frases
    # claramente ajenas. Cuando el título es explícito conserva la prioridad.
    fuentes_indirectas = {}
    for campo, contenido in list(hallazgos.items()):
        if campo == 'Equipamiento utilizado':
            continue
        propios = []
        for fragmento in _segmentos_contextuales(contenido):
            if (campo == 'Método de procesamiento' and
                    _coincide_homologo('Información clínica', fragmento) and
                    not _coincide_homologo('Método de procesamiento', fragmento)):
                if 'Información clínica' not in dedicados:
                    fuentes_indirectas.setdefault('Información clínica', []).append(fragmento)
                continue
            if (campo == 'Información clínica' and
                    _coincide_homologo('Método de procesamiento', fragmento) and
                    not _coincide_homologo('Información clínica', fragmento)):
                if 'Método de procesamiento' not in dedicados:
                    fuentes_indirectas.setdefault('Método de procesamiento', []).append(fragmento)
                continue
            destino = _clasificar_fragmento(fragmento)
            # Respetar el apartado explícito cuando la frase describe su
            # acción principal aunque mencione incidentalmente otra técnica.
            if campo == 'Procesamiento de muestras' and re.search(
                    r'\b(?:se procesan|procesamiento de muestras|procesar las muestras)\b',
                    _normalizar(fragmento)):
                destino = campo
            if destino and destino != campo:
                if destino not in dedicados:
                    fuentes_indirectas.setdefault(destino, []).append(fragmento)
                continue
            propios.append(fragmento)
        if propios:
            hallazgos[campo] = '\n\n'.join(propios)[:3000]
        else:
            hallazgos.pop(campo, None)
    # Sin apartado específico: reconocer frases fuertemente evidenciadas.
    # Las frases dentro de secciones ajenas ya fueron separadas arriba.
    candidatos = {}
    for apartado, linea in _fragmentos_homologos(texto):
        if _seccion_excluida_homologos(apartado):
            continue
        if _es_encabezado(linea):
            continue
        destino = _clasificar_fragmento(linea)
        if destino and destino not in dedicados and destino not in hallazgos:
            candidatos.setdefault(destino, []).append(linea)
    for destino, fragmentos in fuentes_indirectas.items():
        if destino not in dedicados:
            candidatos.setdefault(destino, []).extend(fragmentos)
    for destino, fragmentos in candidatos.items():
        if destino in hallazgos:
            continue
        unicos = list(dict.fromkeys(fragmentos))
        valor = '\n\n'.join(unicos[:5])[:3000]
        if 25 <= len(valor) <= 3000:
            hallazgos[destino] = valor
    return hallazgos




# Complementos prudentes para los cinco campos priorizados. Las expresiones se
# aplican sobre texto normalizado, pero los resultados conservan las frases
# literales del documento. Una coincidencia NO autoriza resumir ni reinterpretar.
_CAMPOS_HOMOLOGOS = (
    'Información clínica', 'Muestra requerida y recipiente',
    'Procesamiento de muestras', 'Estabilidad de la muestra',
    'Método de procesamiento',
)

_EXCLUIR_SECCION_HOMOLOGA = frozenset({
    'preparacion de reactivos', 'materiales y reactivos', 'preparacion de calibradores',
    'control de calidad', 'calibracion', 'referencias bibliograficas',
    'bibliografia', 'bibliografia y referencias', 'anexo', 'anexos',
    'registros', 'mantenimiento', 'limpieza del equipo', 'bioseguridad',
    'interpretacion de resultados', 'valores de referencia',
    'limite de deteccion', 'limites de deteccion', 'valores criticos',
    'plazo entrega resultados', 'resultados', 'validacion de resultados',
})

_PATRONES_HOMOLOGOS = {
    'Información clínica': (
        # Finalidad de la prueba, no la técnica analítica.
        r'\b(?:utilidad|aplicacion(?:es)?|significado|indicacion(?:es)?|interpretacion) clinica[s]?\b',
        r'\b(?:se utiliza|se emplea|se usa|es util|sirve|se indica|esta indicada|'
        r'permite|ayuda|contribuye|es de utilidad)\b.{0,100}'
        r'\b(?:diagnostic|detect|evalu|seguimient|monitoriz|control|enfermedad|'
        r'patolog|inflamatori|infecci|tromb|riesgo|deficiencia)\w*',
        r'\b(?:determinacion|medicion|prueba|analisis|examen)\b.{0,85}'
        r'\b(?:diagnostic|seguimient|monitoriz|evaluacion|deteccion|apoyo clinico)\w*',
        r'\b(?:concentracion|niveles?|actividad)\b.{0,100}'
        r'\b(?:aumenta|disminuye|elevad[oa]s?|se relaciona|se asocia)\b'
        r'.{0,100}\b(?:enfermedad|inflamaci|inflamatori|infecci|patologia|trastorno)\w*',
        r'\b(?:biomarcador|marcador)\b.{0,75}\b(?:inflamaci|infecci|enfermedad|'
        r'patologia|diagnostic)\w*',
    ),
    'Muestra requerida y recipiente': (
        r'\b(?:tipo de muestra|tipo de especimen|muestra requerida|material biologico)\b',
        r'\b(?:se requiere|se utiliza|utilizar|recolectar|obtener|recoger|se necesita)\b'
        r'.{0,65}\b(?:muestra|plasma|suero|sangre|orina|liquido cefalorraquideo|hisopado)\b',
        r'\b(?:muestra|sangre|plasma|suero|orina)\b.{0,70}'
        r'\b(?:tubo(?:s)?|recipiente(?:s)?|contenedor(?:es)?|anticoagulante|citrato|edta|heparina)\b',
        r'\b(?:se utiliza|se requiere|se necesita)\b.{0,50}\b(?:citrato de sodio|edta|heparina)\b',
        r'\b(?:tubos?|recipientes?|contenedores?)\b.{0,65}'
        r'\b(?:citrato|edta|heparina|tapa (?:azul|celeste|roja|morada)|sangre|plasma|suero|orina)\b',
    ),
    'Procesamiento de muestras': (
        r'\b(?:muestras?|tubos?|plasma|suero|sangre)\b.{0,110}'
        r'\b(?:centrifug|separacion|separar|aliquot|homogeneiz|pretratamiento|desproteiniza)\w*',
        r'\b(?:centrifug|separacion|separar|aliquot|homogeneiz|pretratamiento)\w*'
        r'.{0,100}\b(?:muestras?|tubos?|plasma|suero|sangre)\b',
    ),
    'Estabilidad de la muestra': (
        r'\b(?:estabilidad|conservacion|almacenamiento)\b.{0,75}'
        r'\b(?:muestra|plasma|suero|sangre|orina|especimen)\w*',
        r'\b(?:muestra|plasma|suero|sangre|orina|especimen)\w*'
        r'.{0,110}\b(?:conserv|almacen|mantiene|estable|refriger|congel)\w*',
        r'\b(?:plasma|suero|sangre|orina|muestra)\b.{0,90}'
        r'\b(?:hasta|durante|por)\s+\d+\s*(?:h|horas?|dias?|meses?)\b',
    ),
    'Método de procesamiento': (
        r'\b(?:principio|fundamento|metodo|tecnica|metodologia)\b.{0,70}'
        r'\b(?:ensayo|analisis|determinacion|medicion|prueba|inmuno|coagul|'
        r'fotometr|enzimat|cromatograf|espectrofotometr|fluorescen|aglutinacion)\w*',
        r'\b(?:inmunoturbidimet|turbidimet|quimioluminiscen|electroquimioluminiscen|'
        r'coagulomet|nefelomet|cromatograf|espectrofotomet|fotometr|enzimat|'
        r'citometria|inmunoensayo|elisa|eclia|aglutinacion|colorimet|'
        r'potenciomet|electroforesis|pcr\b)\w*',
        r'\b(?:determinacion|medicion|ensayo|analisis)\b.{0,85}'
        r'\b(?:mediante|por)\s+(?:un\s+)?(?:metodo|tecnica|inmuno|quimio|coagulo|'
        r'fotometr|espectro|enzimat|cromato)\w*',
    ),
}
_PATRONES_HOMOLOGOS = {
    campo: tuple(re.compile(patron) for patron in patrones)
    for campo, patrones in _PATRONES_HOMOLOGOS.items()
}

_REACTIVOS_HOMOLOGOS = re.compile(
    r'\b(?:reactivos?|calibradores?|controles? de calidad|solucion(?:es)? de lavado)\b')
_TECNICOS_NO_CLINICOS = re.compile(
    r'\b(?:centrifug|incubacion|inmunoturbidimet|turbidimet|quimioluminiscen|'
    r'coagulomet|espectrofotomet|fotometr|pipetear|analizador|'
    r'preparacion de reactivos)\w*')


# «Características» suele ser una descripción del examen, pero también puede
# contener propiedades puramente analíticas. No tratar una cabecera genérica
# como autorización para copiar el apartado entero a Información clínica.
_TITULOS_DESCRIPCION_CLINICA = frozenset({
    'caracteristicas', 'caracteristicas del examen',
    'caracteristicas de la prueba', 'descripcion del examen',
    'descripcion de la prueba', 'descripcion clinica',
})
_TITULOS_CONTEXTO_CLINICO = _TITULOS_DESCRIPCION_CLINICA | frozenset({
    'introduccion', 'generalidades', 'antecedentes',
    'antecedentes generales', 'descripcion general', 'finalidad',
    'descripcion de la prueba', 'descripcion del examen',
})
_DESCRIPCION_CLINICA = (
    re.compile(r'\b(?:examen|prueba|test|analisis|determinacion|estudio)\b'
               r'.{0,105}\b(?:mide|cuantifica|identifica|detecta|evalua|refleja|'
               r'determina|investiga|examina)\w*\b'),
    re.compile(r'\b(?:proteina|hormona|anticuerpo|enzima|metabolito|biomarcador|'
               r'marcador|lipoproteina|antigeno|fibrina|fibrinogeno|glucosa|'
               r'hemoglobina|hemoglobinas|acido urico|dimeros? d|'
               r'producto de degradacion)\b.{0,125}'
               r'\b(?:es|son|indica|indican|refleja|reflejan|representa|'
               r'se relaciona|se asocia|participa|permiten|permite|aumenta|disminuye)\b'),
    re.compile(r'\b(?:es|son|representa|corresponde a|se considera)\b.{0,135}'
               r'\b(?:proteina|hormona|anticuerpo|enzima|metabolito|biomarcador|'
               r'marcador|lipoproteina|antigeno|producto de degradacion|'
               r'metabolismo|fibrina|coagulacion|respuesta inmunitaria)\b'),
    re.compile(r'\b(?:el|la|los|las)\s+[a-z0-9 ]{3,80}\b(?:es|son|refleja|indica|'
               r'participa|se correlaciona)\b.{0,130}'
               r'\b(?:reserva|funcion|metabolismo|proceso|estado|organismo|'
               r'enfermedad|riesgo|patologia|concentracion|anticuerpos?|' 
               r'inflamacion|respuesta|produccion|degradacion|actividad|presencia)\w*'),
    re.compile(r'\b(?:permite|ayuda|contribuye|se relaciona|se asocia|'
               r'es indicativo|se utiliza|se emplea)\b.{0,125}'
               r'\b(?:funcion|organismo|paciente|salud|trastorno|alteracion|'
               r'condicion|riesgo|diagnostic|seguimient|enfermedad|'
               r'evaluacion|inflamacion|infeccion|patologia)\w*\b'),
)


def _frase_describe_examen(frase):
    """Admite explicaciones biológicas explícitas sin inventar utilidad clínica."""
    normal = _normalizar(frase)
    # No confundir características del rendimiento o instrucciones de trabajo
    # con la descripción del examen, aunque mencionen un analito.
    if _REACTIVOS_HOMOLOGOS.search(normal) or _TECNICOS_NO_CLINICOS.search(normal):
        return False
    if re.search(r'\b(?:coeficiente de variacion|limite de deteccion|'
                 r'linealidad|calibracion|precision analitica|'
                 r'intervalo de medicion|tubo|centrifugacion|'
                 r'volumen de muestra)\b', normal):
        return False
    return any(p.search(normal) for p in _DESCRIPCION_CLINICA)


# Permite conservar más de una frase clínica del mismo apartado, sin
# transformar una introducción completa en información clínica ni reformularla.
_CLINICA_CONTINUA = re.compile(
    r'\b(?:se caracteriza|se asocia|se vincula|se relaciona|'
    r'complicaciones?|trombosis|trombotic|enfermedades?|autoinmun|'
    r'anticuerpos?|diagnostic|evaluacion|seguimiento|embarazo|gestacion|'
    r'abortos?|preeclampsia|riesgo|manifestaciones?|sintomas?|'
    r'patologia|fetal|pronostico|lupus|inmunitari|respuesta inflamatori)\w*',
    re.I)


def _frase_complementaria_clinica(frase):
    normal = _normalizar(frase)
    return (len(normal) >= 28 and len(normal) <= 500
            and not _REACTIVOS_HOMOLOGOS.search(normal)
            and not _TECNICOS_NO_CLINICOS.search(normal)
            and not re.search(r'\b(?:pipetear|centrifugar|control de calidad|'
                              r'coeficiente de variacion|preparacion de muestras|'
                              r'calibracion|volumen minimo)\b', normal)
            and bool(_CLINICA_CONTINUA.search(normal)))


def _seleccionar_contexto_clinico(frases, *, max_frases=4, limite=1250):
    """Frases originales, en orden, con continuidad de tema y máximo legible."""
    elegidas = []
    contexto_iniciado = False
    for frase in frases:
        relevante = _coincide_homologo('Información clínica', frase, descripcion=True)
        complemento = contexto_iniciado and _frase_complementaria_clinica(frase)
        if relevante or complemento:
            if len(elegidas) >= max_frases or sum(len(x) for x in elegidas) + len(frase) > limite:
                break
            if frase not in elegidas:
                elegidas.append(frase)
            contexto_iniciado = True
        elif contexto_iniciado:
            # No unir frases distantes tras instrucciones analíticas ajenas.
            break
    return ('\n\n'.join(elegidas) if any(frase.lstrip().startswith(('•', '●', '-', '–')) for frase in elegidas)
            else ' '.join(elegidas))


def _filtrar_caracteristicas_clinicas(valor):
    """Conserva frases literales sobre el examen, omite pasos técnicos."""
    partes = _segmentos_contextuales(valor)
    return _seleccionar_contexto_clinico(partes)


def _prioridad_cabecera_clinica(titulo):
    if titulo == 'informacion clinica':
        return 0
    if titulo in _TITULOS_DESCRIPCION_CLINICA:
        return 2
    return 1


def _fragmentos_homologos(texto):
    """Fragmentos con su apartado real; nunca continuar tras otro encabezado."""
    apartado = ''
    lineas = []
    grupos = []

    def cerrar():
        if lineas:
            contenido = _limpiar_fragmento(lineas)
            if contenido:
                for fragmento in _segmentos_contextuales(contenido):
                    grupos.append((apartado, fragmento))
        lineas.clear()

    for linea in _reconstruir_lineas_ocr(texto):
        titulo = _es_encabezado(linea)
        resto = None
        if titulo is None:
            partido = _titulo_y_texto_en_misma_linea(linea)
            if partido is not None:
                titulo, _, resto = partido
        if titulo is not None:
            cerrar()
            apartado = titulo
            if resto:
                lineas.append(resto)
        else:
            lineas.append(linea)
    cerrar()
    return grupos


def _seccion_excluida_homologos(titulo):
    return (titulo in _EXCLUIR_SECCION_HOMOLOGA
            or any(titulo.startswith(x + ' ') for x in _EXCLUIR_SECCION_HOMOLOGA))


def _coincide_homologo(campo, frase, *, descripcion=False):
    normal = _normalizar(frase)
    if len(normal) < 24 or len(normal) > 1000:
        return False
    if campo == 'Información clínica':
        if _TECNICOS_NO_CLINICOS.search(normal) and not re.search(
                r'\b(?:diagnostic|evalu|seguimient|monitoriz|inflamaci|infecci|'
                r'enfermedad|patolog|riesgo|tromb)\w*', normal):
            return False
    if campo in ('Muestra requerida y recipiente', 'Procesamiento de muestras',
                 'Estabilidad de la muestra'):
        # No confundir almacenamiento/preparación de reactivos con muestras.
        if _REACTIVOS_HOMOLOGOS.search(normal) and not re.search(
                r'\b(?:muestras?|plasma|suero|sangre|orina|tubo(?:s)?)\b', normal):
            return False
    if campo == 'Estabilidad de la muestra':
        # La estabilidad necesita condiciones o duración, no solo el título.
        if not re.search(
                r'\b(?:\d+\s*(?:h|horas?|dias?|meses?)|'
                r'\d+\s*(?:a|y|-)\s*\d+\s*(?:c|grados)|'
                r'\d+\s*(?:°\s*)?c\b|temperatura|refriger|congel|ambiente|'
                r'estable\s+durante)\w*', normal):
            return False
    if campo == 'Muestra requerida y recipiente':
        if re.search(r'\b(?:centrifug|separa|centrifugacion)\w*', normal) and not re.search(
                r'\b(?:tubo|recipiente|se requiere|utilizar|se necesita|'
                r'tipo de muestra|tipo de especimen|anticoagulante|citrato|edta)\b', normal):
            return False
    if any(p.search(normal) for p in _PATRONES_HOMOLOGOS[campo]):
        return True
    return bool(campo == 'Información clínica' and descripcion
                and _frase_describe_examen(frase))


def _completar_homologos_focalizados(texto, hallazgos, titulos):
    """Apartado específico > homólogo > frase contextual explícita.

    Solo se añaden campos ausentes. El fragmento procede literalmente de un
    apartado del archivo y nunca se completa con conocimiento médico externo.
    """
    dedicados = {campo for _, campo in titulos
                if campo in _CAMPOS_HOMOLOGOS
                and (campo != 'Información clínica' or campo in hallazgos)}
    fuentes = {}
    candidatos = {campo: [] for campo in _CAMPOS_HOMOLOGOS}
    for titulo, fragmento in _fragmentos_homologos(texto):
        if _seccion_excluida_homologos(titulo):
            continue
        encabezado = _campo_de_titulo(titulo)
        for campo in _CAMPOS_HOMOLOGOS:
            if campo in dedicados or campo in hallazgos:
                continue
            if not _coincide_homologo(
                    campo, fragmento,
                    descripcion=(campo == 'Información clínica'
                                 and titulo in _TITULOS_CONTEXTO_CLINICO)):
                continue
            # Una descripción de calibración no es el método del examen.
            if (campo == 'Método de procesamiento'
                    and titulo in _TITULOS_DESCRIPCION_CLINICA
                    and re.search(r'\b(?:calibr|mantenimient|control de calidad)\w*',
                                  _normalizar(fragmento))):
                continue
            # Cuando el texto esté dentro de un apartado explícito AJENO,
            # exigir que la frase tenga marcadores inequívocos del destino.
            # Evitar el ruido técnico de objetivos, reactivos y resultados.
            if campo == 'Información clínica' and titulo and (encabezado and encabezado != campo):
                if not re.search(r'\b(?:diagnostic|evalu|seguimient|monitoriz|'
                                 r'inflamaci|infecci|enfermedad|patolog|riesgo|tromb)\w*',
                                 _normalizar(fragmento)):
                    continue
            prioridad = 0 if (not encabezado and titulo in {
                'introduccion', 'generalidades', 'antecedentes', 'antecedentes generales',
                'finalidad', 'descripcion de la prueba', 'descripcion general'}) else 1
            if campo == 'Método de procesamiento' and titulo in {
                    'introduccion', 'generalidades', 'antecedentes', 'finalidad'}:
                prioridad = 1
            candidatos[campo].append((prioridad, titulo, fragmento))

    for campo, opciones in candidatos.items():
        if not opciones or campo in hallazgos or campo in dedicados:
            continue
        menor = min(opcion[0] for opcion in opciones)
        # Escoger solo el mejor nivel de fuente; no fusionar material de
        # distintas secciones si es inconsistente o repetido.
        mejores = [item for item in opciones if item[0] == menor]
        primer_titulo = mejores[0][1]
        textos = list(dict.fromkeys(frase for _, titulo, frase in mejores
                                   if titulo == primer_titulo))[:4]
        if campo == 'Información clínica':
            # Una frase que justifica la utilidad clínica habilita las frases
            # complementarias contiguas del mismo apartado, no de otro.
            del_mismo_apartado = [frase for titulo, frase in
                                   _fragmentos_homologos(texto) if titulo == primer_titulo]
            contenido = (_seleccionar_contexto_clinico(del_mismo_apartado)
                         or '\n\n'.join(textos)[:1250])
        else:
            contenido = '\n\n'.join(textos)[:3000]
        if len(contenido) >= 25:
            hallazgos[campo] = contenido
            nombres_fuente = {
                'introduccion': 'Introducción',
                'descripcion general': 'Descripción general',
                'fundamento': 'Fundamento',
                'generalidades': 'Generalidades',
                'antecedentes': 'Antecedentes',
                'caracteristicas': 'Características',
            }
            fuentes[campo] = (nombres_fuente.get(primer_titulo,
                                                  primer_titulo.capitalize()) if primer_titulo else
                             'Texto contextual del documento')
    return hallazgos, fuentes

def analizar_manual(texto):
    """Retorna propuestas y diagnóstico verificable de los encabezados OCR.

    Se descartan silenciosamente NO: el diagnóstico registra los títulos
    detectados pero demasiado cortos/largos, para explicar un resultado vacío.
    """
    lineas = _reconstruir_lineas_ocr(texto)
    hallazgos = {}
    titulos = []
    omitidos = []
    activo = None
    contenido = []
    titulo_activo = ''
    prioridad_clinica = 99
    fuente_clinica = ''

    def cerrar():
        nonlocal activo, contenido, prioridad_clinica, fuente_clinica
        if activo is not None:
            valor = _limpiar_fragmento(contenido)
            if activo == 'Información clínica' and titulo_activo in _TITULOS_DESCRIPCION_CLINICA:
                valor = _filtrar_caracteristicas_clinicas(valor)
            if 25 <= len(valor) <= 3000:
                if activo == 'Información clínica':
                    prioridad = _prioridad_cabecera_clinica(titulo_activo)
                    if prioridad < prioridad_clinica:
                        hallazgos[activo] = valor
                        prioridad_clinica = prioridad
                        fuente_clinica = titulo_activo
                elif activo not in hallazgos:
                    hallazgos[activo] = valor
            else:
                omitidos.append((activo, len(valor)))
        activo, contenido = None, []

    for original in lineas:
        titulo = _es_encabezado(original)
        primera_frase = None
        campo_inline = None
        if titulo is None:
            separada = _titulo_y_texto_en_misma_linea(original)
            if separada is not None:
                titulo, campo_inline, primera_frase = separada
        if titulo is not None:
            cerrar()
            campo = campo_inline or _campo_de_titulo(titulo)
            titulos.append((titulo, campo))
            if campo is not None:
                activo = campo
                titulo_activo = titulo
                if primera_frase:
                    contenido.append(primera_frase)
            continue
        if activo is not None:
            contenido.append(original)
    cerrar()

    modelos = []
    patron_modelo = re.compile(
        r'\b(?:cobas\s+[tce]\s*[-]?\s*\d{3,4}|'
        r'sysmex\s+xn\s*[-]?\s*\d{3,4}|'
        r'roller\s*[-]?\s*\d{2}[a-z]{0,3})\b', re.I)
    for coincidencia in patron_modelo.finditer(str(texto or '')):
        modelo = re.sub(r'\s+', ' ', coincidencia.group(0)).strip()
        if modelo.casefold() not in {v.casefold() for v in modelos}:
            modelos.append(modelo)
    if modelos and 'Equipamiento utilizado' not in hallazgos:
        hallazgos['Equipamiento utilizado'] = '\n'.join(modelos[:3])
    hallazgos = _repartir_contexto(texto, hallazgos, titulos)
    # Los cinco campos priorizados requieren evidencia reforzada en frases
    # sin encabezado. Otros campos mantienen la lógica previa del extractor.
    for campo in _CAMPOS_HOMOLOGOS:
        if campo not in {c for _, c in titulos if c}:
            hallazgos.pop(campo, None)
    hallazgos, fuentes_homologas = _completar_homologos_focalizados(
        texto, hallazgos, titulos)
    if (fuente_clinica in _TITULOS_DESCRIPCION_CLINICA
            and 'Información clínica' in hallazgos):
        fuentes_homologas.setdefault('Información clínica',
                                    {'caracteristicas': 'Características',
                                     'descripcion clinica': 'Descripción clínica',
                                     'descripcion de la prueba': 'Descripción de la prueba',
                                     'descripcion del examen': 'Descripción del examen'
                                     }.get(fuente_clinica, fuente_clinica.capitalize()))
    # Los manuales pueden citar bibliografía, pero no están autorizados a
    # completar «Referencias» sin selección explícita del usuario.
    hallazgos.pop('Referencias', None)
    return {
        'campos': hallazgos,
        'titulos': titulos,
        'omitidos': omitidos,
        'caracteres_ocr': len(str(texto or '').strip()),
        'fuentes_homologas': fuentes_homologas,
    }


def extraer_campos_manual(texto):
    """Propone únicamente valores explícitos para cotejo humano."""
    return analizar_manual(texto)['campos']


def filtrar_celdas_ocupadas(candidatos, valores_existentes):
    """Nunca proponer la sustitución de un campo con valor ya validado."""
    ocupadas = {_normalizar(k) for k, v in (valores_existentes or {}).items()
                if v is not None and str(v).strip()}
    return {k: v for k, v in (candidatos or {}).items()
            if _normalizar(k) not in ocupadas and str(v or '').strip()}


def es_manual_de_procedimiento_textual(texto):
    """Detecta manuales digitales inequívocos; no clasifica insertos de fabricante.

    La modalidad de documentos escaneados conserva su comportamiento: esta
    comprobación solo limita cuándo se añaden homólogos en el flujo digital.
    """
    cabecera = str(texto or '')[:3500]
    normal = _normalizar(cabecera)
    if len(normal) < 70:
        return False
    # Exigir que se presente como título. Una mención a otro manual dentro
    # del texto de una ficha del fabricante no basta para reclasificarla.
    for linea in cabecera.splitlines()[:25]:
        encabezado = _normalizar(linea)
        if len(encabezado) > 150:
            continue
        if re.match(r'^(?:manual de procedimientos?|procedimiento operativo estandar|'
                    r'procedimiento normalizado de trabajo|procedimiento para (?:la |el )?'
                    r'(?:deteccion|determinacion|medicion|procesamiento))\b', encabezado):
            return True
    titulos = [encabezado for linea in _reconstruir_lineas_ocr(cabecera)
               if (encabezado := _es_encabezado(linea))]
    familias = {_campo_de_titulo(titulo) for titulo in titulos}
    return ('desarrollo' in titulos and len(titulos) >= 3
            and 'Recepción de muestras' in familias
            and 'Procesamiento de muestras' in familias)


def completar_homologos_manual_digital(texto, datos):
    """Rellena únicamente los cinco campos ausentes de un MANUAL digital.

    Las etiquetas explícitas extraídas por la ruta estándar ganan siempre.
    El reconocimiento de un manual no cambia los nombres ni contenidos de la
    base de datos y no habilita el resto de campos documentales controlados.
    """
    resultado = dict(datos or {})
    if not es_manual_de_procedimiento_textual(texto):
        return resultado
    propuestas = extraer_campos_manual(texto)
    for campo in _CAMPOS_HOMOLOGOS:
        if not str(resultado.get(campo, '') or '').strip() and propuestas.get(campo):
            resultado[campo] = propuestas[campo]
    return resultado
