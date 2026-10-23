import io, os, re, tempfile, subprocess, shutil, csv, json, hashlib, copy, xml.etree.ElementTree as ET
from datetime import datetime
from zoneinfo import ZoneInfo
import openpyxl
import fitz  # PyMuPDF
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

APP_TITLE = 'Herramienta de Extracción de Datos de Fichas Técnicas'

# --- GESTIÓN DE LABORATORIOS ---
LABORATORIOS_PREDETERMINADOS = [
    "Redlab",
    "Barnafi Krause",
    "Complejo Hospitalario Dr. Sótero del Río",
    "Dr. Luis Calvo Mackenna",
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
    "Hematología": {"template": "plantilla_hematologia.xlsx"},
    "Bioquímica": {"template": "plantilla_bioquimica.xlsx"},
    "Hormonas e inmunidad": {"template": None},
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
    resultado = {area: list(guardados.get(area, [])) for area in AREA_CONFIG} if isinstance(guardados, dict) else {a: []
                                                                                                                   for a
                                                                                                                   in
                                                                                                                   AREA_CONFIG}
    for area, config in AREA_CONFIG.items():
        ruta = os.path.join(os.path.dirname(__file__), config["template"] or "")
        if config["template"] and os.path.exists(ruta) and not (isinstance(guardados, dict) and area in guardados):
            try:
                wb = load_workbook(ruta, read_only=True, data_only=True)
                resultado[area] = list(wb.sheetnames)
                wb.close()
            except Exception:
                resultado[area] = []
    return {a: sorted(list(dict.fromkeys(str(x).strip() for x in xs if str(x).strip())), key=str.casefold) for a, xs in
            resultado.items()}


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
    """Elimina citas bibliográficas de superíndices (ej: 'hígado.1', 'embarazo3,4')

    sin afectar cifras decimales (3.2 %, 0.109 M) ni cantidades operativas.
    """
    if not texto:
        return ""

    s = str(texto)
    # 1. Citas explícitas entre corchetes o paréntesis tipo [1], [1,2], (15)
    s = re.sub(r"\[\d+(?:\s*,\s*\d+)*\]", "", s)
    s = re.sub(r"\(\d+(?:\s*,\s*\d+)*\)", "", s)

    # 2. Citas pegadas directamente a una letra (ej: "embarazo3,4", "fase aguda2")
    s = re.sub(r"(?<=[a-zA-ZáéíóúñÁÉÍÓÚÑ])\d+(?:\s*,\s*\d+)*(?=\s|[.,;:!?\n]|$)", "", s)

    # 3. Citas pegadas a puntuación que está PRECEDIDA POR UNA LETRA (ej: "hígado.1", "trombóticas.11,12")
    s = re.sub(r"(?<=[a-zA-ZáéíóúñÁÉÍÓÚÑ][.,;:])\d+(?:\s*,\s*\d+)*(?=\s|[.,;:!?\n]|$)", "", s)

    # 4. Limpieza de espacios dobles y comas descolgadas
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
    """Quita espacios accidentales, líneas de ruido, citas bibliográficas y conserva saltos útiles."""
    if valor is None:
        return ""

    valor_sin_citas = quitar_citas_superindices(valor)
    lineas = [re.sub(r"[ \t]+", " ", linea).strip() for linea in valor_sin_citas.replace("\r", "").split("\n")]
    salida = []
    blancos = 0
    for linea in lineas:
        if not linea or es_linea_basura_o_encabezado(linea, examen_activo):
            blancos += 1
            if blancos == 1 and salida:
                salida.append("")
            continue
        blancos = 0
        if re.fullmatch(r"https?://\S+|www\.\S+", linea, flags=re.I):
            continue
        salida.append(linea)
    return convertir_subindices("\n".join(salida).strip())


# Estructura base de respaldo
FIELD_ALIASES_BASE = {
    'Código': ['codigo examen', 'codigo', 'código examen', 'código', 'Código del Examen'],
    'Nombre LIS': ['Nombre LIS', 'Derivación', 'nombre del examen', 'examen', 'nombre examen', 'Nombres del Examen'],
    'Laboratorio de procesamiento': ['laboratorio de procesamiento', 'laboratorio'],
    'Recepción de muestras': ['recepcion de muestras', 'recepción de muestras', 'recepcion'],
    'Procesamiento de muestras': ['procesamiento de muestras', 'procesamiento de muetsras', 'procesamiento'],
    'Plazo entrega resultados': ['tiempo de respuesta', 'plazo entrega resultados', 'plazo de entrega', 'tat',
                                 'tiempo procesamiento'],
    'Preparación del paciente': ['preparacion paciente', 'preparación paciente', 'preparacion del paciente',
                                 'preparación del paciente'],
    'Procedimiento de toma de muestras': ['Procedimiento', 'Procedimiento de toma de muestras',
                                          'Procedimiento de examen', 'Procedimiento del examen'],
    'Muestra requerida y recipiente': ['tipo de muestra', 'muestra requerida', 'muestra', 'tipo y volumen de muestra',
                                       'material biologico', 'material biológico', 'volumen requerido',
                                       'obtencion y preparacion de las muestras',
                                       'obtención y preparación de las muestras'],
    'Rotulación': ['Rotulación', 'Rotulacion'],
    'Estabilidad de la muestra': ['estabilidad muestra', 'estabilidad de la muestra', 'Conservación y estabilidad',
                                  'conservacion y estabilidad'],
    'Condiciones de envío al Laboratorio': ['Condiciones de envío al Laboratorio', 'temperatura de transporte',
                                            'condiciones de envio', 'condiciones de envío', 'envio al laboratorio',
                                            'envío al laboratorio', 'transporte'],
    'Criterios de rechazo': ['Rechazo', 'Criterios de rechazo', 'Criterios rechazo', 'Criterios del rechazo'],
    'Método de procesamiento': ['metodo utilizado', 'método utilizado', 'metodo de procesamiento',
                                'método de procesamiento', 'metodologia', 'metodología', 'metodo', 'principio del test',
                                'principio', 'principio de la prueba', 'principio del ensayo'],
    'Test incluidos': ['test incluidos', 'pruebas incluidas', 'examenes incluidos', 'exámenes incluidos'],
    'Valores de referencia': ['valor de referencia', 'valores de referencia', 'valores normales',
                              'rango de referencia'],
    'Valor crítico': ['valor critico', 'valor crítico', 'valores criticos', 'valores críticos'],
    'Límite de detección': ['limite de deteccion', 'límite de detección', 'limite deteccion', 'límite deteccion',
                            'límite de cuantificación', 'limite de cuantificacion', 'loq'],
    'Equipamiento utilizado': ['equipamiento utilizado', 'equipo utilizado', 'instrumentacion', 'instrumentación',
                               'equipo', 'system', 'sistema', 'informacion del sistema', 'información del sistema'],
    'Información clínica': ['caracteristicas', 'características', 'utilidad clinica', 'utilidad clínica',
                            'informacion clinica', 'información clínica'],
    'Referencias': ['referencias', 'bibliografia', 'bibliografía', 'fuentes'],
    'Actualizado por': ['actualizado por'],
    'Fecha actualización': ['fecha actualizacion', 'fecha actualización', 'fecha de actualización'],
    'Autorizado por': ['autorizado por'],
}

FIELD_ALIASES_POR_AREA = {
    "Microbiología": {
        'Código': ['codigo examen', 'codigo', 'código examen', 'código', 'cod_micro', 'código microbiología'],
        'Nombre LIS': ['Nombre LIS', 'Derivación', 'nombre del examen', 'examen', 'cultivo', 'agente',
                       'microorganismo'],
        'Laboratorio de procesamiento': ['laboratorio de procesamiento', 'laboratorio', 'sección microbiología',
                                         'lab micro'],
        'Recepción de muestras': ['recepcion de muestras', 'recepción de muestras', 'recepcion', 'recepción cultivo'],
        'Procesamiento de muestras': ['procesamiento de muestras', 'procesamiento', 'siembra', 'incubación'],
        'Plazo entrega resultados': ['tiempo de respuesta', 'plazo entrega resultados', 'plazo de entrega', 'tat',
                                     'días de incubación'],
        'Preparación del paciente': ['preparacion paciente', 'preparación paciente', 'preparacion del paciente',
                                     'sin antibióticos previa toma'],
        'Procedimiento de toma de muestras': ['Procedimiento', 'Procedimiento de toma de muestras', 'toma de muestra',
                                              'técnica aséptica', 'hisopado'],
        'Muestra requerida y recipiente': ['tipo de muestra', 'muestra requerida', 'muestra', 'medio de transporte',
                                           'stuart', 'frasco estéril', 'obtencion y preparacion de las muestras',
                                           'obtención y preparación de las muestras'],
        'Rotulación': ['Rotulación', 'Rotulacion', 'etiquetado muestra'],
        'Estabilidad de la muestra': ['estabilidad muestra', 'estabilidad de la muestra', 'conservación cultivo',
                                      'conservacion y estabilidad', 'conservación y estabilidad'],
        'Condiciones de envío al Laboratorio': ['Condiciones de envío al Laboratorio', 'temperatura de transporte',
                                                'envio al laboratorio', 'transporte ambiente'],
        'Criterios de rechazo': ['Rechazo', 'Criterios de rechazo', 'muestra contaminada', 'frasco derramado'],
        'Método de procesamiento': ['metodo utilizado', 'método utilizado', 'metodología', 'método',
                                    'cultivo tradicional', 'maldi-tof', 'antibiograma', 'principio del test',
                                    'principio', 'principio de la prueba'],
        'Test incluidos': ['test incluidos', 'pruebas incluidas', 'tinción gram', 'identificación y antibiograma'],
        'Valores de referencia': ['valor de referencia', 'valores de referencia', 'desarrollo bacteriano',
                                  'sin desarrollo'],
        'Valor crítico': ['valor critico', 'valor crítico', 'hemocultivo positivo', 'patógeno crítico'],
        'Límite de detección': ['limite de deteccion', 'límite de detección', 'sensibilidad analítica',
                                'límite de cuantificación', 'limite de cuantificacion', 'loq'],
        'Equipamiento utilizado': ['equipamiento utilizado', 'equipo utilizado', 'incubadora', 'estufa',
                                   'sistema automatizado micro', 'system', 'sistema', 'informacion del sistema',
                                   'información del sistema'],
        'Información clínica': ['caracteristicas', 'características', 'utilidad clinica', 'utilidad clínica',
                                'informacion clinica', 'información clínica'],
        'Referencias': ['referencias', 'bibliografia', 'bibliografía'],
        'Actualizado por': ['actualizado por'],
        'Fecha actualización': ['fecha actualizacion', 'fecha actualización'],
        'Autorizado por': ['autorizado por'],
    },
    "Hematología": {
        'Código': ['codigo examen', 'codigo', 'código examen', 'código', 'cod_hemato', 'código hematología'],
        'Nombre LIS': ['Nombre LIS', 'Derivación', 'nombre del examen', 'examen', 'hemograma', 'recuento', 'frotis'],
        'Laboratorio de procesamiento': ['laboratorio de procesamiento', 'laboratorio', 'sección hematología',
                                         'lab hemato'],
        'Recepción de muestras': ['recepcion de muestras', 'recepción de muestras', 'recepcion', 'recepción tubo edta'],
        'Procesamiento de muestras': ['procesamiento de muestras', 'procesamiento', 'homogenización',
                                      'frotis sanguíneo'],
        'Plazo entrega resultados': ['tiempo de respuesta', 'plazo entrega resultados', 'plazo de entrega', 'tat',
                                     'tiempo procesamiento'],
        'Preparación del paciente': ['preparacion paciente', 'preparación paciente', 'preparacion del paciente',
                                     'ayuno previo'],
        'Procedimiento de toma de muestras': ['Obtención y preparación de las muestras', 'Specimen collection and preparation', 'Procedimiento', 'Procedimiento de toma de muestras', 'punción venosa'],
        'Muestra requerida y recipiente': ['Specimen collection and preparation', 'tipo de muestra', 'muestra requerida', 'tubo edta', 'tubo tapa morada',
                                           'sangre total', 'obtencion y preparacion de las muestras',
                                           'obtención y preparación de las muestras'],
        'Rotulación': ['Rotulación', 'Rotulacion'],
        'Estabilidad de la muestra': ['Estabilidad:', 'Stability:', 'estabilidad muestra', 'estabilidad de la muestra', 'estabilidad frotis',
                                      'conservacion y estabilidad', 'conservación y estabilidad'],
        'Condiciones de envío al Laboratorio': ['Condiciones de envío al Laboratorio', 'temperatura de transporte',
                                                'envio ambiente', 'sin congelar'],
        'Criterios de rechazo': ['Rechazo', 'Criterios de rechazo', 'muestra coagulada', 'coágulos', 'hemólisis'],
        'Método de procesamiento': ['Test principle', 'Assay principle', 'metodo utilizado', 'método utilizado', 'metodología', 'citometría de flujo',
                                    'impedancia', 'microscopía', 'principio del test', 'principio',
                                    'principio de la prueba'],
        'Test incluidos': ['test incluidos', 'pruebas incluidas', 'fórmula leucocitaria', 'recuento de plaquetas',
                           'índices eritrocitarios'],
        'Valores de referencia': ['Valores teóricos', 'Expected values', 'Reference intervals', 'Valores treóricos', 'Expected values', ' Reference intervals', 'valor de referencia', 'valores de referencia', 'valores normales',
                                  'rangos hematológicos'],
        'Valor crítico': ['valor critico', 'valor crítico', 'plaquetopenia severa', 'blastos en frotis'],
        'Límite de detección': ['Limites e intervalos', 'Measuring range', 'limite de deteccion', 'límite de detección', 'límite de cuantificación',
                                'limite de cuantificacion', 'loq'],
        'Equipamiento utilizado': ['equipamiento utilizado', 'equipo utilizado', 'analizador hematológico',
                                   'contador celular', 'system', 'sistema', 'informacion del sistema',
                                   'información del sistema'],
        'Información clínica': ['Uso previsto', 'summary', 'caracteristicas', 'características', 'utilidad clinica', 'utilidad clínica',
                                'informacion clinica', 'información clínica'],
        'Referencias': ['referencias', 'bibliografia', 'bibliografía'],
        'Actualizado por': ['actualizado por'],
        'Fecha actualización': ['fecha actualizacion', 'fecha actualización'],
        'Autorizado por': ['autorizado por'],
    },
    "Bioquímica": {
        'Código': ['codigo examen', 'codigo', 'código examen', 'código', 'cod_bioq', 'código bioquímica'],
        'Nombre LIS': ['Nombre LIS', 'Derivación', 'nombre del examen', 'examen', 'analito', 'perfil bioquímico'],
        'Laboratorio de procesamiento': ['laboratorio de procesamiento', 'laboratorio', 'sección bioquímica',
                                         'lab bioq'],
        'Recepción de muestras': ['recepcion de muestras', 'recepción de muestras', 'recepcion', 'recepción suero'],
        'Procesamiento de muestras': ['procesamiento de muestras', 'procesamiento', 'centrifugación',
                                      'separación de suero'],
        'Plazo entrega resultados': ['tiempo de respuesta', 'plazo entrega resultados', 'plazo de entrega', 'tat'],
        'Preparación del paciente': ['preparacion paciente', 'preparación paciente', 'ayuno 8 horas', 'ayuno 12 horas'],
        'Procedimiento de toma de muestras': ['Procedimiento', 'Procedimiento de toma de muestras', 'punción venosa'],
        'Muestra requerida y recipiente': ['tipo de muestra', 'muestra requerida', 'suero', 'plasma', 'tubo tapa roja',
                                           'tubo gel separador', 'obtencion y preparacion de las muestras',
                                           'obtención y preparación de las muestras'],
        'Rotulación': ['Rotulación', 'Rotulacion'],
        'Estabilidad de la muestra': ['estabilidad muestra', 'estabilidad de la muestra', 'estabilidad suero',
                                      'conservacion y estabilidad', 'conservación y estabilidad'],
        'Condiciones de envío al Laboratorio': ['Condiciones de envío al Laboratorio', 'temperatura de transporte',
                                                'refrigerado 2-8°c'],
        'Criterios de rechazo': ['Rechazo', 'Criterios de rechazo', 'suero lipémico', 'suero ictérico', 'hemólisis ++'],
        'Método de procesamiento': ['metodo utilizado', 'método utilizado', 'espectrofotometría', 'fotometría',
                                    'enzimático', 'principio del test', 'principio', 'principio de la prueba'],
        'Test incluidos': ['test incluidos', 'pruebas incluidas', 'perfil hepático', 'perfil lipídico',
                           'función renal'],
        'Valores de referencia': ['valor de referencia', 'valores de referencia', 'valores normales',
                                  'rangos de referencia'],
        'Valor crítico': ['valor critico', 'valor crítico', 'hipoglicemia severa', 'hiperkalemia'],
        'Límite de detección': ['limite de deteccion', 'límite de detección', 'límite de cuantificación',
                                'limite de cuantificacion', 'loq'],
        'Equipamiento utilizado': ['equipamiento utilizado', 'equipo utilizado', 'autoanalizador bioquímico', 'system',
                                   'sistema', 'informacion del sistema', 'información del sistema'],
        'Información clínica': ['caracteristicas', 'características', 'utilidad clinica', 'utilidad clínica',
                                'informacion clinica', 'información clínica'],
        'Referencias': ['referencias', 'bibliografia', 'bibliografía'],
        'Actualizado por': ['actualizado por'],
        'Fecha actualización': ['fecha actualizacion', 'fecha actualización'],
        'Autorizado por': ['autorizado por'],
    },
    "Hormonas e inmunidad": {
        'Código': ['codigo examen', 'codigo', 'código examen', 'código', 'cod_hormonas', 'código inmunología'],
        'Nombre LIS': ['Nombre LIS', 'Derivación', 'nombre del examen', 'examen', 'inmunoensayo', 'hormona',
                       'anticuerpo'],
        'Laboratorio de procesamiento': ['laboratorio de procesamiento', 'laboratorio', 'sección hormonas',
                                         'lab inmunología'],
        'Recepción de muestras': ['recepcion de muestras', 'recepción de muestras', 'recepcion'],
        'Procesamiento de muestras': ['procesamiento de muestras', 'procesamiento', 'centrifugación de muestra'],
        'Plazo entrega resultados': ['tiempo de respuesta', 'plazo entrega resultados', 'plazo de entrega', 'tat'],
        'Preparación del paciente': ['preparacion paciente', 'preparación paciente', 'horario de toma específico',
                                     'ritmo circadiano'],
        'Procedimiento de toma de muestras': ['Procedimiento', 'Procedimiento de toma de muestras', 'punción venosa'],
        'Muestra requerida y recipiente': ['tipo de muestra', 'muestra requerida', 'suero', 'tubo tapa roja',
                                           'tubo gel separador', 'obtencion y preparacion de las muestras',
                                           'obtención y preparación de las muestras'],
        'Rotulación': ['Rotulación', 'Rotulacion'],
        'Estabilidad de la muestra': ['estabilidad muestra', 'estabilidad de la muestra', 'estabilidad en congelación',
                                      'conservacion y estabilidad', 'conservación y estabilidad'],
        'Condiciones de envío al Laboratorio': ['Condiciones de envío al Laboratorio', 'temperatura de transporte',
                                                'transporte congelado', 'refrigerado'],
        'Criterios de rechazo': ['Rechazo', 'Criterios de rechazo', 'muestra descongelada', 'hemólisis intensa'],
        'Método de procesamiento': ['metodo utilizado', 'método utilizado', 'elisa', 'quimioluminiscencia', 'eclia',
                                    'ria', 'inmunofluorometría', 'principio del test', 'principio',
                                    'principio de la prueba'],
        'Test incluidos': ['test incluidos', 'pruebas incluidas', 'perfil tiroideo', 'panel hormonal'],
        'Valores de referencia': ['valor de referencia', 'valores de referencia', 'rangos según edad/sexo',
                                  'fase folicular/lútea'],
        'Valor crítico': ['valor critico', 'valor crítico', 'alerta endocrina'],
        'Límite de detección': ['limite de deteccion', 'límite de detección', 'sensibilidad funcional',
                                'límite de cuantificación', 'limite de cuantificacion', 'loq'],
        'Equipamiento utilizado': ['equipamiento utilizado', 'equipo utilizado', 'analizador de inmunoensayo', 'system',
                                   'sistema', 'informacion del sistema', 'información del sistema'],
        'Información clínica': ['caracteristicas', 'características', 'utilidad clinica', 'utilidad clínica',
                                'informacion clinica', 'información clínica'],
        'Referencias': ['referencias', 'bibliografia', 'bibliografía'],
        'Actualizado por': ['actualizado por'],
        'Fecha actualización': ['fecha actualizacion', 'fecha actualización'],
        'Autorizado por': ['autorizado por'],
    },
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
    """Obtiene el diccionario de alias del área solicitada o la estructura de respaldo."""
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
    """Extrae texto de un PDF utilizando PyMuPDF filtrando dinámicamente márgenes y ruido."""
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
        import pytesseract
        from PIL import Image

        doc = fitz.open(stream=uploaded_bytes, filetype="pdf")
        parts = []
        for page in doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            parts.append(pytesseract.image_to_string(img, lang="spa+eng"))
        return "\n\n".join(parts), "PyMuPDF (OCR)"
    except Exception as e:
        raise RuntimeError("No se pudo extraer texto del PDF con PyMuPDF ni OCR.") from e


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
                found[current] = (found.get(current, '') + ('\n' if found.get(current) else '') + clean_value(
                    remainder)).strip()
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
    if not data.get('Código'):
        m = re.search(r'\b\d{6,}(?:[-.]\d+)?\b', text)
        if m:
            data['Código'] = m.group(0)

    # 1. Búsqueda y extracción limpia de modelos de equipamiento (ej. cobas t 511, cobas t 711, cobas c 501, cobas pure, etc.)
    equipos_encontrados = re.findall(r'\bcobas\s+(?:[a-zA-Z]\s*)?\d{3,4}\b|\bcobas\s+(?:pure|pro)\b', text, flags=re.I)
    if equipos_encontrados:
        equipos_limpios = []
        for eq in equipos_encontrados:
            eq_norm = re.sub(r'\s+', ' ', eq).strip()
            if eq_norm.casefold() not in [x.casefold() for x in equipos_limpios]:
                equipos_limpios.append(eq_norm)
        data['Equipamiento utilizado'] = "\n".join(equipos_limpios)
    elif data.get('Equipamiento utilizado'):
        lineas_limpias = []
        for line in data['Equipamiento utilizado'].split('\n'):
            l = line.strip()
            if l and not re.search(r'^(?:ACN|Nombre abreviado|Información|\d{4,5}|utilizar con|plasma|UniCal|UniCon)',
                                   l, re.I):
                lineas_limpias.append(l)
        data['Equipamiento utilizado'] = "\n".join(lineas_limpias)

    # 2. Detección inteligente del Método de procesamiento si no se capturó por encabezado
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
    """Crea una ficha nueva usando los campos del área correspondiente cuando el examen aún no existe."""
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
            actualizado_por='', autorizado_por=''):
    if field_aliases is None:
        field_aliases = FIELD_ALIASES_BASE
    if not nombre_examen:
        raise ValueError('Selecciona un examen desde el menú del área antes de analizar el PDF.')
    ok, mensaje = nombre_examen_valido(nombre_examen)
    if not ok:
        raise ValueError(mensaje)
    text, method = extract_pdf_text(uploaded_bytes, filename, nombre_examen)
    data = infer_name_code(text, parse_labeled_text(text, field_aliases, nombre_examen))

    fecha_actual = datetime.now().strftime('%d/%m/%Y %H:%M')

    # Asignar explícitamente los campos metadata elegidos por el usuario y la fecha actual
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

    # Asegurar la inserción exacta en las filas del Excel
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

    # Registro de extracción
    log_name = 'Registro extracción'
    if log_name in wb.sheetnames:
        del wb[log_name]
    log = wb.create_sheet(log_name)
    log.append(['Área', 'Examen seleccionado', 'Documento', 'Método de extracción', 'Fecha de procesamiento'])
    log.append(['', nombre_examen, filename, method, datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
    log.append([])
    log.append(['Campo', 'Información extraída'])
    for field in field_aliases:
        log.append([field, data.get(field, '')])
    for row in log.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                cell.value = limpiar_valor_excel(cell.value, nombre_examen)
            cell.alignment = Alignment(vertical='top', wrap_text=True)
    ajustar_ancho_columnas(log)

    wb.active = ws

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue(), data, nombre_examen, score, reason, method, text


# --- ACCESO BÁSICO ---
DEFAULT_USERNAME = "CRSCODUOC"
DEFAULT_PASSWORD = "CRSCODUOC"


def check_login():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if st.session_state["authenticated"]:
        return True

    st.session_state.pop("modal_area_activa", None)

    _, col_login, _ = st.columns([1, 1.2, 1])

    with col_login:
        st.markdown("<h1 style='text-align: center;'>Acceso a la herramienta</h1>", unsafe_allow_html=True)
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


# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title=APP_TITLE, page_icon="🧪", layout="wide")

if not check_login():
    st.stop()

st.title(APP_TITLE)
st.caption("documento de derivación → extracción → revisión → exportación")

# Estados de la aplicación
if "examenes_por_area" not in st.session_state:
    st.session_state["examenes_por_area"] = cargar_examenes_por_area()
if "area_activa" not in st.session_state:
    st.session_state["area_activa"] = None
if "examen_activo" not in st.session_state:
    st.session_state["examen_activo"] = None


def abrir_modal_area(area):
    st.session_state["area_activa"] = area
    st.session_state["modal_area_activa"] = area


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
                st.session_state.pop("modal_area_activa", None)
                st.rerun()
            else:
                st.warning("Selecciona un examen o agrega uno nuevo.")
    with c2:
        if st.button("Limpiar selección", key=f"limpiar_examen_{area}", use_container_width=True):
            st.session_state["area_activa"] = area
            st.session_state["examen_activo"] = None
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
                if guardar_examenes_por_area(st.session_state["examenes_por_area"]):
                    st.session_state.pop("modal_area_activa", None)
                    st.success(f"Se quitó «{quitar}» del menú de selección.")
                    st.rerun()
                else:
                    st.error("No se pudo guardar el cambio.")
        else:
            st.info("No hay exámenes para eliminar.")


# --- SIDEBAR ---
with st.sidebar:
    col_izq, col_logo, col_der = st.columns([0.5, 4, 0.5])
    with col_logo:
        if os.path.exists(os.path.join(os.path.dirname(__file__), "CRSCO.png")):
            st.image(os.path.join(os.path.dirname(__file__), "CRSCO.png"), width=240)
    st.divider()
    st.header("🧪 Áreas del laboratorio")
    st.caption("Selecciona un área para abrir su menú de exámenes.")
    for area in AREA_CONFIG:
        st.button(
            area,
            key=f"abrir_area_{area}",
            use_container_width=True,
            type="primary" if st.session_state.get("area_activa") == area else "secondary",
            on_click=abrir_modal_area,
            args=(area,),
        )
    st.divider()
    st.markdown(f"**Área activa:** {st.session_state.get('area_activa') or 'Sin seleccionar'}")
    st.markdown(f"**Examen activo:** {st.session_state.get('examen_activo') or 'Sin seleccionar'}")
    st.divider()
    st.header("📁 Archivos")
    template = st.file_uploader("Excel personalizado (opcional)", type=["xlsx"], key="template_uploader")
    st.caption("Si no cargas uno, se utilizará la plantilla incluida para el área seleccionada.")
    if st.button("Cerrar sesión"):
        st.session_state["authenticated"] = False
        st.session_state.pop("result", None)
        st.session_state.pop("modal_area_activa", None)
        st.session_state["area_activa"] = None
        st.rerun()

# Modal del área activa
if st.session_state.get("modal_area_activa") in AREA_CONFIG:
    modal_examenes_area(st.session_state["modal_area_activa"])

# --- VISTA PRINCIPAL ---
if st.session_state.get("area_activa") is None:
    st.write("")
    st.info("### 👋 ¡Bienvenido/a a la Herramienta de Extracción de Datos!")
    st.markdown(
        """
        Para comenzar a trabajar, selecciona un **área del laboratorio** en el panel lateral de la izquierda.

        **Pasos para el procesamiento:**
        1. 🧪 **Selecciona un área** (Microbiología, Hematología, Bioquímica, etc.).
        2. 📋 **Selecciona o añade el examen** en el menú desplegable del área.
        3. 📄 **Carga la ficha técnica (PDF)** para extraer y revisar la información automáticamente.
        """
    )
else:
    tab_pdf, tab_export, tab_guide = st.tabs([
        "Extraer información desde un PDF",
        "Exportar a otro formato",
        "Guía de desarrollo seguro",
    ])

    with tab_pdf:
        st.subheader("Extraer información desde un PDF")
        st.write("Selecciona el examen en el menú del área correspondiente y carga el documento técnico.")
        area_activa = st.session_state.get("area_activa")
        examen_activo = st.session_state.get("examen_activo")
        st.info(f"Área seleccionada: **{area_activa}** · Examen: **{examen_activo or 'sin seleccionar'}**")

        st.divider()
        st.markdown("##### ⚙️ Parámetros operacionales de la ficha")

        # 3 Columnas directas con menú desplegable (Selectbox)
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

        # Sección integrada para administrar opciones de las tres listas
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

        pdf = st.file_uploader("Cargar ficha técnica", type=["pdf"], help="PDF con la información del examen.",
                               key="pdf_uploader")

        if pdf:
            if st.button("🔎 Analizar documento", type="primary", key="analyze_pdf"):
                if not examen_activo:
                    st.warning("Primero abre el menú del área en el panel lateral y selecciona un examen.")
                elif not st.session_state.get("laboratorio_activo"):
                    st.warning("Selecciona un laboratorio de procesamiento antes de analizar.")
                else:
                    try:
                        config = AREA_CONFIG[area_activa]
                        aliases_area = get_field_aliases(area_activa)
                        ruta_plantilla = os.path.join(os.path.dirname(__file__), config["template"] or "")
                        if template is not None:
                            template_bytes = template.getvalue()
                        elif config["template"] and os.path.exists(ruta_plantilla):
                            with open(ruta_plantilla, "rb") as f:
                                template_bytes = f.read()
                        else:
                            wb_vacio = openpyxl.Workbook()
                            ws_vacio = wb_vacio.active
                            ws_vacio.title = examen_activo
                            for fila, campo in enumerate(aliases_area, start=1):
                                ws_vacio.cell(row=fila, column=1, value=campo)
                                ws_vacio.cell(row=fila, column=2,
                                              value=examen_activo if norm(campo) == norm("Nombre LIS") else "")
                            buf = io.BytesIO();
                            wb_vacio.save(buf);
                            template_bytes = buf.getvalue()

                        result = process(
                            pdf.getvalue(), pdf.name, template_bytes,
                            st.session_state.get("laboratorio_activo", ""),
                            examen_activo, aliases_area,
                            st.session_state.get("actualizado_activo", ""),
                            st.session_state.get("autorizado_activo", "")
                        )
                        st.session_state["result"] = result
                        st.session_state["filename"] = pdf.name
                        st.session_state["area_resultado"] = area_activa
                    except Exception as e:
                        st.error(f"No se pudo analizar el documento: {e}")
        else:
            st.info("Carga un PDF para comenzar la extracción.")

        if "result" in st.session_state:
            out, data, target, score, reason, method, extracted_text = st.session_state["result"]
            area_res = st.session_state.get("area_resultado", area_activa)
            aliases_area = get_field_aliases(area_res)
            st.subheader("Resultado de extracción")
            st.metric("Campos encontrados", sum(bool(v) for v in data.values()))

            rows = []
            for field in aliases_area:
                value = data.get(field, "")
                if field in ["Nombre LIS", "Laboratorio de procesamiento", "Actualizado por", "Autorizado por"]:
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
            df = pd.DataFrame(rows)


            def highlight_found(row):
                if row["Estado"].startswith("✓"):
                    return ["background-color: #B7F7C1; color: #124D20; font-weight: bold"] * len(row)
                return ["background-color: #FFF0B3; color: #6B4D00"] * len(row)


            st.markdown(
                """
                <div style="font-size: 0.85rem; color: #6c757d; line-height: 1.6;">
                    <span style="background-color: #d4edda; color: #155724; padding: 2px 6px; border-radius: 4px; font-weight: bold;">Verde:</span> Dato encontrado o ingresado por el usuario / sistema.
                    <br><br>
                    <span style="background-color: #fff3cd; color: #856404; padding: 2px 6px; border-radius: 4px; font-weight: bold;">Amarillo:</span> Dato no encontrado o pendiente de ingreso.
                </div>
                """,
                unsafe_allow_html=True
            )

            st.dataframe(
                df.style.apply(highlight_found, axis=1),
                use_container_width=True,
                hide_index=True,
            )

            st.subheader("Descargar Excel actualizado")
            st.write(
                "El Excel conserva las fichas de la plantilla del área y actualiza la ficha del examen seleccionado.")
            st.download_button(
                "Descargar Excel actualizado (.xlsx)", type="primary",
                data=out,
                file_name=f"Fichas_tecnicas_{st.session_state.get('area_resultado', 'Area').replace(' ', '_')}_actualizadas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_updated_xlsx",
            )

            with st.expander("Ver texto bruto extraído"):
                st.text_area("Texto", extracted_text, height=350, key="raw_extracted_text")

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
