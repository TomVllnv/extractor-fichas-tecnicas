"""Regresiones de rendimiento: no volver a leer PDF OCR para generar el Excel."""
import ast
import io
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
APP = (ROOT / 'app.py').read_text(encoding='utf-8')
TREE = ast.parse(APP)
PROCESS = next(n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name == 'process')


class FuentesSimuladas:
    CAMPOS_CONTROLADOS = {'Recepción de muestras', 'Procesamiento de muestras'}


def _ns():
    llamadas = []
    def extraer(*args, **kwargs):
        llamadas.append('ocr')
        return 'Método de procesamiento: dRVVT', 'RapidOCR'

    def aplicar(ws, datos):
        for row in ws.iter_rows():
            key = row[0].value
            if key in datos:
                row[1].value = datos[key]
        return []

    def buscar(wb, nombre):
        return wb[nombre] if nombre in wb.sheetnames else None

    ns = {
        '__name__': 'prueba_ocr', 'io': io, 'datetime': datetime, 'ZoneInfo': ZoneInfo,
        'fd': FuentesSimuladas(), 'load_workbook': openpyxl.load_workbook,
        'Alignment': openpyxl.styles.Alignment,
        'FIELD_ALIASES_BASE': {'Método de procesamiento': [], 'Recepción de muestras': []},
        'nombre_examen_valido': lambda _: (True, ''),
        'buscar_hoja_por_examen': buscar,
        'norm': lambda s: str(s or '').lower().strip(),
        'normalizar_etiquetas_ficha': lambda ws: None,
        'extract_pdf_text': extraer,
        'consolidar_documentos_pdf': lambda *a: (_ for _ in ()).throw(AssertionError('OCR repetido en cola')),
        'parse_labeled_text': lambda text, *args: {'Método de procesamiento': 'dRVVT'},
        'infer_name_code': lambda text, data: dict(data),
        'es_alias_hematologia': lambda _: True,
        'limpiar_valor_excel': lambda v, *args: v,
        'ajustar_ancho_columnas': lambda ws: None,
        'ajustar_alto_fila_con_listas': lambda ws, fila: None,
        'apply_to_sheet': aplicar,
    }
    exec(compile(ast.Module(body=[PROCESS], type_ignores=[]), str(ROOT / 'app.py'), 'exec'), ns)
    return ns, llamadas


def _template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'AC. LUPICO'
    ws.append(['Método de procesamiento', None])
    ws.append(['Recepción de muestras', 'Texto antiguo'])
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def test_pdf_ya_leido_no_dispara_segundo_ocr_y_conserva_excel():
    ns, llamadas = _ns()
    original = {'Método de procesamiento': 'dRVVT', 'Recepción de muestras': 'no extraer'}
    out, data, _, _, _, method, raw = ns['process'](
        b'pdf', 'Ac_lupico.pdf', _template(), nombre_examen='AC. LUPICO',
        datos_pdf_preextraidos=(original, 'texto OCR revisado', 'RapidOCR'))
    assert llamadas == []
    assert method == 'RapidOCR'
    assert raw == 'texto OCR revisado'
    assert original['Recepción de muestras'] == 'no extraer'
    assert data['Método de procesamiento'] == 'dRVVT'
    assert 'Recepción de muestras' not in data
    sheet = openpyxl.load_workbook(io.BytesIO(out)).active
    assert sheet['B1'].value == 'dRVVT'
    assert sheet['B2'].value is None


def test_proceso_directo_sigue_leyendo_pdf_normalmente():
    ns, llamadas = _ns()
    _, data, _, _, _, method, raw = ns['process'](
        b'pdf', 'Ac_lupico.pdf', _template(), nombre_examen='AC. LUPICO')
    assert llamadas == ['ocr']
    assert method == 'RapidOCR'
    assert raw == 'Método de procesamiento: dRVVT'
    assert data['Método de procesamiento'] == 'dRVVT'


def test_cola_de_varios_pdf_reutiliza_consolidacion():
    ns, llamadas = _ns()
    _, data, *_ = ns['process'](
        b'pdf', 'primero.pdf', _template(), nombre_examen='AC. LUPICO',
        documentos_adicionales=[('segundo.pdf', b'pdf')],
        datos_pdf_preextraidos=({'Método de procesamiento': 'dRVVT'}, 'consolidado', 'PDF1 + PDF2'))
    assert llamadas == []
    assert data['Método de procesamiento'] == 'dRVVT'


def test_interfaz_pasa_la_lectura_previa_y_muestra_avance():
    assert 'datos_pdf_preextraidos=(parsed_pdf_data, raw_text, metodo_lectura_pdf)' in APP
    assert 'progreso=_actualizar_progreso_pdf' in APP
    assert 'estado_lectura_pdf.progress(' in APP
    assert 'estado_lectura_pdf.empty()' in APP
