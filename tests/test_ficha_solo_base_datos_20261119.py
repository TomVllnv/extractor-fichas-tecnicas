"""Validaciones del camino institucional SIN analizar ni simular un PDF."""
import ast
import io
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
APP = (ROOT / 'app.py').read_text(encoding='utf-8')
TREE = ast.parse(APP)


def _aislar(nombre, extra=None):
    node = next(n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name == nombre)
    ns = dict(extra or {})
    exec(compile(ast.fix_missing_locations(ast.Module(body=[node], type_ignores=[])),
                 str(ROOT / 'app.py'), 'exec'), ns)
    return ns[nombre]


def _plantilla():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'AC. LUPICO'
    for fila, campo in enumerate(('Código', 'Preparación del paciente', 'Método de procesamiento',
                                   'Nombre LIS', 'Laboratorio de procesamiento',
                                   'Actualizado por', 'Autorizado por', 'Fecha actualización'), 1):
        ws.cell(fila, 1, campo)
        ws.cell(fila, 2, 'Anterior' if campo == 'Preparación del paciente' else None)
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def _proceso_sin_pdf():
    import sys
    sys.path.insert(0, str(ROOT))
    from test_ocr_sin_doble_lectura_20261116 import _ns
    ns, llamadas = _ns()
    ns['fd'].CAMPOS_CONTROLADOS = {'Preparación del paciente'}
    ns['FIELD_ALIASES_BASE'] = {'Código': [], 'Preparación del paciente': [],
                                'Método de procesamiento': [], 'Nombre LIS': [],
                                'Laboratorio de procesamiento': [], 'Actualizado por': [],
                                'Autorizado por': [], 'Fecha actualización': []}
    return ns, llamadas


def test_no_copia_valores_antiguos_como_si_fueran_institucionales():
    construir = _aislar('datos_para_ficha_sin_pdf')
    assert construir({}, {}) == {}
    assert construir({'Código': '0312', 'Preparación del paciente': 'Ayuno de 6 horas'},
                     {'Preparación del paciente': 'Ayuno de 8 horas'}) == {
                         'Código': '0312', 'Preparación del paciente': 'Ayuno de 8 horas'}
    assert construir({'Código': '', 'Preparación del paciente': 'pendiente'},
                     {'Preparación del paciente': 'Ayuno de 8 horas'})['Preparación del paciente'] == 'Ayuno de 8 horas'


def test_generar_solo_fuentes_institucionales_no_invoca_lectura_ni_ocr():
    ns, llamadas = _proceso_sin_pdf()
    salida, datos, _, _, _, metodo, texto = ns['process'](
        b'', '', _plantilla(), 'Laboratorio de procesamiento', 'AC. LUPICO',
        ns['FIELD_ALIASES_BASE'], 'Encargado', 'Revisor',
        datos_sobrescribir={'Código': '03012', 'Preparación del paciente': 'Ayuno de 8 horas'},
        datos_pdf_preextraidos=({}, '', 'Sin análisis de PDF'),
        manual_procedimiento=False,
    )
    assert llamadas == []
    assert metodo == 'Sin análisis de PDF' and texto == ''
    assert datos['Código'] == '03012'
    assert datos['Preparación del paciente'] == 'Ayuno de 8 horas'
    assert 'Método de procesamiento' not in datos
    hoja = openpyxl.load_workbook(io.BytesIO(salida)).active
    filas = {hoja.cell(i, 1).value: hoja.cell(i, 2).value for i in range(1, hoja.max_row + 1)}
    assert filas['Código'] == '03012'
    assert filas['Preparación del paciente'] == 'Ayuno de 8 horas'


def test_existe_ruta_independiente_sin_documentos_y_con_revision():
    assert '"Omitir análisis y trabajar solo con la base de datos"' in APP
    assert '"Generar ficha con datos institucionales"' in APP
    assert '"Volver al análisis de documentos"' in APP
    assert 'if modo_solo_bd:\n            pdfs = []' in APP
    assert 'datos_pdf_preextraidos=({}, "", "Sin análisis de PDF")' in APP
    assert '"procedencia_modalidad"] = "Solo base de datos"' in APP
    assert 'and (pdfs or modo_solo_bd)' in APP
    assert 'def _fuentes_para_referencias' in APP
    assert 'if not pdfs_activos:\n        return [], []' not in APP
