"""Regresiones contra OCR imperfecto en manuales y preservación de datos previos."""
import ast
import io
import sys
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import manuales_procedimiento as mp


# Variantes que aparecen al reconocer páginas escaneadas de los manuales:
# títulos partidos, fallos de puntuación, I/l y párrafos en la misma línea.
OCR_INESTABLE = """
Código: LC-PM-17
Página: 5 de 11
6. DESARROLLO
6.1.
RECEPCI0N DE
LAS MUESTRAS
Las muestras son recepcionadas por el personal del laboratorio y se verifica su identificación.
Posteriormente se registra la muestra en el sistema de información.
6,2
PREPARACION DE LAS
MUESTRAS
El tubo se centrifuga de acuerdo al procedimiento descrito en este mismo documento.
Se separa el plasma para realizar la determinación correspondiente.
6.3. PREPARACION DE REACTIVOS
Los reactivos se conservan refrigerados y NO corresponden al procesamiento de muestras.
"""


def test_ocr_con_titulos_partidos_y_numeros_danados():
    resultado = mp.analizar_manual(OCR_INESTABLE)
    assert set(resultado['campos']) == {'Recepción de muestras', 'Procesamiento de muestras'}
    assert 'sistema de información' in resultado['campos']['Recepción de muestras']
    assert 'Se separa el plasma' in resultado['campos']['Procesamiento de muestras']
    assert 'reactivos se conservan' not in resultado['campos']['Procesamiento de muestras']
    assert not resultado['omitidos']


def test_ocr_con_caracter_l_que_sustituye_i():
    ocr = '''6.1 RECEPClON DE LAS MUESTRAS
    Las muestras se reciben y comprueban antes de ingresarlas al laboratorio.
    6.2 PREPARACION DE LAS MUESTRAS
    Las muestras son centrifugadas siguiendo el procedimiento establecido.
    6.3 PREPARACION DE REACTIVOS
    Este texto no corresponde a las muestras.'''
    campos = mp.extraer_campos_manual(ocr)
    assert 'Recepción de muestras' in campos
    assert 'Procesamiento de muestras' in campos
    assert 'Este texto' not in campos['Procesamiento de muestras']


def test_ocr_con_encabezado_y_texto_en_mismo_renglon():
    ocr = '''6.1. RECEPCION DE LAS MUESTRAS: El personal debe identificar, recibir y registrar las muestras en el laboratorio.
    6.2. PREPARACION DE LAS MUESTRAS Se separa el plasma por centrifugación antes de realizar la determinación.
    6.3. PREPARACION DE REACTIVOS Se prepara el reactivo y se controla antes del análisis.'''
    campos = mp.extraer_campos_manual(ocr)
    assert 'Recepción de muestras' in campos
    assert 'Procesamiento de muestras' in campos
    assert 'reactivo' not in campos['Procesamiento de muestras'].lower()


def test_modo_escaneado_extrae_frases_especificas_sin_encabezados():
    texto = '''Para la recepción de muestras se debe cumplir el protocolo.
    Las muestras son centrifugadas antes de analizarse.
    Se utiliza citrato de sodio al 3,2% en todo momento.'''
    resultado = mp.extraer_campos_manual(texto)
    assert 'Recepción de muestras' in resultado
    assert 'Procesamiento de muestras' in resultado
    assert 'Muestra requerida y recipiente' in resultado
    # La modalidad convencional NO llama a este extractor; usa
    # infer_name_code / parse_labeled_text por separado.
    fuente = (ROOT / 'app.py').read_text(encoding='utf-8')
    assert 'if manual_procedimiento else' in fuente


def test_modo_manual_no_vacia_valor_documental_validado_en_el_excel():
    tree = ast.parse((ROOT / 'app.py').read_text(encoding='utf-8'))
    funcion = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == 'process')
    import test_manuales_procedimiento_20261116 as anterior
    ns = anterior._contexto()
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.title = 'AC. LUPICO'
    hoja.append(['Recepción de muestras', None])
    hoja.append(['Procesamiento de muestras', None])
    hoja.append(['Preparación del paciente', 'Indicación previamente validada'])
    hoja.append(['Rotulación', 'Norma general validada'])
    hoja.append(['Equipamiento utilizado', 'Cobas T511 ya revisado'])
    hoja.append(['Nombre LIS', 'AC. LUPICO'])
    hoja.append(['Fecha actualización', '12/11/2026 02:35'])
    buf = io.BytesIO()
    libro.save(buf)
    datos = mp.extraer_campos_manual(OCR_INESTABLE)
    datos['Equipamiento utilizado'] = 'Equipo adicional no confirmado'
    out, data, *_ = ns['process'](
        b'pdf', 'manual.pdf', buf.getvalue(), nombre_examen='AC. LUPICO',
        datos_pdf_preextraidos=(datos, OCR_INESTABLE, 'RapidOCR'),
        datos_sobrescribir=None, manual_procedimiento=True)
    hoja_final = openpyxl.load_workbook(io.BytesIO(out), data_only=True).active
    valores = {hoja_final.cell(i, 1).value: hoja_final.cell(i, 2).value
               for i in range(1, hoja_final.max_row + 1)}
    assert 'se registra la muestra' in valores['Recepción de muestras']
    assert 'Se separa el plasma' in valores['Procesamiento de muestras']
    # La nueva extracción crea propuestas actuales, sin reciclar el contenido
    # anterior; ese contenido permanece en los bytes de plantilla para
    # mostrarlo al desmarcar «Aplicar».
    assert 'Preparación del paciente' not in data
    assert 'Rotulación' not in data
    assert valores['Equipamiento utilizado'] == 'Equipo adicional no confirmado'
    assert data['Equipamiento utilizado'] == 'Equipo adicional no confirmado'
    assert valores['Fecha actualización'] != '12/11/2026 02:35'
    import revision_valores
    antiguos = revision_valores.valores_anteriores(buf.getvalue(), 'AC. LUPICO',
                                                    lambda t: str(t).strip().lower())
    assert antiguos['preparación del paciente'] == 'Indicación previamente validada'
    assert antiguos['equipamiento utilizado'] == 'Cobas T511 ya revisado'


def test_modo_manual_preserva_diagnostico_interno_sin_panel_visible():
    texto = (ROOT / 'app.py').read_text(encoding='utf-8')
    assert 'diagnostico_manual_actual = (mp.analizar_manual(raw_text)' in texto
    assert '"Diagnóstico de lectura del documento escaneado"' not in texto
    assert '"Descargar texto reconocido por OCR (.txt)"' not in texto
    assert 'if etiqueta in fd.CAMPOS_CONTROLADOS:' in texto
