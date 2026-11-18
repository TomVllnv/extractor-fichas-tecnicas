"""Pruebas de la separación de modalidad, propuestas y comparación con Excel previo."""
from __future__ import annotations
import io
from pathlib import Path

import fitz
import openpyxl
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
import sys
sys.path.insert(0, str(ROOT))
import deteccion_pdf as dp
import manuales_procedimiento as mp
import revision_valores as rv
import procedencia_campos as pc

APP = (ROOT / 'app.py').read_text(encoding='utf-8')


def _pdf_escaneado():
    imagen = Image.new('RGB', (360, 190), 'white')
    png = io.BytesIO()
    imagen.save(png, format='PNG')
    pdf = fitz.open()
    hoja = pdf.new_page()
    hoja.insert_image(fitz.Rect(20, 20, 350, 220), stream=png.getvalue())
    resultado = pdf.tobytes()
    pdf.close()
    return resultado


def test_deteccion_documento_escaneado_sin_ocr():
    tipo, cantidad, total = dp.inspeccionar_pdf(_pdf_escaneado())
    assert (tipo, cantidad, total) == ('escaneado', 1, 1)


def test_deteccion_documento_textual_y_mixto():
    pdf = fitz.open()
    page = pdf.new_page()
    page.insert_text((30, 40), 'Documento de texto digital, ' * 12)
    assert dp.inspeccionar_pdf(pdf.tobytes())[0] == 'texto'
    other = fitz.open(stream=_pdf_escaneado(), filetype='pdf')
    other.insert_pdf(pdf)
    assert dp.inspeccionar_pdf(other.tobytes())[0] == 'mixto'
    other.close()
    pdf.close()


def test_contexto_sin_encabezado_reconoce_estabilidad_y_rechazo():
    texto = ('La muestra se conserva durante 24 horas entre 2 y 8 °C.\n\n'
             'Se rechazan muestras que presentan hemólisis visible.\n\n'
             'No está permitido confundir el tiempo de reacción con el plazo de entrega.')
    valores = mp.extraer_campos_manual(texto)
    assert 'Estabilidad de la muestra' in valores
    assert 'Criterios de rechazo' in valores
    assert 'Plazo entrega resultados' not in valores


def test_contexto_encabezado_propio_prioritario():
    texto = '''6.1. RECEPCIÓN DE MUESTRAS
Se registra la recepción en LIS.
Se realiza centrifugación a 3600 RPM durante 15 minutos.
6.2. PROCESAMIENTO DE MUESTRAS
Las muestras se procesan según el procedimiento aprobado por el laboratorio.'''
    valores = mp.extraer_campos_manual(texto)
    assert 'registra la recepción' in valores['Recepción de muestras']
    assert 'centrifugación' not in valores['Recepción de muestras']
    assert 'se procesan' in valores['Procesamiento de muestras']
    assert '3600 RPM' not in valores['Procesamiento de muestras']


def test_ocr_no_autocompleta_referencias():
    texto = '6.8. REFERENCIAS BIBLIOGRÁFICAS\nDocumento de laboratorio 2025.\n'
    assert 'Referencias' not in mp.extraer_campos_manual(texto)
    assert 'data.pop("Referencias", None)' in APP
    assert 'parsed_pdf_data.pop("Referencias", None)' in APP
    assert 'btn_aplicar_referencias_apa' in APP


def test_plantilla_anterior_se_mantiene_como_fuente_comparacion():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'AC. LUPICO'
    ws.append(['Código', 'CRS-11'])
    ws.append(['Estabilidad de la muestra', '8 horas'])
    ws.append(['Referencias', 'Libro anterior'])
    salida = io.BytesIO()
    wb.save(salida)
    anterior = rv.valores_anteriores(salida.getvalue(), 'AC. LUPICO', lambda x: str(x).lower())
    assert anterior['código'] == 'CRS-11'
    assert anterior['estabilidad de la muestra'] == '8 horas'
    assert anterior['referencias'] == 'Libro anterior'
    assert not rv.valores_anteriores(salida.getvalue(), 'OTRO', lambda x: str(x).lower())


def test_fuente_seleccionada_prioridad_con_discrepancia():
    propuesta, discrepancia = rv.decidir_propuesta('Muestra del OCR', 'Muestra de Cartera')
    assert discrepancia and propuesta == 'Muestra de Cartera'
    propuesta, discrepancia = rv.decidir_propuesta('Solo OCR', None)
    assert not discrepancia and propuesta == 'Solo OCR'


def test_procedencia_real_ocr_en_campo_institucional():
    fuentes, efectiva = pc.procedencia_campo(
        'Preparación del paciente', 'Ayuno 4 horas', '',
        {'Preparación del paciente': 'Ayuno 4 horas'}, {}, set(), True, {})
    assert efectiva == pc.PDF_ESCANEADO
    assert fuentes[0] == pc.PDF_ESCANEADO
    _, efectiva = pc.procedencia_campo(
        'Preparación del paciente', 'Texto antiguo', '',
        {'Preparación del paciente': 'Ayuno 4 horas'}, {}, set(), True, {})
    assert efectiva is None


def test_interfaz_contiene_comparacion_visual_y_leyenda_naranja():
    assert '"Valor anterior": st.session_state.get("valores_anteriores_ficha"' in APP
    assert "params.data['Actualizar'] === false" in APP
    assert '#D99B60' in APP
    assert 'Campo completado con valor anterior' in APP
    assert 'Sin valor anterior' in APP
    assert "'Referencias';" in APP  # Referencias no editable salvo selección APA.


def test_modo_normal_no_hace_ocr_ni_cambia_modalidad_sin_confirmacion():
    inicio = APP.index('def extract_pdf_text(')
    fin = APP.index('def clean_value(', inicio)
    seccion = APP[inicio:fin]
    assert 'if manual_procedimiento and len(' in seccion
    assert 'oe.leer_pagina_escaneada(' in seccion
    assert 'result, _ = engine(imagen_rgb)' not in seccion
    assert 'clase_pdf' in APP
    assert 'elige «Documentos escaneados»' in APP


def test_se_eliminan_ambos_textos_solicitados():
    assert 'no cambia los tres estados ni sus colores' not in APP
    assert 'Solo se propondrán datos explícitos del documento escaneado' not in APP
    assert 'Fuentes de mayor a menor prioridad. Se destaca la utilizada en el examen.' in APP
