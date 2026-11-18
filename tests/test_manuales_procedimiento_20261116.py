"""El OCR de los manuales se interpreta sin inventar valores ni pisar fichas previas."""
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

APP = (ROOT / "app.py").read_text(encoding="utf-8")
TREE = ast.parse(APP)
PROCESS = next(n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name == "process")

# La sección de la página 5 se conserva como muestra OCR para esta regresión.
MANUAL_OCR = """
PROCEDIMIENTO PARA LA DETECCION DE ANTICOAGULANTE LUPICO EN COAGULOMETRO COBAS T511
Laboratorio Clínico CRSCO
Página: 5 de 11
6. DESARROLLO
6.1. RECEPCION DE LAS MUESTRAS
Este examen aplica a muestras recibidas de pacientes del CRSCO y HSO.
Las muestras son recepcionadas y revisadas por el técnico paramédico del laboratorio.
Posteriormente, TENS hará la recepción de la muestra en el LIS.
Para la determinación se deben tomar 2 tubos celestes con citrato de sodio al 3,2%.
6.2. PREPARACION DE LAS MUESTRAS
Una vez recepcionadas las muestras, se inicia el proceso de centrifugación.
Tubo 1: se centrifuga a 4000 rpm por 15 min y se entrega a la sección para procesamiento.
Tubo 2: se somete a doble centrifugación a 3600 rpm por 15 minutos para separar el plasma.
6.3. PREPARACION DE REACTIVOS
Los reactivos se deben mantener refrigerados hasta su uso.
6.4. PROCEDIMIENTO
Primero controlar la técnica de acuerdo con el procedimiento correspondiente.
"""


def test_encabezados_numerados_de_manual_extraen_dos_campos():
    datos = mp.extraer_campos_manual(MANUAL_OCR)
    assert 'Recepción de muestras' in datos
    assert 'Procesamiento de muestras' in datos
    assert 'recepcionadas' in datos['Recepción de muestras']
    assert '3600 rpm' in datos['Procesamiento de muestras']
    assert 'reactivos' not in datos['Procesamiento de muestras'].casefold()
    assert 'Preparación del paciente' not in datos
    assert 'Valores de referencia' not in datos
    assert 'Procedimiento de toma de muestras' not in datos
    assert datos['Equipamiento utilizado'].upper() == 'COBAS T511'


def test_manual_no_debe_inventar_valores_por_prosa_generica():
    assert mp.extraer_campos_manual('''
    6. PROCEDIMIENTO
    El ensayo se realizó con muestras de plasma.
    Los valores fueron procesados y medidos en el analizador.
    ''') == {}


def test_filtro_solo_celdas_vacias_ignora_contenido_validado():
    datos = {'Recepción de muestras':'Nuevo texto', 'Procesamiento de muestras':'Centrifugar'}
    validados = {'Recepción de muestras':'Ya revisado', 'Procesamiento de muestras':None}
    assert mp.filtrar_celdas_ocupadas(datos, validados) == {'Procesamiento de muestras':'Centrifugar'}


class FuentesSimuladas:
    CAMPOS_CONTROLADOS = ('Recepción de muestras','Procesamiento de muestras','Preparación del paciente')
    CAMPOS_MANUAL = ('Recepción de muestras','Procesamiento de muestras')


def _contexto():
    ns = {
        '__name__':'prueba_manual',
        'io':io,'datetime':datetime,'ZoneInfo':ZoneInfo,
        'fd':FuentesSimuladas(), 'mp':mp,
        'load_workbook':openpyxl.load_workbook,
        'Alignment':openpyxl.styles.Alignment,
        'FIELD_ALIASES_BASE':{'Recepción de muestras':[], 'Procesamiento de muestras':[]},
        'nombre_examen_valido':lambda _: (True,''),
        'buscar_hoja_por_examen':lambda wb,n: wb[n] if n in wb.sheetnames else None,
        'norm':lambda s: str(s or '').lower().strip(),
        'normalizar_etiquetas_ficha':lambda ws: None,
        'infer_name_code': lambda text, data: data,
        'limpiar_valor_excel':lambda v,*args:v,
        'aplicar_formato_normas_excel':lambda *a:a[0],
        'es_alias_hematologia':lambda _:True,
        'ajustar_ancho_columnas':lambda _:None,
        'ajustar_alto_fila_con_listas':lambda *args:None,
    }
    def _apply(ws,data):
        for row in ws.iter_rows():
            if row[0].value in data:
                row[1].value=data[row[0].value]
        return []
    ns['apply_to_sheet']=_apply
    exec(compile(ast.Module(body=[PROCESS],type_ignores=[]),str(ROOT/'app.py'),'exec'),ns)
    return ns


def _template():
    wb = openpyxl.Workbook()
    ws=wb.active
    ws.title='AC. LUPICO'
    ws.append(['Recepción de muestras','Revisión anterior aprobada'])
    ws.append(['Procesamiento de muestras',None])
    ws.append(['Preparación del paciente','Indicaciones del CRS'])
    ws.append(['Nombre LIS','AC. LUPICO'])
    ws.append(['Fecha actualización','12/11/2026 02:35'])
    buf=io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_manual_propone_datos_nuevos_y_respeta_prioridad_cartera():
    ns=_contexto()
    datos=mp.extraer_campos_manual(MANUAL_OCR)
    out, data, *_ = ns['process'](
        b'pdf','manual.pdf',_template(),nombre_examen='AC. LUPICO',
        datos_pdf_preextraidos=(datos,MANUAL_OCR,'RapidOCR'),
        datos_sobrescribir={'Preparación del paciente':'Indicaciones desde cartera'},
        manual_procedimiento=True)
    wb=openpyxl.load_workbook(io.BytesIO(out))
    ws=wb['AC. LUPICO']
    assert 'recepcionadas' in ws['B1'].value
    assert ws['B2'].value and '3600 rpm' in ws['B2'].value
    assert ws['B3'].value == 'Indicaciones desde cartera'
    assert 'recepcionadas' in data['Recepción de muestras']
    assert data['Procesamiento de muestras']
    assert data['Nombre LIS']=='AC. LUPICO'
    assert data['Fecha actualización'] != '12/11/2026 02:35'
    assert ws['B5'].value == data['Fecha actualización']


def test_normal_ficha_roche_conserva_prohibicion_campos_manual():
    ns=_contexto()
    datos={'Recepción de muestras':'Nunca desde inserto','Procesamiento de muestras':'No autorizado'}
    out, data, *_=ns['process'](
        b'pdf','roche.pdf',_template(),nombre_examen='AC. LUPICO',
        datos_pdf_preextraidos=(datos,'texto','PyMuPDF'),
        manual_procedimiento=False)
    ws=openpyxl.load_workbook(io.BytesIO(out)).active
    assert 'Recepción de muestras' not in data
    assert 'Procesamiento de muestras' not in data
    assert ws['B1'].value is None and ws['B2'].value is None


def test_interfaz_optativa_con_paginas_y_prioridad_de_fuentes():
    assert '"Documentos escaneados"' in APP
    assert 'rango_paginas=rango_paginas_manual' in APP
    assert 'manual_procedimiento=es_manual_procedimiento' in APP
    assert 'datos_sobrescribir=datos_sobrescribir' in APP


def test_ocr_separa_numero_y_titulo_de_seccion():
    texto = '6.1.\nRECEPCION DE LAS MUESTRAS\nLas muestras son recibidas y revisadas por el laboratorio.\n6.2\nPREPARACION DE LAS MUESTRAS\nSe centrifugan las muestras de plasma para análisis.\n6.3. PREPARACION DE REACTIVOS\nLos reactivos permanecen refrigerados.'
    datos=mp.extraer_campos_manual(texto)
    assert 'Recepción de muestras' in datos
    assert 'Procesamiento de muestras' in datos
    assert 'refrigerados' not in datos['Procesamiento de muestras']


def test_ocr_paginas_acotadas_y_sin_vuelta_png(monkeypatch):
    import fitz
    import re
    import types
    import numpy as np
    n_llamadas = []

    class MotorOCRSimulado:
        def __call__(self, imagen):
            n_llamadas.append(tuple(imagen.shape))
            return [([0, 0, 1, 1], '6.1. RECEPCION DE LAS MUESTRAS: Se revisa la muestra.', 0.98)], 0.1

    monkeypatch.setitem(sys.modules, 'rapidocr_onnxruntime',
                        types.SimpleNamespace(RapidOCR=MotorOCRSimulado))
    doc = fitz.open()
    for _ in range(3):
        doc.new_page(width=595, height=842)
    pdf = doc.tobytes()
    doc.close()

    extractor = next(n for n in TREE.body
                     if isinstance(n, ast.FunctionDef) and n.name == 'extract_pdf_text')
    import ocr_documentos_escaneados as oe
    monkeypatch.setattr(oe, '_tesseract_disponible', lambda: '')
    ns = {
        'fitz':fitz, 'io':io,'re':re, 'oe':oe,
        '_superindices_visuales_en_pagina':lambda pagina:{},
        'es_linea_basura_o_encabezado':lambda *args:False,
        '_membrete_clinico_hematologia':lambda *args:set(),
    }
    exec(compile(ast.Module(body=[extractor],type_ignores=[]),
                 str(ROOT/'app.py'),'exec'),ns)
    progreso=[]
    texto, metodo = ns['extract_pdf_text'](
        pdf,'manual.pdf', manual_procedimiento=True,
        rango_paginas=(2,2),
        progreso=lambda *args: progreso.append(args))
    assert len(n_llamadas)==1
    assert n_llamadas[0][-1]==3
    assert n_llamadas[0][0] < 842 * 200 / 72  # 160 dpi, no 200 dpi
    assert progreso == [(0,1,False),(1,1,True)]
    assert 'RECEPCIÓN DE LAS MUESTRAS' in texto
    assert 'RapidOCR' in metodo
