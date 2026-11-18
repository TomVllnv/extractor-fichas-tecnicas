"""Cobertura de la modalidad única, fuentes ordenadas y OCR espaciado."""
import sys
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import manuales_procedimiento as mp
import ocr_documentos_escaneados as oe
import procedencia_campos as pc
import fuentes_documentales as fd

APP = (ROOT / 'app.py').read_text(encoding='utf-8')


def test_exactamente_dos_modalidades_y_sin_selector_tipo_documento():
    assert 'modalidades_validas = ("Fichas técnicas", "Documentos escaneados")' in APP
    assert '"Un PDF por examen (cola actual)"' not in APP
    assert '"Varios PDF para un mismo examen"' not in APP
    assert '"Ficha técnica del fabricante"' not in APP
    assert '"Manual de procedimiento del laboratorio"' not in APP
    assert '"Tipo de documento para analizar"' not in APP
    assert 'es_manual_procedimiento = modo_procesamiento == "Documentos escaneados"' in APP


def test_procedencia_de_todos_los_campos_mismo_orden_de_tabla():
    seccion = APP.split('with st.expander("Procedencia de campos documentales"', 1)[1].split('st.markdown(\'<div class="cuaderno-procedencia">\'', 1)[0]
    assert 'for campo in aliases_area:' in seccion
    assert 'pc.procedencia_campo(' in seccion
    assert 'cuaderno-fuente-seleccionada' in seccion
    assert 'for campo in fd.CAMPOS_CONTROLADOS:' not in seccion


def test_orden_fuentes_respetado_y_negrita_correcta():
    fuentes, elegida = pc.procedencia_campo(
        'Preparación del paciente', 'No aplica', 'Cartera de Acreditación · Específico',
        {}, {}, set(), False, {'Preparación del paciente': 'No aplica'})
    assert fuentes == list(fd.PRIORIDAD_FUENTES_POR_CAMPO['Preparación del paciente'])
    assert elegida == 'Cartera de acreditación'
    fuentes, elegida = pc.procedencia_campo('Código', 'A2', '', {'Código':'A2'}, {}, set(), False)
    assert elegida == pc.PDF_TECNICO
    assert fuentes.index(pc.CRS) < fuentes.index(pc.PDF_TECNICO)
    fuentes, elegida = pc.procedencia_campo('Nombre LIS', 'AC. LUPICO')
    assert elegida == pc.EXAMEN  # valor real viene de la selección, no inventar PDF
    fuentes, elegida = pc.procedencia_campo('Recepción de muestras','Texto', '',{'Recepción de muestras':'Texto'}, {}, set(),True)
    assert elegida == pc.PDF_ESCANEADO
    assert fuentes == [pc.PDF_ESCANEADO]


def test_una_ficha_previa_no_atribuida_erroneamente_a_fuente_nueva():
    _, elegida = pc.procedencia_campo('Preparación del paciente', 'Anterior',
        'Cartera de Exámenes del CRS · Indicaciones', {}, {}, set(), True,
        {'Preparación del paciente':'Nuevo'})
    assert elegida is None  # un valor previo no revela su fuente documental


def test_ocr_rapid_preserva_lineas_y_separaciones_visuales():
    datos = [([[0,0],[50,0],[50,14],[0,14]],'6.1. RECEPCION DE LAS MUESTRAS',0.99),
             ([[0,18],[150,18],[150,30],[0,30]],'Se revisa la muestra.',0.95),
             ([[0,80],[200,80],[200,95],[0,95]],'6.2. PREPARACION DE LAS MUESTRAS',0.95)]
    texto = oe.reconstruir_lineas_rapidocr(datos)
    assert 'Se revisa la muestra.\n\n6.2.' in texto


def test_ocr_tesseract_no_invocado_en_ruta_ficha_tecnica():
    # Solo la condición explícita de documento escaneado autoriza un motor OCR.
    inicio = APP.index('def extract_pdf_text(')
    fin = APP.index('def clean_value(', inicio)
    seccion = APP[inicio:fin]
    assert 'if manual_procedimiento and len(' in seccion
    assert 'oe.leer_pagina_escaneada(' in seccion
    assert 'result, _ = engine(imagen_rgb)' not in seccion


def test_parrafos_del_manual_conservan_saltos_y_vinetas():
    texto = '''6.1. RECEPCION DE LAS MUESTRAS
Las muestras son revisadas por el técnico.
Posteriormente se registran en LIS.

Se solicita una nueva muestra si no cumple condiciones.
• Verificar la orden.
• Confirmar la muestra.
6.2. PREPARACION DE LAS MUESTRAS
Se centrifuga el tubo.''' 
    datos = mp.extraer_campos_manual(texto)
    recepcion = datos['Recepción de muestras']
    assert 'técnico.\n\nPosteriormente' in recepcion
    assert '\n\nSe solicita' in recepcion
    assert '\n\n• Verificar la orden.' in recepcion
    assert '\n\n• Confirmar la muestra.' in recepcion
    assert 'centrifuga' not in recepcion


def test_puentes_ocr_opcionales_no_exigen_tesseract_en_windows():
    class MotorSimulado:
        def __call__(self, imagen):
            return [([[0,0],[100,0],[100,10],[0,10]],'Texto reconocido.',0.97)], 0.02
    with patch.object(oe,'_tesseract_disponible',return_value=''):
        texto,motor = oe.leer_pagina_escaneada(None, motor_rapid=MotorSimulado())
    assert motor == 'RapidOCR' and texto == 'Texto reconocido.'
