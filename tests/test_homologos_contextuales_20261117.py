"""Cobertura de homólogos clínicos en manuales digitales y escaneados.

Se evita alucinar datos: cada propuesta debe ser una frase literal presente en
el propio documento, sin trasvasar contenido de reactivos/bibliografía.
"""
import ast
import io
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import manuales_procedimiento as mp
import procedencia_campos as pc


INTRODUCCION = '''1. INTRODUCCIÓN
La proteína C reactiva es una proteína de fase aguda cuya concentración aumenta en respuesta a procesos inflamatorios. Su determinación se utiliza como apoyo en la evaluación y seguimiento de enfermedades inflamatorias e infecciosas. La medición se realiza mediante un método inmunoturbidimétrico automatizado.
2. PREPARACIÓN DE REACTIVOS
El reactivo se conserva durante 12 horas a 2-8 °C.
3. REFERENCIAS BIBLIOGRÁFICAS
En otra publicación se describen valores de referencia de 2021.'''

MEZCLA = '''PROCEDIMIENTO PARA LA DETERMINACIÓN DE PROTEÍNA C REACTIVA
1. DESCRIPCIÓN GENERAL
Las muestras de sangre deberán recolectarse en tubos con citrato de sodio. Posteriormente, se centrifugarán para obtener plasma pobre en plaquetas. El plasma obtenido puede conservarse durante 4 horas a temperatura ambiente, según las condiciones indicadas para este procedimiento.
2. PREPARACIÓN DE REACTIVOS
Los reactivos se mantienen refrigerados a 2-8 °C por 2 días.
3. PROCEDIMIENTO
Se realiza el análisis con un reactivo listo para usar.'''


def test_introduccion_no_copia_metodo_ni_reactivos_a_informacion_clinica():
    r = mp.analizar_manual(INTRODUCCION)
    campos = r['campos']
    assert campos['Información clínica'].count('inflamatorias') == 1
    assert 'concentración aumenta' in campos['Información clínica']
    assert 'método inmunoturbidimétrico' not in campos['Información clínica']
    assert campos['Método de procesamiento'] == 'La medición se realiza mediante un método inmunoturbidimétrico automatizado.'
    assert 'Estabilidad de la muestra' not in campos
    assert 'Valores de referencia' not in campos
    assert r['fuentes_homologas']['Información clínica'] == 'Introducción'
    assert r['fuentes_homologas']['Método de procesamiento'] == 'Introducción'


def test_divide_un_parrafo_en_tres_campos_sin_arrastrar_reactivos():
    r = mp.analizar_manual(MEZCLA)
    campos = r['campos']
    assert 'tubos con citrato' in campos['Muestra requerida y recipiente']
    assert 'centrifugarán' in campos['Procesamiento de muestras']
    assert '4 horas a temperatura ambiente' in campos['Estabilidad de la muestra']
    for campo in ('Muestra requerida y recipiente', 'Procesamiento de muestras', 'Estabilidad de la muestra'):
        assert 'reactivos se mantienen' not in campos[campo]
        assert r['fuentes_homologas'][campo] == 'Descripción general'


def test_fundamento_clinico_se_separa_de_metodo():
    t = '''1. FUNDAMENTO
La determinación se utiliza para apoyar el diagnóstico de infecciones. El análisis se efectúa mediante inmunoturbidimetría en un sistema automatizado.
2. CONTROL DE CALIDAD
Los reactivos se mantienen refrigerados.'''
    v = mp.analizar_manual(t)
    assert 'diagnóstico de infecciones' in v['campos']['Información clínica']
    assert 'diagnóstico de infecciones' not in v['campos']['Método de procesamiento']
    assert 'inmunoturbidimetría' in v['campos']['Método de procesamiento']
    assert v['fuentes_homologas']['Información clínica'] == 'Fundamento'


def test_seccion_explicita_gana_sobre_introduccion():
    t = '''1. INTRODUCCIÓN
Este examen se utiliza para el seguimiento de enfermedades infecciosas.
2. INFORMACIÓN CLÍNICA
Se indica para apoyar el diagnóstico de enfermedades inflamatorias.
3. FUNDAMENTO
La detección se realiza por inmunoturbidimetría.'''
    v = mp.extraer_campos_manual(t)
    assert 'enfermedades inflamatorias' in v['Información clínica']
    assert 'enfermedades infecciosas' not in v['Información clínica']


def test_documento_dudoso_y_reactivos_sin_rellenos_inventados():
    t = '''1. INTRODUCCIÓN
El reactivo se conserva durante 4 horas a temperatura ambiente. Se centrifuga el control durante 10 minutos.
2. PREPARACIÓN DE REACTIVOS
Los tubos para reactivos se preparan con agua destilada.
3. BIBLIOGRAFÍA
Las muestras se centrifugan 10 min en una investigación distinta.'''
    v = mp.extraer_campos_manual(t)
    assert not any(k in v for k in ('Información clínica', 'Estabilidad de la muestra',
                                   'Procesamiento de muestras', 'Muestra requerida y recipiente'))


def test_identifica_manual_digital_pero_no_un_inserto_comercial():
    assert mp.es_manual_de_procedimiento_textual(MEZCLA)
    assert not mp.es_manual_de_procedimiento_textual('''REACTIVO COBAS
Método analítico
La medición se realiza mediante inmunoturbidimetría.
Conservación 2 a 8 °C. Datos del fabricante.''')


def test_fallback_manual_digital_rellena_solo_vacios_y_cinco_campos():
    anteriores = {'Información clínica': 'Dato explícito revisado',
                  'Método de procesamiento': '', 'Código': 'CRS-001'}
    v = mp.completar_homologos_manual_digital(MEZCLA, anteriores)
    assert v['Información clínica'] == 'Dato explícito revisado'
    assert v['Código'] == 'CRS-001'
    assert 'Muestra requerida y recipiente' in v
    assert 'Procesamiento de muestras' in v
    assert 'Estabilidad de la muestra' in v
    assert 'Recepción de muestras' not in v
    assert anteriores == {'Información clínica': 'Dato explícito revisado', 'Método de procesamiento': '', 'Código': 'CRS-001'}
    assert mp.completar_homologos_manual_digital(INTRODUCCION, {}) == {}


def test_procedencia_digital_manual_no_cambia_prioridad_institucional():
    fuentes, _ = pc.procedencia_campo('Muestra requerida y recipiente', 'Suero', manual_digital=True)
    assert fuentes == ['Cartera de acreditación', pc.PDF_TECNICO]
    fuentes, _ = pc.procedencia_campo('Procesamiento de muestras', 'Centrifugación', manual_digital=True)
    assert fuentes == [pc.PDF_TECNICO]
    fuentes, _ = pc.procedencia_campo('Preparación del paciente', 'Ayuno', manual_digital=True)
    assert pc.PDF_TECNICO not in fuentes
    fuentes, _ = pc.procedencia_campo('Muestra requerida y recipiente', 'Suero')
    assert fuentes == ['Cartera de acreditación']


def test_integracion_extraccion_digital_manual_sin_lectura_optica():
    import test_manuales_procedimiento_20261116 as previo
    espacio = previo._contexto()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'EXAMEN DE PRUEBA'
    campos = ('Información clínica', 'Muestra requerida y recipiente',
              'Procesamiento de muestras', 'Estabilidad de la muestra',
              'Método de procesamiento', 'Preparación del paciente', 'Nombre LIS')
    for campo in campos:
        ws.append([campo, None])
    b = io.BytesIO(); wb.save(b)
    datos = mp.completar_homologos_manual_digital(MEZCLA, {})
    out, result, *_ = espacio['process'](
        b'pdf', 'manual_digital.pdf', b.getvalue(),
        nombre_examen='EXAMEN DE PRUEBA',
        datos_pdf_preextraidos=(datos, MEZCLA, 'PyMuPDF'),
        datos_sobrescribir={'Muestra requerida y recipiente': 'Muestra desde Cartera'},
        manual_procedimiento=False)
    libro = openpyxl.load_workbook(io.BytesIO(out), data_only=True)
    valores = {r[0].value: r[1].value for r in libro.active.iter_rows() if r[0].value}
    assert valores['Muestra requerida y recipiente'] == 'Muestra desde Cartera'
    assert 'centrifugarán' in valores['Procesamiento de muestras']
    assert '4 horas' in valores['Estabilidad de la muestra']
    assert 'Preparación del paciente' not in result


def test_introduccion_con_mayuscula_inicial_es_encabezado_independiente():
    texto = '''Introducción
Esta determinación se utiliza para el seguimiento de infecciones en pacientes.
Materiales y reactivos
El reactivo debe mantenerse refrigerado durante 2 horas.'''
    campos = mp.extraer_campos_manual(texto)
    assert 'seguimiento de infecciones' in campos['Información clínica']
    assert 'reactivo' not in campos['Información clínica']
    assert 'Estabilidad de la muestra' not in campos


def test_mencion_de_otro_manual_en_ficha_tecnica_no_habilita_digital():
    texto = '''FICHA DEL FABRICANTE
Información sobre la prueba analítica y las instrucciones de uso.
Consulte el manual de procedimientos de su laboratorio para más detalles.
Método de procesamiento
La determinación se realiza mediante inmunoturbidimetría.'''
    assert not mp.es_manual_de_procedimiento_textual(texto)


def test_encabezados_clinicos_normales_e_introduccion_en_misma_linea():
    texto = '''Introducción: Esta determinación se utiliza para la evaluación de enfermedades infecciosas.
Información clínica
El ensayo permite el seguimiento de procesos inflamatorios.
Fundamento
El método se basa en inmunoturbidimetría para medir el analito.'''
    datos = mp.extraer_campos_manual(texto)
    assert datos['Información clínica'] == 'El ensayo permite el seguimiento de procesos inflamatorios.'
    assert 'inmunoturbidimetría' in datos['Método de procesamiento']
    assert 'enfermedades infecciosas' not in datos['Información clínica']
