"""Regresiones de lectura óptica (viñetas) e información clínica de manuales.

Solo se comprueba información presente en el texto; no hay reconstrucción clínica
por conocimiento externo ni necesidad de base de datos institucional.
"""
import io
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import lenguaje_ocr as lectura
import manuales_procedimiento as manual


def test_e_guion_al_inicio_de_renglon_se_convierte_en_vineta():
    original = ("2. RECEPCIÓN DE MUESTRAS\n"
                "Las muestras deben identificarse correctamente.\n"
                "e- Verificar la identidad del paciente al recibir la muestra.\n"
                "e - Comprobar la concordancia con la orden del examen.\n"
                "3. PREPARACIÓN DE REACTIVOS\n"
                "El reactivo se almacena en su envase.")
    corregido = lectura.corregir_texto_lectura_optica(original)
    assert "• Verificar la identidad" in corregido
    assert "• Comprobar la concordancia" in corregido
    assert not any(x in corregido for x in ("e- Verificar", "e - Comprobar"))
    resultado = manual.extraer_campos_manual(corregido)
    assert resultado['Recepción de muestras'].count('• ') == 2
    assert 'reactivo se almacena' not in resultado['Recepción de muestras']


def test_vineta_reconocida_en_linea_aislada_se_une_con_su_texto():
    text = ("1. RECEPCIÓN DE MUESTRAS\n"
            "El personal comprueba la identificación del tubo de muestra.\n"
            "e-\nVerificar todos los datos identificatorios antes de recibir la muestra.\n"
            "2. REACTIVOS\nSolo se incluyen datos de reactivos.")
    corregido = lectura.corregir_texto_lectura_optica(text)
    valor = manual.extraer_campos_manual(corregido)['Recepción de muestras']
    assert '• Verificar todos los datos' in valor
    assert '\ne-' not in valor


def test_e_guion_legitimo_y_vinetas_existentes_se_respetan():
    texto = ('Se envía por e-mail y se etiqueta como e-123.\n'
             'e-coli es la forma escrita del ejemplo.\n'
             'e-mail correo electrónico institucional\n'
             '• No cambiar una viñeta existente.\n'
             'e-12 código específico\n'
             'El valor e- es parte de la oración.\n')
    salida = lectura.corregir_texto_lectura_optica(texto)
    assert salida == texto


def test_caracteristicas_es_homologo_clinico_sin_incluir_calibracion():
    texto = '''1. CARACTERÍSTICAS
La transferrina es una proteína transportadora de hierro que refleja su disponibilidad en el organismo.
El analizador se calibra mediante fotometría y utiliza un reactivo.
2. PREPARACIÓN DE REACTIVOS
El reactivo se conserva durante 4 horas a 2-8 grados.'''
    reporte = manual.analizar_manual(texto)
    clinica = reporte['campos']['Información clínica']
    assert 'transferrina es una proteína' in clinica
    assert 'analizador' not in clinica
    assert 'reactivo' not in clinica
    assert reporte['fuentes_homologas']['Información clínica'] == 'Características'
    assert 'Estabilidad de la muestra' not in reporte['campos']


def test_caracteristicas_inline_acepta_descripcion_de_la_prueba():
    texto = '''CARACTERÍSTICAS: El examen cuantifica la concentración de ferritina en sangre para evaluar la reserva de hierro.
MATERIALES Y REACTIVOS
Los calibradores deben conservarse refrigerados.'''
    clinica = manual.extraer_campos_manual(texto)['Información clínica']
    assert 'ferritina en sangre' in clinica
    assert 'calibradores' not in clinica


def test_informacion_clinica_especifica_gana_a_caracteristicas_sin_importar_orden():
    for partes in [(
            '''1. CARACTERÍSTICAS
La prueba mide la ferritina en sangre para conocer sus niveles.''',
            '''2. INFORMACIÓN CLÍNICA
La ferritina permite evaluar las reservas de hierro.'''),
            (
            '''1. INFORMACIÓN CLÍNICA
La ferritina permite evaluar las reservas de hierro.''',
            '''2. CARACTERÍSTICAS
La prueba mide la ferritina en sangre para conocer sus niveles.'''),
    ]:
        clinica = manual.extraer_campos_manual('\n'.join(partes))['Información clínica']
        assert clinica == 'La ferritina permite evaluar las reservas de hierro.'


def test_introduccion_descriptiva_rellena_cuando_caracteristicas_solo_tecnica():
    texto = '''1. INTRODUCCIÓN
La ferritina sérica se correlaciona con las reservas de hierro del organismo.
2. CARACTERÍSTICAS
El analizador utiliza una técnica de fotometría con calibradores.
3. MATERIALES Y REACTIVOS
Conservar los reactivos refrigerados.'''
    reporte = manual.analizar_manual(texto)
    assert reporte['campos']['Información clínica'] == (
        'La ferritina sérica se correlaciona con las reservas de hierro del organismo.')
    assert reporte['fuentes_homologas']['Información clínica'] == 'Introducción'


def test_sin_verdadera_descripcion_no_inventa_informacion_clinica():
    texto = '''1. CARACTERÍSTICAS
El equipo utiliza un método de espectrofotometría con reactivos y calibradores.
2. CARACTERÍSTICAS ANALÍTICAS
El coeficiente de variación es inferior al 5 %.
3. PROCEDIMIENTO
Se deben pipetear los reactivos con precaución.'''
    datos = manual.extraer_campos_manual(texto)
    assert 'Información clínica' not in datos


def test_caracteristicas_clinicas_con_vinetas_llegan_al_excel():
    import test_manuales_procedimiento_20261116 as soporte
    texto_ocr = '''1. CARACTERÍSTICAS
Este examen permite evaluar la función tiroidea en pacientes con sospecha de alteraciones.
e- El análisis se utiliza para el seguimiento de enfermedad tiroidea.
e- Contribuye a la evaluación de pacientes con trastornos hormonales.
2. PREPARACIÓN DE REACTIVOS
Los reactivos requieren enfriamiento.'''
    texto_corregido = lectura.corregir_texto_lectura_optica(texto_ocr)
    datos = manual.extraer_campos_manual(texto_corregido)
    assert datos['Información clínica'].count('• ') == 2
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'EXAMEN PRUEBA'
    ws.append(['Información clínica', None])
    ws.append(['Nombre LIS', None])
    buf = io.BytesIO(); wb.save(buf)
    ctx = soporte._contexto()
    result, propuestos, *_ = ctx['process'](
        b'pdf', 'procedimiento.pdf', buf.getvalue(), nombre_examen='EXAMEN PRUEBA',
        datos_pdf_preextraidos=(datos, texto_corregido, 'Tesseract'),
        manual_procedimiento=True)
    valor = openpyxl.load_workbook(io.BytesIO(result), data_only=True).active['B1'].value
    assert valor == propuestos['Información clínica']
    assert valor.count('• ') == 2
    assert 'e-' not in valor
    assert 'reactivos' not in valor.lower()
