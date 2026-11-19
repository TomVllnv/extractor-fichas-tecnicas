"""Regresiones de nomenclatura, procedencia, fechas y equipamiento."""
from __future__ import annotations

import io
from pathlib import Path
import sys

import fitz

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import referencias_apa as apa
import lenguaje_ocr as lo
import equipamientos as eq
import procedencia_campos as pc
import fuentes_documentales as fd

APP = (ROOT / "app.py").read_text(encoding="utf-8")


def _pdf_primera_pagina(texto=""):
    pdf = fitz.open()
    hoja = pdf.new_page()
    if texto:
        hoja.insert_text((30, 100), texto)
    datos = pdf.tobytes()
    pdf.close()
    return datos


def test_equipo_variantes_del_mismo_modelo_de_cualquier_marca():
    for par in [
        ("COBAS T511", "Cobas T 511"),
        ("Sysmex XN-1000", "SYSMEX XN 1000"),
        ("Architect i2000SR", "ARCHITECT I 2000 SR"),
        ("Atellica IM 1600", "Atellica IM-1600"),
    ]:
        assert eq.equipamientos_equivalentes(*par)
    assert not eq.equipamientos_equivalentes("Cobas T500", "Cobas T511")
    assert not eq.equipamientos_equivalentes("Sysmex XN-1000", "Sysmex XN-2000")


def test_repeticiones_se_quitan_sin_cambiar_nombre_original_ni_orden():
    assert eq.deduplicar_equipamiento("Cobas T511\nCOBAS T 511\nSysmex XN-1000") == (
        "Cobas T511\nSysmex XN-1000")
    assert eq.deduplicar_equipamiento("Atellica IM 1600, ATELLICA IM-1600") == (
        "Atellica IM 1600")
    assert eq.deduplicar_equipamiento("Cobas T511, con 8 posiciones disponibles") == (
        "Cobas T511, con 8 posiciones disponibles")


def test_fechas_del_contenido_y_no_del_nombre():
    pdf = _pdf_primera_pagina("Version 00 - Septiembre 2025")
    ref = apa.fuente_pdf("nombre_2019.pdf", pdf)
    assert ref.anio == "2025"
    assert "(2025)" in apa.referencia_apa(ref)
    pdf_vacio = _pdf_primera_pagina("Texto de laboratorio sin fecha")
    ref_vacia = apa.fuente_pdf("Ac_lupico_2025.pdf", pdf_vacio)
    assert ref_vacia.anio == ""
    assert "(s. f.)" in apa.referencia_apa(ref_vacia)


def test_pdf_escaneado_aprovecha_texto_ya_leido_sin_nuevo_ocr():
    vacio = _pdf_primera_pagina()
    ref = apa.fuente_pdf("registro_2021.pdf", vacio,
                          texto_reconocido="AC LUPICO V00 Septiembre 2025")
    assert ref.anio == "2025"
    ref_errata = apa.fuente_pdf("registro.pdf", vacio,
                          texto_reconocido="Versión 00 Septiembre 2O25")
    assert ref_errata.anio == "2025"


def test_ignora_anios_no_documentales_y_futuros():
    assert apa._anio_explicito("Próxima revisión: septiembre de 2027") == ""
    assert apa._anio_explicito("Hemoglobina 2025, valor 4.5") == ""
    assert apa._anio_explicito("septiembre de 2025") == ""
    assert apa._anio_explicito("Fecha de emisión: 15/09/2025") == "2025"
    assert apa._anio_explicito("Fecha de publicación: 2025-09-15") == "2025"


def test_correccion_lexica_segura_ocr():
    original = "Área de Bi0quimica. Hemat0logía. Lab0ratorio."
    corregido = lo.corregir_texto_lectura_optica(original)
    assert "Bioquímica" in corregido
    assert "Hematología" in corregido
    assert "Laboratorio" in corregido
    protegido = "Cobas T511; Sysmex XN-1000; Código 031007-003; 3600 RPM; 2025"
    assert lo.corregir_texto_lectura_optica(protegido) == protegido
    assert lo.corregir_texto_lectura_optica("Inmun0logia") == "Inmunología"


def test_prioridades_ocr_y_eleccion_usuario():
    fuentes, aplicada = pc.procedencia_campo(
        'Preparación del paciente', 'Ayuno 4 horas',
        'Cartera de Exámenes del CRS · Indicaciones',
        {'Preparación del paciente': 'Ayuno 4 horas'}, {}, set(), True,
        {'Preparación del paciente': 'Ayuno 4 horas'})
    assert fuentes[0] == 'Escáner'
    assert aplicada == 'Escáner'
    fuentes, aplicada = pc.procedencia_campo(
        'Preparación del paciente', 'Ayuno 4 horas',
        'Cartera de Exámenes del CRS · Indicaciones',
        {'Preparación del paciente': 'Ayuno 4 horas'}, {}, set(), True,
        {'Preparación del paciente': 'Ayuno 4 horas'}, None,
        {'Preparación del paciente'})
    assert aplicada == pc.CRS
    fuentes_codigo, _ = pc.procedencia_campo('Código', '123', es_escaneado=True)
    assert fuentes_codigo[0] == pc.CRS
    assert fuentes_codigo[1] == pc.PDF_ESCANEADO
    assert pc.procedencia_campo('Referencias', 'Cita', es_escaneado=True)[0] == [pc.REFERENCIAS]
    assert pc.procedencia_campo('Nombre LIS', 'Examen', es_escaneado=True)[0] == [pc.EXAMEN]
    assert pc.procedencia_campo('Preparación del paciente', 'Ayuno')[0] == list(
        fd.PRIORIDAD_FUENTES_POR_CAMPO['Preparación del paciente'])


def test_leyenda_completa_y_columna_estado_abreviada():
    assert 'Naranja:</span> Campo completado con valor anterior.' in APP
    assert "? 'Valor anterior' : 'Sin valor anterior';" in APP
    assert 'zona_deshacer_estable' in APP
    assert 'st.container(height=62, border=False, key="zona_deshacer_estable")' in APP


def test_etiquetas_sin_tecnicismos_para_usuario():
    assert '"Páginas del documento por analizar con lectura óptica"' in APP
    assert '"Páginas del documento por analizar con OCR"' not in APP
    assert '"Guardar nuevas Normas en SQLite"' not in APP
    assert pc.PDF_ESCANEADO == 'Escáner'


def test_analisis_conjunto_no_cruza_fechas_entre_pdf():
    import ast
    nodo = next(n for n in ast.parse(APP).body
                if isinstance(n, ast.FunctionDef) and
                n.name == '_texto_optico_para_referencia')
    scope = {'re': __import__('re')}
    exec(compile(ast.Module(body=[nodo], type_ignores=[]), 'app.py', 'exec'), scope)
    extraer = scope['_texto_optico_para_referencia']
    union = ('===== primero.pdf =====\nVersión 00 Septiembre 2025\n\n'
             '===== segundo.pdf =====\nFecha de emisión: marzo de 2024')
    assert apa._anio_explicito(extraer(union, 'primero.pdf')) == '2025'
    assert apa._anio_explicito(extraer(union, 'segundo.pdf')) == '2024'
    assert extraer(union, 'tercero.pdf') == ''
