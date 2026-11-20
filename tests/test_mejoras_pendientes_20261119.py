"""Pruebas dirigidas de las mejoras en preparación, aún sin distribuir."""
from pathlib import Path
import ast
import re
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import equipamientos as eq
from lenguaje_ocr import corregir_texto_lectura_optica


APP = Path(__file__).resolve().parents[1].joinpath('app.py').read_text()


def funcion_pura(nombre, globals_extra=None):
    cuerpo = ast.parse(APP)
    fn = next(n for n in cuerpo.body if isinstance(n, ast.FunctionDef) and n.name == nombre)
    tree = ast.fix_missing_locations(ast.Module(body=[fn], type_ignores=[]))
    namespace = {"re": re, "eq": eq}
    namespace.update(globals_extra or {})
    exec(compile(tree, '<funcion-aislada>', 'exec'), namespace)
    return namespace[nombre]


def test_el_equipo_detectado_no_oculta_el_modulo():
    extraer = funcion_pura('infer_name_code')
    assert extraer('Equipo COBAS T511 y Cobas T 511', {})['Equipamiento utilizado'] == 'COBAS T511'
    assert extraer('Equipo COBAS T511 y Cobas T500', {})['Equipamiento utilizado'] == 'COBAS T511\nCobas T500'


def test_deduplicacion_sin_coincidencia_preexistente():
    extraer = funcion_pura('infer_name_code')
    assert extraer('No se informa equipamiento', {'Equipamiento utilizado': 'Cobas T511\nCOBAS T 511'})['Equipamiento utilizado'] == 'Cobas T511'


def test_vinetas_mal_reconocidas():
    texto = '+ Centrifugar la muestra\ne- Mantener la muestra refrigerada'
    salida = corregir_texto_lectura_optica(texto)
    assert salida == '• Centrifugar la muestra\n• Mantener la muestra refrigerada'


def test_caracteres_sensibles_sin_cambios():
    texto = 'Na+\n++\n+5 °C\n+ 5 °C\n+ Suero\nA+B'
    assert corregir_texto_lectura_optica(texto) == texto


def test_segundos_se_recortan_sin_notas():
    extraer = funcion_pura('extraer_solo_valor_referencia')
    assert extraer('15 a 25 s. Interpretar según contexto clínico') == '15 a 25 s'
    assert extraer('24-36 segundos. Estos valores corresponden al fabricante') == '24-36 segundos'
    assert extraer('≤ 12,5 seg según cada laboratorio') == '≤ 12,5 seg'
    assert extraer('No reactivo') == 'No reactivo'


def test_guia_y_tabs_actualizadas():
    assert 'Guía básica para usuario' in APP
    assert 'Guía de desarrollo seguro' not in APP
    assert 'pestanas_principales' in APP
    assert 'dblclick' in APP and 'pointermove' in APP
