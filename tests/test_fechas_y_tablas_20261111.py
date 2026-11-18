"""Regresiones visuales y de nombres de descarga sin base de datos institucional."""
import ast
from datetime import datetime
from pathlib import Path
import re
from zoneinfo import ZoneInfo

ROOT = Path(__file__).resolve().parents[1]
CODIGO = (ROOT / 'app.py').read_text(encoding='utf-8')
ARBOL = ast.parse(CODIGO)
FUNCIONES = ('fecha_hoy_nombre_archivo', 'nombre_ficha_tecnica', 'nombre_descarga_excel_generado')
namespace = {'datetime': datetime, 'ZoneInfo': ZoneInfo, 're': re}
exec(compile(ast.Module(body=[n for n in ARBOL.body
    if isinstance(n, ast.FunctionDef) and n.name in FUNCIONES], type_ignores=[]),
    str(ROOT / 'app.py'), 'exec'), namespace)


def test_fecha_archivos_dd_mm_aaaa_y_valida_para_windows():
    fecha = namespace['fecha_hoy_nombre_archivo']()
    assert re.fullmatch(r'\d{2}-\d{2}-\d{4}', fecha)
    datetime.strptime(fecha, '%d-%m-%Y')
    assert '/' not in fecha


def test_nombre_de_ficha_con_area_y_fecha_de_extraccion():
    nombre = namespace['nombre_ficha_tecnica']
    assert nombre('Hematología y coagulación', '11-11-2026') == (
        'Ficha_Técnica_Hematología_y_coagulación_Actualizada_11-11-2026.xlsx')
    assert nombre('Bioquímica', '09-12-2026') == (
        'Ficha_Técnica_Bioquímica_Actualizada_09-12-2026.xlsx')


def test_descarga_nombres_historicos_sin_modificar_archivo_original():
    convertir = namespace['nombre_descarga_excel_generado']
    assert convertir('Ficha_Tecnica_Actualizada_Hematologia_2026-11-11.xlsx') == (
        'Ficha_Técnica_Hematologia_Actualizada_11-11-2026.xlsx')
    assert convertir('Ficha_Técnica_Actualizada_Bioquímica_2026-12-01.xlsx') == (
        'Ficha_Técnica_Bioquímica_Actualizada_01-12-2026.xlsx')
    assert convertir('Ficha_Técnica_Hematología_y_coagulación_Actualizada_11-11-2026.xlsx') == (
        'Ficha_Técnica_Hematología_y_coagulación_Actualizada_11-11-2026.xlsx')
    assert convertir('plantilla_hematologia_2026-11-11.xlsx') == 'plantilla_hematologia_2026-11-11.xlsx'
    assert convertir(None) is None
    assert convertir('Ficha_Tecnica_Actualizada_Hematologia_11-11-2026.xlsx') == (
        'Ficha_Técnica_Hematologia_Actualizada_11-11-2026.xlsx')
    assert convertir('plantilla_hematologia.xlsx', area='Hematología y coagulación',
                    fecha_registro='10/11/2026 15:00') == (
        'Ficha_Técnica_Hematología_y_coagulación_Actualizada_10-11-2026.xlsx')
    assert convertir('Ficha_Tecnica_Actualizada_Hematologia_2026-11-11.xlsx',
                    area='Hematología y coagulación', fecha_registro='10/12/2026 15:00') == (
        'Ficha_Técnica_Hematología_y_coagulación_Actualizada_11-11-2026.xlsx')


def test_fecha_del_nombre_proviene_de_la_extraccion_y_se_reutiliza():
    assert 'st.session_state["fecha_ultima_extraccion_ficha"] = fecha_hoy_nombre_archivo()' in CODIGO
    # Descarga normal, guardado en SQLite y cierre de cola tras omitir el último.
    assert CODIGO.count('st.session_state.get("fecha_ultima_extraccion_ficha")') == 3
    assert 'nombre_ficha_tecnica(' in CODIGO
    assert 'file_name=nombre_descarga_excel_generado(' in CODIGO


def test_modal_usuarios_con_tabla_completa_y_formularios_en_tabs():
    assert '@st.dialog("Administración de usuarios", width="large"' in CODIGO
    assert 'st.dataframe(' in CODIGO
    for contenido in ('crear', 'editar', 'clave', 'estado'):
        assert f'with contenido_{contenido}:' in CODIGO
    assert 'with contenido_centro:' not in CODIGO
    for campo in ('Usuario', 'Nombre visible', 'Rol', 'Estado', 'Creado'):
        assert f'"{campo}": st.column_config.TextColumn("{campo}", width=' in CODIGO


def test_resumen_diferencias_usa_anchos_de_encabezado():
    for titulo in ('Campo', 'Tipo de cambio', 'Valor anterior', 'Valor nuevo'):
        assert f'"{titulo}": st.column_config.TextColumn("{titulo}", width=' in CODIGO
    assert "nombre_generado = nombre_ficha_tecnica(" in CODIGO
