"""Casillas por campo, restauración de Excel y diálogos sin datos institucionales."""
import ast
import copy
import io
from pathlib import Path
import re
import unittest

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Alignment

ROOT = Path(__file__).resolve().parents[1]
CODIGO = (ROOT / 'app.py').read_text(encoding='utf-8')
ARBOL = ast.parse(CODIGO)


def entorno_test():
    nombres = {'respetar_campos_no_seleccionados', 'campo_marcado_para_actualizar',
               'buscar_hoja_por_examen', 'progreso_pendiente_cambio_area',
               'aplicar_ediciones_a_excel', 'norm',
               'ajustar_ancho_columnas', 'ajustar_alto_fila_con_listas',
               '_guardar_resolucion_en_ficha'}
    funciones = [n for n in ARBOL.body if isinstance(n, ast.FunctionDef) and n.name in nombres]
    env = {
        'load_workbook': load_workbook, 'io': io, 'pd': pd, 'copy': copy, 're': re,
        'st': type('S', (), {'session_state': {}})(),
        'Alignment': Alignment,
        'texto_normas_con_negritas': lambda campo, valor: str(valor),
        'get_field_aliases': lambda area: ['Nombre LIS', 'Equipamiento utilizado', 'Fecha actualización'],
    }
    exec(compile(ast.Module(body=funciones, type_ignores=[]), str(ROOT/'app.py'), 'exec'), env)
    return env


def libro(nombre, campos, estilos=False):
    wb = Workbook()
    ws = wb.active
    ws.title = nombre
    for fila, (campo, valor) in enumerate(campos, start=1):
        ws.cell(fila, 1, campo)
        ws.cell(fila, 2, valor)
    if estilos:
        ws['B2'].fill = PatternFill('solid', fgColor='FFFF00')
        ws['B2'].alignment = Alignment(wrap_text=True)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


class TestCasillasExcel(unittest.TestCase):
    def setUp(self):
        self.funcs = entorno_test()
        self.anterior = libro('Hemograma', [
            ('Nombre LIS', 'Hemograma'), ('Preparación del paciente', 'Ayuno anterior'),
            ('Procesamiento de muestras', 'Método antiguo'),
            ('Referencias', 'Referencias anteriores')], estilos=True)
        self.nuevo = libro('Hemograma', [
            ('Nombre LIS', 'Hemograma'), ('Preparación del paciente', 'Nueva preparación'),
            ('Procesamiento de muestras', 'Método nuevo'),
            ('Referencias', 'Referencias nuevas')])

    def test_todo_marcado_conserva_excel_nuevo_sin_reescritura(self):
        resultado = self.funcs['respetar_campos_no_seleccionados'](
            self.nuevo, self.anterior, 'Hemograma', {'Preparación del paciente': True})
        self.assertIs(resultado, self.nuevo)

    def test_desmarcar_una_fila_restaurar_valor_y_formato_sin_cambiar_otras(self):
        f = self.funcs['respetar_campos_no_seleccionados']
        salida = f(self.nuevo, self.anterior, 'Hemograma', {
            'Preparación del paciente': False,
            'Procesamiento de muestras': True,
            'Referencias': True})
        ws = load_workbook(io.BytesIO(salida)).active
        self.assertEqual(ws['B2'].value, 'Ayuno anterior')
        self.assertEqual(ws['B2'].fill.fgColor.rgb, '00FFFF00')
        self.assertEqual(ws['B3'].value, 'Método nuevo')
        self.assertEqual(ws['B4'].value, 'Referencias nuevas')
        # Volver a marcar la casilla reconstruye siempre desde el candidato no modificado.
        salida2 = f(self.nuevo, self.anterior, 'Hemograma', {'Preparación del paciente': True})
        self.assertEqual(load_workbook(io.BytesIO(salida2)).active['B2'].value,
                         'Nueva preparación')

    def test_si_el_examen_es_nuevo_el_campo_desmarcado_queda_vacio(self):
        anterior = libro('Otra ficha', [('Campo', 'Antiguo')])
        salida = self.funcs['respetar_campos_no_seleccionados'](
            self.nuevo, anterior, 'Hemograma', {'Referencias': False})
        self.assertIsNone(load_workbook(io.BytesIO(salida)).active['B4'].value)

    def test_sin_plantilla_inicial_un_campo_desmarcado_no_se_reemplaza(self):
        salida = self.funcs['respetar_campos_no_seleccionados'](
            self.nuevo, None, 'Hemograma', {'Referencias': False})
        self.assertIsNone(load_workbook(io.BytesIO(salida)).active['B4'].value)


    def test_aplicar_y_restaurar_funciona_con_etiquetas_con_dos_puntos(self):
        previo = libro('Hemograma', [
            ('Autorizado por:', 'Encargado anterior'),
            ('Fecha actualización:', '01/01/2026')])
        editado = pd.DataFrame([
            {'Campo': 'Autorizado por', 'Información extraída': '', 'Estado': 'Pendiente de ingreso'},
            {'Campo': 'Fecha actualización', 'Información extraída': '10/11/2026', 'Estado': 'Modificado por usuario'},
        ])
        salida = self.funcs['aplicar_ediciones_a_excel'](previo, editado, 'Hemograma')
        ws = load_workbook(io.BytesIO(salida)).active
        self.assertIn(ws['B1'].value, ('', None))
        self.assertEqual(ws['B2'].value, '10/11/2026')
        restaurado = self.funcs['respetar_campos_no_seleccionados'](
            salida, previo, 'Hemograma', {'Autorizado por': True, 'Fecha actualización': False})
        ws2 = load_workbook(io.BytesIO(restaurado)).active
        self.assertIn(ws2['B1'].value, ('', None))
        self.assertEqual(ws2['B2'].value, '01/01/2026')

    def test_discrepancia_validada_reemplaza_equipamiento_ya_extraido(self):
        anterior = libro('FIBRINOGENO', [
            ('Nombre LIS', 'FIBRINOGENO'),
            ('Equipamiento utilizado:', 'cobas t 511\ncobas t 711')])
        revision = pd.DataFrame([{
            'Campo': 'Equipamiento utilizado',
            'Información extraída': 'cobas t 511',
            'Estado': 'Encontrado',
        }])
        aplicar = self.funcs['aplicar_ediciones_a_excel']
        # El cambio normal no debe aplanar datos enriquecidos por accidente.
        sin_confirmar = aplicar(anterior, revision, 'FIBRINOGENO')
        self.assertEqual(load_workbook(io.BytesIO(sin_confirmar)).active['B2'].value,
                         'cobas t 511\ncobas t 711')
        # Una decisión explícita del modal debe actualizar el campo.
        con_confirmacion = aplicar(anterior, revision, 'FIBRINOGENO',
                                  campos_forzar={'Equipamiento utilizado'})
        hoja = load_workbook(io.BytesIO(con_confirmacion)).active
        self.assertEqual(hoja['B2'].value, 'cobas t 511')
        self.assertTrue(hoja['B2'].alignment.wrap_text)
        self.assertGreater(hoja.column_dimensions['B'].width, 12)
        self.assertIsNotNone(hoja.row_dimensions[2].height)
        self.assertIn('campos_forzar=campos_resueltos', CODIGO)

    def test_modal_guarda_en_excel_la_eleccion_de_un_solo_equipo(self):
        estado = self.funcs['st'].session_state
        estado.clear()
        estado.update({
            'result_bytes': libro('FIBRINOGENO', [
                ('Nombre LIS', 'FIBRINOGENO'),
                ('Equipamiento utilizado:', 'cobas t 511\ncobas t 711'),
                ('Fecha actualización:', '10/11/2026')]),
            'examen_activo': 'FIBRINOGENO',
            'area_resultado': 'Hematología y coagulación',
            'user_edited_fields': set(),
        })
        datos = {'Nombre LIS': 'FIBRINOGENO', 'Equipamiento utilizado': 'cobas t 511',
                 'Fecha actualización': '10/11/2026'}
        self.funcs['_guardar_resolucion_en_ficha'](datos, {'Equipamiento utilizado'})
        hoja = load_workbook(io.BytesIO(estado['result_bytes'])).active
        self.assertEqual(hoja['B2'].value, 'cobas t 511')
        self.assertEqual(estado['data_extracted']['Equipamiento utilizado'], 'cobas t 511')

    def test_autoajuste_python_recalcula_filas_y_columnas_en_excel(self):
        anterior = libro('FIBRINOGENO', [
            ('Nombre LIS', 'FIBRINOGENO'),
            ('Procesamiento de muestras:', 'Texto breve'),
            ('Equipamiento utilizado:', 'cobas t 511\ncobas t 711')])
        fila_larga = ('Las muestras se procesan siguiendo criterios institucionales '
                      'y se verifican los datos de la solicitud. ') * 9
        revision = pd.DataFrame([{
            'Campo': 'Procesamiento de muestras',
            'Información extraída': fila_larga,
            'Estado': 'Modificado por usuario',
        }])
        salida = self.funcs['aplicar_ediciones_a_excel'](anterior, revision, 'FIBRINOGENO')
        hoja = load_workbook(io.BytesIO(salida)).active
        self.assertEqual(hoja['B2'].value, fila_larga.strip())
        self.assertGreater(hoja.column_dimensions['B'].width, 12)
        self.assertGreater(hoja.row_dimensions[2].height, hoja.row_dimensions[1].height)
        self.assertTrue(hoja['B2'].alignment.wrap_text)

    def test_encabezados_respetan_alineacion_por_columna(self):
        self.assertIn("headerClass='ag-header-left'", CODIGO)
        self.assertEqual(CODIGO.count("headerClass='ag-header-left'"), 2)
        self.assertGreaterEqual(CODIGO.count("headerClass='ag-header-centered'"), 2)
        self.assertIn('".ag-header-left .ag-header-cell-label"', CODIGO)


    def test_checkbox_false_tambien_si_llega_como_cadena(self):
        f = self.funcs['campo_marcado_para_actualizar']
        self.assertFalse(f('false'))
        self.assertFalse(f(False))
        self.assertTrue(f(True))
        self.assertTrue(f('true'))

    def test_codigo_usa_misma_version_en_descarga_sqlite_y_siguiente_examen(self):
        self.assertIn('"excel_para_actualizar"] = validar_excel_para_descarga(', CODIGO)
        self.assertIn('st.session_state["result_bytes"] = excel_con_confirmaciones', CODIGO)
        self.assertIn('respetar_campos_no_seleccionados(', CODIGO)
        self.assertIn('cellRenderer=casilla_actualizacion', CODIGO)
        self.assertIn('modal_confirmar_cambio_area()', CODIGO)

    def test_autoajuste_python_sin_macros_y_archivos_unificados(self):
        self.assertFalse((ROOT/'AutoajustarCampos.bas').exists())
        self.assertNotIn('AutoajustarCampos.bas', CODIGO)
        self.assertIn('ajustar_ancho_columnas(ws)', CODIGO)
        self.assertIn('ajustar_alto_fila_con_listas(ws, fila)', CODIGO)
        self.assertFalse((ROOT/'AutoajustarColumnas.bas').exists())
        for nombre in ('GUIA_TRAZABILIDAD.txt', 'INSTRUCCIONES_ACTUALIZACION.txt',
                       'LEEME_PRIMERO_NO_REEMPLAZAR_SQLITE.txt'):
            self.assertFalse((ROOT/nombre).exists())
        self.assertTrue((ROOT/'Updates.txt').exists())


if __name__ == '__main__':
    unittest.main()
