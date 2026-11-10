"""Pruebas de seguridad al importar Hematología y preservar viñetas en XLSX."""
import ast
import io
import re
import sys
import unicodedata
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import fuentes_documentales as fd

TREE = ast.parse((ROOT / 'app.py').read_text(encoding='utf-8'))


def _funciones():
    names = {'conservar_fichas_al_reconstruir_hematologia', 'ajustar_alto_fila_con_listas',
             'buscar_hoja_por_examen', 'aplicar_ediciones_a_excel',
             'texto_normas_con_negritas', 'aplicar_formato_normas_excel'}
    nodes = [node for node in TREE.body if isinstance(node, ast.FunctionDef) and node.name in names]
    def norm(text):
        text = unicodedata.normalize('NFD', str(text or '').casefold())
        text = ''.join(char for char in text if not unicodedata.combining(char))
        return re.sub(r'[^a-z0-9]+', ' ', text).strip()
    st = type('MockStreamlit', (), {'session_state': {
        'origen_campos_documentales': {'Preparación del paciente': 'Normas Generales de Toma de Muestras'},
        'texto_referencias_auto': None, 'fuentes_referencias_aplicadas': None,
    }})()
    env = {'io': io, 're': re, 'pd': pd, 'fd': fd, 'Alignment': Alignment,
           'load_workbook': load_workbook, 'norm': norm, 'st': st,
           'apa': type('MockAPA', (), {'referencias_excel_con_cursiva': lambda x: str(x)})()}
    exec(compile(ast.fix_missing_locations(ast.Module(body=nodes, type_ignores=[])),
                 'app.py', 'exec'), env)
    return env


def _bytes(wb):
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


class TestPlantillaReconstruidaYVinetas(unittest.TestCase):
    def test_reemplazo_sin_perder_campos_ni_examenes_agregados(self):
        fun = _funciones()
        nuevo = Workbook(); n = nuevo.active; n.title = 'FIBRINOGENO'
        n['A1'], n['B1'] = 'Código', '0301021-001'
        n['A2'], n['B2'] = 'Nombre LIS', 'FIBRINOGENO'
        n['A3'] = 'Preparación del paciente'; n['A4'] = 'Rotulacion'
        viejo = Workbook(); v = viejo.active; v.title = 'FIBRINOGENO'
        v['A1'], v['B1'] = 'Código', '0301021-001'
        v['A2'], v['B2'] = 'Nombre LIS', 'FIBRINOGENO'
        v['A3'], v['B3'] = 'Preparación del paciente', 'Antes de la toma\n  • Uno\n  • Dos'
        v['A4'], v['B4'] = 'Rotulación', 'Datos mínimos\n  • Identificación'
        extra = viejo.create_sheet('EXAMEN NUEVO'); extra['A1'], extra['B1'] = 'Nombre LIS', 'EXAMEN NUEVO'
        final = load_workbook(io.BytesIO(fun['conservar_fichas_al_reconstruir_hematologia'](
            _bytes(nuevo), _bytes(viejo))), rich_text=True)
        self.assertEqual(str(final['FIBRINOGENO']['B3'].value).count('•'), 2)
        self.assertEqual(str(final['FIBRINOGENO']['B4'].value).count('•'), 1)
        self.assertIn('EXAMEN NUEVO', final.sheetnames)

    def test_tabla_sin_editar_no_transforma_vinetas_en_prosa(self):
        fun = _funciones()
        wb = Workbook(); sh = wb.active; sh.title = 'FIBRINOGENO'
        sh['A1'] = 'Preparación del paciente'
        sh['B1'] = 'Antes de la toma\n  • Primera\n  • Segunda\nDocumentación requerida\n  • Tercera'
        pre = fun['aplicar_formato_normas_excel'](
            _bytes(wb), 'FIBRINOGENO',
            {'Preparación del paciente': 'Normas Generales de Toma de Muestras'})
        vista = pd.DataFrame([{'Campo': 'Preparación del paciente',
                               'Información extraída': 'Antes de la toma • Primera • Segunda • Tercera',
                               'Estado': 'Encontrado'}])
        result = fun['aplicar_ediciones_a_excel'](pre, vista, 'FIBRINOGENO')
        self.assertEqual(pre, result)
        celda = load_workbook(io.BytesIO(result), rich_text=True).active['B1']
        self.assertIn('\n  • Primera', str(celda.value))
        self.assertIn('\n  • Segunda', str(celda.value))
        self.assertIn('\n  • Tercera', str(celda.value))
        self.assertIsInstance(celda.value, str)

    def test_edicion_manual_sigue_actualizando_lista(self):
        fun = _funciones()
        wb = Workbook(); sh = wb.active; sh.title = 'FIBRINOGENO'
        sh['A1'], sh['B1'] = 'Preparación del paciente', 'Texto antiguo'
        datos = pd.DataFrame([{'Campo': 'Preparación del paciente',
                               'Información extraída': 'Antes de la toma\n  • Instrucción nueva',
                               'Estado': 'Modificado por usuario'}])
        final = fun['aplicar_ediciones_a_excel'](_bytes(wb), datos, 'FIBRINOGENO')
        self.assertIn('\n  • Instrucción nueva', str(load_workbook(io.BytesIO(final)).active['B1'].value))

if __name__ == '__main__':
    unittest.main()
