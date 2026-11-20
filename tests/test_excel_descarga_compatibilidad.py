"""Prevención de fórmulas accidentales y validación del libro XLSX descargado."""
import ast
import io
import re
import sys
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import referencias_apa as apa
import compatibilidad_excel as ce

CODIGO = (ROOT / 'app.py').read_text(encoding='utf-8')
ARBOL = ast.parse(CODIGO)


def funciones_reales():
    nombres = {
        'buscar_hoja_por_examen', 'apply_to_sheet',
        'aplicar_ediciones_a_excel', 'actualizar_campo_excel',
        'validar_excel_para_descarga',
    }
    seleccion = [n for n in ARBOL.body if isinstance(n, ast.FunctionDef) and n.name in nombres]
    sesion = {'origen_campos_documentales': {}, 'texto_referencias_auto': None,
              'fuentes_referencias_aplicadas': None}
    st = type('StreamlitFalso', (), {'session_state': sesion})()
    entorno = {
        'st': st, 'io': io, 'ET': __import__('xml.etree.ElementTree', fromlist=['ElementTree']),
        'zipfile': __import__('zipfile'), 'pd': pd, 'Alignment': Alignment,
        'ce': ce,
        'load_workbook': load_workbook, 'apa': apa,
        'norm': lambda value: re.sub(r'\W+', ' ', str(value or '').casefold()).strip(),
        'limpiar_valor_excel': lambda value, examen=None: str(value or '').strip(),
        'get_field_aliases': lambda area: ['Nombre LIS', 'Preparación del paciente', 'Valores de referencia'],
        'ajustar_ancho_columnas': lambda hoja: None,
        'ajustar_alto_fila_con_listas': lambda hoja, fila: None,
        'texto_normas_con_negritas': lambda campo, valor: valor,
    }
    exec(compile(ast.fix_missing_locations(ast.Module(body=seleccion, type_ignores=[])),
                 'app.py', 'exec'), entorno)
    return entorno


def libro_inicial():
    libro = Workbook()
    hoja = libro.active
    hoja.title = 'FIBRINOGENO'
    hoja['A1'] = 'Nombre LIS'
    hoja['B1'] = 'FIBRINOGENO'
    hoja['A2'] = 'Preparación del paciente'
    hoja['A3'] = 'Valores de referencia'
    hoja['A4'] = 'Referencias'
    archivo = io.BytesIO()
    libro.save(archivo)
    return archivo.getvalue()


class TestExcelDescargaCompatibilidad(unittest.TestCase):
    def test_edicion_literal_no_crea_formula_invalida(self):
        funciones = funciones_reales()
        valor = '=Texto clínico con signos; no es una fórmula de Excel'
        datos = pd.DataFrame([{'Campo': 'Preparación del paciente', 'Información extraída': valor}])
        salida = funciones['aplicar_ediciones_a_excel'](libro_inicial(), datos, 'FIBRINOGENO')
        celda = load_workbook(io.BytesIO(salida)).active['B2']
        self.assertEqual(celda.value, valor)
        self.assertEqual(celda.data_type, 's')
        self.assertIs(salida, funciones['validar_excel_para_descarga'](salida))

    def test_actualizacion_literal_protege_valor_igual(self):
        funciones = funciones_reales()
        salida = funciones['actualizar_campo_excel'](
            libro_inicial(), 'FIBRINOGENO', 'Valores de referencia', '=15 mg/dL')
        celda = load_workbook(io.BytesIO(salida)).active['B3']
        self.assertEqual(celda.value, '=15 mg/dL')
        self.assertEqual(celda.data_type, 's')

    def test_aplicacion_automatica_no_crea_formula(self):
        funciones = funciones_reales()
        libro = Workbook()
        hoja = libro.active
        hoja['A1'] = 'Preparación del paciente'
        funciones['apply_to_sheet'](hoja, {'Preparación del paciente': '=Ayuno según indicación'})
        self.assertEqual(hoja['B1'].data_type, 's')
        self.assertEqual(hoja['B1'].value, '=Ayuno según indicación')

    def test_sin_ediciones_conserva_bytes_sin_serializaciones(self):
        funciones = funciones_reales()
        original = libro_inicial()
        ediciones = pd.DataFrame([{'Campo': 'Nombre LIS', 'Información extraída': 'FIBRINOGENO'}])
        final = funciones['aplicar_ediciones_a_excel'](original, ediciones, 'FIBRINOGENO')
        self.assertEqual(final, original)

    def test_rechaza_paquete_zip_corrupto(self):
        funciones = funciones_reales()
        with self.assertRaises(ValueError):
            funciones['validar_excel_para_descarga'](b'Esto no es un Excel')

    def test_rechaza_xml_interno_corrupto(self):
        funciones = funciones_reales()
        import zipfile
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w') as z:
            z.writestr('[Content_Types].xml', '<Types')
            z.writestr('_rels/.rels', '<Relationships/>')
            z.writestr('xl/workbook.xml', '<workbook/>')
        with self.assertRaises(ValueError):
            funciones['validar_excel_para_descarga'](buf.getvalue())


if __name__ == '__main__':
    unittest.main()
