"""Regresiones aisladas: portada/copyright bibliográfico y alturas de Excel."""
import ast
import io
import re
from pathlib import Path
import unittest

import fitz
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment

import referencias_apa as apa

APP = (Path(__file__).resolve().parents[1] / 'app.py').read_text(encoding='utf-8')


def _ajustador_real():
    arbol = ast.parse(APP)
    nodo = next(n for n in arbol.body if isinstance(n, ast.FunctionDef)
                and n.name == 'ajustar_alto_fila_con_listas')
    espacio = {'Alignment': Alignment, 're': re}
    exec(compile(ast.fix_missing_locations(ast.Module(body=[nodo], type_ignores=[])),
                 'app.py', 'exec'), espacio)
    return espacio['ajustar_alto_fila_con_listas']


def _pdf_fabricado(titulo='Fibrinogen', roche=True):
    doc = fitz.open()
    primera = doc.new_page(width=612, height=792)
    primera.insert_text((36, 35), '071534735004.0', fontsize=8)
    primera.insert_text((36, 90), titulo, fontsize=29)
    primera.insert_text((470, 95), 'cobas', fontsize=30)
    primera.insert_text((36, 155), 'Información del sistema', fontsize=12)
    primera.insert_text((36, 190), 'ID del sistema 07 2005 9', fontsize=10)
    ultima = doc.new_page(width=612, height=792)
    ultima.insert_text((35, 700), 'Información adicional sobre este documento.', fontsize=10)
    if roche:
        ultima.insert_text((35, 715), '2024, Roche Diagnostics', fontsize=10)
        ultima.insert_text((35, 730), 'Roche Diagnostics GmbH, Sandhofer Strasse 116, D-68305 Mannheim', fontsize=10)
        ultima.insert_text((35, 745), 'www.roche.com', fontsize=10)
    data = doc.tobytes()
    doc.close()
    return data


class TestReferenciaPdfRoche(unittest.TestCase):
    def test_titulo_de_portada_y_editorial_ultima_pagina(self):
        f = apa.fuente_pdf('07153473500.pdf', _pdf_fabricado(), 'FIBRINOGENO')
        self.assertEqual(f.titulo, 'Fibrinogen')
        self.assertEqual(f.autor, 'Roche Diagnostics')
        self.assertEqual(f.anio, '2024')
        self.assertIn('Sandhofer Strasse 116', f.datos_editoriales)
        self.assertEqual(f.url, 'https://www.roche.com')
        self.assertFalse(f.titulo_provisional)
        cita = apa.referencia_apa(f)
        self.assertIn('Roche Diagnostics. (2024). Fibrinogen', cita)
        self.assertIn('D-68305 Mannheim', cita)
        self.assertIn('https://www.roche.com', cita)
        self.assertIn('<em>Fibrinogen</em>', apa.referencia_apa_html(f))
        self.assertIsInstance(apa.referencias_excel_con_cursiva([f]), str)

    def test_otro_examen_obtiene_su_propio_titulo_y_no_inventa_editorial(self):
        f = apa.fuente_pdf('glucosa.pdf', _pdf_fabricado('Glucose', roche=False), 'GLUCOSA')
        self.assertEqual(f.titulo, 'Glucose')
        self.assertFalse(f.autor)
        self.assertFalse(f.anio)  # '07 2005 9' es un ID, no fecha de publicación.
        self.assertFalse(f.datos_editoriales)
        self.assertFalse(f.url)
        self.assertIn('(s. f.)', apa.referencia_apa(f))


class TestFilasXlsx(unittest.TestCase):
    def test_parrafo_largo_sin_saltos_de_linea_tiene_altura(self):
        wb = Workbook()
        ws = wb.active
        ws.column_dimensions['B'].width = 42
        ws['A3'] = 'Procesamiento de muestras'
        contenido = ('Procesar la muestra de acuerdo con el procedimiento indicado, '
                     'verificar el estado y registrar los resultados del control. ') * 8
        ws['B3'] = contenido
        ws.row_dimensions[3].height = 15
        _ajustador_real()(ws, 3)
        self.assertTrue(ws['B3'].alignment.wrap_text)
        self.assertGreater(ws.row_dimensions[3].height, 130)
        buf = io.BytesIO()
        wb.save(buf)
        recuperado = load_workbook(io.BytesIO(buf.getvalue()))
        self.assertEqual(recuperado.active['B3'].value, contenido)
        self.assertGreater(recuperado.active.row_dimensions[3].height, 130)

    def test_una_fila_corta_y_una_con_vinetas_preservan_valor(self):
        wb = Workbook()
        ws = wb.active
        ws.column_dimensions['B'].width = 48
        ws['B1'] = 'Sin complicaciones'
        ws['B2'] = 'Indicaciones\n  • Primer punto\n  • Segundo punto\n  • Tercer punto'
        fn = _ajustador_real()
        fn(ws, 1)
        fn(ws, 2)
        self.assertLess(ws.row_dimensions[1].height or 15, 40)
        self.assertGreater(ws.row_dimensions[2].height, 60)
        self.assertEqual(ws['B2'].value.count('•'), 3)


class TestModalSinRepeticion(unittest.TestCase):
    def test_el_campo_se_muestra_una_sola_vez(self):
        inicio = APP.index('def modal_discrepancias_unificadas():')
        seccion = APP[inicio:APP.index('        respuestas[campo] =', inicio)]
        self.assertIn('titulo = "Valor extraído del PDF:"', seccion)
        self.assertIn('titulo = "Valor en Cartera SQLite:"', seccion)
        self.assertNotIn('Valor extraído del PDF ({campo})', seccion)
        self.assertNotIn('Valor en Cartera SQLite ({campo})', seccion)
        self.assertIn('Campo evaluado:', seccion)
        self.assertIn('_boton_valor_discrepancia(', seccion)
