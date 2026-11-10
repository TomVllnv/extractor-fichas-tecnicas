"""Pruebas aisladas: AG Grid, limpieza Normas y formato XLSX de la ficha."""
import ast
import io
import re
import unicodedata
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment

import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fuentes_documentales as fd

APP = (Path(__file__).resolve().parents[1] / 'app.py').read_text(encoding='utf-8')
ARBOL = ast.parse(APP)


def funciones_excel():
    nombres = {'texto_normas_con_negritas', 'aplicar_formato_normas_excel',
               'aplicar_ediciones_a_excel', 'buscar_hoja_por_examen',
               'ajustar_alto_fila_con_listas', 'limpiar_valor_excel'}
    seleccion = [n for n in ARBOL.body if isinstance(n, ast.FunctionDef) and n.name in nombres]
    from openpyxl import load_workbook
    import hashlib
    espacio = {
        're': re, 'io': io, 'pd': pd, 'fd': fd, 'Alignment': Alignment, 'load_workbook': load_workbook,
        'norm': lambda t: re.sub(r'\s+', ' ', ''.join(c for c in unicodedata.normalize('NFD', str(t or '').casefold()) if unicodedata.category(c) != 'Mn')).strip(),
        'quitar_citas_superindices': lambda t: str(t),
        'convertir_subindices': lambda t: str(t),
        'es_linea_basura_o_encabezado': lambda *args: False,
        'st': type('Falso', (), {'session_state': {'origen_campos_documentales': {
            'Preparación del paciente': 'Normas Generales de Toma de Muestras'
        }}})(),
        'apa': type('ApaFalso', (), {'referencias_excel_con_cursiva': lambda x: str(x)})(),
    }
    exec(compile(ast.fix_missing_locations(ast.Module(body=seleccion, type_ignores=[])),
                 'app.py', 'exec'), espacio)
    return espacio


class TestNormasExcel(unittest.TestCase):
    def test_editor_largo_tiene_limite_real_de_excel_no_200(self):
        self.assertIn("cellEditorParams={'maxLength': 32767", APP)
        self.assertIn("cellEditor='agLargeTextCellEditor'", APP)

    def test_documento_antiguo_no_repite_parrafo_y_conserva_11_vinetas(self):
        muestra = ("Sangre venosa por sistema al vacío\n"
                   "Higiene de manos entre pacientes. Verificar que etiquetas y orden correspondan. "
                   "Seleccionar una vena y un sitio indemne. Mezclar por inversión suave. "
                   "Eliminar material cortopunzante sin recapsular. "
                   "Comprimir, colocar parche y registrar al responsable.\n"
                   "Preparación\n"
                   "● Higiene de manos entre pacientes.\n"
                   "● Verificar que etiquetas y orden correspondan.\n"
                   "● Seleccionar una vena y un sitio indemne.\n"
                   "● Mezclar por inversión suave.\n"
                   "● Eliminar material cortopunzante sin recapsular.\n"
                   "● Comprimir, colocar parche y registrar al responsable.\n")
        resultado = fd.limpiar_formato_normas('Procedimiento de toma de muestras', muestra)
        self.assertEqual(resultado.count('Higiene de manos entre pacientes'), 1)
        self.assertIn('\nPreparación\n', resultado)
        self.assertEqual(resultado.count('•'), 6)

    def test_reconstruye_paciente_sin_basura_ni_lineas_cortadas(self):
        texto = ('Regla institucional\n\nIndicaciones institucionales aplicables\n\n'
                 'a la atención ambulatoria.\n\nAntes de la toma\n\n'
                 '  ● Si toma un medicamento,\n\n    tráigalo consigo.\n\n'
                 '  ● No fume antes de los exámenes.\n\n')
        resultado = fd.limpiar_formato_normas('Preparación del paciente', texto)
        self.assertIn('Indicaciones institucionales aplicables a la atención ambulatoria.', resultado)
        self.assertIn('  • Si toma un medicamento, tráigalo consigo.', resultado)
        self.assertEqual(resultado.count('•'), 2)
        self.assertEqual(fd.limpiar_formato_normas('Preparación del paciente', resultado), resultado)

    def test_sangria_y_secuencia_de_vinetas_sobrevive_limpieza_excel(self):
        ff = funciones_excel()
        original = ('Antes de la toma\n'
                    '● Primera indicación clínica.\n'
                    '   ● Segunda indicación clínica.\n'
                    '\nDocumentación requerida\n'
                    '● Tercera indicación clínica.')
        limpio = ff['limpiar_valor_excel'](original)
        self.assertEqual(limpio.count('•'), 3)
        self.assertLess(limpio.index('Primera'), limpio.index('Segunda'))
        self.assertLess(limpio.index('Segunda'), limpio.index('Tercera'))
        self.assertIn('\n  • Primera', limpio)
        self.assertIn('\n  • Segunda', limpio)
        self.assertIn('\n  • Tercera', limpio)
        self.assertEqual(limpio.count('Documentación requerida'), 1)

    def test_excel_alto_ajustado_muestra_vinetas_y_titulos(self):
        ff = funciones_excel()
        wb = Workbook()
        ws = wb.active
        ws.title = 'GLUCOSA'
        ws.column_dimensions['B'].width = 55
        ws['A1'] = 'Preparación del paciente'
        original = ('Antes de la toma\n' +
                    ''.join(f'  • Indicación {n}: texto clínico sin alterar.\n' for n in range(1, 10)) +
                    '\nDocumentación requerida\n  • Orden del examen.')
        ws['B1'] = original
        archivo = io.BytesIO(); wb.save(archivo)
        result = ff['aplicar_formato_normas_excel'](
            archivo.getvalue(), 'GLUCOSA',
            {'Preparación del paciente': 'Normas Generales de Toma de Muestras'})
        salido = load_workbook(io.BytesIO(result), rich_text=True).active
        esperado = fd.limpiar_formato_normas('Preparación del paciente', original)
        self.assertEqual(str(salido['B1'].value), esperado)
        self.assertEqual(str(salido['B1'].value).count('•'), 10)
        self.assertGreater(salido.row_dimensions[1].height, 120)
        self.assertTrue(salido['B1'].alignment.wrap_text)
        self.assertIsInstance(salido['B1'].value, str)
        self.assertEqual(str(salido['B1'].value).count('Indicación'), 9)

    def test_hematologia_rotulacion_sin_tilde_conserva_vinetas_y_negritas(self):
        """La plantilla de Hematología usa «Rotulacion» sin acento."""
        ff = funciones_excel()
        wb = Workbook()
        ws = wb.active
        ws.title = 'FIBRINOGENO'
        ws['A1'] = 'Rotulacion'
        ws['B1'] = 'Datos mínimos\n  • Nombre y apellidos.\n  • Fecha de nacimiento.'
        otra = wb.create_sheet('HEMOGRAMA')
        otra['A1'], otra['B1'] = 'Rotulacion', 'Sin cambios'
        bf = io.BytesIO(); wb.save(bf)
        res = ff['aplicar_formato_normas_excel'](
            bf.getvalue(), 'FIBRINOGENO',
            {'Rotulación': 'Normas Generales de Toma de Muestras'})
        verificado = load_workbook(io.BytesIO(res), rich_text=True)
        celda = verificado['FIBRINOGENO']['B1']
        self.assertEqual(str(celda.value).count('•'), 2)
        self.assertIsInstance(celda.value, str)
        self.assertTrue(celda.alignment.wrap_text)
        self.assertEqual(verificado['HEMOGRAMA']['B1'].value, 'Sin cambios')

    def test_tabla_revision_conserva_diseno_previo_sin_ajuste_de_filas(self):
        inicio = APP.index('gb.configure_column(\n                "Información extraída",')
        fin = APP.index('gb.configure_column(\n                "Estado",', inicio)
        configuracion = APP[inicio:fin]
        self.assertNotIn('wrapText=True', configuracion)
        self.assertNotIn('whiteSpace', configuracion)
        self.assertIn('autoHeight=True', configuracion)
        self.assertIn("'maxLength': 32767", configuracion)

    def test_excel_preserva_negritas_y_edicion_sin_cortarla(self):
        ff = funciones_excel()
        wb = Workbook()
        ws = wb.active
        ws.title = 'GLUCOSA'
        ws['A1'] = 'Preparación del paciente'
        original = ('Regla institucional\nIndicaciones institucionales aplicables.\n'
                    'Antes de la toma\n  • No fume antes del examen.')
        ws['B1'] = original
        archivo = io.BytesIO(); wb.save(archivo)
        formateado = ff['aplicar_formato_normas_excel'](
            archivo.getvalue(), 'GLUCOSA',
            {'Preparación del paciente': 'Normas Generales de Toma de Muestras'})
        wb2 = load_workbook(io.BytesIO(formateado), rich_text=True)
        celda = wb2.active['B1']
        self.assertEqual(str(celda.value), fd.limpiar_formato_normas('Preparación del paciente', original))
        self.assertIsInstance(celda.value, str)
        self.assertIn('• No fume', str(celda.value))
        nuevo = str(celda.value)
        filas = pd.DataFrame([{'Campo': 'Preparación del paciente', 'Información extraída': nuevo}])
        despues = ff['aplicar_ediciones_a_excel'](formateado, filas, 'GLUCOSA')
        celda3 = load_workbook(io.BytesIO(despues), rich_text=True).active['B1']
        self.assertIsInstance(celda3.value, str)
        # El editor y la escritura real soportan textos sustancialmente >200 caracteres.
        texto_largo = nuevo + '\n' + ('Detalle clínico de ejemplo. ' * 500)
        filas.iloc[0, 1] = texto_largo
        editado = ff['aplicar_ediciones_a_excel'](despues, filas, 'GLUCOSA')
        valor = load_workbook(io.BytesIO(editado), rich_text=True).active['B1'].value
        self.assertEqual(str(valor), texto_largo.strip())
        self.assertGreater(len(str(valor)), 10000)


if __name__ == '__main__':
    unittest.main()
