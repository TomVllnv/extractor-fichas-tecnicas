"""Verifica las fichas de Hormonas e inmunidad / Derivados en SQLite."""
import ast
import io
import re
import sqlite3
import unittest
from pathlib import Path
import openpyxl

ROOT = Path(__file__).resolve().parents[1]


def funciones_hojas():
    arbol = ast.parse((ROOT / 'app.py').read_text(encoding='utf-8'))
    nombres = {'nombre_examen_valido', 'titulo_hoja_para_examen', 'buscar_hoja_por_examen',
               'norm', 'actualizar_campo_excel', 'crear_hoja_examen'}
    nodos = [n for n in arbol.body if isinstance(n, ast.FunctionDef) and n.name in nombres]
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment
    espacio = {'re': re, 'io': io, 'load_workbook': load_workbook,
               'Alignment': Alignment, 'FIELD_ALIASES_BASE': {},
               'limpiar_valor_excel': lambda valor, nombre: valor,
               'ajustar_ancho_columnas': lambda hoja: None,
               'ajustar_alto_fila_con_listas': lambda hoja, fila: None}
    exec(compile(ast.Module(body=nodos,type_ignores=[]),str(ROOT / 'app.py'),'exec'),espacio)
    return espacio

class TestNuevasAreas(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with sqlite3.connect(ROOT/'base_de_datos.db') as con:
            cls.registros = dict(con.execute('SELECT area,excel_bytes FROM plantillas_area'))
            cls.areas = {a:[r[0] for r in con.execute('SELECT nombre_examen FROM examenes_area WHERE area=?',(a,)).fetchall()]
                         for a in ('Derivados','Hormonas e inmunidad')}
        cls.fn=funciones_hojas()

    def test_areas_y_cantidad_de_fichas(self):
        self.assertEqual(len(self.areas['Derivados']),205)
        self.assertEqual(len(self.areas['Hormonas e inmunidad']),14)
        self.assertEqual(len(set(self.areas['Derivados'])),205)
        self.assertEqual(len(set(self.areas['Hormonas e inmunidad'])),14)
        self.assertNotIn('EXAMEN ESPECIAL UMT',self.areas['Derivados'])
        self.assertNotIn('EX. ESPECIAL CALVO',self.areas['Derivados'])

    def test_plantillas_xlsx_completas_y_fieles_a_descripcion(self):
        for area,nombre_laboratorio,esperadas in [('Derivados','DERIVADO EN CONVENIO',205),
                                                   ('Hormonas e inmunidad','HORMONAS E INMUNOLOGÍA',14)]:
            with self.subTest(area=area):
                wb=openpyxl.load_workbook(io.BytesIO(self.registros[area]),read_only=True,data_only=True)
                self.assertEqual(len(wb.sheetnames),esperadas)
                self.assertEqual(len(set(n.casefold() for n in wb.sheetnames)),esperadas)
                self.assertTrue(all(len(n)<=31 and not re.search(r'[\\/*?:\[\]]',n) for n in wb.sheetnames))
                full_names=[]
                for ws in wb.worksheets:
                    nombre=ws['B2'].value
                    full_names.append(nombre)
                    self.assertTrue(ws['B1'].value, f'Falta código en {nombre}')
                    self.assertEqual(ws['A2'].value,'Nombre LIS')
                    self.assertEqual(ws['B3'].value,nombre_laboratorio)
                    self.assertEqual(ws['A17'].value,'Valor crítico')
                    self.assertEqual(ws['A18'].value,'Límite de detección ')
                    self.assertEqual(ws['A24'].value,'Autorizado por:')
                self.assertEqual(set(full_names),set(self.areas[area]))
                wb.close()

    def test_hojas_con_nombres_largos_y_caracteres_prohibidos(self):
        examen='ANTIGENO PROSTATICO ESPECIFICO TOTAL - HOSMET'
        wb=openpyxl.load_workbook(io.BytesIO(self.registros['Derivados']),read_only=True)
        self.assertTrue(self.fn['nombre_examen_valido'](examen)[0])
        ws=self.fn['buscar_hoja_por_examen'](wb,examen)
        self.assertIsNotNone(ws)
        self.assertEqual(ws['B2'].value,examen)
        self.assertNotEqual(ws.title,examen)
        ws2=self.fn['buscar_hoja_por_examen'](wb,'PANEL MOLECULAR GINE/URO (HCM)')
        self.assertIsNotNone(ws2)
        self.assertNotIn('/',ws2.title)
        wb.close()

    def test_actualizar_un_examen_largo_sin_crear_hoja_duplicada(self):
        examen='ANTIGENO PROSTATICO ESPECIFICO TOTAL - HOSMET'
        anterior=self.registros['Derivados']
        despues=self.fn['actualizar_campo_excel'](anterior,examen,'Valor crítico','Revisión manual de prueba')
        wb=openpyxl.load_workbook(io.BytesIO(despues),read_only=True)
        self.assertEqual(len(wb.sheetnames),205)
        hoja=self.fn['buscar_hoja_por_examen'](wb,examen)
        self.assertIsNotNone(hoja)
        self.assertEqual(hoja['B17'].value,'Revisión manual de prueba')
        wb.close()

if __name__=='__main__':
    unittest.main()
