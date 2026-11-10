"""Pruebas específicas de la tercera fuente documental almacenada en SQLite."""
from pathlib import Path
import sys
import tempfile
import sqlite3
import shutil
import hashlib
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import fuentes_documentales as fd


class TestManualToma(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.db = Path(self.tmp.name) / 'copia.db'
        shutil.copyfile(ROOT / 'base_de_datos.db', self.db)
        fd.inicializar_fuentes(str(self.db))

    def test_recepcion_fragmento_exacto_sin_basura_ni_titulos(self):
        with sqlite3.connect(self.db) as c:
            texto = c.execute("SELECT contenido FROM manual_toma_campos WHERE campo=?", ('Recepción de muestras',)).fetchone()[0]
        self.assertTrue(texto.startswith('Las muestras deben ser recepcionadas por un técnico paramédico'))
        self.assertTrue(texto.endswith('cumplen con lo estipulado por el Laboratorio del CRSCO.'))
        self.assertIn('Servicio de Salud Metropolitano Oriente', texto)
        self.assertIn('procedimiento exámenes VIH', texto)
        self.assertIn('Anexo 3', texto)
        self.assertNotIn('X. RECEPCION', texto)
        for basura in ('Código :', 'Cod. Caract.', 'Vigencia:', 'Página :', 'Página:', '57 de 75', '56 \n'):
            self.assertNotIn(basura, texto)

    def test_procesamiento_fragmento_exacto_sin_titulo_ni_pagina_siguiente(self):
        with sqlite3.connect(self.db) as c:
            texto = c.execute("SELECT contenido FROM manual_toma_campos WHERE campo=?", ('Procesamiento de muestras',)).fetchone()[0]
        self.assertTrue(texto.startswith('Las muestras son rotuladas en la Toma de Muestras'))
        self.assertTrue(texto.endswith('paciente en la secretaría de laboratorio.'))
        self.assertIn('Sistema Informático de Laboratorio', texto)
        self.assertIn('Técnico Paramédico de Laboratorio', texto)
        self.assertNotIn('2. MUESTRAS PROCEDENTES', texto)
        self.assertNotIn('3. MUESTRAS PROCEDENTES', texto)
        self.assertNotIn('Página', texto)

    def test_valores_para_examen_no_vinculado_son_solo_normas_y_manual(self):
        valores, origen, prestacion, estado = fd.obtener_valores_fuentes(str(self.db),'NOMBRE DE PRUEBA QUE NO EXISTE', 'Microbiología')
        self.assertIsNone(prestacion)
        self.assertEqual(estado,'sin coincidencia exacta')
        self.assertEqual(set(valores),set(fd.CAMPOS_NORMAS) | set(fd.CAMPOS_MANUAL))
        for campo in fd.CAMPOS_MANUAL:
            self.assertEqual(origen[campo], 'Manual de toma de muestras')

    def test_actualizacion_invalida_conserva_pdf_original_y_campos(self):
        antes = fd.documento_fuente(str(self.db),fd.MANUAL_ID,True)
        with sqlite3.connect(self.db) as c:
            previos = c.execute('SELECT campo,contenido FROM manual_toma_campos ORDER BY campo').fetchall()
        with self.assertRaises(ValueError):
            fd.guardar_manual_toma(str(self.db),b'Esto no es un documento', 'nuevo.pdf')
        with self.assertRaises(ValueError):
            fd.guardar_manual_toma(str(self.db),b'PDF invalido','nuevo.docx')
        despues=fd.documento_fuente(str(self.db),fd.MANUAL_ID,True)
        self.assertEqual(antes['sha256'],despues['sha256'])
        self.assertEqual(antes['nombre_archivo'],despues['nombre_archivo'])
        with sqlite3.connect(self.db) as c:
            self.assertEqual(previos,c.execute('SELECT campo,contenido FROM manual_toma_campos ORDER BY campo').fetchall())

    def test_reimportacion_valida_actualiza_nombre_dinamico(self):
        fuente=fd.documento_fuente(str(self.db),fd.MANUAL_ID,True)
        self.assertEqual(fd.guardar_manual_toma(str(self.db),fuente['archivo'],'Manual_actualizado_2027.pdf'),2)
        despues=fd.documento_fuente(str(self.db),fd.MANUAL_ID,True)
        self.assertEqual(despues['nombre_archivo'],'Manual_actualizado_2027.pdf')
        self.assertEqual(despues['registros'],2)
        self.assertEqual(despues['sha256'],hashlib.sha256(fuente['archivo']).hexdigest())
        self.assertTrue(despues['fecha_actualizacion'])

    def test_pdf_cargado_no_sobrescribe_los_campos_controlados(self):
        codigo=(ROOT/'app.py').read_text(encoding='utf-8')
        self.assertIn('if campo not in fd.CAMPOS_CONTROLADOS',codigo)
        self.assertIn('if str(ws.cell(fila, 1).value or "").strip() in fd.CAMPOS_CONTROLADOS',codigo)
        self.assertEqual(len(fd.CAMPOS_CONTROLADOS),8)
        self.assertNotIn('Recepción de muestras',fd.MAPA_ACREDITACION)
        self.assertNotIn('Procesamiento de muestras',fd.MAPA_ACREDITACION)

    def test_genera_excel_con_manual_y_no_con_texto_pdf(self):
        """Prueba aislada de la ruta real de llenado del Excel, sin Streamlit."""
        import ast
        import copy
        import io
        from datetime import datetime
        from zoneinfo import ZoneInfo
        import openpyxl
        import re
        codigo=(ROOT/'app.py').read_text(encoding='utf-8')
        arbol=ast.parse(codigo)
        funcion=next(n for n in arbol.body if isinstance(n, ast.FunctionDef) and n.name=='process')
        mod=ast.Module(body=[funcion],type_ignores=[])
        ns={'__name__':'test_manual_proceso', 'io':io, 'datetime':datetime, 'ZoneInfo':ZoneInfo,
            'fd':fd, 'load_workbook':openpyxl.load_workbook,
            'Alignment':openpyxl.styles.Alignment,
            'FIELD_ALIASES_BASE':{'Recepción de muestras':[], 'Procesamiento de muestras':[], 'Otro campo':[]},
            'nombre_examen_valido':lambda x:(True,''),
            'buscar_hoja_por_examen':lambda wb,n: wb[n] if n in wb.sheetnames else None,
            'norm':fd.normalizar,
            'extract_pdf_text':lambda *args:('PDF contiene Recepción de muestras: VALOR INADECUADO. Procesamiento de muestras: VALOR INADECUADO','texto'),
            'parse_labeled_text':lambda *args:{'Recepción de muestras':'Valor incorrecto del PDF',
                                               'Procesamiento de muestras':'Valor incorrecto del PDF',
                                               'Otro campo':'Valor correcto del PDF'},
            'infer_name_code':lambda raw,d:d,
            'es_alias_hematologia':lambda x:False,
            'limpiar_valor_excel':lambda v,*a:v,
            'ajustar_ancho_columnas':lambda ws:None,
            'ajustar_alto_fila_con_listas':lambda ws, fila:None}

        def aplicar(ws, datos):
            cambios=[]
            for row in ws.iter_rows():
                campo=str(row[0].value or '')
                if campo in datos:
                    row[1].value=datos[campo]
                    cambios.append((campo,datos[campo]))
            return cambios
        ns['apply_to_sheet']=aplicar
        exec(compile(mod,str(ROOT/'app.py'),'exec'),ns)
        wb=openpyxl.Workbook()
        ws=wb.active
        ws.title='EXAMEN PRUEBA'
        for idx,campo in enumerate(['Recepción de muestras','Procesamiento de muestras','Otro campo'],1):
            ws.cell(idx,1,campo)
            ws.cell(idx,2,'texto obsoleto')
        buffer=io.BytesIO()
        wb.save(buffer)
        manual,_,_,_=fd.obtener_valores_fuentes(str(self.db),'EXAMEN PRUEBA','Microbiología')
        resultado,datos,*_=ns['process'](b'pdf falso','prueba.pdf',buffer.getvalue(),
                                nombre_examen='EXAMEN PRUEBA',datos_sobrescribir=manual)
        self.assertEqual(datos['Recepción de muestras'],manual['Recepción de muestras'])
        self.assertEqual(datos['Procesamiento de muestras'],manual['Procesamiento de muestras'])
        self.assertEqual(datos['Otro campo'],'Valor correcto del PDF')
        out=openpyxl.load_workbook(io.BytesIO(resultado))['EXAMEN PRUEBA']
        self.assertEqual(out.cell(1,2).value,manual['Recepción de muestras'])
        self.assertEqual(out.cell(2,2).value,manual['Procesamiento de muestras'])
        self.assertEqual(out.cell(3,2).value,'Valor correcto del PDF')

    def test_interfaz_tercera_fuente(self):
        codigo=(ROOT/'app.py').read_text(encoding='utf-8')
        self.assertIn('st.expander("Manual de toma de muestras"',codigo)
        self.assertIn('key="descargar_manual_toma_original"',codigo)
        self.assertIn('key="actualizar_manual_toma_archivo"',codigo)
        self.assertIn('key="guardar_manual_toma_nuevo"',codigo)
        self.assertIn("formatear_fecha_fuente(meta_manual['fecha_actualizacion'])",codigo)
        self.assertIn("meta_manual['nombre_archivo']",codigo)

if __name__=='__main__':
    unittest.main()
