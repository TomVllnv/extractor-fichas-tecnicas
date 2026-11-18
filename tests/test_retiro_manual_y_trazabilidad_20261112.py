"""Regresión del retiro del Manual, respetando Normas y las fichas existentes."""
from pathlib import Path
import ast
import io
import sqlite3
import sys
import tempfile
import unittest
from datetime import datetime
from zoneinfo import ZoneInfo
import openpyxl

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
import fuentes_documentales as fd

APP=(ROOT/'app.py').read_text(encoding='utf-8')

class TestRetiroManual(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.db=str(Path(self.tmp.name)/'datos.db')
        fd.inicializar_fuentes(self.db)
        with sqlite3.connect(self.db) as con:
            con.execute("INSERT INTO normas_generales_campos(campo, contenido) VALUES (?, ?)",
                        ('Rotulación', 'Normas originales sin cambios'))
            con.execute("INSERT INTO documentos_fuente VALUES (?,?,?,?,?,?,?)", (
                'normas_generales','normas-originales.docx','application/docx',b'normas',
                'hashnormas','2026-11-01',4))
            con.execute("INSERT INTO documentos_fuente VALUES (?,?,?,?,?,?,?)", (
                fd.MANUAL_ID,'manual.pdf','application/pdf',b'manual-confidencial',
                'hashmanual','2026-11-01',2))
            con.execute("CREATE TABLE manual_toma_campos (campo TEXT PRIMARY KEY, contenido TEXT NOT NULL)")
            con.executemany("INSERT INTO manual_toma_campos VALUES (?, ?)", [
                ('Recepción de muestras','TEXTO MANUAL ANTERIOR'),
                ('Procesamiento de muestras','OTRO TEXTO MANUAL ANTERIOR')])

    def test_migracion_elimina_solo_manual_y_se_puede_repetir(self):
        self.assertTrue(fd.hay_manual_heredado(self.db))
        with sqlite3.connect(self.db) as con:
            normas_antes=con.execute("SELECT * FROM documentos_fuente WHERE fuente_id='normas_generales'").fetchone()
            contenido_normas=con.execute("SELECT * FROM normas_generales_campos").fetchall()
        fd.inicializar_fuentes(self.db)
        self.assertFalse(fd.hay_manual_heredado(self.db))
        fd.inicializar_fuentes(self.db)  # No reimporta la fuente tras un reinicio.
        with sqlite3.connect(self.db) as con:
            self.assertIsNone(con.execute("SELECT * FROM documentos_fuente WHERE fuente_id=?", (fd.MANUAL_ID,)).fetchone())
            self.assertIsNone(con.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='manual_toma_campos'").fetchone())
            self.assertEqual(con.execute("SELECT * FROM documentos_fuente WHERE fuente_id='normas_generales'").fetchone(), normas_antes)
            self.assertEqual(con.execute("SELECT * FROM normas_generales_campos").fetchall(),contenido_normas)

    def test_respaldo_previo_preserva_manual_antiguo(self):
        import gestion_compartida as gp
        if fd.hay_manual_heredado(self.db):
            copia=gp.respaldar(self.db)
        fd.inicializar_fuentes(self.db)
        self.assertTrue(Path(copia).exists())
        with sqlite3.connect(copia) as antiguo:
            self.assertEqual(antiguo.execute(
                "SELECT archivo FROM documentos_fuente WHERE fuente_id=?",(fd.MANUAL_ID,)).fetchone()[0],
                b'manual-confidencial')
            self.assertIsNotNone(antiguo.execute(
                "SELECT 1 FROM sqlite_master WHERE name='manual_toma_campos'").fetchone())
        self.assertIsNone(fd.documento_fuente(self.db,fd.MANUAL_ID))
        self.assertIsNotNone(fd.documento_fuente(self.db,fd.NORMAS_ID))

    def test_no_valores_automaticos_manual_ni_alias_pdf_en_ninguna_area(self):
        fd.inicializar_fuentes(self.db)
        valores, origen, _, _ = fd.obtener_valores_fuentes(self.db,'Examen no presente','Bioquímica')
        self.assertEqual(valores['Rotulación'],'Normas originales sin cambios')
        for campo in fd.CAMPOS_MANUAL:
            self.assertIn(campo,fd.CAMPOS_CONTROLADOS)
            self.assertNotIn(campo,valores)
            self.assertNotIn(campo,origen)
            self.assertEqual(fd.PRIORIDAD_FUENTES_POR_CAMPO[campo],())
        arbol=ast.parse(APP)
        for node in arbol.body:
            if isinstance(node,ast.Assign) and len(node.targets)==1:
                target=node.targets[0]
                if isinstance(target,ast.Name) and target.id.startswith('FIELD_ALIASES_') and target.id!='FIELD_ALIASES_POR_AREA':
                    aliases=ast.literal_eval(node.value)
                    for campo in fd.CAMPOS_MANUAL:
                        self.assertIn(campo,aliases)
                        self.assertEqual(aliases[campo],[],(target.id,campo))

    def test_rellenar_excel_deja_campos_vacios_aunque_pdf_o_excel_tengan_texto(self):
        """Incluso con PDF y Excel heredados, los valores automáticos se excluyen."""
        arbol=ast.parse(APP)
        proceso=next(n for n in arbol.body if isinstance(n,ast.FunctionDef) and n.name=='process')
        ns={'__name__':'test_fuente_retirada', 'io':io, 'datetime':datetime, 'ZoneInfo':ZoneInfo,
            'fd':fd, 'load_workbook':openpyxl.load_workbook,
            'Alignment':openpyxl.styles.Alignment,
            'FIELD_ALIASES_BASE':{'Recepción de muestras':[], 'Procesamiento de muestras':[], 'Otro campo':[]},
            'nombre_examen_valido':lambda n:(True,''),
            'buscar_hoja_por_examen':lambda wb,n: wb[n] if n in wb.sheetnames else None,
            'norm':fd.normalizar,
            'normalizar_etiquetas_ficha':lambda ws:None,
            'extract_pdf_text':lambda *args, **kwargs:('PDF con textos institucionales anteriores','texto'),
            'parse_labeled_text':lambda *args:{'Recepción de muestras':'TEXTO PDF ERRÓNEO',
                                                'Procesamiento de muestras':'TEXTO PDF ERRÓNEO',
                                                'Otro campo':'Dato válido'},
            'infer_name_code':lambda raw,d:d,
            'es_alias_hematologia':lambda x:False,
            'limpiar_valor_excel':lambda v,*a:v,
            'ajustar_ancho_columnas':lambda ws:None,
            'ajustar_alto_fila_con_listas':lambda ws, fila:None}
        def aplicar(ws,datos):
            cambios=[]
            for fila in ws.iter_rows():
                campo=str(fila[0].value or '')
                if campo in datos:
                    fila[1].value=datos[campo]
                    cambios.append((campo,datos[campo]))
            return cambios
        ns['apply_to_sheet']=aplicar
        exec(compile(ast.Module(body=[proceso],type_ignores=[]),str(ROOT/'app.py'),'exec'),ns)
        wb=openpyxl.Workbook()
        ws=wb.active
        ws.title='EXAMEN PRUEBA'
        for i,campo in enumerate(['Recepción de muestras','Procesamiento de muestras','Otro campo'],1):
            ws.cell(i,1,campo)
            ws.cell(i,2,'Valor antiguo guardado en Excel')
        buf=io.BytesIO();wb.save(buf)
        fd.inicializar_fuentes(self.db)
        docs,_,_,_=fd.obtener_valores_fuentes(self.db,'EXAMEN PRUEBA','Bioquímica')
        out,datos,*_=ns['process'](b'pdf','ejemplo.pdf',buf.getvalue(),nombre_examen='EXAMEN PRUEBA',datos_sobrescribir=docs)
        self.assertNotIn('Recepción de muestras',datos)
        self.assertNotIn('Procesamiento de muestras',datos)
        sheet=openpyxl.load_workbook(io.BytesIO(out))['EXAMEN PRUEBA']
        self.assertIsNone(sheet.cell(1,2).value)
        self.assertIsNone(sheet.cell(2,2).value)
        self.assertEqual(sheet.cell(3,2).value,'Dato válido')

    def test_sidebar_y_trazabilidad_sin_textos_obsoletos(self):
        self.assertIn('with st.expander("Normas generales de toma de muestras", expanded=False):',APP)
        self.assertNotIn('with st.expander("Manual de toma de muestras"',APP)
        self.assertNotIn('descargar_manual_toma_original',APP)
        self.assertNotIn('guardar_manual_toma_nuevo',APP)
        self.assertNotIn('Los cinco registros más recientes. Revisión: versión del recurso tras guardar el cambio.', APP)
        self.assertNotIn('st.caption("Eventos recientes.")',APP)
        self.assertIn('_tabla_trazabilidad("modificaciones", 5)',APP)
        self.assertIn('_tabla_trazabilidad("actividad", 5)',APP)

if __name__ == '__main__':
    unittest.main()
