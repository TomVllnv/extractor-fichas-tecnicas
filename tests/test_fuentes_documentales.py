"""Pruebas de regresión de las fuentes SQLite, sin ejecutar Streamlit."""
import os
import sys
import json
import shutil
import sqlite3
import tempfile
import unittest
import zipfile
from io import BytesIO

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import fuentes_documentales as fd


class TestFuentesDocumentales(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.db = os.path.join(self.temp.name,'test.db')
        shutil.copyfile(os.path.join(os.path.dirname(__file__), '..', 'base_de_datos.db'),self.db)
        fd.inicializar_fuentes(self.db)

    def tearDown(self):
        self.temp.cleanup()

    def test_recursos_y_cartera_original(self):
        with sqlite3.connect(self.db) as con:
            self.assertEqual(con.execute('SELECT COUNT(*) FROM cartera_examenes').fetchone()[0],393)
            self.assertEqual(con.execute('SELECT COUNT(*) FROM prestaciones_acreditacion').fetchone()[0],374)
            self.assertEqual(con.execute('SELECT COUNT(*) FROM normas_generales_campos').fetchone()[0],4)
            self.assertIsNone(con.execute("SELECT name FROM sqlite_master WHERE name='manual_toma_campos'").fetchone())
        for fuente in (fd.NORMAS_ID,fd.ACREDITACION_ID):
            meta = fd.documento_fuente(self.db,fuente,True)
            self.assertTrue(meta['archivo'])
            self.assertEqual(len(meta['sha256']),64)

    def test_prestacion_exacta_prioriza_cartera(self):
        val,org,pre,mensaje = fd.obtener_valores_fuentes(self.db,'VITAMINA D','Bioquímica')
        self.assertEqual(mensaje,'exacta')
        self.assertEqual(pre['prestacion'],'VITAMINA D')
        self.assertNotEqual(val['Preparación del paciente'], 'No aplica')
        self.assertTrue(org['Preparación del paciente'].startswith('Normas Generales'))
        self.assertIn('Suero',val['Muestra requerida y recipiente'])
        self.assertEqual(len(val),6)
        for c in fd.CAMPOS_NORMAS_PRIORITARIAS:
            self.assertEqual(org[c], 'Normas Generales de Toma de Muestras')
        for c in ('Muestra requerida y recipiente', 'Condiciones de envío al Laboratorio'):
            self.assertIn('Cartera de Acreditación', org[c])
        self.assertTrue(all(c not in val and c not in org for c in fd.CAMPOS_MANUAL))

    def test_pendiente_usa_normas_no_texto_pendiente(self):
        val,org,pre,_ = fd.obtener_valores_fuentes(self.db,'17 - Hidroxiprogesterona','Derivados')
        self.assertIsNotNone(pre)
        self.assertEqual(len(val),5)
        self.assertNotIn('Muestra requerida y recipiente',val)
        self.assertTrue(org['Preparación del paciente'].startswith('Normas Generales'))
        self.assertIn('Higiene de manos',val['Procedimiento de toma de muestras'])
        self.assertIn('Normas Generales',org['Rotulación'])

    def test_sin_match_no_aplica_datos_de_otra_prestacion(self):
        val,org,pre,msg = fd.obtener_valores_fuentes(self.db,'examen desconocido','Microbiología')
        self.assertIsNone(pre)
        self.assertEqual(msg,'sin coincidencia exacta')
        self.assertEqual(set(val),set(fd.CAMPOS_NORMAS))
        self.assertTrue(all('Normas' in org[c] for c in fd.CAMPOS_NORMAS))
        self.assertTrue(all(c not in val and c not in org for c in fd.CAMPOS_MANUAL))

    def test_area_incompatible_no_se_autovincula(self):
        val,orig,pre,msg = fd.obtener_valores_fuentes(self.db,'VITAMINA D','Microbiología')
        self.assertIsNone(pre)
        self.assertIn('otra área',msg)
        self.assertEqual(len(val),4)

    def test_vinculo_manual(self):
        with sqlite3.connect(self.db) as con:
            x = con.execute("SELECT id FROM prestaciones_acreditacion WHERE prestacion='VITAMINA D'").fetchone()[0]
        val,orig,pre,msg = fd.obtener_valores_fuentes(self.db,'NOMBRE DE PRUEBA','Microbiología',x)
        self.assertEqual(pre['prestacion'],'VITAMINA D')
        self.assertEqual(msg,'manual')
        self.assertNotEqual(val['Preparación del paciente'], 'No aplica')
        self.assertTrue(orig['Preparación del paciente'].startswith('Normas Generales'))

    def test_actualizacion_invalida_no_borra_documentos(self):
        antes = fd.documento_fuente(self.db,fd.ACREDITACION_ID,True)
        with self.assertRaises(Exception):
            fd.guardar_cartera_acreditacion(self.db,b'no es excel','corrupto.xlsx')
        despues = fd.documento_fuente(self.db,fd.ACREDITACION_ID,True)
        self.assertEqual(antes['sha256'],despues['sha256'])
        self.assertEqual(despues['registros'],374)
        n_antes = fd.documento_fuente(self.db,fd.NORMAS_ID,True)
        with self.assertRaises(ValueError):
            fd.guardar_normas(self.db,b'no es pdf','incorrecto.txt')
        self.assertEqual(fd.documento_fuente(self.db,fd.NORMAS_ID,True)['sha256'],n_antes['sha256'])

    def test_docx_actualizable_sin_numeros_visibles(self):
        texto = [('Preparación del paciente','Preparación institucional'),
                 ('Procedimiento de toma de muestras','Procedimiento institucional'),
                 ('Rotulación','Rotulación institucional'),
                 ('Criterios de rechazo','Rechazo institucional')]
        ns='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        xml='<w:document xmlns:w="'+ns+'"><w:body>'
        for title,content in texto:
            xml+=f'<w:p><w:r><w:t>{title}</w:t></w:r></w:p>'
            xml+=f'<w:p><w:r><w:t>{content}</w:t></w:r></w:p>'
        xml+='</w:body></w:document>'
        out=BytesIO()
        with zipfile.ZipFile(out,'w') as z:
            z.writestr('word/document.xml',xml)
        self.assertEqual(fd.guardar_normas(self.db,out.getvalue(),'actualizacion.docx'),4)
        val,_,_,_=fd.obtener_valores_fuentes(self.db,'no presente','Bioquímica')
        self.assertEqual(val['Rotulación'],'Rotulación institucional')
        self.assertEqual(fd.documento_fuente(self.db,fd.NORMAS_ID)['nombre_archivo'],'actualizacion.docx')


if __name__ == '__main__':
    unittest.main()
