"""Regresión de prioridad clínica y migración del área sin reemplazar SQLite."""
import io
import json
import os
import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from openpyxl import Workbook
import cartera_crs as crs
import fuentes_documentales as fd


class PrioridadPreparacionTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.db = os.path.join(self.temp.name, 'prueba.db')
        fd.inicializar_fuentes(self.db)
        with sqlite3.connect(self.db) as con:
            con.execute("INSERT INTO normas_generales_campos(campo, contenido) VALUES (?, ?)",
                        ('Preparación del paciente', 'Recomendación general de Normas'))
            con.execute("INSERT INTO normas_generales_campos(campo, contenido) VALUES (?, ?)",
                        ('Rotulación', 'Rotulación general'))
            datos = {'Preparación del paciente': 'Ayuno de 8 horas', 'Rotulación': 'No aplica'}
            estados = {'Preparación del paciente': 'Específico', 'Rotulación': 'Específico'}
            con.execute('''INSERT INTO prestaciones_acreditacion(fila_excel,prestacion,prestacion_norm,area,area_norm,datos_json,estados_json)
                           VALUES (1,'GLUCOSA','glucosa','Bioquímica','bioquimica',?,?)''',
                        (json.dumps(datos), json.dumps(estados)))

    def tearDown(self):
        self.temp.cleanup()

    def _set_preparacion_acreditacion(self, valor, estado='Específico'):
        with sqlite3.connect(self.db) as con:
            r = con.execute('SELECT datos_json,estados_json FROM prestaciones_acreditacion').fetchone()
            datos, estados = json.loads(r[0]), json.loads(r[1])
            datos['Preparación del paciente'] = valor
            estados['Preparación del paciente'] = estado
            con.execute('UPDATE prestaciones_acreditacion SET datos_json=?,estados_json=?',
                        (json.dumps(datos),json.dumps(estados)))

    def test_crs_tiene_maxima_prioridad(self):
        vals, origen, _, _ = fd.obtener_valores_fuentes(self.db,'GLUCOSA','Bioquímica',
                                                        indicaciones_cartera='Ayuno de 12 horas')
        self.assertEqual(vals['Preparación del paciente'], 'Ayuno de 12 horas')
        self.assertEqual(origen['Preparación del paciente'], 'Cartera de Exámenes del CRS · Indicaciones')
        self.assertEqual(vals['Rotulación'], 'Rotulación general')  # Normas tiene prioridad en este campo.

    def test_acreditacion_prioriza_normas_si_crs_vacio(self):
        vals, origen, _, _ = fd.obtener_valores_fuentes(self.db,'GLUCOSA','Bioquímica')
        self.assertEqual(vals['Preparación del paciente'], 'Ayuno de 8 horas')
        self.assertTrue(origen['Preparación del paciente'].startswith('Cartera de Acreditación'))

    def test_no_aplica_acreditacion_es_un_valor_documentado_prioritario(self):
        for contenido in ('No aplica', 'No aplicable', 'N/A', 'NA', '  no APLICA.  '):
            with self.subTest(contenido=contenido):
                self._set_preparacion_acreditacion(contenido)
                vals, origen, _, _ = fd.obtener_valores_fuentes(self.db,'GLUCOSA','Bioquímica')
                self.assertEqual(vals['Preparación del paciente'], contenido)
                self.assertTrue(origen['Preparación del paciente'].startswith('Cartera de Acreditación'))

    def test_no_aplica_con_estado_incompleto_no_sustituye_normas(self):
        for estado in ('Pendiente', 'Incompleto', 'No encontrado', 'Sin documentar'):
            with self.subTest(estado=estado):
                self._set_preparacion_acreditacion('No aplica', estado)
                vals, origen, _, _ = fd.obtener_valores_fuentes(self.db,'GLUCOSA','Bioquímica')
                self.assertEqual(vals['Preparación del paciente'], 'Recomendación general de Normas')
                self.assertEqual(origen['Preparación del paciente'], 'Normas Generales de Toma de Muestras')

    def test_no_aplica_en_normas_sigue_descartandose_sin_acreditacion(self):
        with sqlite3.connect(self.db) as con:
            con.execute('UPDATE normas_generales_campos SET contenido=? WHERE campo=?',
                        ('No aplica', 'Preparación del paciente'))
        self._set_preparacion_acreditacion('Pendiente', 'Pendiente')
        vals, origen, _, _ = fd.obtener_valores_fuentes(self.db, 'GLUCOSA', 'Bioquímica')
        self.assertNotIn('Preparación del paciente', vals)
        self.assertNotIn('Preparación del paciente', origen)

    def test_no_aplica_acreditacion_no_desplaza_indicaciones_crs(self):
        self._set_preparacion_acreditacion('No aplica', 'Heredado')
        vals, origen, _, _ = fd.obtener_valores_fuentes(
            self.db, 'GLUCOSA', 'Bioquímica', indicaciones_cartera='Ayuno 8 horas')
        self.assertEqual(vals['Preparación del paciente'], 'Ayuno 8 horas')
        self.assertEqual(origen['Preparación del paciente'], 'Cartera de Exámenes del CRS · Indicaciones')

    def test_no_aplica_acreditacion_no_altera_prioridad_rotulacion(self):
        self._set_preparacion_acreditacion('No aplica')
        vals, origen, _, _ = fd.obtener_valores_fuentes(self.db, 'GLUCOSA', 'Bioquímica')
        self.assertEqual(vals['Rotulación'], 'Rotulación general')
        self.assertEqual(origen['Rotulación'], 'Normas Generales de Toma de Muestras')

    def test_no_aplica_importado_desde_encabezado_real_del_excel(self):
        libro = Workbook()
        hoja = libro.active
        hoja.title = 'Cartera_Completa'
        encabezados = ['ID', 'Código', 'Prestación', 'Área']
        for nombre_columna in fd.MAPA_ACREDITACION.values():
            encabezados.extend([nombre_columna, nombre_columna.replace('· texto', '· estado')])
        hoja.append(encabezados)
        datos_fila = ['1', 'LAB01', 'GLUCOSA', 'Bioquímica']
        for campo in fd.MAPA_ACREDITACION:
            datos_fila.extend(['No aplica' if campo == 'Preparación del paciente' else '',
                               'Específico'])
        hoja.append(datos_fila)
        memoria = io.BytesIO()
        libro.save(memoria)
        self.assertEqual(fd.guardar_cartera_acreditacion(self.db, memoria.getvalue(), 'cartera.xlsx'), 1)
        vals, origen, _, _ = fd.obtener_valores_fuentes(self.db, 'GLUCOSA', 'Bioquímica')
        self.assertEqual(vals['Preparación del paciente'], 'No aplica')
        self.assertTrue(origen['Preparación del paciente'].startswith('Cartera de Acreditación'))

    def test_pendiente_acreditacion_usa_normas(self):
        self._set_preparacion_acreditacion('Pendiente incorporar la instrucción específica del centro de derivación. '
                                          'El manual CRS indica consultar los manuales/instructivos del laboratorio externo '
                                          'para estas prestaciones.')
        vals, _, _, _ = fd.obtener_valores_fuentes(self.db,'GLUCOSA','Bioquímica')
        self.assertEqual(vals['Preparación del paciente'], 'Recomendación general de Normas')

    def test_aviso_generico_acreditacion_usa_normas(self):
        self._set_preparacion_acreditacion(fd.AVISO_PREPARACION_GENERAL_ACREDITACION)
        vals, origen, _, _ = fd.obtener_valores_fuentes(self.db, 'GLUCOSA', 'Bioquímica')
        self.assertEqual(vals['Preparación del paciente'], 'Recomendación general de Normas')
        self.assertEqual(origen['Preparación del paciente'], 'Normas Generales de Toma de Muestras')

    def test_aviso_generico_acreditacion_se_ignora_con_espacios_y_tildes_variables(self):
        texto = ("  APLICAR requisitos generales de preparacion del paciente del CRS;\n"
                 " verificar  ademas si la prestación tiene una indicación ESPECÍFICA.  ")
        self._set_preparacion_acreditacion(texto)
        vals, origen, _, _ = fd.obtener_valores_fuentes(self.db, 'GLUCOSA', 'Bioquímica')
        self.assertEqual(vals['Preparación del paciente'], 'Recomendación general de Normas')
        self.assertIn('Normas', origen['Preparación del paciente'])

    def test_crs_prevalece_aunque_acreditacion_tenga_el_aviso_generico(self):
        self._set_preparacion_acreditacion(fd.AVISO_PREPARACION_GENERAL_ACREDITACION)
        vals, origen, _, _ = fd.obtener_valores_fuentes(
            self.db, 'GLUCOSA', 'Bioquímica', indicaciones_cartera='Ayuno de 12 horas')
        self.assertEqual(vals['Preparación del paciente'], 'Ayuno de 12 horas')
        self.assertIn('Cartera de Exámenes del CRS', origen['Preparación del paciente'])

    def test_acreditacion_con_instrucciones_extra_no_se_descarta_por_error(self):
        texto = fd.AVISO_PREPARACION_GENERAL_ACREDITACION + ' Ayuno de 8 horas.'
        self._set_preparacion_acreditacion(texto)
        vals, origen, _, _ = fd.obtener_valores_fuentes(self.db, 'GLUCOSA', 'Bioquímica')
        self.assertEqual(vals['Preparación del paciente'], texto)
        self.assertIn('Cartera de Acreditación', origen['Preparación del paciente'])

    def test_aviso_de_acreditacion_no_se_filtra_de_normas_ni_crs(self):
        aviso = fd.AVISO_PREPARACION_GENERAL_ACREDITACION
        self.assertTrue(fd.preparacion_valida(aviso))
        self.assertFalse(fd.preparacion_valida(aviso, 'Heredado', fuente_acreditacion=True))
        vals, origen, _, _ = fd.obtener_valores_fuentes(
            self.db, 'GLUCOSA', 'Bioquímica', indicaciones_cartera=aviso)
        self.assertEqual(vals['Preparación del paciente'], aviso)
        self.assertIn('Cartera de Exámenes del CRS', origen['Preparación del paciente'])

    def test_estado_pendiente_acreditacion_usa_normas(self):
        self._set_preparacion_acreditacion('Ayuno de 10 horas','Pendiente')
        vals, _, _, _ = fd.obtener_valores_fuentes(self.db,'GLUCOSA','Bioquímica')
        self.assertEqual(vals['Preparación del paciente'], 'Recomendación general de Normas')

    def test_crs_no_aplica_es_contenido_del_primer_documento(self):
        # Indicaciones CRS mantiene su máxima prioridad, incluso con «No aplica».
        vals, origen, _, _ = fd.obtener_valores_fuentes(self.db,'GLUCOSA','Bioquímica',
                                                        indicaciones_cartera='No aplica')
        self.assertEqual(vals['Preparación del paciente'], 'No aplica')
        self.assertIn('Cartera de Exámenes',origen['Preparación del paciente'])


class MigracionCarteraTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.db = os.path.join(self.temp.name,'antigua.db')
        fd.inicializar_fuentes(self.db)
        with sqlite3.connect(self.db) as con:
            con.execute('''CREATE TABLE cartera_examenes (
                id INTEGER PRIMARY KEY AUTOINCREMENT, nombre_examen TEXT UNIQUE,
                indicaciones TEXT, codigo_examen TEXT, tecnologia TEXT,
                entrega_normal TEXT, entrega_urgente TEXT)''')
            con.execute("INSERT INTO cartera_examenes(nombre_examen,indicaciones) VALUES('GLUCOSA','Ayuno 8 h')")
            con.execute("INSERT INTO cartera_examenes(nombre_examen,indicaciones) VALUES('UROCULTIVO','Sin ayuno')")
            con.execute("CREATE TABLE usuarios_aplicacion(usuario TEXT PRIMARY KEY, clave_derivada BLOB)")
            con.execute("INSERT INTO usuarios_aplicacion VALUES('ADMIN',X'123456')")

    def tearDown(self):
        self.temp.cleanup()

    def test_migracion_sin_excel_no_inventa_areas_ni_elimina_usuarios(self):
        self.assertEqual(crs.migrar_laboratorio_cartera(self.db),0)
        with sqlite3.connect(self.db) as con:
            self.assertIn('laboratorio', {x[1] for x in con.execute('PRAGMA table_info(cartera_examenes)')})
            self.assertEqual(con.execute('SELECT COUNT(*) FROM cartera_examenes').fetchone()[0],2)
            self.assertEqual(con.execute('SELECT clave_derivada FROM usuarios_aplicacion WHERE usuario="ADMIN"').fetchone()[0],b'\x124V')
            self.assertIsNone(con.execute('SELECT laboratorio FROM cartera_examenes WHERE nombre_examen="GLUCOSA"').fetchone()[0])

    def test_recupera_area_del_original_y_puede_seleccionar_por_area(self):
        book = Workbook()
        hoja = book.active
        hoja.append(['LABORATORIO','DESCRIPCION','INDICACIONES'])
        hoja.append(['BIOQUÍMICA','GLUCOSA','Ayuno 8 h'])
        hoja.append(['MICROBIOLOGÍA','UROCULTIVO','Sin ayuno'])
        buffer=io.BytesIO();book.save(buffer)
        with sqlite3.connect(self.db) as con:
            con.execute('''INSERT INTO documentos_fuente(fuente_id,nombre_archivo,tipo_mime,archivo,sha256,fecha_actualizacion,registros)
                         VALUES ('cartera_examenes','crs.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',?,'test','2026-11-07',2)''',
                        (buffer.getvalue(),))
        self.assertEqual(crs.migrar_laboratorio_cartera(self.db),2)
        self.assertEqual(crs.migrar_laboratorio_cartera(self.db),0)  # idempotente
        self.assertEqual(crs.coincidencia_exacta_cartera(self.db,'glucosa','Bioquímica'),'GLUCOSA')
        self.assertIsNone(crs.coincidencia_exacta_cartera(self.db,'glucosa','Microbiología'))
        candidatos = crs.candidatos_cartera(self.db,'glucosa','Bioquímica')
        self.assertEqual(candidatos[0][0], 'GLUCOSA')
        self.assertEqual(candidatos[0][1], 'BIOQUÍMICA')
        with sqlite3.connect(self.db) as con:
            self.assertEqual(con.execute('SELECT COUNT(*) FROM cartera_examenes').fetchone()[0],2)
            self.assertEqual(con.execute('SELECT COUNT(*) FROM usuarios_aplicacion').fetchone()[0],1)
