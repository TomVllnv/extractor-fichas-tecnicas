"""Regresiones sin dependencias de Streamlit (la app es una interfaz interactiva)."""
import ast
import hashlib
import os
from pathlib import Path
import re
import shutil
import sqlite3
import tempfile
import unittest
import unicodedata
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import openpyxl
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fuentes_documentales as fd

ROOT = Path(__file__).resolve().parents[1]
APP_SOURCE = (ROOT / 'app.py').read_text(encoding='utf-8')


def seleccionar_funciones():
    """Compilar únicamente funciones de negocio, sin arrancar la interfaz Streamlit."""
    names = {
        'init_db', 'obtener_lista_cartera_db', 'obtener_nombre_plantilla_area_db',
        'guardar_plantilla_area_db', 'guardar_cartera_db',
        'cargar_examenes_por_area', 'obtener_plantilla_area_db',
        'guardar_examenes_por_area', 'norm', 'clave_orden_alfabetico',
        'nombres_hojas_plantilla', 'actualizar_lista_area_conexion',
    }
    root = ast.parse(APP_SOURCE)
    selected = [n for n in root.body if isinstance(n, ast.FunctionDef) and n.name in names]
    code = compile(ast.Module(body=selected, type_ignores=[]), str(ROOT / 'app.py'), 'exec')
    return code


class TestInterfazYMetadatos(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name)
        self.db = self.path / 'base_de_datos.db'
        shutil.copyfile(ROOT / 'base_de_datos.db', self.db)
        self.ns = dict(os=os, sqlite3=sqlite3, hashlib=hashlib, datetime=datetime,
                       ZoneInfo=ZoneInfo, pd=pd, re=re, unicodedata=unicodedata,
                       load_workbook=openpyxl.load_workbook, fd=fd,
                       DB_PATH=str(self.db), __file__=str(self.path / 'app.py'),
                       AREA_CONFIG={
                           'Microbiología': {'template': 'plantilla_microbiologia.xlsx'},
                           'Hematología y coagulación': {'template': 'plantilla_hematologia.xlsx'},
                           'Bioquímica': {'template': 'plantilla_bioquimica.xlsx'},
                           'Hormonas e inmunidad': {'template': None},
                           'Derivados': {'template': None},
                       })
        import io
        self.ns['io'] = io
        exec(seleccionar_funciones(), self.ns)
        self.ns['init_db']()
        fd.inicializar_fuentes(str(self.db))

    def test_393_y_nombres_que_comienzan_en_numero(self):
        nombres = self.ns['obtener_lista_cartera_db']()
        self.assertEqual(len(nombres), 393)
        self.assertIn('17-HIDROXIPROGESTERONA', nombres)
        self.assertIn('17-BETA-ESTRADIOL', nombres)

    def test_migracion_no_borra_plantillas_y_nombres_verificables(self):
        with sqlite3.connect(self.db) as c:
            self.assertIn('nombre_archivo', {r[1] for r in c.execute('PRAGMA table_info(plantillas_area)')})
            self.assertEqual(c.execute('SELECT COUNT(*) FROM plantillas_area').fetchone()[0], 5)
            self.assertEqual(c.execute('SELECT COUNT(*) FROM cartera_examenes').fetchone()[0], 393)
            self.assertEqual(c.execute('SELECT COUNT(*) FROM prestaciones_acreditacion').fetchone()[0], 374)
            self.assertEqual(c.execute('SELECT COUNT(*) FROM normas_generales_campos').fetchone()[0], 4)
        self.assertEqual(self.ns['obtener_nombre_plantilla_area_db']('Microbiología'), 'plantilla_microbiologia.xlsx')
        self.assertEqual(self.ns['obtener_nombre_plantilla_area_db']('Bioquímica'), 'plantilla_bioquimica.xlsx')
        # No inventar un nombre para la plantilla modificada cuyo contenido ya difiere.
        self.assertIsNone(self.ns['obtener_nombre_plantilla_area_db']('Hematología y coagulación'))

    def test_actualizacion_plantilla_actualiza_nombre_sin_alterar_otras(self):
        # La plantilla vive en SQLite: no dependemos de un Excel suelto.
        original, _ = self.ns['obtener_plantilla_area_db']('Hematología y coagulación')
        self.assertTrue(original)
        ok,_ = self.ns['guardar_plantilla_area_db']('Hematología y coagulación',original,'hematologia_2026.xlsx')
        self.assertTrue(ok)
        self.assertEqual(self.ns['obtener_nombre_plantilla_area_db']('Hematología y coagulación'),'hematologia_2026.xlsx')
        ok,_ = self.ns['guardar_plantilla_area_db']('Hematología y coagulación',original,'hematologia_2027.xlsx')
        self.assertTrue(ok)
        self.assertEqual(self.ns['obtener_nombre_plantilla_area_db']('Hematología y coagulación'),'hematologia_2027.xlsx')
        with sqlite3.connect(self.db) as c:
            self.assertEqual(c.execute('SELECT COUNT(*) FROM cartera_examenes').fetchone()[0],393)

    def test_actualizacion_cartera_guarda_nombre_y_original(self):
        original = 'Bytes originales de prueba: la app guarda el original, no una exportación'.encode('utf-8')
        df = pd.DataFrame({'DESCRIPCION': ['Examen de control', 'Otro estudio'],
                           'Codigo Examen': ['A01','B02']})
        ok,_ = self.ns['guardar_cartera_db'](df, 'Cartera_2026.xlsx', original)
        self.assertTrue(ok)
        meta = fd.documento_fuente(str(self.db),'cartera_examenes',incluir_archivo=True)
        self.assertEqual(meta['nombre_archivo'],'Cartera_2026.xlsx')
        self.assertEqual(meta['archivo'], original)
        self.assertEqual(meta['registros'],2)
        self.assertEqual(len(self.ns['obtener_lista_cartera_db']()),2)
        ok,_ = self.ns['guardar_cartera_db'](df,'Cartera_2027.xlsx','Nueva versión'.encode('utf-8'))
        self.assertTrue(ok)
        meta = fd.documento_fuente(str(self.db),'cartera_examenes',incluir_archivo=True)
        self.assertEqual(meta['nombre_archivo'],'Cartera_2027.xlsx')
        self.assertEqual(meta['archivo'], 'Nueva versión'.encode('utf-8'))

    def test_cartera_vacia_no_elimina_datos_existentes(self):
        bad = pd.DataFrame({'DESCRIPCION':['',None,'123']})
        ok, _ = self.ns['guardar_cartera_db'](bad, 'invalido.xlsx',b'invalido')
        self.assertFalse(ok)
        with sqlite3.connect(self.db) as c:
            self.assertEqual(c.execute('SELECT COUNT(*) FROM cartera_examenes').fetchone()[0],393)

    def test_formato_fechas_fuentes_documentales(self):
        """Las dos fechas nuevas se muestran como día/mes/año sin mutar SQLite."""
        arbol = ast.parse(APP_SOURCE)
        funcion = next(n for n in arbol.body if isinstance(n, ast.FunctionDef)
                       and n.name == 'formatear_fecha_fuente')
        espacio = {'datetime': datetime, 'ZoneInfo': ZoneInfo}
        exec(compile(ast.Module(body=[funcion], type_ignores=[]), str(ROOT/'app.py'), 'exec'), espacio)
        formato = espacio['formatear_fecha_fuente']
        self.assertEqual(formato('2026-11-05T23:46:12-03:00'), '05/11/2026 23:46')
        self.assertEqual(formato('2026-11-06T02:46:12+00:00'), '05/11/2026 23:46')
        self.assertEqual(formato(None), 'Sin registro')
        self.assertEqual(formato('fecha sin formato'), 'Fecha no disponible')
        self.assertIn('formatear_fecha_fuente(meta_normas', APP_SOURCE)
        self.assertIn('formatear_fecha_fuente(meta_acr', APP_SOURCE)
        self.assertNotIn("['fecha_actualizacion'][:16].replace('T',' ')", APP_SOURCE)

    def test_normas_independientes_sin_manual(self):
        """Normas generales continúa editable sin el Manual retirado."""
        inicio = APP_SOURCE.index('with st.expander("Normas generales de toma de muestras", expanded=False):')
        fin = APP_SOURCE.index('    st.header("Trazabilidad", anchor=False)', inicio)
        bloque = APP_SOURCE[inicio:fin]
        self.assertNotIn('st.expander("Manual de toma de muestras"', APP_SOURCE)
        self.assertNotIn('tab_manual, tab_normas', APP_SOURCE)
        for nombre in ('descargar_manual_toma_original', 'actualizar_manual_toma_archivo',
                       'guardar_manual_toma_nuevo', 'solicitar_cambio("manual"'):
            self.assertNotIn(nombre, APP_SOURCE)
        for nombre in ('descargar_normas_original', 'actualizar_normas_archivo',
                       'guardar_normas_nuevas', 'solicitar_cambio("normas"'):
            self.assertIn(nombre, bloque)
        self.assertIn("meta_normas['nombre_archivo']", bloque)

    def test_carteras_agrupadas_en_orden_alfabetico_y_descargas_formato_oracion(self):
        """Dos grupos principales, pestañas alfabéticas y botones de descarga con formato oración."""
        inicio = APP_SOURCE.index('with st.expander("Cartera de exámenes y acreditación", expanded=cartera_expirada):')
        fin = APP_SOURCE.index('with st.expander("Normas generales de toma de muestras", expanded=False):', inicio)
        bloque = APP_SOURCE[inicio:fin]
        self.assertIn('tab_acreditacion, tab_examenes = st.tabs(["Acreditación", "Exámenes"])', bloque)
        self.assertIn('with tab_acreditacion:', bloque)
        self.assertIn('with tab_examenes:', bloque)
        self.assertNotIn('with st.expander("Cartera de exámenes",', APP_SOURCE)
        self.assertNotIn('with st.expander("Cartera de acreditación",', APP_SOURCE)
        self.assertLess(inicio, fin)
        acreditacion = bloque.split('with tab_acreditacion:', 1)[1].split('with tab_examenes:', 1)[0]
        examenes = bloque.split('with tab_examenes:', 1)[1]
        for k in ('descargar_acreditacion_original', 'actualizar_acreditacion_archivo',
                  'guardar_acreditacion_nueva'):
            self.assertIn(k, acreditacion)
            self.assertNotIn(k, examenes)
        for k in ('dl_cartera_btn', 'cartera_uploader_sidebar', 'btn_save_cartera'):
            self.assertIn(k, examenes)
            self.assertNotIn(k, acreditacion)
        self.assertIn('with st.expander("Normas generales de toma de muestras", expanded=False):', APP_SOURCE)
        # Solo las etiquetas de los botones de descarga, no el nombre real del archivo.
        arbol = ast.parse(APP_SOURCE)
        esperados = {
            'descargar_acreditacion_original': 'Descargar cartera de acreditación original',
            'dl_cartera_btn': 'Descargar cartera de exámenes',
            'descargar_normas_original': 'Descargar documento original de normas',
        }
        encontrados = set()
        for nodo in ast.walk(arbol):
            if not isinstance(nodo, ast.Call) or not isinstance(nodo.func, ast.Attribute):
                continue
            if nodo.func.attr != 'download_button':
                continue
            opts = {kw.arg: kw.value for kw in nodo.keywords}
            key = opts.get('key')
            if isinstance(key, ast.Constant) and key.value in esperados:
                encontrados.add(key.value)
                self.assertEqual(nodo.args[0].value, esperados[key.value])
            if isinstance(key, ast.JoinedStr) or (isinstance(key, ast.Constant) and
                                                  str(key.value).startswith('dl_area_btn_')):
                self.assertTrue(any(isinstance(x, ast.Constant) and
                                    str(x.value).startswith('Descargar ficha técnica (')
                                    for x in nodo.args[0].values))
        self.assertEqual(encontrados, set(esperados))

    def test_orden_base_datos_areas_antes_de_fuentes_documentales(self):
        """No alterar el orden de las áreas ni el contenido de las pestañas."""
        inicio_db = APP_SOURCE.index('    st.header("Base de datos", anchor=False)')
        area = APP_SOURCE.index('    for area_item in AREA_CONFIG:', inicio_db)
        cartera = APP_SOURCE.index('with st.expander("Cartera de exámenes y acreditación", expanded=cartera_expirada):', inicio_db)
        normas = APP_SOURCE.index('with st.expander("Normas generales de toma de muestras", expanded=False):', inicio_db)
        salida = APP_SOURCE.index('    if st.button("Cerrar sesión"):', inicio_db)
        self.assertLess(inicio_db, area)
        self.assertLess(area, cartera)
        self.assertLess(cartera, normas)
        self.assertLess(normas, salida)
        self.assertIn('for area_item in AREA_CONFIG:', APP_SOURCE[area:cartera])
        self.assertIn('tab_acreditacion, tab_examenes = st.tabs(["Acreditación", "Exámenes"])', APP_SOURCE[cartera:normas])
        self.assertIn('descargar_normas_original', APP_SOURCE[normas:salida])

    def test_administrar_encargados_sin_pestana_duplicada(self):
        """Se conserva una lista compartida para los dos campos de la ficha."""
        self.assertIn('st.tabs(["Laboratorios", "Encargados"])', APP_SOURCE)
        self.assertNotIn('st.tabs(["Laboratorios", "Actualizado por", "Autorizado por"])', APP_SOURCE)
        self.assertNotIn('with tab_m_aut:', APP_SOURCE)
        self.assertIn('new_act = st.text_input("Añadir encargado"', APP_SOURCE)
        self.assertEqual(APP_SOURCE.count('tabla="encargados", nombres='), 2)
        # Los campos de la ficha técnica original se conservan.
        self.assertIn('"Actualizado por",', APP_SOURCE)
        self.assertIn('"Autorizado por",', APP_SOURCE)
        self.assertIn('return _listar_catalogo("encargados")', APP_SOURCE)

    def test_selectores_limpiables_y_placeholders_en_espanol(self):
        parsed = ast.parse(APP_SOURCE)
        selectores = [n for n in ast.walk(parsed) if isinstance(n,ast.Call)
                     and isinstance(n.func,ast.Attribute) and n.func.attr=='selectbox']
        self.assertEqual(len(selectores),17)  # Se eliminaron los dos menús redundantes de la Cartera del CRS
        for selector in selectores:
            kw = {x.arg:x.value for x in selector.keywords}
            self.assertIn('index',kw,f'Selector en línea {selector.lineno} sin index=None')
            self.assertIsInstance(kw['index'],ast.Constant)
            self.assertIsNone(kw['index'].value)
            self.assertIn('placeholder',kw)
            self.assertIsInstance(kw['placeholder'],ast.Constant)
            self.assertTrue(kw['placeholder'].value.startswith('Selecciona'))
        self.assertNotIn('content: "Cargar";',APP_SOURCE)
        self.assertNotIn('[data-testid="stFileUploaderDropzone"] button', APP_SOURCE)
        self.assertIn('placeholder="Selecciona una prestación (opcional)"', APP_SOURCE)
        self.assertIn('Documento almacenado:',APP_SOURCE)


if __name__=='__main__':
    unittest.main()
