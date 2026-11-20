"""Regresiones: la prestación detectada debe verse en el selector de Streamlit."""
from pathlib import Path
import json
import sqlite3
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import fuentes_documentales as fd


class TestSelectorAcreditacionSincronizado(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.db = str(Path(self.tmp.name) / 'prueba.db')
        fd.inicializar_fuentes(self.db)
        self.registros = [
            ('FIBRINOGENO', 'HEMATOLOGÍA Y COAGULACIÓN'),
            ('FIBRINOGENO – HOSMET', 'DERIVADO EN CONVENIO'),
            ('HEMOGRAMA', 'HEMATOLOGÍA Y COAGULACIÓN'),
        ]
        self._insertar()
        with sqlite3.connect(self.db) as con:
            self.ids = {r[1]:r[0] for r in con.execute(
                'SELECT id,prestacion FROM prestaciones_acreditacion')}
        self.opciones = sorted(self.ids.values())

    def _insertar(self):
        with sqlite3.connect(self.db) as con:
            con.executemany('''INSERT INTO prestaciones_acreditacion (
                fila_excel, identificador, codigo, prestacion, prestacion_norm,
                area, area_norm, datos_json, estados_json
            ) VALUES (?,?,?,?,?,?,?,?,?)''', [
                (i, str(i), str(i), nombre, fd.normalizar(nombre), area,
                 fd.area_equivalente(area), json.dumps({}), json.dumps({}))
                for i, (nombre, area) in enumerate(self.registros, 1)
            ])

    def sincronizar(self, state, examen='FIBRINOGENO', area='Hematología y coagulación',
                    revision=0, opciones=None):
        return fd.sincronizar_vinculo_acreditacion_en_sesion(
            state, self.db, examen, area,
            self.opciones if opciones is None else opciones, revision
        )

    def test_primer_rerun_preselecciona_exacto_y_usa_mismo_id_para_extraccion(self):
        state = {'acreditacion_manual_selector': None}
        valor = self.sincronizar(state)
        self.assertEqual(valor, self.ids['FIBRINOGENO'])
        self.assertEqual(state['acreditacion_manual_selector'], self.ids['FIBRINOGENO'])
        self.assertEqual(state['acreditacion_manual_id'], self.ids['FIBRINOGENO'])
        self.assertEqual(state['acreditacion_coincidencia_automatica'], 'exacta')
        _, _, vinculo, _ = fd.obtener_valores_fuentes(
            self.db, 'FIBRINOGENO', 'Hematología y coagulación', valor)
        self.assertEqual(vinculo['prestacion'], 'FIBRINOGENO')

    def test_repara_selector_vacio_obsoleto_incluso_sin_cambio_examen(self):
        state = {'acreditacion_manual_selector': None}
        self.sincronizar(state)
        state['acreditacion_manual_selector'] = None
        self.assertEqual(self.sincronizar(state), self.ids['FIBRINOGENO'])

    def test_un_candidato_similar_no_se_vincula_automaticamente(self):
        state = {}
        self.assertIsNone(self.sincronizar(state, examen='FIBRINOGEN'))
        self.assertIsNone(state['acreditacion_manual_selector'])

    def test_cambiar_manual_a_variante_se_respeta_entre_reruns(self):
        state = {}
        self.sincronizar(state)
        state['acreditacion_manual_selector'] = self.ids['FIBRINOGENO – HOSMET']
        state['acreditacion_seleccion_confirmada_manualmente'] = True
        self.assertEqual(self.sincronizar(state), self.ids['FIBRINOGENO – HOSMET'])
        self.assertIsNone(state['acreditacion_coincidencia_automatica'])

    def test_borrar_vinculo_manual_no_lo_restaurara(self):
        state = {}
        self.sincronizar(state)
        state['acreditacion_manual_selector'] = None
        state['acreditacion_seleccion_confirmada_manualmente'] = True
        self.assertIsNone(self.sincronizar(state))
        self.assertIsNone(state['acreditacion_manual_id'])

    def test_cambiar_examen_recalcula_y_no_hereda_seleccion(self):
        state = {}
        self.sincronizar(state)
        state['acreditacion_manual_selector'] = self.ids['FIBRINOGENO – HOSMET']
        state['acreditacion_seleccion_confirmada_manualmente'] = True
        self.assertEqual(self.sincronizar(state, examen='HEMOGRAMA'), self.ids['HEMOGRAMA'])
        self.assertFalse(state['acreditacion_seleccion_confirmada_manualmente'])

    def test_area_incompatible_no_autovincula(self):
        state = {}
        self.assertIsNone(self.sincronizar(state, area='Microbiología'))

    def test_si_cambia_revision_preserva_seleccion_manual_disponible(self):
        state = {}
        self.sincronizar(state)
        state['acreditacion_manual_selector'] = self.ids['FIBRINOGENO – HOSMET']
        state['acreditacion_seleccion_confirmada_manualmente'] = True
        self.assertEqual(self.sincronizar(state, revision=2), self.ids['FIBRINOGENO – HOSMET'])

    def test_estado_heredado_manual_se_respeta(self):
        state = {'acreditacion_manual_selector': self.ids['FIBRINOGENO – HOSMET'],
                 'acreditacion_manual_id': self.ids['FIBRINOGENO – HOSMET']}
        self.assertEqual(self.sincronizar(state), self.ids['FIBRINOGENO – HOSMET'])
        self.assertTrue(state['acreditacion_seleccion_confirmada_manualmente'])

    def test_si_desaparece_alternativa_manual_retorna_exacto(self):
        state = {}
        self.sincronizar(state)
        state['acreditacion_manual_selector'] = self.ids['FIBRINOGENO – HOSMET']
        state['acreditacion_seleccion_confirmada_manualmente'] = True
        self.assertEqual(self.sincronizar(state, revision=1,
                opciones=[self.ids['FIBRINOGENO']]), self.ids['FIBRINOGENO'])

    def test_varios_exactos_elige_area_o_permanece_vacio_si_ambigua(self):
        with sqlite3.connect(self.db) as con:
            con.execute('''INSERT INTO prestaciones_acreditacion (
                fila_excel, identificador, codigo, prestacion, prestacion_norm,
                area, area_norm, datos_json, estados_json
            ) VALUES (?,?,?,?,?,?,?,?,?)''',
                (99,'99','99','FIBRINÓGENO',fd.normalizar('FIBRINÓGENO'),
                 'BIOQUÍMICA',fd.area_equivalente('BIOQUÍMICA'),'{}','{}'))
            bio = con.execute('SELECT id FROM prestaciones_acreditacion WHERE fila_excel=99').fetchone()[0]
        opts = self.opciones + [bio]
        self.assertEqual(self.sincronizar({},opciones=opts), self.ids['FIBRINOGENO'])
        self.assertEqual(self.sincronizar({},area='Bioquímica',opciones=opts), bio)
        self.assertIsNone(self.sincronizar({},area='Microbiología',opciones=opts))

    def test_app_coloca_el_valor_antes_del_selectbox_y_declara_on_change(self):
        codigo = (ROOT/'app.py').read_text(encoding='utf8')
        segmento = codigo[codigo.index('st.markdown("##### Emparejamiento con **Cartera de acreditación**")'):]
        self.assertLess(segmento.index('fd.sincronizar_vinculo_acreditacion_en_sesion('),
                        segmento.index('identificador_manual = st.selectbox('))
        self.assertIn('on_change=_cambiar_selector_acreditacion', segmento)
        self.assertIn('acreditacion_coincidencia_automatica', segmento)

if __name__ == '__main__':
    unittest.main()
