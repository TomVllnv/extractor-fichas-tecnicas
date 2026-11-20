"""Pruebas del reinicio por cambio de área, sin instalar ni ejecutar Streamlit."""
import ast
from pathlib import Path
import unittest

RUTA = Path(__file__).resolve().parents[1] / 'app.py'
TEXTO = RUTA.read_text(encoding='utf-8')
ARBOL = ast.parse(TEXTO)


def funciones_aisladas():
    nombres = {'reiniciar_trabajo_por_cambio_area', 'seleccionar_area',
               'progreso_pendiente_cambio_area'}
    asignaciones = {'CLAVES_TRABAJO_AREA', 'AREA_CONFIG'}
    nodos = []
    for nodo in ARBOL.body:
        if isinstance(nodo, ast.FunctionDef) and nodo.name in nombres:
            nodos.append(nodo)
        if isinstance(nodo, ast.Assign) and any(isinstance(t, ast.Name) and t.id in asignaciones for t in nodo.targets):
            nodos.append(nodo)
    espacio = {}
    exec(compile(ast.Module(body=nodos, type_ignores=[]), str(RUTA), 'exec'), espacio)
    return espacio


class StreamlitFalso:
    def __init__(self, estado):
        self.session_state = estado


class TestReinicioArea(unittest.TestCase):
    def setUp(self):
        self.f = funciones_aisladas()
        self.estado = {
            'area_activa': 'Hematología y coagulación',
            'examen_activo': 'Hemograma',
            'examen_resultado': 'Hemograma',
            'data_extracted': {'Valores de referencia': '193-412 mg/dL'},
            'user_edited_fields': {'Valores de referencia'},
            'result_bytes': b'EXCEL_GENERADO',
            'extracted_text': 'texto de un PDF anterior',
            'orden_nombres_pdf': ['hematologia.pdf', 'hemo2.pdf'],
            'cola_pdf_indice': 1,
            'modo_procesamiento_pdf': 'Varios PDF para un mismo examen',
            'sel_cartera_main': 'Hemograma',
            'acreditacion_manual_selector': 92,
            'sel_lab_main': 'Redlab',
            'sel_act_main': 'Operador',
            'sel_aut_main': 'Supervisor',
            'nombre_nuevo_laboratorio': 'Laboratorio de prueba',
            'nombre_nuevo_actualizado': 'Nombre temporal',
            'nombre_nuevo_autorizado': 'Nombre temporal',
            'selector_examen_Hematología y coagulación': 'Hemograma',
            'selector_examen_Microbiología': 'Urocultivo',
            'modal_confirmar_planilla_abierto': True,
            'revision_plantilla_pendiente': 12,
            'aggrid_key_counter': 3,
            'generacion_cargas_area': 5,
            # Los datos centrales de usuarios y SQLite nunca se borran.
            'usuario_acceso': 'ADMIN',
            'authenticated': True,
            'revision_sesion': 8,
            'examenes_por_area': {'Microbiología': ['Urocultivo']},
            'id_ultimo_cambio_compartido': 16,
        }
        self.f['seleccionar_area'].__globals__['st'] = StreamlitFalso(self.estado)
        self.f['reiniciar_trabajo_por_cambio_area'].__globals__['st'] = StreamlitFalso(self.estado)

    def test_cambiar_area_limpia_todo_el_trabajo_y_mantiene_usuario(self):
        self.f['seleccionar_area']('Microbiología')
        # Los campos previamente completados impiden borrar sin confirmación.
        self.assertEqual(self.estado['area_activa'], 'Hematología y coagulación')
        self.assertEqual(self.estado['cambio_area_pendiente']['destino'], 'Microbiología')
        self.assertEqual(self.estado['result_bytes'], b'EXCEL_GENERADO')
        # Simulación de pulsar el botón azul Continuar en el diálogo.
        self.f['reiniciar_trabajo_por_cambio_area']()
        self.estado['area_activa'] = 'Microbiología'
        self.estado.pop('cambio_area_pendiente', None)
        self.assertEqual(self.estado['area_activa'], 'Microbiología')
        self.assertEqual(self.estado['generacion_cargas_area'], 6)
        self.assertEqual(self.estado['aggrid_key_counter'], 4)
        self.assertEqual(self.estado['cola_pdf_indice'], 0)
        for clave in ('examen_activo', 'examen_resultado', 'data_extracted',
                      'result_bytes', 'extracted_text', 'modo_procesamiento_pdf',
                      'sel_cartera_main', 'acreditacion_manual_selector',
                      'sel_lab_main', 'sel_act_main', 'sel_aut_main',
                      'nombre_nuevo_laboratorio', 'nombre_nuevo_actualizado',
                      'nombre_nuevo_autorizado',
                      'selector_examen_Hematología y coagulación',
                      'selector_examen_Microbiología',
                      'modal_confirmar_planilla_abierto', 'revision_plantilla_pendiente'):
            with self.subTest(clave=clave):
                self.assertNotIn(clave, self.estado)
        self.assertEqual(self.estado['usuario_acceso'], 'ADMIN')
        self.assertTrue(self.estado['authenticated'])
        self.assertIn('examenes_por_area', self.estado)
        self.assertEqual(self.estado['revision_sesion'], 8)

    def test_cambio_sin_archivos_ni_campos_es_directo(self):
        self.estado = {'area_activa': 'Microbiología', 'aggrid_key_counter': 0,
                       'generacion_cargas_area': 0, 'usuario_acceso': 'ADMIN'}
        self.f['seleccionar_area'].__globals__['st'] = StreamlitFalso(self.estado)
        self.f['reiniciar_trabajo_por_cambio_area'].__globals__['st'] = StreamlitFalso(self.estado)
        self.f['seleccionar_area']('Bioquímica')
        self.assertEqual(self.estado['area_activa'], 'Bioquímica')
        self.assertNotIn('cambio_area_pendiente', self.estado)

    def test_pdf_adjunto_pide_confirmacion_aunque_no_haya_campos(self):
        self.estado['examen_activo'] = None
        self.estado.pop('data_extracted', None)
        self.estado.pop('result_bytes', None)
        self.estado['pdf_uploader_5'] = [object()]
        self.f['seleccionar_area']('Microbiología')
        pendiente = self.estado['cambio_area_pendiente']
        self.assertEqual(self.estado['area_activa'], 'Hematología y coagulación')
        self.assertTrue(pendiente['archivos'])

    def test_pulsar_la_misma_area_no_descarta_trabajo(self):
        inicial = self.estado.copy()
        self.f['seleccionar_area']('Hematología y coagulación')
        self.assertEqual(self.estado, inicial)

    def test_no_acepta_areas_que_no_existen(self):
        with self.assertRaises(ValueError):
            self.f['seleccionar_area']('Laboratorio desconocido')
        self.assertEqual(self.estado['area_activa'], 'Hematología y coagulación')

    def test_uploaders_se_regeneran_en_vez_de_manipular_el_archivo(self):
        self.assertIn('key=f"pdf_uploader_{st.session_state.get(', TEXTO)
        self.assertIn('key=f"excel_exportacion_uploader_{st.session_state.get(', TEXTO)
        self.assertIn('generacion_cargas_area', TEXTO)
        self.assertNotIn('st.session_state["pdf_uploader"] =', TEXTO)
        self.assertNotIn('st.session_state.pop("pdf_uploader"', TEXTO)

    def test_el_cambio_se_ejecuta_como_callback_de_los_botones(self):
        self.assertIn('on_click=seleccionar_area,', TEXTO)
        self.assertIn('args=(area,),', TEXTO)


if __name__ == '__main__':
    unittest.main()
