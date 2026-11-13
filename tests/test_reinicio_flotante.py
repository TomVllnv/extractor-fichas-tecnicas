"""Pruebas del reinicio rápido de dos pulsaciones sin ejecutar Streamlit."""
import ast
from pathlib import Path
import types
import unittest

PATH = Path(__file__).resolve().parents[1] / 'app.py'
SOURCE = PATH.read_text(encoding='utf-8')
TREE = ast.parse(SOURCE)
NOMBRES = {
    'reiniciar_trabajo_por_cambio_area', '_iniciar_reinicio_flotante',
    '_confirmar_reinicio_flotante', '_pulsar_reinicio_flotante',
    '_actualizar_tiempo_reinicio_flotante', '_vincular_candidato_cartera',
    '_seleccionar_candidato_acreditacion'
}
CONSTANTES = {'AREA_CONFIG', 'CLAVES_TRABAJO_AREA',
              'SEGUNDOS_CONFIRMACION_REINICIO', 'SEGUNDOS_FLASH_CONFIRMACION'}


def aislar_funciones(estado, reloj):
    nodos = []
    for nodo in TREE.body:
        if isinstance(nodo, ast.FunctionDef) and nodo.name in NOMBRES:
            nodos.append(nodo)
        if isinstance(nodo, ast.Assign) and any(
                isinstance(x, ast.Name) and x.id in CONSTANTES for x in nodo.targets):
            nodos.append(nodo)
    env = {'st': types.SimpleNamespace(session_state=estado),
           'time': types.SimpleNamespace(monotonic=lambda: reloj[0])}
    exec(compile(ast.Module(body=nodos, type_ignores=[]), str(PATH), 'exec'), env)
    return env


class ReinicioFlotanteTest(unittest.TestCase):
    def setUp(self):
        self.reloj = [100.0]
        self.estado = {
            'area_activa': 'Microbiología', 'usuario_acceso': 'ADMIN',
            'authenticated': True, 'examen_activo': 'Urocultivo',
            'result_bytes': b'excel temporal', 'orden_nombres_pdf': ['examen.pdf'],
            'acreditacion_manual_selector': 123, 'generacion_cargas_area': 9,
            'aggrid_key_counter': 2, 'examenes_por_area': {'Microbiología': ['Urocultivo']},
        }
        self.f = aislar_funciones(self.estado,self.reloj)

    def test_primera_pulsacion_solo_activa_giro(self):
        self.f['_pulsar_reinicio_flotante']()
        self.assertEqual(self.estado['estado_reinicio_flotante'],'confirmar')
        self.assertEqual(self.estado['vencimiento_reinicio_flotante'],103.0)
        self.assertIn('result_bytes',self.estado)

    def test_dos_pulsaciones_inmediatas_reinician_sesion_sin_cambiar_area(self):
        self.f['_pulsar_reinicio_flotante']()
        self.reloj[0] += 0.015  # 15 ms: antes incluso de un rerun de la página
        self.f['_pulsar_reinicio_flotante']()
        self.assertEqual(self.estado['estado_reinicio_flotante'],'completado')
        self.assertIsNone(self.estado['fin_flash_reinicio_flotante'])
        self.assertTrue(self.estado['reinicio_flotante_refrescar_pantalla'])
        self.assertEqual(self.estado['area_activa'],'Microbiología')
        self.assertTrue(self.estado['authenticated'])
        self.assertEqual(self.estado['generacion_cargas_area'],10)
        self.assertNotIn('result_bytes',self.estado)
        self.assertNotIn('orden_nombres_pdf',self.estado)
        self.assertIn('examenes_por_area',self.estado)

    def test_segunda_pulsacion_dentro_de_los_tres_segundos(self):
        self.f['_pulsar_reinicio_flotante']()
        self.reloj[0] += 2.9
        self.f['_pulsar_reinicio_flotante']()
        self.assertEqual(self.estado['estado_reinicio_flotante'],'completado')

    def test_expiracion_sin_perder_datos(self):
        self.f['_pulsar_reinicio_flotante']()
        self.reloj[0] += 3.01
        self.assertEqual(self.f['_actualizar_tiempo_reinicio_flotante'](),'inactivo')
        self.assertIn('examen_activo',self.estado)
        self.assertIn('result_bytes',self.estado)

    def test_clic_fuera_del_plazo_abre_nueva_confirmacion_y_no_borra(self):
        self.f['_pulsar_reinicio_flotante']()
        self.reloj[0] += 3.2
        self.f['_pulsar_reinicio_flotante']()
        self.assertEqual(self.estado['estado_reinicio_flotante'],'confirmar')
        self.assertIn('result_bytes',self.estado)

    def test_giro_y_confirmacion_visual_sin_textos_adicionales(self):
        self.assertIn('position: fixed !important;', SOURCE)
        self.assertIn('right: 22px;', SOURCE)
        self.assertIn('bottom: 78px;', SOURCE)
        self.assertIn('right: 12px; bottom: 96px;', SOURCE)
        self.assertIn('background: #f8fafc', SOURCE)
        self.assertIn('color: #111827', SOURCE)
        self.assertIn('reinicio_flecha_giratoria 0.9s linear infinite', SOURCE)
        self.assertIn('reinicio_confirmacion_suave 1.5s', SOURCE)
        self.assertIn('help="Pulsa dos veces para reiniciar"', SOURCE)
        self.assertIn('key="boton_reinicio_flotante"',SOURCE)
        self.assertIn('on_click=_pulsar_reinicio_flotante',SOURCE)
        self.assertIn('@st.fragment\ndef _reinicio_flotante_rapido():', SOURCE)
        self.assertNotIn('run_every=1.0', SOURCE)
        self.assertIn('animation-iteration-count: 3 !important;', SOURCE)
        self.assertNotIn('reinicio_flotante_cancelado',SOURCE)
        self.assertIn('if sugeridas:', SOURCE)
        self.assertIn('_seleccionar_candidato_acreditacion, identificador', SOURCE)

    def test_avisos_compactos(self):
        self.assertIn('div[data-testid="stAlert"]', SOURCE)
        self.assertIn('width: fit-content !important;', SOURCE)
        self.assertIn('max-width: 100% !important;', SOURCE)

    def test_sugerencia_crs_y_acreditacion(self):
        self.f['_vincular_candidato_cartera']('VITAMINA D')
        self.assertEqual(self.estado['cartera_examen_emparejado'],'VITAMINA D')
        self.estado['_acreditacion_reset_pendiente'] = True
        self.f['_seleccionar_candidato_acreditacion'](143)
        self.assertFalse(self.estado['_acreditacion_reset_pendiente'])
        self.assertEqual(self.estado['acreditacion_manual_id'],143)


if __name__ == '__main__':
    unittest.main()
