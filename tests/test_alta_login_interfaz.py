"""Regresión estructural de la nueva opción de alta sin inicio de sesión."""
import ast
from pathlib import Path
import unittest

SOURCE = (Path(__file__).resolve().parents[1] / 'app.py').read_text(encoding='utf-8')
TREE = ast.parse(SOURCE)


class InterfazAltaTest(unittest.TestCase):
    def test_login_muestra_crear_usuario_y_modal_protegido(self):
        self.assertIn('"Crear nuevo usuario", use_container_width=True',SOURCE)
        self.assertIn('@st.dialog("Crear nuevo usuario", on_dismiss=_cerrar_alta_usuario_login)',SOURCE)
        self.assertIn('"Usuario del administrador"',SOURCE)
        self.assertIn('"Contraseña del administrador", type="password"',SOURCE)
        self.assertIn('rul.verificar_administrador(DB_PATH, admin_usuario, admin_clave)',SOURCE)
        self.assertIn('rul.autorizacion_vigente(DB_PATH, autorizacion)',SOURCE)
        self.assertIn('rul.crear_encargado_autorizado(',SOURCE)

    def test_requiere_clave_repetida_usa_respaldo_y_no_guarda_admin_en_session_state(self):
        self.assertIn('"Repetir contraseña nueva", type="password"',SOURCE)
        self.assertIn('rul.validar_datos_nueva_cuenta(',SOURCE)
        self.assertIn('gp.respaldar(DB_PATH)',SOURCE)
        self.assertNotIn('st.session_state["admin_clave"]',SOURCE)
        self.assertNotIn('st.session_state["alta_admin_clave"]',SOURCE)

    def test_reinicio_con_un_unico_boton_y_ayuda(self):
        self.assertEqual(SOURCE.count('key="boton_reinicio_flotante"'),1)
        self.assertIn('help="Pulsa dos veces para reiniciar"',SOURCE)
        self.assertIn('on_click=_pulsar_reinicio_flotante',SOURCE)
        self.assertIn('st.rerun(scope="app")',SOURCE)
        self.assertIn('st.container(key="reinicio_flotante")',SOURCE)
        self.assertIn('estado-reinicio-completado',SOURCE)


if __name__ == '__main__':
    unittest.main()
