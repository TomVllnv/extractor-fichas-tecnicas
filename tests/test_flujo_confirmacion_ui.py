"""Verifica rutas de escritura de interfaz sin requerir el paquete Streamlit."""
from pathlib import Path
import ast
import unittest

APP = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf8")


class TestConfirmacionUI(unittest.TestCase):
    def test_no_existe_clave_compartida_hardcodeada(self):
        self.assertNotIn('DEFAULT_PASSWORD = "CRSCODUOC"', APP)
        self.assertIn('gp.autenticar(DB_PATH, username, password)', APP)
        self.assertIn('gp.usuario_habilitado(DB_PATH', APP)

    def test_modal_exige_checkbox_y_detecta_revision(self):
        inicio = APP.index('def modal_confirmar_cambio_compartido():')
        fin = APP.index('# --- DEFINICIÓN DE DIÁLOGOS EMERGENTES', inicio)
        texto = APP[inicio:fin]
        self.assertIn('st.checkbox(', texto)
        self.assertIn('disabled=not marcado or conflicto', texto)
        self.assertIn('vigente["revision"] != p["revision"]', texto)
        self.assertIn('anterior[\'usuario\']', texto)

    def test_sidebar_solicita_cambio_y_no_guarda_directamente(self):
        sidebar = APP.split('# --- SIDEBAR ---',1)[1].split('# --- VISTA PRINCIPAL ---',1)[0]
        for tipo in ('"area"', '"acreditacion"', '"cartera"', '"manual"', '"normas"'):
            self.assertIn('solicitar_cambio(' + tipo, sidebar)
        for nombre in ('guardar_plantilla_area_db(', 'guardar_cartera_db(',
                       'fd.guardar_normas(', 'fd.guardar_manual_toma(', 'fd.guardar_cartera_acreditacion('):
            self.assertNotIn(nombre, sidebar)

    def test_editor_de_examenes_catalogos_y_ficha_generada_protegidos(self):
        self.assertIn('solicitar_cambio("examenes", "area:" + area,', APP)
        self.assertIn('solicitar_cambio("catalogo", "catalogo:laboratorios"', APP)
        self.assertIn('solicitar_cambio("catalogo", "catalogo:encargados"', APP)
        self.assertIn('validado = st.checkbox(', APP)
        self.assertIn('gp.ejecutar_cambio(DB_PATH, recurso, st.session_state["usuario_acceso"]', APP)

    def test_acceso_admin_y_separador_de_trazabilidad(self):
        self.assertIn('with st.expander("Acceso", expanded=False)', APP)
        self.assertIn('gp.rol_usuario(DB_PATH, st.session_state.get("usuario_acceso"))', APP)
        self.assertIn('if rol_actual == "administrador":', APP)
        self.assertIn('@st.dialog("Administración de usuarios", on_dismiss=_cerrar_administracion_por_x)', APP)
        self.assertIn('gp.es_administrador(DB_PATH, actor)', APP)
        self.assertIn('gp.administrar_cuenta(', APP)
        self.assertIn('st.header("Trazabilidad")', APP)
        self.assertLess(APP.index('st.header("Base de datos (SQLite)")'),
                        APP.index('st.header("Trazabilidad")'))
        self.assertLess(APP.index('st.header("Trazabilidad")'),
                        APP.index('with st.expander("Historial de modificaciones"'))

    def test_usa_disco_configurable_y_historial(self):
        self.assertIn('DB_PATH = gp.ruta_base()', APP)
        self.assertIn('"Historial de modificaciones"', APP)
        self.assertIn('"Actividad de usuarios"', APP)


if __name__ == "__main__":
    unittest.main()
