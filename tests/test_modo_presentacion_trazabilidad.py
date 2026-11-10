"""Regresión para la ocultación reversible del historial en una demostración."""
from pathlib import Path
import ast
import unittest

APP = (Path(__file__).resolve().parents[1] / 'app.py').read_text(encoding='utf-8')


class TestModoPresentacion(unittest.TestCase):
    def test_historial_y_actividad_condicionales(self):
        ast.parse(APP)
        sidebar = APP.split('# --- SIDEBAR ---', 1)[1].split('# --- VISTA PRINCIPAL ---', 1)[0]
        self.assertIn('key="modo_presentacion"', sidebar)
        bloque = sidebar.split('if not st.session_state.get("modo_presentacion", False):', 1)[1]
        bloque = bloque.split('    if st.button("Cerrar sesión"):', 1)[0]
        self.assertIn('st.header("Trazabilidad")', bloque)
        self.assertIn('gp.historial(DB_PATH, limite=30)', bloque)
        self.assertIn('gp.historial_actividad(DB_PATH, limite=25)', bloque)

    def test_no_se_borran_auditorias(self):
        self.assertNotIn('DELETE FROM auditoria_cambios', APP)
        self.assertNotIn('DELETE FROM eventos_acceso', APP)
        self.assertIn('Los registros de auditoría y de acceso se conservan en SQLite.', APP)


if __name__ == '__main__':
    unittest.main()
