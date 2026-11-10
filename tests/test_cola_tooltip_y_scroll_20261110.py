"""Regresiones de mensajería, ayuda condicional y scroll sin cronómetro."""
import ast
from pathlib import Path
import unittest

APP = (Path(__file__).resolve().parents[1] / 'app.py').read_text(encoding='utf-8')
TREE = ast.parse(APP)


class MensajeColaYTareasDeInterfaz(unittest.TestCase):
    def test_mensaje_colectivo_solo_para_lotes(self):
        self.assertIn('if total_pdfs > 1:', APP)
        self.assertIn('¡Has completado la cola de exámenes! Descarga o actualiza la documentación.', APP)
        self.assertIn('elif total_pdfs == 1:', APP)
        self.assertIn('¡Has completado el examen! Descarga o actualiza la documentación.', APP)
        self.assertNotIn('Puedes descargar el Excel consolidado con el botón de la izquierda.', APP)

    def test_tooltip_condicional_al_estado_de_discrepancias(self):
        partes = APP.split('with col_update:\n', 1)
        self.assertEqual(len(partes), 2)
        bloque = partes[1].split('with col_next:', 1)[0]
        self.assertIn('st.session_state.get("discrepancias_pdf")', bloque)
        self.assertIn('st.session_state.get("conflictos_pendientes")', bloque)
        self.assertIn('disabled=discrepancias_sin_revisar', bloque)
        self.assertIn('if discrepancias_sin_revisar else None', bloque)
        self.assertIn('Comprobar discrepancias antes de actualizar.', bloque)

    def test_fragmento_flotante_no_se_ejecuta_periodicamente(self):
        boton = next(n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name == '_reinicio_flotante_rapido')
        self.assertEqual(len(boton.decorator_list), 1)
        self.assertIsInstance(boton.decorator_list[0], ast.Attribute)
        self.assertEqual(boton.decorator_list[0].attr, 'fragment')
        self.assertNotIn('run_every=', APP)
        self.assertIn('animation-iteration-count: 3 !important;', APP)
        self.assertIn('reinicio_flotante_intento', APP)
        self.assertIn('reinicio_flecha_giratoria_alternativa', APP)

    def test_no_altera_config_de_pdf_y_referencias(self):
        self.assertIn('@st.fragment\ndef _visor_pdf_rapido', APP)
        self.assertIn('apa.ordenar_y_unir(refs_pdf)', APP)
        self.assertIn('validar_excel_para_descarga(', APP)


if __name__ == '__main__':
    unittest.main()
