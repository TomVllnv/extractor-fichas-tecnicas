"""Pruebas específicas de la compacidad, visor rápido y tooltips al pasar el mouse."""
import ast
import base64
import hashlib
import unittest
from pathlib import Path
from types import SimpleNamespace

import fitz

ROOT = Path(__file__).resolve().parents[1]
APP = (ROOT / "app.py").read_text(encoding="utf-8")
ESTILOS = (ROOT / "coincidencias_ui.py").read_text(encoding="utf-8")


class TestVistaCompacta(unittest.TestCase):
    def test_ambas_carteras_ocupan_solo_mitad_y_sin_lineas(self):
        self.assertEqual(APP.count('columna_sugerencias, _ = st.columns([1, 1], gap="small")'), 2)
        self.assertIn('key="lista_coincidencias_crs"', APP)
        self.assertIn('key="lista_coincidencias_acreditacion"', APP)
        self.assertNotIn('st.markdown(\'<hr class="separador-emparejamientos">\'', APP)
        self.assertIn('gap: 0.28rem !important', APP)
        self.assertIn('st.html(cui.estilos_coincidencia(', APP)
        self.assertIn('text-align: left !important', ESTILOS)

    def test_titulo_oscuro_con_fondo_legible_y_tarjeta_seleccionable(self):
        self.assertIn('titulo_oscuro = "#12518a"', APP)
        self.assertIn('else "#166534"', APP)
        self.assertIn('fondo_titulo = "#e8f4ff"', APP)
        self.assertIn('else "#e6f8ed"', APP)
        self.assertIn('button p strong:first-of-type', APP)
        self.assertIn('st.button(etiqueta, key=clave, use_container_width=True, on_click=on_click', APP)
        self.assertNotIn('key=f"elegir_', APP)

    def test_screentips_solo_al_pasarse_mouse_sin_temporizador(self):
        self.assertIn('[role="tooltip"]', APP)
        self.assertIn('animation: none !important;', APP)
        self.assertIn('transition: none !important;', APP)
        self.assertIn('pointer-events: none !important;', APP)
        self.assertNotIn('ocultar_screentip_4s', APP)

    def test_botones_clave_sin_ayuda_emergente(self):
        self.assertNotIn('help="Descarga el libro Excel con todas las hojas', APP)
        self.assertNotIn('help="Resuelve primero las discrepancias pendientes.', APP)
        self.assertIn('key="btn_continuar_siguiente_examen", help=None', APP)

    def test_visor_es_fragmento_y_solo_hay_flechas_laterales(self):
        self.assertIn('@st.fragment\ndef _visor_pdf_rapido', APP)
        self.assertIn('_visor_pdf_rapido(pdf_activo, modo_conjunto, total_pdfs, pdfs_ordenados, idx_actual)', APP)
        self.assertEqual(APP.count('st.button("❮"'), 1)
        self.assertEqual(APP.count('st.button("❯"'), 1)
        self.assertIn('pdf_preview_imagen_cache', APP)
        self.assertNotIn('@st.cache_data(show_spinner=False, max_entries=60)\ndef _imagen_pagina_pdf_cacheada', APP)


class TestCacheVisor(unittest.TestCase):
    def test_cache_de_paginas_recuerda_imagenes_en_sesion_sin_modificar_pdf(self):
        raiz = ast.parse(APP)
        nombres = {"_total_paginas_pdf_cacheadas", "_imagen_pagina_pdf_cacheada"}
        defs = [n for n in raiz.body if isinstance(n, ast.FunctionDef) and n.name in nombres]
        self.assertEqual(len(defs), 2)
        sesion = {}
        scope = {"st": SimpleNamespace(session_state=sesion), "fitz": fitz,
                 "hashlib": hashlib, "base64": base64}
        exec(compile(ast.Module(body=defs, type_ignores=[]), "app.py", "exec"), scope)
        documento = fitz.open()
        for _ in range(2):
            pagina = documento.new_page(width=280, height=350)
            pagina.insert_text((18, 34), f"PÁGINA {pagina.number + 1}")
        original = documento.tobytes()
        documento.close()
        self.assertEqual(scope["_total_paginas_pdf_cacheadas"](original), 2)
        uno = scope["_imagen_pagina_pdf_cacheada"](original, 0, 130)
        dos = scope["_imagen_pagina_pdf_cacheada"](original, 1, 130)
        self.assertNotEqual(uno[0], dos[0])
        self.assertIs(uno, scope["_imagen_pagina_pdf_cacheada"](original, 0, 130))
        self.assertEqual(len(sesion["pdf_preview_imagen_cache"]), 2)
        self.assertEqual(original[:4], b"%PDF")


if __name__ == '__main__':
    unittest.main()
