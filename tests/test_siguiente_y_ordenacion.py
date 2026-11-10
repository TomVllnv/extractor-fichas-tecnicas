"""Regresiones focalizadas: orden alfabético, avance y cargador opcional."""
import ast
from pathlib import Path
import sqlite3
import unittest
import unicodedata

ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT / "app.py").read_text(encoding="utf-8")
TREE = ast.parse(SOURCE)


def obtener_clave():
    node = next(n for n in TREE.body if isinstance(n, ast.FunctionDef)
                and n.name == "clave_orden_alfabetico")
    ns = {"unicodedata": unicodedata}
    exec(compile(ast.Module(body=[node], type_ignores=[]), str(ROOT / "app.py"), "exec"), ns)
    return ns["clave_orden_alfabetico"]


class TestOrdenYAvance(unittest.TestCase):
    def test_acentos_ordenados_segun_alfabeto_espanol(self):
        clave = obtener_clave()
        self.assertEqual(sorted(["JUEGO", "ÍNDICE ANDROGENICO LIBRE", "HEMOGRAMA", "NITROGENO"], key=clave),
                         ["HEMOGRAMA", "ÍNDICE ANDROGENICO LIBRE", "JUEGO", "NITROGENO"])
        self.assertEqual(clave("ÁCIDO")[0], clave("ACIDO")[0])

    def test_derivados_indice_no_al_final(self):
        with sqlite3.connect(ROOT / "base_de_datos.db") as con:
            nombres = [fila[0] for fila in con.execute("SELECT nombre_examen FROM examenes_area WHERE area='Derivados'")]
        orden = sorted(nombres, key=obtener_clave())
        indice = orden.index("ÍNDICE ANDROGENICO LIBRE")
        self.assertEqual(len(orden), 205)
        self.assertGreater(indice, 0)
        self.assertLess(indice, len(orden) - 1)
        self.assertTrue(orden[indice - 1][0].upper() in "HI")
        self.assertTrue(orden[indice + 1][0].upper() in "IJ")
        self.assertIn("key=clave_orden_alfabetico", SOURCE)

    def test_boton_azul_y_contador_de_documento_actual(self):
        self.assertIn('f"Continuar al siguiente examen ({idx_actual + 1}/{total_pdfs})"', SOURCE)
        self.assertIn('st.session_state["cola_pdf_indice"] = idx_actual + 1', SOURCE)
        self.assertIn('with st.container(key="contenedor_continuar_examen"):', SOURCE)
        self.assertIn('background-color: #2563eb !important;', SOURCE)
        self.assertNotIn('Finalizar y pasar al examen', SOURCE)
        self.assertNotIn('idx_actual + 2}/{total_pdfs}', SOURCE)

    def test_sin_carga_opcional_de_plantilla_en_barra_lateral(self):
        self.assertNotIn('Cargar Excel del área (opcional)', SOURCE)
        self.assertNotIn('template_uploader', SOURCE)
        self.assertNotIn('elif template is not None:', SOURCE)
        self.assertIn('elif plantilla_bytes_db:', SOURCE)
        self.assertIn('template_exportacion = st.file_uploader(', SOURCE)
        self.assertIn('key=f"excel_exportacion_uploader_{st.session_state.get(', SOURCE)


if __name__ == "__main__":
    unittest.main()
