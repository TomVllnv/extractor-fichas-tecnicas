"""Verifica datos reales de portada y contraportada para el campo Referencias."""
from __future__ import annotations

import io
from pathlib import Path
import unittest

import fitz
import referencias_apa as apa


def fabricar_pdf(titulo: str, con_roche: bool = True, metadata: bool = False) -> bytes:
    doc = fitz.open()
    portada = doc.new_page(width=612, height=792)
    portada.insert_text((38, 35), "071534735004.0", fontsize=8)
    portada.insert_text((38, 90), titulo, fontsize=30)
    portada.insert_text((472, 90), "cobas", fontsize=31)
    portada.insert_text((38, 170), "Uso previsto", fontsize=11)
    portada.insert_text((38, 190), "Datos clínicos y analíticos del documento.", fontsize=10)
    final = doc.new_page(width=612, height=792)
    if con_roche:
        final.insert_text((38, 705), "2024, Roche Diagnostics", fontsize=10)
        final.insert_text((38, 724), "Roche Diagnostics GmbH, Sandhofer Strasse 116, D-68305 Mannheim", fontsize=10)
        final.insert_text((38, 743), "www.roche.com", fontsize=10)
    if metadata:
        doc.set_metadata({"title": "Documento genérico", "author": "User"})
    contenido = doc.tobytes()
    doc.close()
    return contenido


class TestRocheSoloParaReferencias(unittest.TestCase):
    def test_recupera_titulo_autor_fecha_editorial_y_dominio(self):
        fuente = apa.fuente_pdf("07153473500.pdf", fabricar_pdf("Fibrinogen", metadata=True), "FIBRINOGENO")
        self.assertEqual(fuente.titulo, "Fibrinogen")
        self.assertEqual(fuente.autor, "Roche Diagnostics")
        self.assertEqual(fuente.anio, "2024")
        self.assertEqual(fuente.datos_editoriales, "Roche Diagnostics GmbH")
        self.assertEqual(fuente.url, "roche.com")
        self.assertFalse(fuente.titulo_provisional)
        self.assertEqual(fuente.tipo, "")
        texto = apa.ordenar_y_unir([fuente])
        self.assertIn("Roche Diagnostics. (2024). Fibrinogen.", texto)
        self.assertIn("Roche Diagnostics GmbH", texto)
        self.assertIn("roche.com", texto)
        self.assertNotIn("www.", texto)
        self.assertNotIn("Sandhofer", texto)
        self.assertNotIn("Mannheim", texto)
        self.assertNotIn("https://", texto)
        self.assertNotIn("[Ficha técnica de", texto)
        self.assertNotIn("[Documento PDF]", texto)
        self.assertNotIn("FIBRINOGENO", texto)
        self.assertIn("<em>Fibrinogen</em>", apa.ordenar_y_unir_html([fuente]))
        self.assertEqual(apa.referencias_excel_con_cursiva([fuente]), texto)

    def test_urls_institucionales_se_muestran_como_dominio(self):
        fuente = apa.FuenteBibliografica(
            "institucional", "Manual de muestras", autor="CRS", anio="2024",
            url="https://www.ejemplo.cl/archivos/guia.pdf",
        )
        ref = apa.referencia_apa(fuente)
        self.assertTrue(ref.endswith("ejemplo.cl"))
        self.assertNotIn("www.", ref)
        self.assertNotIn("https://", ref)

    def test_cambia_titulo_según_pdf_y_no_inventa_resto(self):
        fuente = apa.fuente_pdf("ficha_19482.pdf", fabricar_pdf("Glucose", False), "GLUCOSA")
        self.assertEqual(fuente.titulo, "Glucose")
        self.assertEqual(fuente.autor, "")
        self.assertEqual(fuente.anio, "")
        self.assertEqual(fuente.datos_editoriales, "")
        self.assertEqual(fuente.url, "")
        self.assertEqual(fuente.tipo, "")
        self.assertIn("(s. f.)", apa.referencia_apa(fuente))
        self.assertNotIn("GLUCOSA", apa.referencia_apa(fuente))

    def test_sin_texto_valido_mantiene_referencia_provisional(self):
        doc = fitz.open()
        doc.new_page()
        datos = doc.tobytes()
        doc.close()
        fuente = apa.fuente_pdf("ficha_desconocida.pdf", datos, "UN EXAMEN")
        self.assertEqual(fuente.titulo, "ficha desconocida")
        self.assertTrue(fuente.titulo_provisional)
        self.assertEqual(fuente.tipo, "")
        self.assertEqual(fuente.anio, "")

    def test_cache_bibliografica_se_actualiza_y_modo_presentacion_fuera(self):
        codigo = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
        self.assertIn('"bibliografia_v4"', codigo)
        self.assertNotIn('modo_presentacion', codigo)
        self.assertNotIn('Modo presentación', codigo)


if __name__ == "__main__":
    unittest.main()
