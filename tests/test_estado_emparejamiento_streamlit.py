"""Regresiones de la sincronización del selectbox sin necesidad de Streamlit."""
import sqlite3
import tempfile
import unittest
from pathlib import Path

import cartera_crs as crs


class TestEstadoVinculoCRS(unittest.TestCase):
    def setUp(self):
        temporal = tempfile.TemporaryDirectory()
        self.addCleanup(temporal.cleanup)
        self.db = str(Path(temporal.name) / "test.db")
        with sqlite3.connect(self.db) as con:
            con.execute("CREATE TABLE cartera_examenes (nombre_examen TEXT UNIQUE, laboratorio TEXT)")
            con.executemany("INSERT INTO cartera_examenes VALUES (?,?)", [
                ("FIBRINOGENO", "HEMATOLOGÍA Y COAGULACIÓN"),
                ("FIBRINOGENO – HOSMET", "DERIVADO EN CONVENIO"),
                ("HEMOGLOBINA", "HEMATOLOGÍA Y COAGULACIÓN"),
            ])
        self.opciones = ["FIBRINOGENO", "FIBRINOGENO – HOSMET", "HEMOGLOBINA"]
        self.examen = "FIBRINOGENO"
        self.area = "Hematología y coagulación"

    def sincronizar(self, estado, examen=None, area=None, revision=0, opciones=None):
        return crs.sincronizar_vinculo_en_sesion(
            estado, self.db, self.examen if examen is None else examen,
            self.area if area is None else area,
            self.opciones if opciones is None else opciones, revision)

    def test_primera_entrada_vincula_100_y_ofrece_variante(self):
        estado = {"sel_cartera_main": None}
        self.assertEqual(self.sincronizar(estado), "FIBRINOGENO")
        self.assertEqual(estado["cartera_examen_emparejado"], "FIBRINOGENO")
        self.assertEqual(crs.candidatos_variantes_cartera(self.db,self.examen,self.area)[0][0],
                         "FIBRINOGENO – HOSMET")

    def test_repara_selector_vacio_con_examen_ya_activo(self):
        estado = {"contexto_emparejamiento_cartera": self.examen,
                  "sel_cartera_main": None}
        self.assertEqual(self.sincronizar(estado), "FIBRINOGENO")
        estado["sel_cartera_main"] = None  # Estado obsoleto por recarga de UI.
        self.assertEqual(self.sincronizar(estado), "FIBRINOGENO")

    def test_no_reemplaza_variante_elegida_manualmente(self):
        estado = {}
        self.sincronizar(estado)
        estado["sel_cartera_main"] = "FIBRINOGENO – HOSMET"
        estado["cartera_seleccion_confirmada_manualmente"] = True
        self.assertEqual(self.sincronizar(estado), "FIBRINOGENO – HOSMET")
        self.assertEqual(estado["cartera_examen_emparejado"], "FIBRINOGENO – HOSMET")

    def test_dejar_sin_vinculo_manualmente_no_reactiva_exacto(self):
        estado = {}
        self.sincronizar(estado)
        estado["sel_cartera_main"] = None
        estado["cartera_seleccion_confirmada_manualmente"] = True
        self.assertIsNone(self.sincronizar(estado))
        self.assertEqual(estado["cartera_examen_emparejado"], "Sin vincular")

    def test_cambiar_examen_o_area_recalcula(self):
        estado = {}
        self.sincronizar(estado)
        estado["sel_cartera_main"] = "FIBRINOGENO – HOSMET"
        estado["cartera_seleccion_confirmada_manualmente"] = True
        self.assertEqual(self.sincronizar(estado, examen="HEMOGLOBINA"), "HEMOGLOBINA")
        self.assertFalse(estado["cartera_seleccion_confirmada_manualmente"])
        # La excepción existente de nombre 100 % único + variante permite
        # vincular incluso si el área de trabajo es distinta.
        self.assertEqual(self.sincronizar(estado, area="Microbiología"), "FIBRINOGENO")

    def test_migracion_desde_contexto_antiguo_respeta_vinculo_manual_valido(self):
        estado = {"contexto_emparejamiento_cartera": "FIBRINOGENO",
                  "sel_cartera_main": "FIBRINOGENO – HOSMET"}
        self.assertEqual(self.sincronizar(estado), "FIBRINOGENO – HOSMET")
        self.assertTrue(estado["cartera_seleccion_confirmada_manualmente"])

    def test_revision_cartera_resincroniza_pero_preserva_eleccion_manual(self):
        estado = {}
        self.sincronizar(estado)
        self.assertEqual(self.sincronizar(estado, revision=1), "FIBRINOGENO")
        estado["sel_cartera_main"] = "FIBRINOGENO – HOSMET"
        estado["cartera_seleccion_confirmada_manualmente"] = True
        self.assertEqual(self.sincronizar(estado, revision=2), "FIBRINOGENO – HOSMET")

    def test_no_vincula_difusos_en_ausencia_de_exacto(self):
        estado = {}
        self.assertIsNone(self.sincronizar(estado, examen="FIBRINOGEN"))
        self.assertEqual(estado["cartera_examen_emparejado"], "Sin vincular")

    def test_recalcula_exacto_si_se_elimina_alternativa_del_catalogo(self):
        estado = {}
        self.sincronizar(estado)
        estado["sel_cartera_main"] = "FIBRINOGENO – HOSMET"
        estado["cartera_seleccion_confirmada_manualmente"] = True
        self.assertEqual(self.sincronizar(estado, revision=1,
                                         opciones=["FIBRINOGENO", "HEMOGLOBINA"]), "FIBRINOGENO")


if __name__ == "__main__":
    unittest.main()
