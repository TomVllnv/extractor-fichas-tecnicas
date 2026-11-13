"""Regresiones: borrados accidentales y omisión segura de PDFs en cola."""
from pathlib import Path
import io
import sys

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import historial_tabla as ht
import cola_fichas as cf

CODIGO = (ROOT / "app.py").read_text(encoding="utf-8")


def test_borrar_celda_con_suprimir_y_recuperarla_despues_de_rerun():
    historial = ht.contexto_nuevo({}, ("Hematología y coagulación", "PT", 0))
    datos = {"Estabilidad": "8 horas a temperatura ambiente"}
    editados = set()
    borrada = [{"Campo": "Estabilidad", "Información extraída": "", "Estado": "No encontrado"}]
    assert ht.registrar_cambios(historial, datos, editados, borrada)
    datos["Estabilidad"] = ""
    assert ht.aplicar_accion(historial, datos, editados, "deshacer") == ["Estabilidad"]
    assert datos["Estabilidad"] == "8 horas a temperatura ambiente"
    assert "Estabilidad" not in editados
    assert ht.aplicar_accion(historial, datos, editados, "rehacer") == ["Estabilidad"]
    assert datos["Estabilidad"] == ""


def test_restaurar_una_edicion_manual_previa_y_separar_examenes():
    hist = ht.contexto_nuevo({}, ("Área", "Examen 1", 0))
    data = {"Equipamiento utilizado": "cobas 711"}
    editados = {"Equipamiento utilizado"}
    assert ht.registrar_cambios(hist, data, editados, [{"Campo": "Equipamiento utilizado", "Información extraída": "", "Estado": "No encontrado"}])
    data["Equipamiento utilizado"] = ""
    editados.clear()
    ht.aplicar_accion(hist, data, editados, "deshacer")
    assert data["Equipamiento utilizado"] == "cobas 711"
    assert "Equipamiento utilizado" in editados
    ht.contexto_nuevo(hist, ("Área", "Examen 2", 1))
    assert hist["deshacer"] == [] and hist["rehacer"] == []


def test_varios_cambios_y_limite_20_y_rehacer_se_invalida_si_editas():
    hist = ht.contexto_nuevo({}, ("Área", "A", 0))
    data = {"Campo": "v0"}
    for i in range(1, 24):
        nuevo = f"v{i}"
        assert ht.registrar_cambios(hist, data, set(), [{"Campo": "Campo", "Información extraída": nuevo, "Estado": "Modificado por usuario"}])
        data["Campo"] = nuevo
    assert len(hist["deshacer"]) == 20
    assert ht.aplicar_accion(hist, data, set(), "deshacer")
    assert data["Campo"] == "v22"
    assert len(hist["rehacer"]) == 1
    assert ht.registrar_cambios(hist, data, set(), [{"Campo": "Campo", "Información extraída": "nuevo", "Estado": "Modificado por usuario"}])
    assert hist["rehacer"] == []
    assert not ht.registrar_cambios(hist, data, set(), [{"Campo": "Campo", "Información extraída": "v22", "Estado": "Encontrado"}])


def test_una_operacion_en_multiples_celdas_y_marcas_del_teclado():
    hist = ht.contexto_nuevo({}, ("Área", "A", 0))
    data = {"Estabilidad": "1 h", "Tipo de muestra": "Plasma"}
    filas = [{"Campo": k, "Información extraída": "", "Estado": "No encontrado"} for k in data]
    assert ht.registrar_cambios(hist, data, set(), filas)
    data.update({campo: "" for campo in data})
    assert set(ht.aplicar_accion(hist, data, set(), "deshacer")) == set(data)
    assert data == {"Estabilidad": "1 h", "Tipo de muestra": "Plasma"}
    assert ht.comando_desde_filas([{"__historial_accion": "rehacer:1881"}]) == "rehacer"
    assert ht.comando_desde_filas([{"__historial_accion": ""}]) is None


def test_ignora_checkbox_sin_cambiar_texto_y_comandos_invalidos():
    hist = ht.contexto_nuevo({}, ("A", "E", 0))
    assert not ht.registrar_cambios(hist, {"Campo": "Valor"}, set(), [
        {"Campo": "Campo", "Información extraída": "Valor", "Actualizar": False, "Estado": "Encontrado"}])
    assert not ht.aplicar_accion(hist, {"Campo": "Valor"}, set(), "deshacer")
    assert not ht.comando_desde_filas([{"__historial_accion": "eliminar:123"}])


def _libro(valores):
    wb = Workbook()
    for i, (hoja, campos) in enumerate(valores.items()):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = hoja
        for fila, (campo, valor) in enumerate(campos.items(), start=1):
            ws.cell(fila, 1, campo)
            ws.cell(fila, 2, valor)
    salida = io.BytesIO()
    wb.save(salida)
    wb.close()
    return salida.getvalue()


def test_omision_intermedia_y_final_preservan_bytes_y_no_escriben_db():
    bytes_previos = _libro({"Examen 1": {"Estabilidad": "4 h"}})
    estado = {"result_bytes": bytes_previos, "cola_pdf_procesados": ["primero.pdf"]}
    assert not cf.omitir_documento(estado, 1, 4, "segundo.pdf")
    assert estado["cola_pdf_indice"] == 2
    assert estado["result_bytes"] is bytes_previos
    assert cf.omitir_documento(estado, 3, 4, "cuarto.pdf")
    assert estado["cola_pdf_finalizada_omision"] is True
    assert estado["cola_pdf_indice"] == 4
    assert cf.resumen_cola(estado, 4)["puede_guardar"] is True
    assert estado["result_bytes"] is bytes_previos
    assert estado["cola_pdf_omitidos"] == ["segundo.pdf", "cuarto.pdf"]


def test_omitir_todos_no_inventa_resultados_ni_habilita_guardado():
    estado = {}
    cf.omitir_documento(estado, 0, 2, "a.pdf")
    cf.omitir_documento(estado, 1, 2, "b.pdf")
    assert not cf.resumen_cola(estado, 2)["puede_guardar"]
    for indice, total in ((0, 1), (-1, 2), (2, 2)):
        try:
            cf.omitir_documento({}, indice, total, "imposible.pdf")
        except ValueError:
            pass
        else:
            assert False, "No se debe permitir omitir en esta posición."


def test_compara_todas_las_hojas_y_no_incluye_omitidas():
    original = _libro({"Exam A": {"Campo": "Original"}, "Exam B": {"Campo": "Sin cambios"}})
    nuevo = _libro({"Exam A": {"Campo": "Actualizado"}, "Exam B": {"Campo": "Sin cambios"}, "Exam C": {"Campo": "Nuevo"}})
    comparacion = cf.comparar_libros_completos(original, nuevo)
    assert comparacion["hojas_antes"] == 2
    assert comparacion["hojas_despues"] == 3
    assert comparacion["es_nueva_hoja"]
    assert len(comparacion["diferencias"]) == 2
    assert not any("Exam B" in d["Campo"] for d in comparacion["diferencias"])
    assert any(d["Campo"] == "Exam A · Campo" for d in comparacion["diferencias"])


def test_interfaz_tiene_atajos_omision_y_estado_visible():
    assert 'grid_options[\'onCellKeyDown\'] = JsCode(' in CODIGO
    assert "primerNodo.setDataValue('__historial_accion'" in CODIGO
    assert 'campos_forzar=st.session_state.pop("campos_restaurados_tabla", None)' in CODIGO
    assert 'on_click=_aplicar_historial_tabla, args=("deshacer",)' in CODIGO
    assert 'on_click=_aplicar_historial_tabla, args=("rehacer",)' not in CODIGO
    assert 'btn_rehacer_tabla' not in CODIGO
    assert "estadoUndo.pila.push([cambio])" in CODIGO
    assert 'reload_data=False' in CODIGO
    assert 'st.rerun()\n            ht.registrar_cambios(' not in CODIGO
    assert '"Omitir último examen y finalizar cola"' in CODIGO
    assert 'if total_pdfs > 1 and "data_extracted" not in st.session_state:' in CODIGO
    assert 'cf.comparar_libros_completos(' in CODIGO
    assert 'st.subheader("Referencias", anchor=False)' in CODIGO
    assert 'with st.expander("Referencias bibliográficas", expanded=True)' in CODIGO
    assert "sidebar-examen-activo" in CODIGO and "sidebar-examen-sin-seleccionar" in CODIGO
    assert "st.caption(\"Al cambiar de área se descartarán" not in CODIGO


def test_restauracion_no_se_queda_solo_en_tabla_tambien_reescribe_excel():
    """Un estado Encontrado normalmente no sobrescribe Excel salvo al deshacer."""
    import ast
    import types
    import pandas as pd
    import openpyxl
    from openpyxl.styles import Alignment

    nodo = next(n for n in ast.parse(CODIGO).body
                if isinstance(n, ast.FunctionDef) and n.name == "aplicar_ediciones_a_excel")
    espacio = {
        "io": io, "pd": pd, "load_workbook": openpyxl.load_workbook,
        "buscar_hoja_por_examen": lambda wb, nombre: wb[nombre],
        "norm": lambda txt: str(txt).strip().casefold(),
        "st": types.SimpleNamespace(session_state={}),
        "Alignment": Alignment,
        "ajustar_alto_fila_con_listas": lambda *args: None,
        "ajustar_ancho_columnas": lambda *args: None,
        "apa": types.SimpleNamespace(referencias_excel_con_cursiva=lambda v: v),
    }
    exec(compile(ast.Module(body=[nodo], type_ignores=[]), str(ROOT / "app.py"), 'exec'), espacio)
    actualizador = espacio["aplicar_ediciones_a_excel"]
    libro_borrado = _libro({"PT": {"Estabilidad": ""}})
    filas = pd.DataFrame([{"Campo": "Estabilidad", "Información extraída": "8 h",
                           "Estado": "Encontrado"}])
    # Protege los valores de extracción sin editar durante un rerun normal.
    sin_forzar = actualizador(libro_borrado, filas, "PT")
    assert sin_forzar == libro_borrado
    # Al presionar Ctrl+Z, este campo SÍ debe recuperarse en el archivo.
    restaurado = actualizador(libro_borrado, filas, "PT", campos_forzar=["Estabilidad"])
    ws = openpyxl.load_workbook(io.BytesIO(restaurado), data_only=True)["PT"]
    assert ws.cell(1, 2).value == "8 h"
