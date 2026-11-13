"""Regresión visual y funcional de los controles al terminar una cola por omisión."""
import ast
from pathlib import Path

CODIGO = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
ARBOL = ast.parse(CODIGO)


def _control(clave):
    for nodo in ast.walk(ARBOL):
        if not isinstance(nodo, ast.Call) or not isinstance(nodo.func, ast.Attribute):
            continue
        if nodo.func.attr not in ("button", "download_button"):
            continue
        argumentos = {kw.arg: kw.value for kw in nodo.keywords}
        if isinstance(argumentos.get("key"), ast.Constant) and argumentos["key"].value == clave:
            return nodo, argumentos
    raise AssertionError(f"No se encontró el control {clave}")


def _evaluar(nodo, **variables):
    expresion = ast.fix_missing_locations(ast.Expression(nodo))
    return eval(compile(expresion, "<test tooltip>", "eval"), {}, variables)


def test_botones_de_omision_tienen_mismos_textos_colores_y_anchos_referenciales():
    for clave_omision, clave_habitual in (
        ("descarga_cola_tras_omision", "download_updated_xlsx"),
        ("guardar_cola_tras_omision", "btn_trigger_save_template"),
    ):
        widget_omision, args_omision = _control(clave_omision)
        widget_habitual, args_habitual = _control(clave_habitual)
        assert widget_omision.func.attr == widget_habitual.func.attr
        assert ast.unparse(widget_omision.args[0]) == ast.unparse(widget_habitual.args[0])
        assert ast.unparse(args_omision["type"]) == ast.unparse(args_habitual["type"])
        assert ast.literal_eval(args_omision["use_container_width"]) is True
        assert ast.literal_eval(args_habitual["use_container_width"]) is True
        if widget_omision.func.attr == "download_button":
            assert ast.literal_eval(args_omision["on_click"]) == "ignore"
            assert ast.unparse(args_omision["mime"]) == ast.unparse(args_habitual["mime"])


def test_botones_omision_centrados_y_referencia_habitual_sin_cambiar():
    fragmento_omision = CODIGO.split('if resumen["puede_guardar"]:', 1)[1].split(
        "        elif pdfs and not exceso_archivos:", 1
    )[0]
    assert '_, col_final_descarga, col_final_actualiza, _ = st.columns(\n                    [1, 2, 2, 1], gap="small"' in fragmento_omision
    assert "with col_final_descarga:" in fragmento_omision
    assert "with col_final_actualiza:" in fragmento_omision
    fragmento_habitual = CODIGO.split('st.subheader("Opciones de guardado y avance de cola"', 1)[1]
    assert "col_down, col_update, col_next = st.columns(3)" in fragmento_habitual


def test_aviso_es_exclusivo_de_omision_cuando_hay_discrepancias_sin_revisar():
    _, args_omision = _control("guardar_cola_tras_omision")
    _, args_habitual = _control("btn_trigger_save_template")
    ayuda_omision = args_omision["help"]
    assert _evaluar(ayuda_omision, hay_discrepancias=True, excel_final=b"xls") == (
        "Deshabilitado por discrepancias sin analizar."
    )
    assert _evaluar(ayuda_omision, hay_discrepancias=False, excel_final=b"xls") is None
    assert _evaluar(ayuda_omision, hay_discrepancias=False, excel_final=None) == (
        "No hay un Excel válido para actualizar."
    )
    assert "Deshabilitado por discrepancias sin analizar." not in ast.unparse(args_habitual["help"])
    assert "Comprobar discrepancias antes de actualizar." in ast.unparse(args_habitual["help"])
    assert _evaluar(args_omision["disabled"], excel_final=b"xls", hay_discrepancias=True)
    assert not _evaluar(args_omision["disabled"], excel_final=b"xls", hay_discrepancias=False)
