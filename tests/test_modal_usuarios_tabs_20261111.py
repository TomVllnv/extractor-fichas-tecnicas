"""Regresión: los formularios no deben escapar de sus pestañas."""
import ast
from pathlib import Path

MODULO = ast.parse((Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8"))
MODAL = next(n for n in MODULO.body if isinstance(n, ast.FunctionDef) and n.name == "modal_administracion_usuarios")


def test_cada_formulario_crea_su_columna_dentro_de_su_pestana():
    for nombre_pestana, nombre_contenedor in (("tab_crear", "contenido_crear"),
                                              ("tab_editar", "contenido_editar"),
                                              ("tab_clave", "contenido_clave"),
                                              ("tab_estado", "contenido_estado")):
        pestañas = [n for n in MODAL.body if isinstance(n, ast.With)
                    and isinstance(n.items[0].context_expr, ast.Name)
                    and n.items[0].context_expr.id == nombre_pestana]
        assert len(pestañas) == 1
        hijos = pestañas[0].body
        assert isinstance(hijos[0], ast.Assign)
        assert isinstance(hijos[0].value, ast.Call)
        assert isinstance(hijos[0].value.func, ast.Attribute)
        assert hijos[0].value.func.attr == "columns"
        assert isinstance(hijos[1], ast.With)
        assert isinstance(hijos[1].items[0].context_expr, ast.Name)
        assert hijos[1].items[0].context_expr.id == nombre_contenedor


def test_tabla_de_usuarios_fuera_de_pestanas_y_modal_ancho():
    source = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
    encabezado = source.split('def modal_administracion_usuarios():')[0].splitlines()[-1]
    assert 'width="large"' in encabezado
    assert len([n for n in ast.walk(MODAL) if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)
                and n.func.attr == 'dataframe']) == 1
