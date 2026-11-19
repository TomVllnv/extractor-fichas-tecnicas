"""Color del estado activo: aplica al valor, no a la etiqueta."""
from pathlib import Path
import re

CODIGO = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")


def test_color_unicamente_en_valores_de_examen():
    reglas = dict((nombre, cuerpo) for nombre, cuerpo in re.findall(
        r"(?m)^\s*(\.sidebar-examen-(?:activo|sin-seleccionar)(?: span)?)\s*\{([^}]*)\}",
        CODIGO,
    ))
    assert "color:" not in reglas[".sidebar-examen-activo"]
    assert "color:" not in reglas[".sidebar-examen-sin-seleccionar"]
    assert "color:#50e89a" in reglas[".sidebar-examen-activo span"]
    assert "color:#ffb45b" in reglas[".sidebar-examen-sin-seleccionar span"]


def test_etiqueta_sin_color_propietario_y_textos_sin_cambios():
    assert "'<strong>Examen activo:</strong> '" in CODIGO
    assert "'<strong>Examen activo:</strong> <span>Sin seleccionar</span></div>'" in CODIGO
    assert "f'<span>{html.escape(str(examen_en_uso))}</span></div>'" in CODIGO
