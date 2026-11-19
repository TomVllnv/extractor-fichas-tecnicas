"""El aviso de finalización por omisión se centra sin alterar botones."""
import ast
from pathlib import Path


APP = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
SEGMENTO = APP.split('if resumen["puede_guardar"]:', 1)[1].split(
    "        elif pdfs and not exceso_archivos:", 1
)[0]


def test_aviso_omision_sigue_siendo_info_y_se_centra_antes_de_los_botones():
    assert '_, col_aviso_guardado, _ = st.columns([1, 4, 1], gap="small")' in SEGMENTO
    assert 'with col_aviso_guardado:' in SEGMENTO
    assert 'with st.container(key="aviso_guardado_omision"):' in SEGMENTO
    mensaje = "Los exámenes procesados siguen en la ficha consolidada. Omitir un PDF no altera ni borra los avances anteriores."
    assert SEGMENTO.count(f'st.info("{mensaje}")') == 1
    assert "[data-testid='stAlert'] p { text-align: center; }" in SEGMENTO
    assert SEGMENTO.index(f'st.info("{mensaje}")') < SEGMENTO.index("_, col_final_descarga")


def test_botones_omision_conservan_anchuras_y_acciones():
    assert '_, col_final_descarga, col_final_actualiza, _ = st.columns(\n                    [1, 2, 2, 1], gap="small"' in SEGMENTO
    assert 'key="descarga_cola_tras_omision", on_click="ignore"' in SEGMENTO
    assert 'key="guardar_cola_tras_omision"' in SEGMENTO
    ast.parse(APP)
