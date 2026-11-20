"""El centrado del encabezado por omisión es aislado y no altera botones."""
from pathlib import Path
import ast

APP = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
SEGMENTO = APP.split('if resumen["puede_guardar"]:', 1)[1].split(
    "        elif pdfs and not exceso_archivos:", 1
)[0]


def test_titulo_centrado_aislado_con_estilo_nativo():
    assert 'with st.container(key="titulo_guardado_omision"):' in SEGMENTO
    assert 'st.subheader("Guardar avances de la cola", anchor=False)' in SEGMENTO
    assert "st-key-titulo_guardado_omision h3 { text-align: center; width: 100%; }" in SEGMENTO
    assert "[data-testid='stHeadingWithActionElements'] { justify-content: center; width: 100%; }" in SEGMENTO
    assert SEGMENTO.index('with st.container(key="titulo_guardado_omision")') < SEGMENTO.index('_, col_aviso_guardado, _')
    ast.parse(APP)


def test_aviso_y_botones_permanecen_sin_cambios():
    assert 'with st.container(key="aviso_guardado_omision"):' in SEGMENTO
    assert 'st.info("Los exámenes procesados siguen en la ficha consolidada. Omitir un PDF no altera ni borra los avances anteriores.")' in SEGMENTO
    assert '_, col_final_descarga, col_final_actualiza, _ = st.columns(' in SEGMENTO
    assert 'key="descarga_cola_tras_omision", on_click="ignore"' in SEGMENTO
    assert 'key="guardar_cola_tras_omision"' in SEGMENTO
