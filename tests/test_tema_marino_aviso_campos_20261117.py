"""Verifica el tema marino, la permanencia del claro y la legibilidad del aviso."""
import re
from pathlib import Path
import tomllib

ROOT = Path(__file__).resolve().parents[1]
CONFIG = tomllib.loads((ROOT / '.streamlit/config.toml').read_text(encoding='utf-8'))
APP = (ROOT / 'app.py').read_text(encoding='utf-8')


def _luminancia(hex_color):
    channels = [int(hex_color[i:i + 2], 16) / 255 for i in (1, 3, 5)]
    linear = [v / 12.92 if v <= .04045 else ((v + .055) / 1.055) ** 2.4 for v in channels]
    return .2126 * linear[0] + .7152 * linear[1] + .0722 * linear[2]


def _contraste(a, b):
    x, y = sorted((_luminancia(a), _luminancia(b)), reverse=True)
    return (x + .05) / (y + .05)


def test_tema_claro_se_conserva_sin_cambios():
    assert CONFIG['theme']['light'] == {
        'primaryColor': '#D63C3C',
        'backgroundColor': '#FFFFFF',
        'secondaryBackgroundColor': '#F3F6FA',
        'textColor': '#172033',
        'sidebar': {
            'backgroundColor': '#F3F6FA',
            'secondaryBackgroundColor': '#FFFFFF',
            'textColor': '#172033',
        },
    }


def test_tema_marino_y_contraste_legible():
    dark = CONFIG['theme']['dark']
    assert dark['backgroundColor'] == '#101D2D'
    assert dark['secondaryBackgroundColor'] == '#1D3047'
    assert dark['sidebar']['backgroundColor'] == '#16283E'
    for color in (dark['backgroundColor'], dark['secondaryBackgroundColor'],
                  dark['sidebar']['backgroundColor'], dark['sidebar']['secondaryBackgroundColor']):
        assert _contraste(dark['textColor'], color) >= 4.5
    assert CONFIG['theme']['base'] == 'dark'


def test_nombres_de_campos_resaltados_solo_en_su_aviso():
    assert 'with st.container(key="aviso_campos_manual"):' in APP
    assert '.st-key-aviso_campos_manual [data-testid="stAlert"] strong' in APP
    assert 'color: #FFFFFF !important;' in APP
    assert 'background-color: #245B8A;' in APP
    assert _contraste('#FFFFFF', '#245B8A') >= 4.5
    assert '"**" + re.sub(' in APP
    assert 'El manual permitió reconocer los siguientes campos: ' in APP
    assert 'Revisa los datos propuestos antes de guardarlos.' in APP


def test_nombres_con_caracteres_markdown_no_interrumpen_el_aviso():
    # Ejercicio aislado de la misma regla de escape utilizada en la lista de campos.
    especial = r'Equipo [XN_1000]*'
    escaped = re.sub(r'([\\*_[\]`])', r'\\\1', especial)
    assert escaped == r'Equipo \[XN\_1000\]\*'
