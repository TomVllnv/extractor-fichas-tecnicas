"""Comprobaciones independientes de BD para el cambio entre modos claro y oscuro."""
from pathlib import Path
import tomllib

ROOT = Path(__file__).resolve().parents[1]


def _config():
    config = ROOT / ".streamlit" / "config.toml"
    assert config.is_file()
    return tomllib.loads(config.read_text(encoding="utf-8"))


def test_modo_oscuro_marino_contraste_adecuado():
    t = _config()["theme"]
    assert t["dark"]["backgroundColor"] == "#101D2D"
    assert t["dark"]["secondaryBackgroundColor"] == "#1D3047"
    assert t["dark"]["textColor"] == "#EAF1F8"
    assert t["dark"]["sidebar"]["backgroundColor"] == "#16283E"
    assert t["dark"]["sidebar"]["secondaryBackgroundColor"] == "#233B55"


def test_modo_claro_puede_elegirse_y_es_legible():
    t = _config()["theme"]
    assert "light" in t and "dark" in t
    assert t["light"]["backgroundColor"] == "#FFFFFF"
    assert t["light"]["textColor"] == "#172033"
    assert t["light"]["sidebar"]["backgroundColor"] == "#F3F6FA"
    assert t["light"]["sidebar"]["textColor"] == "#172033"


def test_tema_no_contiene_opciones_de_servidor_o_base_de_datos():
    options = _config()
    assert set(options) == {"theme"}
    assert not (ROOT / ".streamlit" / "secrets.toml").exists()


def test_version_streamlit_acepta_dos_temas():
    requirements = (ROOT / "requirements.txt").read_text(encoding="utf-8")
    assert "streamlit>=1.51.0" in requirements
