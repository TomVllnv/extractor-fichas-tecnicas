"""Regresión visual del bloque de instrucciones de revisión."""
from html.parser import HTMLParser
from pathlib import Path

APP = Path(__file__).resolve().parents[1] / "app.py"


class _AdviceHTML(HTMLParser):
    def __init__(self):
        super().__init__()
        self.items = []
        self.bold = []
        self.in_li = False
        self.in_b = False

    def handle_starttag(self, tag, attrs):
        if tag == "li":
            self.in_li = True
            self.items.append("")
        elif tag == "b":
            self.in_b = True
            self.bold.append("")

    def handle_endtag(self, tag):
        if tag == "li":
            self.in_li = False
        elif tag == "b":
            self.in_b = False

    def handle_data(self, data):
        if self.in_li:
            self.items[-1] += data
        if self.in_b:
            self.bold[-1] += data


def test_consejos_tienen_tres_vinetas_con_textos_solicitados():
    src = APP.read_text(encoding="utf-8")
    block = src.split('<ul style="margin: 0; padding-left: 1.25rem; list-style-type: disc;">', 1)[1].split("</ul>", 1)[0]
    parser = _AdviceHTML()
    parser.feed(block)
    assert parser.items == [
        "Las casillas de la columna Aplicar indican qué campos reemplazar en el Excel al descargar. Desmarca un campo para conservar el valor anterior.",
        "Haz doble clic sobre cualquier celda de la columna Información extraída para editar su contenido directamente en la tabla. El editor admite textos extensos.",
        "Presionar Ctrl + Z o hacer clic en Deshacer permitirá revertir los cambios.",
    ]
    assert parser.bold == ["Aplicar", "Información extraída", "Ctrl + Z", "Deshacer"]


def test_consejos_conservan_estilo_del_bloque_original():
    src = APP.read_text(encoding="utf-8")
    assert '<div style="font-size: 0.85rem; color: #6c757d; line-height: 1.6; margin-bottom: 12px;">' in src
    assert "Presionar Ctrl + Z para deshacer cambios" not in src
    assert "Ctrl + Y para rehacer" not in src
