"""Detección rápida y no destructiva de páginas escaneadas, sin ejecutar OCR."""
from __future__ import annotations
import fitz


def inspeccionar_pdf(datos):
    """Devuelve (clase, páginas_sin_texto, total); la clasificación es orientativa."""
    with fitz.open(stream=datos, filetype='pdf') as documento:
        total = documento.page_count
        imagen_sin_texto = 0
        texto = 0
        for pagina in documento:
            caracteres = sum(len(w[4]) for w in pagina.get_text('words'))
            if caracteres >= 65:
                texto += 1
                continue
            bloques = pagina.get_text('dict').get('blocks', [])
            tiene_imagen = any(b.get('type') == 1 for b in bloques)
            if tiene_imagen:
                imagen_sin_texto += 1
        if imagen_sin_texto == 0:
            return 'texto', 0, total
        if texto == 0 and imagen_sin_texto >= max(1, total * .6):
            return 'escaneado', imagen_sin_texto, total
        return 'mixto', imagen_sin_texto, total
