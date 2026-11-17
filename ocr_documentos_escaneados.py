"""OCR especializado de documentos escaneados: no se usa en fichas técnicas.

Tesseract con idioma español es preferente por su segmentación en palabras.
Si no existe la aplicación de escritorio/binario, RapidOCR sigue disponible.
No se insertan espacios DENTRO de una palabra reconocida por conjetura.
"""
from __future__ import annotations

import re
from functools import lru_cache


@lru_cache(maxsize=1)
def _tesseract_disponible():
    try:
        import pytesseract
        idiomas = pytesseract.get_languages(config="")
        idioma = "spa+eng" if "spa" in idiomas and "eng" in idiomas else (
            "spa" if "spa" in idiomas else "eng" if "eng" in idiomas else "")
        return idioma
    except (ImportError, OSError, RuntimeError):
        return ""


def reconstruir_lineas_rapidocr(resultados):
    """Reconstruye renglones/párrafos según la geometría de cada detección.

    Las detecciones pueden contener una frase larga sin espacios: no se
    corrige la terminología clínica mediante un diccionario improvisado.
    """
    registros = []
    for linea in resultados or ():
        if len(linea) < 2 or not str(linea[1] or "").strip():
            continue
        texto = str(linea[1]).strip()
        try:
            puntos = linea[0]
            xs = [float(p[0]) for p in puntos]
            ys = [float(p[1]) for p in puntos]
            x0, y0, y1 = min(xs), min(ys), max(ys)
            alto = max(y1 - y0, 1.0)
        except (TypeError, ValueError, IndexError):
            x0, y0, alto = 0.0, len(registros) * 20.0, 15.0
        registros.append((y0, x0, alto, texto))
    registros.sort(key=lambda par: (par[0], par[1]))
    resultado = []
    anterior = None
    for y, x, alto, texto in registros:
        if anterior is not None:
            ant_y, ant_x, ant_alto, ant_texto = anterior
            delta = y - ant_y
            # Separación geométrica significativa => nuevo párrafo.
            salto = "\n\n" if delta > 2.0 * max(alto, ant_alto) else "\n"
            resultado.append(salto)
        resultado.append(texto)
        anterior = (y, x, alto, texto)
    return "".join(resultado).strip()


def leer_pagina_escaneada(imagen_rgb, motor_rapid=None):
    """Devuelve (texto, motor) sin OCR redundante por defecto.

    Solo lo llama el modo de documentos escaneados. Tesseract se utiliza
    cuando está disponible; si falla se conserva la lectura RapidOCR.
    """
    idioma = _tesseract_disponible()
    if idioma:
        try:
            import pytesseract
            from PIL import Image
            texto = pytesseract.image_to_string(
                Image.fromarray(imagen_rgb), lang=idioma,
                config="--psm 3 -c preserve_interword_spaces=1",
            ).replace("\r\n", "\n").strip()
            if len(re.sub(r"\s", "", texto)) >= 10:
                return texto, "Tesseract"
        except (ImportError, OSError, RuntimeError, ValueError):
            pass

    if motor_rapid is None:
        from rapidocr_onnxruntime import RapidOCR
        motor_rapid = RapidOCR()
    resultado, _ = motor_rapid(imagen_rgb)
    return reconstruir_lineas_rapidocr(resultado), "RapidOCR"
