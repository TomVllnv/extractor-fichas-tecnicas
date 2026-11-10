"""Construcción conservadora de referencias APA para la ficha técnica.

Nunca confundir fecha de subida con fecha de publicación ni inventar autores.
La referencia provisional se identifica para que el operador pueda completarla.
"""
from __future__ import annotations

import os
import html as _html
import re
import unicodedata
from dataclasses import dataclass

import fitz


@dataclass(frozen=True)
class FuenteBibliografica:
    identificador: str
    titulo: str
    autor: str = ""
    anio: str = ""
    version: str = ""
    tipo: str = ""
    titulo_provisional: bool = False
    datos_editoriales: str = ""
    url: str = ""


def _titulo_archivo(nombre: str) -> str:
    titulo = os.path.basename(str(nombre or "")).strip()
    while re.search(r"\.(?:pdf|xlsx|xls|docx|doc)$", titulo, flags=re.I):
        titulo = re.sub(r"\.(?:pdf|xlsx|xls|docx|doc)$", "", titulo, flags=re.I)
    return re.sub(r"\s+", " ", titulo.replace("_", " ")).strip() or "Documento sin título identificado"


def _anio_explicito(texto: str) -> str:
    for patron in (
        r"(?i)\b(?:fecha(?:\s+de\s+(?:publicaci[oó]n|emisi[oó]n|revisi[oó]n))?|publicado|vigente\s+desde)\s*:?\s*([0-3]?\d)[/-]([01]?\d)[/-]((?:19|20)\d{2})\b",
        r"(?i)\b(?:a[nñ]o\s+de\s+publicaci[oó]n|fecha\s+de\s+publicaci[oó]n)\s*:?\s*((?:19|20)\d{2})\b",
    ):
        encontrado = re.search(patron, texto)
        if encontrado:
            return encontrado.group(encontrado.lastindex)
    return ""


def _version_explicita(texto: str) -> str:
    m = re.search(r"(?i)\bversi[oó]n\s*:?\s*(v?\d+(?:\.\d+){0,4})\b", texto)
    return m.group(1) if m else ""


def _normalizar_espacios(texto):
    return re.sub(r"\s+", " ", str(texto or "")).strip().rstrip(".")


def _titulo_cabecera_pdf(pagina) -> str:
    """Obtiene el título visible, normalmente la tipografía grande de portada.

    Evita usar como título el logo, códigos de catálogo u otros encabezados.
    Si no se identifica de manera razonable, se conserva el título del PDF.
    """
    candidatos = []
    try:
        bloques = pagina.get_text("dict").get("blocks", [])
    except (ValueError, RuntimeError, TypeError):
        return ""
    for bloque in bloques:
        if bloque.get("type") != 0:
            continue
        for linea in bloque.get("lines", []):
            spans = [s for s in linea.get("spans", []) if (s.get("text") or "").strip()]
            if not spans:
                continue
            texto = _normalizar_espacios(" ".join(s["text"] for s in spans))
            tamanio = max(float(s.get("size", 0)) for s in spans)
            x0 = min(float(s.get("bbox", (0, 0, 0, 0))[0]) for s in spans)
            y0 = min(float(s.get("bbox", (0, 0, 0, 0))[1]) for s in spans)
            if (tamanio < 15 or y0 > pagina.rect.height * 0.28
                    or x0 > pagina.rect.width * 0.68
                    or not 4 <= len(texto) <= 95):
                continue
            if (re.fullmatch(r"[0-9.\-_/\s]+", texto)
                    or re.search(r"(?i)^(?:cobas[®™]?|roche|espa[nñ]ol|informaci[oó]n|ref\b|content\b|system\b)", texto)
                    or re.search(r"(?i)^(?:https?://|www\.)", texto)):
                continue
            candidatos.append((tamanio, -y0, texto))
    return max(candidatos, default=(0, 0, ""))[2]


def _anio_copyright_roche(texto: str) -> str:
    """Un año explícitamente asociado a Roche, sin utilizar la fecha de subida."""
    for patron in (
        r"(?i)(?:©|copyright)\s*[,;:]?\s*((?:19|20)\d{2})\s*[,;:]*\s*Roche\s+Diagnostics",
        r"(?i)\b((?:19|20)\d{2})\s*[,;:]?\s*Roche\s+Diagnostics\b",
        r"(?i)Roche\s+Diagnostics(?:\s+GmbH)?\s*[,;:]*\s*(?:©|copyright)\s*((?:19|20)\d{2})\b",
    ):
        encontrado = re.search(patron, texto)
        if encontrado:
            return encontrado.group(1)
    return ""


def _datos_roche_pdf(texto: str) -> tuple[str, str, str, str]:
    """Autor, año, editorial/dirección y URL indicados en el propio PDF."""
    limpio = re.sub(r"\s+", " ", texto.replace("\u00a0", " ")).strip()
    if not re.search(r"(?i)Roche\s+Diagnostics\b", limpio):
        return "", "", "", ""
    autor = "Roche Diagnostics"
    anio = _anio_copyright_roche(limpio)
    # Conservar únicamente el nombre editorial; la dirección postal no forma
    # parte de la referencia solicitada, aunque se encuentre en el PDF.
    editorial = re.search(r"(?i)\bRoche\s+Diagnostics\s+GmbH\b", limpio)
    datos_editoriales = _normalizar_espacios(editorial.group(0)) if editorial else ""
    # En Referencias se presenta solamente el dominio, sin protocolo ni www.
    url = "roche.com" if re.search(r"(?i)\b(?:https?://)?(?:www\.)?roche\.com\b", limpio) else ""
    return autor, anio, datos_editoriales, url


def fuente_pdf(nombre: str, datos: bytes, examen: str = "") -> FuenteBibliografica:
    """Crea una referencia solo para bibliografía; no modifica otros campos."""
    titulo, autor, anio, version, datos_editoriales, url = "", "", "", "", "", ""
    try:
        with fitz.open(stream=datos, filetype="pdf") as doc:
            met = doc.metadata or {}
            titulo_met = _normalizar_espacios(met.get("title"))
            if (titulo_met and titulo_met.lower() not in
                    ("untitled", "document", "documento", "microsoft word")
                    and not re.search(r"\.(pdf|docx?)$", titulo_met, re.I)):
                titulo = titulo_met
            autor_met = _normalizar_espacios(met.get("author"))
            if (autor_met and autor_met.casefold() not in
                    ("user", "usuario", "admin", "administrator", "autor")
                    and not re.search(r"(?:microsoft|adobe|libreoffice|canon|epson)", autor_met, re.I)):
                autor = autor_met
            if doc.page_count:
                titulo_visible = _titulo_cabecera_pdf(doc[0])
                if titulo_visible:
                    titulo = titulo_visible
            # Los avisos de copyright y domicilio suelen aparecer en la última
            # página de fichas Roche. Leer solo portada y páginas finales.
            indices = sorted(set(range(min(2, doc.page_count))) |
                             set(range(max(0, doc.page_count - 2), doc.page_count)))
            textos = {i: doc[i].get_text()[:18000] for i in indices}
            primeras_paginas = "\n".join(textos[i] for i in indices if i < 2)[:26000]
            texto_bibliografico = "\n".join(textos.values())[:64000]
            anio = _anio_explicito(primeras_paginas)
            version = _version_explicita(primeras_paginas)
            autor_roche, anio_roche, datos_editoriales, url = _datos_roche_pdf(texto_bibliografico)
            if autor_roche:
                autor = autor_roche
            if not anio:
                anio = anio_roche
    except (ValueError, RuntimeError, TypeError):
        pass
    return FuenteBibliografica(
        identificador="pdf:" + nombre,
        titulo=titulo or _titulo_archivo(nombre),
        autor=autor,
        anio=anio,
        version=version,
        # No crear una descripción artificial a partir del examen seleccionado.
        # El nombre del examen debe proceder del título visible del propio PDF.
        tipo="",
        titulo_provisional=not bool(titulo),
        datos_editoriales=datos_editoriales,
        url=url,
    )


def fuente_institucional(fuente_id: str, documento: dict) -> FuenteBibliografica:
    """Fuente basada en los metadatos del ORIGINAL, no en fecha_actualizacion."""
    nombre = str(documento.get("nombre_archivo") or "")
    version = ""
    anio = ""
    autor = ""
    titulo = _titulo_archivo(nombre)
    # Solo atribuir la institución cuando el propio PDF la identifica.
    if fuente_id == "manual_toma_muestras" and documento.get("archivo"):
        try:
            with fitz.open(stream=documento["archivo"], filetype="pdf") as pdf:
                encabezado = " ".join(pdf[i].get_text() for i in range(min(4, len(pdf))))
                version = _version_explicita(encabezado)
                anio = _anio_explicito(encabezado)
                if re.search(r"CENTRO\s+DE\s+REFERENCIA\s+DE\s+SALUD", encabezado, re.I):
                    autor = "Centro de Referencia de Salud de Peñalolén Cordillera Oriente"
                if re.search(r"MANUAL\s+DE\s+TOMA\s+DE\s+MUESTRAS", encabezado, re.I):
                    titulo = "Manual de toma de muestras laboratorio clínico del Centro de Referencia de Salud de Peñalolén Cordillera Oriente"
        except (ValueError, RuntimeError):
            pass
    else:
        coincidencia = re.search(r"(?i)(?:^|[_\s-])v(\d+(?:\.\d+){1,4})(?:\b|\.)", nombre)
        if coincidencia:
            version = coincidencia.group(1)
            titulo = re.sub(r"(?i)\s+v" + re.escape(version) + r"$", "", titulo).strip()
    return FuenteBibliografica(
        identificador=fuente_id, titulo=titulo, autor=autor, anio=anio, version=version,
        tipo={"manual_toma_muestras": "Manual institucional", "normas_generales": "Documento institucional",
              "cartera_acreditacion": "Cartera de acreditación", "cartera_examenes": "Cartera de exámenes"}.get(fuente_id, "Documento institucional"),
        titulo_provisional=fuente_id != "manual_toma_muestras" or not autor,
    )


def _complemento_editorial(fuente: FuenteBibliografica) -> str:
    """Datos editoriales verificables, solo en las referencias bibliográficas."""
    partes = []
    if fuente.datos_editoriales:
        partes.append(_normalizar_espacios(fuente.datos_editoriales).rstrip(".") + ".")
    if fuente.url:
        # Mostrar solo el dominio, sin www. ni protocolo.
        dominio = re.sub(r"(?i)^https?://", "", str(fuente.url).strip())
        dominio = re.sub(r"(?i)^www\.", "", dominio)
        dominio = re.split(r"[/?#]", dominio, maxsplit=1)[0]
        if dominio:
            partes.append(dominio)
    return " " + " ".join(partes) if partes else ""


def referencia_apa(fuente: FuenteBibliografica) -> str:
    """Referencia con autor corporativo y datos editoriales cuando se conocen."""
    titulo = _normalizar_espacios(fuente.titulo)
    autor = _normalizar_espacios(fuente.autor)
    anio = fuente.anio if re.fullmatch(r"(?:19|20)\d{2}", fuente.anio or "") else "s. f."
    version = f" (versión {fuente.version})" if fuente.version else ""
    tipo = f" [{fuente.tipo}]" if fuente.tipo else ""
    if autor and unicodedata.normalize("NFD", autor).casefold() != unicodedata.normalize("NFD", titulo).casefold():
        cita = f"{autor}. ({anio}). {titulo}{version}{tipo}."
    else:
        cita = f"{titulo}{version}{tipo}. ({anio})."
    return cita + _complemento_editorial(fuente)


def ordenar_y_unir(fuentes: list[FuenteBibliografica]) -> str:
    unicas = {}
    for fuente in fuentes:
        unicas.setdefault(fuente.identificador, fuente)
    citas = sorted((referencia_apa(f) for f in unicas.values()), key=lambda s: s.casefold())
    return "\n\n".join(citas)




def referencia_apa_markdown(fuente: FuenteBibliografica) -> str:
    """Versión para la vista previa web: título en cursiva Markdown/HTML."""
    titulo = _normalizar_espacios(fuente.titulo)
    autor = _normalizar_espacios(fuente.autor)
    anio = fuente.anio if re.fullmatch(r"(?:19|20)\d{2}", fuente.anio or "") else "s. f."
    version = f" (versión {fuente.version})" if fuente.version else ""
    tipo = f" [{fuente.tipo}]" if fuente.tipo else ""
    if autor and unicodedata.normalize("NFD", autor).casefold() != unicodedata.normalize("NFD", titulo).casefold():
        return f"{autor}. ({anio}). *{titulo}*{version}{tipo}." + _complemento_editorial(fuente)
    return f"*{titulo}*{version}{tipo}. ({anio})." + _complemento_editorial(fuente)


def ordenar_y_unir_markdown(fuentes: list[FuenteBibliografica]) -> str:
    unicas = {}
    for fuente in fuentes:
        unicas.setdefault(fuente.identificador, fuente)
    citas = sorted((referencia_apa_markdown(f) for f in unicas.values()), key=lambda s: re.sub(r'[*]', '', s).casefold())
    return "\n\n".join(citas)

def referencia_apa_html(fuente: FuenteBibliografica) -> str:
    """Vista previa segura: título en cursiva y datos bibliográficos escapados."""
    titulo = _normalizar_espacios(fuente.titulo)
    autor = _normalizar_espacios(fuente.autor)
    anio = fuente.anio if re.fullmatch(r"(?:19|20)\d{2}", fuente.anio or "") else "s. f."
    version = f" (versión {fuente.version})" if fuente.version else ""
    tipo = f" [{fuente.tipo}]" if fuente.tipo else ""
    titulo_it = f"<em>{_html.escape(titulo)}</em>"
    resto = _html.escape(version + tipo)
    if autor and unicodedata.normalize("NFD", autor).casefold() != unicodedata.normalize("NFD", titulo).casefold():
        return f"{_html.escape(autor)}. ({anio}). {titulo_it}{resto}." + _html.escape(_complemento_editorial(fuente))
    return f"{titulo_it}{resto}. ({anio})." + _html.escape(_complemento_editorial(fuente))


def ordenar_y_unir_html(fuentes: list[FuenteBibliografica]) -> str:
    """Una vista previa legible de cada referencia con separación de párrafos."""
    unicas = {}
    for fuente in fuentes:
        unicas.setdefault(fuente.identificador, fuente)
    ordenadas = sorted(unicas.values(), key=lambda f: referencia_apa(f).casefold())
    return "<br><br>".join(referencia_apa_html(f) for f in ordenadas)


def advertencias_fuentes(fuentes: list[FuenteBibliografica]) -> list[str]:
    issues = []
    for f in fuentes:
        faltan = []
        if not f.autor: faltan.append("autor")
        if not f.anio: faltan.append("año de publicación")
        if f.titulo_provisional: faltan.append("título oficial")
        if faltan:
            issues.append(f"{f.titulo}: verificar {', '.join(faltan)}.")
    return issues


def referencias_excel_con_cursiva(fuentes: list[FuenteBibliografica]):
    """Devuelve texto plano para Excel, sin cursivas, priorizando compatibilidad."""
    return ordenar_y_unir(fuentes)
