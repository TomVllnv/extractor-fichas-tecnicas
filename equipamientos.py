"""Comparación conservadora de modelos de equipos, independiente del fabricante.

No se aplican semejanzas difusas a identificadores de modelos: T500 y T511
son dispositivos distintos. Los valores originales se conservan al mostrar.
"""
from __future__ import annotations

import re
import unicodedata


def _fragmentos(texto: str) -> list[str]:
    """Separa listas, sin dividir una descripción en frases incidentales."""
    partes = re.split(r"\s*(?:\r?\n|;|[•●])\s*", str(texto or ""))
    elementos = []
    for parte in partes:
        parte = parte.strip().strip("•● ")
        if not parte:
            continue
        candidatos = [p.strip() for p in parte.split(",")]
        # Las comas solo son separadores de equipos si TODOS los miembros
        # parecen identificadores breves con modelo numérico. Evitar dividir
        # descripciones como «Cobas T511, con 8 posiciones disponibles».
        if (len(candidatos) > 1 and all(
            2 <= len(c) <= 75 and 2 <= len(c.split()) <= 4
            and re.search(r"[a-záéíóúñ]", c, re.I)
            and re.search(r"\d", c)
            and not re.search(r"(?i)\b(?:con|sin|de|para|por|entre|que|tiene|"
                              r"disponibles|posiciones|muestras|tubos|procesa)\b", c)
            for c in candidatos
        )):
            elementos.extend(candidatos)
        else:
            elementos.append(parte)
    return elementos


def clave_equipo(texto: str) -> str:
    """Ignora formato gráfico, pero conserva íntegra la identidad del modelo."""
    normal = unicodedata.normalize("NFKD", str(texto or "").casefold())
    sin_tildes = "".join(c for c in normal if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", "", sin_tildes)


def deduplicar_equipamiento(texto: str) -> str:
    partes = _fragmentos(texto)
    unicos = {}
    for parte in partes:
        clave = clave_equipo(parte)
        if clave:
            unicos.setdefault(clave, parte)
    return "\n".join(unicos.values())


def equipamientos_equivalentes(izquierda: str, derecha: str) -> bool:
    a = {clave_equipo(x) for x in _fragmentos(izquierda)} - {""}
    b = {clave_equipo(x) for x in _fragmentos(derecha)} - {""}
    return a == b
