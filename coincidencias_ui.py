"""Paleta perceptual y estilos para elegir coincidencias, sin modificar su puntuación."""
from __future__ import annotations

import hashlib
import math


def claves_coincidencia(tipo: str, identificador: object) -> str:
    """Clave estable, sin introducir nombres clínicos en selectores CSS."""
    if tipo not in ('crs', 'acreditacion'):
        raise ValueError('Tipo de cartera no reconocido.')
    huella = hashlib.sha256(str(identificador).encode('utf-8')).hexdigest()[:16]
    return f'coincidencia_nombre_{tipo}_{huella}'


def porcentaje_similitud(puntaje: float) -> int:
    numero = float(puntaje)
    return round(max(0.0, min(1.0, numero)) * 100) if math.isfinite(numero) else 0


def matiz_similitud(puntaje: float) -> int:
    """55 % rojo, ~70 % naranja, ~80 % amarillo y 100 % verde."""
    numero = float(puntaje)
    if not math.isfinite(numero):
        return 0
    return round(132 * max(0.0, min(1.0, (numero - 0.55) / 0.45)))


def estilos_coincidencia(clave_nombre: str, puntaje: float) -> str:
    """CSS ligado a la tarjeta clicable; color base según similitud."""
    matiz = matiz_similitud(puntaje)
    return f'''<style>
    .st-key-{clave_nombre} button {{
      width: 100%;
      min-height: 0 !important;
      height: auto !important;
      display: flex !important;
      justify-content: flex-start !important;
      align-items: center !important;
      text-align: left !important;
      background: hsl({matiz} 72% 58%) !important;
      border: 1px solid hsl({matiz} 64% 42%) !important;
      color: #0f172a !important;
      border-radius: 7px;
      padding: 0.14rem 0.5rem !important;
      transition: background-color 140ms ease, color 140ms ease, border-color 140ms ease;
      box-shadow: none !important;
    }}
    .st-key-{clave_nombre} button > div {{
      width: 100% !important;
      display: flex !important;
      justify-content: flex-start !important;
    }}
    .st-key-{clave_nombre} button p {{
      color: inherit !important;
      text-align: left !important;
      white-space: normal !important;
      overflow-wrap: anywhere;
      font-weight: 650;
      line-height: 1.18;
      margin: 0 !important;
      width: 100% !important;
      font-size: 0.92rem !important;
    }}
    .st-key-{clave_nombre} button:hover,
    .st-key-{clave_nombre} button:focus-visible {{
      background: #ffffff !important;
      color: #0f172a !important;
      border-color: hsl({matiz} 64% 42%) !important;
    }}
    .st-key-{clave_nombre} button:hover p,
    .st-key-{clave_nombre} button:focus-visible p {{
      color: #0f172a !important;
    }}
    </style>'''
