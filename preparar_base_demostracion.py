"""Crea una COPIA de SQLite para presentar sin historiales de trazabilidad.

No modifica el archivo original y no se ejecuta al abrir la aplicación.
Ejemplo (desde la carpeta del proyecto):
    python preparar_base_demostracion.py
    python preparar_base_demostracion.py --origen "C:/datos/base_de_datos.db"

Para abrir Streamlit con la copia en PowerShell:
    $env:HERRAMIENTA_DB_PATH = "$PWD\\base_de_datos_demo.db"
    streamlit run app.py

Después de la presentación, abre otra terminal o elimina la variable:
    Remove-Item Env:HERRAMIENTA_DB_PATH
"""
from __future__ import annotations

import argparse
import os
import sqlite3
from contextlib import closing
from pathlib import Path

TABLAS_HISTORIAL = ("auditoria_cambios", "eventos_acceso")


def preparar_copia(origen: Path, destino: Path) -> dict[str, int]:
    """Crea una copia vía backup de SQLite y limpia solo sus historiales."""
    origen = origen.expanduser().resolve()
    destino = destino.expanduser().resolve()
    if origen == destino:
        raise ValueError("El destino debe ser distinto de la base de datos original.")
    if not origen.is_file():
        raise FileNotFoundError(f"No se encontró la base de datos original: {origen}")
    if destino.exists():
        raise FileExistsError(f"Ya existe {destino}. No se sobrescribirá ninguna copia.")
    destino.parent.mkdir(parents=True, exist_ok=True)
    resultado = {}
    try:
        # Solo conexión de lectura al original: se usa backup para soportar WAL.
        with closing(sqlite3.connect(f"file:{origen.as_posix()}?mode=ro", uri=True)) as original:
            with closing(sqlite3.connect(destino)) as copia:
                original.backup(copia)
        with closing(sqlite3.connect(destino)) as copia:
            tablas = {f[0] for f in copia.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            if not set(TABLAS_HISTORIAL).issubset(tablas):
                raise ValueError("La base de datos no tiene ambas tablas de trazabilidad esperadas.")
            with copia:
                for tabla in TABLAS_HISTORIAL:
                    resultado[tabla] = copia.execute(f"SELECT COUNT(*) FROM {tabla}").fetchone()[0]
                    copia.execute(f"DELETE FROM {tabla}")
            comprobacion = copia.execute("PRAGMA integrity_check").fetchone()[0]
            if comprobacion != "ok":
                raise ValueError(f"La verificación de la copia falló: {comprobacion}")
        return resultado
    except BaseException:
        destino.unlink(missing_ok=True)
        raise


def main() -> None:
    parser = argparse.ArgumentParser(description="Crear copia de demostración sin los historiales de trazabilidad.")
    defecto = os.getenv("HERRAMIENTA_DB_PATH") or str(Path(__file__).resolve().parent / "base_de_datos.db")
    parser.add_argument("--origen", type=Path, default=Path(defecto), help="Ubicación de la base de datos REAL.")
    parser.add_argument("--destino", type=Path, default=Path(__file__).resolve().parent / "base_de_datos_demo.db",
                        help="Ruta NUEVA para la copia de demostración.")
    args = parser.parse_args()
    resultado = preparar_copia(args.origen, args.destino)
    print(f"Copia creada: {args.destino.expanduser().resolve()}")
    print("Original intacto. Solo la copia quedó sin trazabilidad:")
    for tabla, total in resultado.items():
        print(f"  {tabla}: {total} registro(s) omitidos en la copia")
    print("Para usarla, establece HERRAMIENTA_DB_PATH a la ruta de la copia antes de iniciar Streamlit.")


if __name__ == "__main__":
    main()
