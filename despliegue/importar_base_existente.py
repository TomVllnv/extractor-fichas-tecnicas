"""Importa una base SQLite existente SIN sobrescribir la base del contenedor.

Usar con la aplicación apagada y el archivo origen montado como solo lectura:
 docker compose run --rm --no-deps -v /origen/en/servidor:/importacion:ro \
   herramienta python despliegue/importar_base_existente.py \
   --origen /importacion/base_de_datos.db

Se utiliza la API backup de SQLite para incluir datos confirmados en modo WAL.
"""
from __future__ import annotations

import argparse
import os
import sqlite3
import uuid
from pathlib import Path


def importar(origen: Path, destino: Path) -> None:
    origen = origen.resolve()
    destino = destino.resolve()
    if not origen.is_file():
        raise RuntimeError(f"No se encontró el archivo de origen: {origen}")
    if origen == destino:
        raise RuntimeError("El origen y el destino deben ser diferentes")
    if destino.exists():
        raise RuntimeError(
            "Ya existe una base de datos en el volumen. NO se ha reemplazado. "
            "Solicita a Informática una migración con respaldo previo."
        )

    destino.parent.mkdir(parents=True, exist_ok=True)
    temporal = destino.parent / (f".importacion_{uuid.uuid4().hex}.tmp")
    try:
        # URI de solo lectura: la base origen debe existir y no puede crearse.
        with sqlite3.connect(f"{origen.as_uri()}?mode=ro", uri=True) as fuente:
            if fuente.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
                raise RuntimeError("La base de origen no superó el control de integridad")
            with sqlite3.connect(str(temporal)) as copia:
                fuente.backup(copia)
                if copia.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
                    raise RuntimeError("La copia no superó el control de integridad")
        temporal.chmod(0o600)
        # Creación atómica: nunca sustituir una base que otro proceso creó.
        os.link(temporal, destino)
        print(f"Base importada correctamente en {destino}")
        print("Se conservaron sus tablas, cuentas y plantillas originales.")
    finally:
        temporal.unlink(missing_ok=True)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--origen", required=True, type=Path)
    args = parser.parse_args()
    ruta = os.environ.get("HERRAMIENTA_DB_PATH", "/datos/base_de_datos.db")
    try:
        importar(args.origen, Path(ruta))
    except (sqlite3.Error, OSError, RuntimeError) as exc:
        parser.exit(1, f"IMPORTACIÓN CANCELADA: {exc}\n")


if __name__ == "__main__":
    main()
