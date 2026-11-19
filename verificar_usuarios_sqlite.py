"""Comprueba con acceso de solo lectura cual SQLite y cuantas cuentas hay."""
from pathlib import Path
import os, sqlite3, sys

base=os.environ.get('HERRAMIENTA_DB_PATH')
p=Path(sys.argv[1]).expanduser().resolve() if len(sys.argv)>1 else (Path(base).expanduser().resolve() if base else Path(__file__).resolve().parent/'base_de_datos.db')
print('Base configurada:', p)
if not p.is_file():
    print('ATENCION: no existe el archivo de base de datos en esta ruta.')
    sys.exit(1)
try:
    # Abrir read-only: no crear ni alterar archivos.
    con=sqlite3.connect(p.as_uri()+'?mode=ro', uri=True)
    try:
        existe=con.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='usuarios_aplicacion'").fetchone()[0]
        if not existe:
            print('ATENCION: la tabla usuarios_aplicacion no existe en esta base.')
        else:
            n=con.execute('SELECT count(*) FROM usuarios_aplicacion').fetchone()[0]
            print('Cuentas encontradas:',n)
            if n == 0:
                print('No hay cuentas en ESTA base. Busca una copia anterior.')
            else:
                print('Las cuentas existen en esta base; comprueba que Streamlit use la misma ruta.')
    finally:
        con.close()
except sqlite3.Error as e:
    print('No se pudo leer esta base:',e)
    sys.exit(1)
