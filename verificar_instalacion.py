"""Diagnóstico seguro de SQLite, cuentas y fuentes, sin revelar contraseñas."""
from gestion_compartida import ruta_base, inicializar, conexion_cerrada, cantidad_usuarios


def main():
    ruta = ruta_base()
    inicializar(ruta)
    print("Herramienta de extracción de datos para laboratorio")
    print("Ubicación SQLite:", ruta)
    with conexion_cerrada(ruta) as con:
        print("Integridad:", con.execute("PRAGMA integrity_check").fetchone()[0])
        print("Modo SQLite:", con.execute("PRAGMA journal_mode").fetchone()[0])
        for tabla in ("cartera_examenes", "prestaciones_acreditacion", "plantillas_area",
                      "normas_generales_campos", "manual_toma_campos", "auditoria_cambios"):
            try:
                n = con.execute(f"SELECT COUNT(*) FROM {tabla}").fetchone()[0]
            except Exception as exc:
                print(tabla, "— no disponible:", exc)
            else:
                print(tabla + ":", n)
        print("Cuentas activas:", cantidad_usuarios(ruta))


if __name__ == "__main__":
    main()
