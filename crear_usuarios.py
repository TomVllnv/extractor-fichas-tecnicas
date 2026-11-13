"""Administra usuarios del laboratorio desde un terminal local del servidor.

Python crear_usuarios.py --usuarios 3
Python crear_usuarios.py --usuarios 1
Python crear_usuarios.py --desactivar operador2
"""
import argparse
import getpass
from gestion_compartida import ruta_base, inicializar, crear_usuario, conexion_cerrada


def main():
    parser = argparse.ArgumentParser(description="Usuarios de la Herramienta de extracción de datos para laboratorio")
    parser.add_argument("--usuarios", type=int, default=3, help="Cantidad de cuentas a crear o actualizar")
    parser.add_argument("--desactivar", metavar="USUARIO", help="Deshabilitar una cuenta sin borrar su historial")
    args = parser.parse_args()
    ruta = ruta_base()
    inicializar(ruta)
    print("Herramienta de extracción de datos para laboratorio — administración de usuarios")
    print("SQLite:", ruta)
    if args.desactivar:
        respuesta = input(f"¿Desactivar el acceso de «{args.desactivar}»? (sí/no): ").strip().casefold()
        if respuesta in ("s", "si", "sí"):
            with conexion_cerrada(ruta) as con:
                con.execute("UPDATE usuarios_aplicacion SET activo=0 WHERE usuario=? COLLATE NOCASE", (args.desactivar,))
            print("Acceso desactivado (el historial se conserva).")
        return
    if not 1 <= args.usuarios <= 20:
        parser.error("--usuarios debe estar entre 1 y 20")
    print("Registra cada encargado con una cuenta y contraseña distintas.")
    for numero in range(1, args.usuarios + 1):
        print(f"\nEncargado {numero}/{args.usuarios}")
        usuario = input("Usuario de acceso: ").strip()
        nombre = input("Nombre visible en historial: ").strip()
        clave = getpass.getpass("Contraseña (mínimo 12 caracteres): ")
        confirmar = getpass.getpass("Repetir contraseña: ")
        if clave != confirmar:
            print("Las contraseñas no coinciden. Esta cuenta no se guardó.")
            continue
        try:
            crear_usuario(ruta, usuario, nombre, clave)
        except (ValueError, OSError) as err:
            print("No se guardó:", err)
        else:
            print("Cuenta creada/actualizada:", usuario)
    print("\nProceso finalizado. No se ha mostrado ni guardado ninguna contraseña en claro.")


if __name__ == "__main__":
    main()
