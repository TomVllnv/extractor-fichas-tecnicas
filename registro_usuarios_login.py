"""Alta de encargados desde la pantalla de acceso, bajo autorización de ADMIN.

No se guarda ninguna contraseña en el estado de sesión. La autorización es
breve, vinculada a la revisión de la sesión del administrador y solo permite
crear cuentas con rol 'encargado'. El alta utiliza la auditoría de SQLite.
"""
from __future__ import annotations

import time

import gestion_compartida as gp

SEGUNDOS_AUTORIZACION = 180


def verificar_administrador(ruta: str, usuario: str, contrasena: str) -> dict | None:
    """Solo una cuenta habilitada con rol administrador autoriza un alta."""
    credencial = gp.autenticar(ruta, usuario, contrasena)
    if not credencial or credencial.get("rol") != "administrador":
        return None
    if not gp.es_administrador(ruta, credencial["usuario"]):
        return None
    return {
        "actor": credencial["usuario"],
        "revision_sesion": credencial["revision_sesion"],
        "vence": time.monotonic() + SEGUNDOS_AUTORIZACION,
    }


def autorizacion_vigente(ruta: str, autorizacion: dict | None) -> bool:
    if not isinstance(autorizacion, dict):
        return False
    actor = autorizacion.get("actor")
    vencimiento = autorizacion.get("vence", 0)
    if not actor or not isinstance(vencimiento, (int, float)) or time.monotonic() >= vencimiento:
        return False
    return (gp.usuario_habilitado(ruta, actor, autorizacion.get("revision_sesion"))
            and gp.es_administrador(ruta, actor))


def validar_datos_nueva_cuenta(usuario: str, nombre_visible: str,
                               contrasena: str, repetir_contrasena: str) -> None:
    """Validación antes de generar respaldos o intentar modificar SQLite."""
    gp._validar_nombre_usuario(usuario)
    gp._validar_nombre_visible(nombre_visible)
    if contrasena != repetir_contrasena:
        raise ValueError("Las contraseñas de la nueva cuenta no coinciden.")
    if not contrasena or len(contrasena) < 12:
        raise ValueError("La contraseña debe tener al menos 12 caracteres.")


def crear_encargado_autorizado(ruta: str, autorizacion: dict, revision: int, *,
                               usuario: str, nombre_visible: str, contrasena: str,
                               repetir_contrasena: str) -> str:
    """Crea una cuenta de encargado, sin permitir roles elevados desde el login.

    La copia previa y posterior de SQLite se gestiona en la interfaz y el registro
    de auditoría se realiza en administrar_cuenta, dentro de la transacción.
    """
    if not autorizacion_vigente(ruta, autorizacion):
        raise PermissionError("La autorización de administrador caducó. Vuelve a identificarte.")
    validar_datos_nueva_cuenta(usuario, nombre_visible, contrasena, repetir_contrasena)
    return gp.administrar_cuenta(
        ruta, autorizacion["actor"], revision, "crear", usuario,
        nombre_visible=nombre_visible, rol="encargado", clave=contrasena,
    )
