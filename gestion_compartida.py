"""Persistencia central y trazabilidad multiusuario de la herramienta del laboratorio.

SQLite permanece en un disco LOCAL persistente. Toda escritura de recursos compartidos
se confirma bajo BEGIN IMMEDIATE y comprobación optimista de revisión.
No guarda contraseñas en claro ni documentos clínicos en el registro de auditoría.
"""
from __future__ import annotations

import hashlib
import hmac
import os
import shutil
import sqlite3
import time
import uuid
from contextlib import contextmanager, closing
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

ZONA = ZoneInfo("America/Santiago")
ITERACIONES = 500_000


class ConflictoConcurrente(RuntimeError):
    """La versión que vio el operador dejó de ser la última."""


def fecha_actual():
    return datetime.now(ZONA).isoformat(timespec="seconds")


def ruta_base():
    """Permite apuntar a /srv/.../datos/base_de_datos.db (volumen persistente)."""
    origen = Path(__file__).resolve().parent / "base_de_datos.db"
    configurada = os.getenv("HERRAMIENTA_DB_PATH", "").strip()
    if not configurada:
        return str(origen)
    destino = Path(configurada).expanduser().resolve()
    if origen == destino:
        return str(origen)
    destino.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    if not destino.exists():
        # Se entrega un SQLite inicial en el ZIP; para el primer uso se copia
        # al disco persistente antes de comenzar a registrar nuevos cambios.
        temporal = destino.parent / ("." + destino.name + "." + uuid.uuid4().hex + ".tmp")
        if origen.exists():
            shutil.copyfile(origen, temporal)
            try:
                # Evita que dos instancias reemplacen un fichero creado por otra.
                os.link(temporal, destino)
            except FileExistsError:
                pass
            finally:
                temporal.unlink(missing_ok=True)
        else:
            with sqlite3.connect(destino):
                pass
    return str(destino)


def conectar(ruta):
    con = sqlite3.connect(ruta, timeout=30, isolation_level=None)
    con.execute("PRAGMA busy_timeout=30000")
    con.execute("PRAGMA foreign_keys=ON")
    return con



@contextmanager
def conexion_cerrada(ruta):
    con = conectar(ruta)
    try:
        yield con
    finally:
        con.close()


def inicializar(ruta):
    with conexion_cerrada(ruta) as con:
        # WAL permite lecturas mientras otra conexión confirma una escritura.
        con.execute("PRAGMA journal_mode=WAL")
        con.execute("PRAGMA synchronous=NORMAL")
        # Migración protegida frente a dos sesiones que se inician a la vez.
        con.execute("BEGIN IMMEDIATE")
        try:
            con.execute("""CREATE TABLE IF NOT EXISTS usuarios_aplicacion (
                usuario TEXT PRIMARY KEY COLLATE NOCASE,
                nombre_visible TEXT NOT NULL,
                sal BLOB NOT NULL,
                clave_derivada BLOB NOT NULL,
                iteraciones INTEGER NOT NULL,
                activo INTEGER NOT NULL DEFAULT 1 CHECK(activo IN (0,1)),
                creado_en TEXT NOT NULL
            )""")
            # Migración compatible: conserva las cuentas, sus claves derivadas y sal.
            # La asignación automática de ADMIN se ejecuta SOLO al agregar la columna.
            columnas_usuario = {fila[1] for fila in con.execute("PRAGMA table_info(usuarios_aplicacion)")}
            migracion_roles = "rol" not in columnas_usuario
            if migracion_roles:
                con.execute("""ALTER TABLE usuarios_aplicacion ADD COLUMN rol TEXT NOT NULL
                            DEFAULT 'encargado' CHECK(rol IN ('administrador','encargado'))""")
                con.execute("""UPDATE usuarios_aplicacion SET rol='administrador'
                            WHERE usuario='ADMIN' COLLATE NOCASE""")
            # Las sesiones se invalidan tras cambios de credenciales o permisos.
            if "sesion_revision" not in {x[1] for x in con.execute("PRAGMA table_info(usuarios_aplicacion)")}:
                con.execute("ALTER TABLE usuarios_aplicacion ADD COLUMN sesion_revision INTEGER NOT NULL DEFAULT 0")
            con.execute("""CREATE TABLE IF NOT EXISTS revisiones_recursos (
                recurso TEXT PRIMARY KEY,
                revision INTEGER NOT NULL DEFAULT 0,
                actualizado_en TEXT,
                actualizado_por TEXT,
                nombre_archivo TEXT
            )""")
            con.execute("""CREATE TABLE IF NOT EXISTS auditoria_cambios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT NOT NULL,
                usuario TEXT NOT NULL,
                recurso TEXT NOT NULL,
                accion TEXT NOT NULL,
                revision_anterior INTEGER NOT NULL,
                revision_nueva INTEGER NOT NULL,
                archivo TEXT,
                detalle TEXT
            )""")
            con.execute("CREATE INDEX IF NOT EXISTS idx_auditoria_fecha ON auditoria_cambios(fecha DESC)")
            con.execute("""CREATE TABLE IF NOT EXISTS eventos_acceso (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT NOT NULL,
                usuario TEXT NOT NULL,
                evento TEXT NOT NULL,
                descripcion TEXT
            )""")
            if migracion_roles:
                admin_preexistente = con.execute("""SELECT usuario FROM usuarios_aplicacion
                    WHERE usuario='ADMIN' COLLATE NOCASE AND rol='administrador'""").fetchone()
                if admin_preexistente:
                    marca = fecha_actual()
                    con.execute("""INSERT OR IGNORE INTO revisiones_recursos
                        (recurso,revision,actualizado_en,actualizado_por)
                        VALUES ('usuarios:administracion',1,?,'Sistema (migración inicial)')""", (marca,))
                    con.execute("""INSERT INTO auditoria_cambios
                        (fecha,usuario,recurso,accion,revision_anterior,revision_nueva,detalle)
                        VALUES (?,'Sistema (migración inicial)','usuarios:administracion',
                            'Asignar rol administrador a ADMIN',0,1,?)""",
                        (marca, 'Se mantuvieron los usuarios y hashes existentes; se habilitó ADMIN.'))
            con.execute("INSERT OR IGNORE INTO revisiones_recursos(recurso,revision) VALUES ('usuarios:administracion',0)")
            con.commit()
        except Exception:
            con.rollback()
            raise
    sincronizar_legado(ruta)


def sincronizar_legado(ruta):
    """Migra metadatos anteriores sin atribuirlos falsamente a un usuario."""
    with conexion_cerrada(ruta) as con:
        tablas = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        if "plantillas_area" in tablas:
            for area, fecha, nombre in con.execute(
                "SELECT area, fecha_actualizacion, nombre_archivo FROM plantillas_area"
            ):
                con.execute("""INSERT OR IGNORE INTO revisiones_recursos
                    (recurso,revision,actualizado_en,actualizado_por,nombre_archivo)
                    VALUES (?,0,?,'No registrado (versión anterior)',?)""",
                    ("area:" + area, fecha, nombre))
        if "documentos_fuente" in tablas:
            for clave, fecha, nombre in con.execute(
                "SELECT fuente_id,fecha_actualizacion,nombre_archivo FROM documentos_fuente"
            ):
                recurso = "fuente:" + clave
                con.execute("""INSERT OR IGNORE INTO revisiones_recursos
                    (recurso,revision,actualizado_en,actualizado_por,nombre_archivo)
                    VALUES (?,0,?,'No registrado (versión anterior)',?)""",
                    (recurso,fecha,nombre))
        if "cartera_metadata" in tablas:
            datos = dict(con.execute("SELECT clave,valor FROM cartera_metadata"))
            if datos.get("ultima_actualizacion"):
                con.execute("""INSERT OR IGNORE INTO revisiones_recursos
                    (recurso,revision,actualizado_en,actualizado_por,nombre_archivo)
                    VALUES ('fuente:cartera_examenes',0,?,
                    'No registrado (versión anterior)',?)""",
                    (datos["ultima_actualizacion"], datos.get("nombre_archivo")))
        for recurso in ["catalogo:laboratorios", "catalogo:encargados"]:
            con.execute("INSERT OR IGNORE INTO revisiones_recursos(recurso,revision) VALUES (?,0)", (recurso,))
        # Las listas por área se versionan individualmente, no como un listado global.
        if "examenes_area" in tablas:
            for (area,) in con.execute("SELECT DISTINCT area FROM examenes_area"):
                con.execute("INSERT OR IGNORE INTO revisiones_recursos(recurso,revision) VALUES (?,0)",
                            ("examenes:" + area,))


def estado_recurso(ruta, recurso):
    with conexion_cerrada(ruta) as con:
        fila = con.execute("SELECT revision,actualizado_en,actualizado_por,nombre_archivo FROM revisiones_recursos WHERE recurso=?", (recurso,)).fetchone()
    if not fila:
        return {"revision": 0, "fecha": None, "usuario": "Sin modificaciones registradas", "archivo": None}
    return {"revision": fila[0], "fecha": fila[1], "usuario": fila[2] or "No registrado (versión anterior)", "archivo": fila[3]}


def mostrar_fecha(fecha):
    if not fecha:
        return "Sin registro"
    try:
        t = datetime.fromisoformat(fecha)
        if t.tzinfo:
            t = t.astimezone(ZONA)
        return t.strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return str(fecha)


def ejecutar_cambio(ruta, recurso, usuario, revision_esperada, accion, escritor,
                   archivo=None, detalle=None):
    """Escritura atómica + control de versión + auditoría en la MISMA transacción.

    escritor(conexion) realiza los cambios sin abrir otra conexión ni confirmar.
    """
    if not usuario or not usuario.strip():
        raise PermissionError("La modificación requiere un usuario identificado.")
    con = conectar(ruta)
    try:
        con.execute("BEGIN IMMEDIATE")
        habilitado = con.execute(
            "SELECT 1 FROM usuarios_aplicacion WHERE usuario=? COLLATE NOCASE AND activo=1",
            (usuario.strip(),)).fetchone()
        if not habilitado:
            raise PermissionError("La cuenta no está habilitada para modificar bases compartidas.")
        if recurso == "usuarios:administracion" and not _es_admin_conexion(con, usuario.strip()):
            raise PermissionError("Solo un administrador puede gestionar las cuentas de acceso.")
        fila = con.execute("SELECT revision FROM revisiones_recursos WHERE recurso=?", (recurso,)).fetchone()
        vigente = fila[0] if fila else 0
        if int(revision_esperada) != vigente:
            raise ConflictoConcurrente(
                "Otro usuario modificó esta información después de que abriste la confirmación. "
                "No se sobrescribió nada. Vuelve a revisar los datos y solicita la actualización nuevamente."
            )
        resultado = escritor(con)
        marca = fecha_actual()
        con.execute("""INSERT INTO revisiones_recursos
            (recurso,revision,actualizado_en,actualizado_por,nombre_archivo)
            VALUES (?,?,?,?,?)
            ON CONFLICT(recurso) DO UPDATE SET revision=excluded.revision,
                actualizado_en=excluded.actualizado_en,
                actualizado_por=excluded.actualizado_por,
                nombre_archivo=COALESCE(excluded.nombre_archivo,revisiones_recursos.nombre_archivo)""",
            (recurso, vigente + 1, marca, usuario.strip(), archivo))
        con.execute("""INSERT INTO auditoria_cambios
            (fecha,usuario,recurso,accion,revision_anterior,revision_nueva,archivo,detalle)
            VALUES (?,?,?,?,?,?,?,?)""",
            (marca, usuario.strip(), recurso, accion, vigente, vigente + 1, archivo, detalle))
        con.commit()
        return resultado
    except Exception:
        con.rollback()
        raise
    finally:
        con.close()


def registrar_acceso(ruta, usuario, evento, descripcion=None):
    with conexion_cerrada(ruta) as con:
        con.execute("INSERT INTO eventos_acceso(fecha,usuario,evento,descripcion) VALUES(?,?,?,?)",
                    (fecha_actual(), usuario, evento, descripcion))



def historial_actividad(ruta, limite=30, desplazamiento=0):
    with conexion_cerrada(ruta) as con:
        return con.execute("""SELECT fecha,usuario,evento,descripcion FROM eventos_acceso
            ORDER BY id DESC LIMIT ? OFFSET ?""",
            (max(1,min(int(limite),200)), max(0,int(desplazamiento)))).fetchall()


def total_historial_actividad(ruta):
    with conexion_cerrada(ruta) as con:
        return con.execute("SELECT COUNT(*) FROM eventos_acceso").fetchone()[0]


def usuario_habilitado(ruta, usuario, revision_sesion=None):
    if not usuario:
        return False
    with conexion_cerrada(ruta) as con:
        fila = con.execute("""SELECT sesion_revision FROM usuarios_aplicacion
            WHERE usuario=? COLLATE NOCASE AND activo=1""", (usuario,)).fetchone()
        if not fila:
            return False
        return revision_sesion is None or int(revision_sesion) == int(fila[0])


def _es_admin_conexion(con, usuario):
    fila = con.execute("""SELECT rol FROM usuarios_aplicacion
        WHERE usuario=? COLLATE NOCASE AND activo=1""", (usuario,)).fetchone()
    return bool(fila and fila[0] == 'administrador')


def rol_usuario(ruta, usuario):
    """Consulta el rol vigente en SQLite en cada interacción (sin confiar en sesión)."""
    if not usuario:
        return None
    with conexion_cerrada(ruta) as con:
        fila = con.execute("""SELECT rol FROM usuarios_aplicacion
            WHERE usuario=? COLLATE NOCASE AND activo=1""", (usuario,)).fetchone()
        return fila[0] if fila else None


def es_administrador(ruta, usuario):
    return rol_usuario(ruta, usuario) == 'administrador'


def listar_usuarios(ruta, solicitante):
    """Datos no sensibles para el panel de administrador, nunca hashes o sales."""
    with conexion_cerrada(ruta) as con:
        if not _es_admin_conexion(con, solicitante):
            raise PermissionError("Solo un administrador puede consultar las cuentas.")
        filas = con.execute("""SELECT usuario,nombre_visible,rol,activo,creado_en
            FROM usuarios_aplicacion ORDER BY usuario COLLATE NOCASE""").fetchall()
    return [dict(zip(('usuario','nombre_visible','rol','activo','creado_en'), f)) for f in filas]


def _validar_nombre_usuario(usuario):
    username = " ".join(str(usuario or "").strip().split())
    if not username or len(username) > 80 or not username.replace('_','').replace('-','').isalnum():
        raise ValueError('El usuario debe tener letras, números, guiones o guion bajo (máximo 80 caracteres).')
    if ' ' in username:
        raise ValueError('El nombre de acceso no puede contener espacios.')
    return username


def _validar_nombre_visible(nombre):
    limpio = " ".join(str(nombre or "").strip().split())
    if not limpio or len(limpio) > 120:
        raise ValueError('Ingresa un nombre visible válido (máximo 120 caracteres).')
    return limpio


def _clave_protegida(clave):
    if not isinstance(clave, str) or len(clave) < 12:
        raise ValueError('La contraseña debe tener al menos 12 caracteres.')
    sal = os.urandom(32)
    return sal, hashlib.pbkdf2_hmac('sha256', clave.encode('utf-8'), sal, ITERACIONES)


def administrar_cuenta(ruta, actor, revision_esperada, operacion, objetivo, *,
                       nombre_visible=None, rol=None, clave=None, activo=None, nuevo_usuario=None):
    """Crea/edita/restablece/desactiva cuentas con auditoría y control de versión.

    No devuelve ni registra contraseñas, sales ni hashes.
    ejecutar_cambio comprueba el rol de actor dentro de la transacción.
    """
    nombre_usuario = _validar_nombre_usuario(objetivo)
    if operacion not in ('crear','editar','clave','estado'):
        raise ValueError('Operación de usuarios no admitida.')
    if operacion in ('crear','editar'):
        nombre_visible = _validar_nombre_visible(nombre_visible)
        if rol not in ('administrador','encargado'):
            raise ValueError('Selecciona un rol válido.')
    if operacion == 'editar':
        nuevo_usuario = _validar_nombre_usuario(nuevo_usuario if nuevo_usuario is not None else nombre_usuario)
    if operacion in ('crear','clave'):
        sal, derivada = _clave_protegida(clave)
    if operacion == 'estado' and not isinstance(activo, bool):
        raise ValueError('Selecciona si la cuenta debe quedar activa o desactivada.')
    acciones = {
        'crear':'Crear usuario', 'editar':'Editar usuario',
        'clave':'Restablecer contraseña', 'estado':'Cambiar estado de usuario'
    }

    def escritor(con):
        existente = con.execute("""SELECT usuario,rol,activo FROM usuarios_aplicacion
            WHERE usuario=? COLLATE NOCASE""", (nombre_usuario,)).fetchone()
        if operacion == 'crear':
            if existente:
                raise ValueError('Ya existe esta cuenta. Usa Editar, Contraseña o Estado.')
            con.execute("""INSERT INTO usuarios_aplicacion
                (usuario,nombre_visible,sal,clave_derivada,iteraciones,activo,creado_en,rol)
                VALUES (?,?,?,?,?,1,?,?)""",
                (nombre_usuario,nombre_visible,sal,derivada,ITERACIONES,fecha_actual(),rol))
        else:
            if not existente:
                raise ValueError('La cuenta seleccionada ya no existe.')
            original, rol_actual, estaba_activo = existente
            nuevo_rol = rol if operacion == 'editar' else rol_actual
            nuevo_activo = int(activo) if operacion == 'estado' else int(estaba_activo)
            # El último administrador activo nunca puede perder el acceso.
            if rol_actual == 'administrador' and int(estaba_activo) == 1 and (
                    nuevo_rol != 'administrador' or nuevo_activo != 1):
                cantidad = con.execute("""SELECT COUNT(*) FROM usuarios_aplicacion
                    WHERE rol='administrador' AND activo=1""").fetchone()[0]
                if cantidad <= 1:
                    raise ValueError('No se puede desactivar ni degradar al último administrador activo.')
            # Evitar cerrar la propia cuenta por error.
            if operacion == 'estado' and original.casefold() == actor.casefold() and not activo:
                raise ValueError('No puedes desactivar tu propia cuenta mientras administras usuarios.')
            if operacion == 'editar':
                cambio_usuario = nuevo_usuario != original
                if cambio_usuario:
                    duplicado = con.execute("""SELECT 1 FROM usuarios_aplicacion
                        WHERE usuario=? COLLATE NOCASE AND usuario<>? COLLATE NOCASE""",
                        (nuevo_usuario, original)).fetchone()
                    if duplicado:
                        raise ValueError('Ya existe otra cuenta con ese nombre de usuario.')
                con.execute("""UPDATE usuarios_aplicacion
                    SET usuario=?,nombre_visible=?,rol=?,sesion_revision=sesion_revision+?
                    WHERE usuario=?""",
                    (nuevo_usuario, nombre_visible, rol,
                     int(cambio_usuario or rol != rol_actual), original))
            elif operacion == 'clave':
                con.execute("""UPDATE usuarios_aplicacion
                    SET sal=?,clave_derivada=?,iteraciones=?,sesion_revision=sesion_revision+1
                    WHERE usuario=?""", (sal, derivada, ITERACIONES, original))
            else:
                con.execute("""UPDATE usuarios_aplicacion
                    SET activo=?,sesion_revision=sesion_revision+1 WHERE usuario=?""",
                    (int(activo), original))
        if operacion == 'editar' and nuevo_usuario != nombre_usuario:
            return f'Editar usuario: {nombre_usuario} → {nuevo_usuario}'
        return f'{acciones[operacion]}: {nombre_usuario}'

    # Solo nombres de acceso en auditoría, jamás contraseñas.
    detalle = f'Cuenta: {nombre_usuario}'
    if operacion == 'editar' and nuevo_usuario != nombre_usuario:
        detalle += f' → {nuevo_usuario}'
    return ejecutar_cambio(ruta, 'usuarios:administracion', actor, revision_esperada,
        acciones[operacion], escritor, detalle=detalle)


def historial(ruta, limite=50, desplazamiento=0):
    with conexion_cerrada(ruta) as con:
        return con.execute("""SELECT fecha,usuario,recurso,accion,revision_nueva,archivo
            FROM auditoria_cambios ORDER BY id DESC LIMIT ? OFFSET ?""",
            (max(1,min(int(limite),200)), max(0,int(desplazamiento)))).fetchall()


def total_historial(ruta):
    with conexion_cerrada(ruta) as con:
        return con.execute("SELECT COUNT(*) FROM auditoria_cambios").fetchone()[0]


def cantidad_usuarios(ruta):
    with conexion_cerrada(ruta) as con:
        return con.execute("SELECT COUNT(*) FROM usuarios_aplicacion WHERE activo=1").fetchone()[0]


def crear_usuario(ruta, usuario, nombre_visible, password):
    username = " ".join(usuario.strip().split())
    nombre = " ".join(nombre_visible.strip().split())
    if not username or len(username)>80 or not username.replace("_", "").replace("-", "").isalnum():
        raise ValueError("El usuario debe contener letras/números, guiones o guion bajo (máximo 80 caracteres).")
    if not nombre or len(nombre)>120:
        raise ValueError("Especifica el nombre visible del encargado.")
    if len(password)<12:
        raise ValueError("La contraseña debe tener al menos 12 caracteres.")
    sal = os.urandom(32)
    clave = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), sal, ITERACIONES)
    with conexion_cerrada(ruta) as con:
        # Herramienta de recuperación local: solo ADMIN puede ser el primer
        # administrador cuando aún no existe ninguno en una instalación nueva.
        administradores = con.execute("""SELECT COUNT(*) FROM usuarios_aplicacion
            WHERE rol='administrador' AND activo=1""").fetchone()[0]
        rol_inicial = ('administrador' if username.casefold() == 'admin' and not administradores
                       else 'encargado')
        con.execute("""INSERT INTO usuarios_aplicacion
            (usuario,nombre_visible,sal,clave_derivada,iteraciones,activo,creado_en,rol)
            VALUES (?,?,?,?,?,1,?,?)
            ON CONFLICT(usuario) DO UPDATE SET nombre_visible=excluded.nombre_visible,
                sal=excluded.sal,clave_derivada=excluded.clave_derivada,
                iteraciones=excluded.iteraciones, activo=1,
                sesion_revision=usuarios_aplicacion.sesion_revision+1""",
            (username, nombre, sal, clave, ITERACIONES, fecha_actual(), rol_inicial))


def autenticar(ruta, usuario, password):
    """Verifica usuario con mayúsculas/minúsculas exactas y contraseña segura."""
    if not usuario or not password:
        return None
    with conexion_cerrada(ruta) as con:
        fila = con.execute("""SELECT usuario,nombre_visible,sal,clave_derivada,iteraciones,rol,sesion_revision
            FROM usuarios_aplicacion WHERE usuario=? COLLATE BINARY AND activo=1""", (usuario.strip(),)).fetchone()
    if not fila:
        # Ejecución con coste similar para evitar distinguir usuarios por tiempo.
        hashlib.pbkdf2_hmac("sha256", password.encode(), b"s" * 32, ITERACIONES)
        return None
    calculada = hashlib.pbkdf2_hmac("sha256", password.encode(), fila[2], int(fila[4]))
    if hmac.compare_digest(calculada, fila[3]):
        return {"usuario": fila[0], "nombre": fila[1], "rol": fila[5], "revision_sesion": fila[6]}
    return None


def respaldar(ruta, retencion=20):
    """Copia consistente usando SQLite Backup API. Fallos no revierten escrituras."""
    original = Path(ruta).resolve()
    backup_dir = Path(os.getenv("HERRAMIENTA_BACKUP_DIR", str(original.parent / "respaldos"))).expanduser()
    backup_dir.mkdir(parents=True, exist_ok=True, mode=0o700)
    hora = datetime.now(ZONA).strftime("%Y%m%d_%H%M%S_%f")
    destino = backup_dir / f"respaldo_{hora}.db"
    temporal = backup_dir / f".{destino.name}.tmp"
    try:
        # sqlite3.Connection como context manager solo confirma o revierte
        # transacciones: NO cierra el archivo. En Windows eso provocaba
        # WinError 32 al intentar renombrar el .db.tmp mientras seguía abierto.
        # closing() libera explícitamente ambos manejadores antes del replace.
        with closing(sqlite3.connect(original, timeout=30)) as src:
            with closing(sqlite3.connect(temporal, timeout=30)) as dst:
                src.backup(dst, pages=512, sleep=0.1)

        # Normalmente ya está libre. Una breve espera cubre antivirus o
        # indexadores de Windows que puedan abrirlo justo al terminar.
        for intento in range(5):
            try:
                temporal.replace(destino)
                break
            except PermissionError as exc:
                if getattr(exc, "winerror", None) not in (32, 33) or intento == 4:
                    raise
                time.sleep(0.2 * (intento + 1))

        todos = sorted(backup_dir.glob("respaldo_*.db"), reverse=True)
        for viejo in todos[max(1, retencion):]:
            try:
                viejo.unlink(missing_ok=True)
            except PermissionError:
                # Una copia antigua abierta no invalida el respaldo nuevo.
                pass
        return str(destino)
    except Exception:
        try:
            temporal.unlink(missing_ok=True)
        except PermissionError:
            # No ocultar el error original si otro proceso mantiene el temporal.
            pass
        raise
