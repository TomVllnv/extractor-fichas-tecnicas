# Despliegue de OCR con Docker (CRS)

Este instructivo es para Informática y los responsables de la aplicación.
El proyecto ya contiene `Dockerfile`, `compose.yaml`, `.dockerignore` y
`.env.example`. **No contiene la base institucional, cuentas, plantillas ni
archivos con contraseñas**. Es indispensable preservar la base de datos que
ya utilizan los encargados.

## 1. Requisitos en el servidor

- Linux de 64 bits con Docker Engine y el complemento Docker Compose v2
  (`docker compose`), con acceso a imágenes y paquetes durante el build.
- Espacio persistente local (SSD) para el volumen Docker, respaldado por TI.
  No montar SQLite sobre SMB/NFS ni ofrecer varias réplicas que compartan
  simultáneamente el mismo archivo SQLite.
- Puertos: Streamlit escucha **8501/TCP dentro del contenedor**.
  Por defecto Compose publica `127.0.0.1:8501`, **solo en el servidor**.
  Para usuarios de la red institucional, configurar un proxy HTTPS en
  **443/TCP** (ejemplo en `despliegue/nginx-ejemplo.conf`).
- Permisos administrativos para el despliegue, supervisión y respaldos.
  No hay que instalar Node.js, Python ni Tesseract en el host: viven en la imagen.

## 2. Primer despliegue, sin una base de datos anterior

Desde una terminal del servidor, en la carpeta **OCR**:

```bash
docker compose build
# Preparar permisos del volumen (solo esta operación se ejecuta como root):
docker compose run --rm --no-deps --user 0 \
  --cap-add CHOWN --cap-add FOWNER herramienta sh -c \
  'chown -R 10001:10001 /datos && chmod 700 /datos'
docker compose run --rm --no-deps herramienta python crear_usuarios.py --usuarios 3
docker compose up -d
docker compose ps
docker compose logs --tail=80 herramienta
```

Durante `crear_usuarios.py` se solicita por terminal el nombre de usuario,
nombre visible y contraseña (no se imprime la contraseña). **La aplicación
requiere una cuenta ADMIN para tareas administrativas**; escribe `ADMIN`
como uno de los usuarios si todavía no existe. La base se inicializa en
`/datos/base_de_datos.db` dentro del volumen persistente. Es una base nueva,
**no trae las carteras ni las plantillas de una instalación anterior**: esas
fuentes deben importarse usando las funciones existentes de la herramienta
si se parte desde cero.

Abre `http://127.0.0.1:8501` **en el servidor** y verifica el inicio de sesión.
La comprobación de salud es `http://127.0.0.1:8501/_stcore/health`.
En un equipo remoto, el acceso se realizará por el dominio HTTPS que configure TI.

## 3. Si ya existen datos y usuarios: migrar ANTES de iniciar

**No sigas el paso 2 para crear una base vacía cuando ya tienes una base
institucional.** Antes de migrar, detén la aplicación anterior de Streamlit,
realiza una copia consistente de SQLite y guarda otro respaldo fuera del servidor.

1. Conserva la base original, por ejemplo
   `/ruta/segura/migracion/base_de_datos.db`. El directorio debe ser legible
   por el usuario del contenedor (UID 10001) y estar montado solo lectura.
2. En `OCR/`, **antes de ejecutar `docker compose up`**, importa la base
   con el siguiente comando (reemplaza la ruta de ejemplo):

```bash
docker compose build
# Preparar permisos del volumen antes de importar:
docker compose run --rm --no-deps --user 0 \
  --cap-add CHOWN --cap-add FOWNER herramienta sh -c \
  'chown -R 10001:10001 /datos && chmod 700 /datos'
docker compose run --rm --no-deps \
  -v /ruta/segura/migracion:/importacion:ro \
  herramienta python despliegue/importar_base_existente.py \
  --origen /importacion/base_de_datos.db
docker compose run --rm --no-deps herramienta python verificar_usuarios_sqlite.py
docker compose up -d
```

El importador usa la función de copia segura de SQLite y **se niega a
sobrescribir** una base existente en el volumen. Si ya se ejecutó antes la
aplicación y se creó una base vacía, detén todo y pide a TI que verifique
qué contiene el volumen **antes de intentar cualquier reemplazo**. Nunca
borres el volumen para «reiniciar» sin haber resguardado los datos.

## 4. Operación diaria

```bash
docker compose ps
docker compose logs -f --tail=100 herramienta
docker compose restart herramienta
docker compose stop herramienta
```

Para crear o actualizar cuentas (por el administrador desde terminal):

```bash
docker compose run --rm --no-deps herramienta python crear_usuarios.py --usuarios 1
```

Los datos se conservan en el volumen Docker `laboratorio_datos`. Su nombre
físico depende del nombre del proyecto Compose (habitualmente
`ocr_laboratorio_datos` si se ejecuta desde `OCR/`). **No cambies el nombre
del directorio, el proyecto Compose ni el identificador del volumen sin
planificar la migración de datos.** `docker compose down` conserva volúmenes;
**`docker compose down -v` puede borrarlos y NO debe utilizarse**.

## 5. Actualización de la herramienta

1. Exporta un respaldo SQLite consistente y copia los respaldos fuera del
   servidor según la política de TI. La carpeta de respaldos dentro del
   mismo volumen **no sustituye** un respaldo independiente.
2. Guarda una copia del código anterior (sin mezclar versiones).
3. Mantén el mismo directorio/proyecto de Compose y el volumen original.
4. Actualiza los archivos del código y ejecuta:

```bash
docker compose build
docker compose up -d --no-deps herramienta
docker compose ps
docker compose logs --tail=80 herramienta
```

No reemplaces manualmente el `.db` durante la actualización. Si debes volver
a una versión anterior, restaura también una base compatible bajo supervisión
de TI si se ejecutaron migraciones de esquema.

## 6. Configuración, acceso y seguridad

- **Variables de entorno:** Compose configura `HERRAMIENTA_DB_PATH` y
  `HERRAMIENTA_BACKUP_DIR` como rutas internas, sin contraseñas. Un archivo
  `.env` **no es obligatorio**. El ejemplo `.env.example` permite cambiar
  el puerto del host si fuese necesario; nunca contiene valores reales.
- **Claves API:** la versión distribuida no requiere claves API externas.
  Si TI agrega integraciones, debe custodiar los secretos fuera del código
  y de la imagen, idealmente mediante el sistema de secretos institucional.
- **HTTPS:** el acceso de usuarios debe publicarse detrás de un proxy HTTPS
  administrado por TI; no exponer directamente 8501 a Internet.
- **Autenticación:** la herramienta usa usuarios almacenados como hashes en
  SQLite; no dispone de SSO institucional integrado. Evaluar aprobación por
  seguridad informática antes de usar información sensible.
- **Permisos:** el proceso de Streamlit se ejecuta sin privilegios de root
  (UID 10001) y el contenedor tiene capacidades Linux retiradas. Solo la
  preparación explícita del volumen se ejecuta momentáneamente como root.
  No se ejecuta como root la aplicación web.
- **Recursos:** dimensionar CPU, RAM, espacio y límites del servicio según
  carga real de PDF/lectura óptica. No se ha medido aún la concurrencia máxima.
- **Documentación:** el `README.md` conserva las instrucciones de uso local.

## 7. Comprobaciones antes de liberar a usuarios

1. `docker compose ps` debe indicar un servicio activo y `healthy`.
2. `docker compose logs --tail=100 herramienta` sin errores recurrentes.
3. Comprobar `docker compose exec herramienta tesseract --list-langs`:
   debe listar `spa`.
4. Probar acceso a la interfaz y login por HTTPS, además de ambos temas.
5. Probar documentos digitales, escaneados, tablas de Derivados y Excel.
6. Confirmar persistencia: reiniciar el contenedor y comprobar que siguen
   los mismos usuarios, carteras, plantillas e historial.
7. Comprobar copia y restauración de respaldos **en un entorno de pruebas**.

**Estado de validación:** los archivos del proyecto se pueden verificar
estáticamente sin Docker; el build real y el acceso HTTPS deben probarse en
un host que tenga Docker Engine/Compose y el proxy institucional disponible.
