# Herramienta de extracción de datos para laboratorio

Aplicación desarrollada con Python y Streamlit para extraer información de documentos PDF y completar fichas técnicas en Excel.

## Uso

1. Instala Python 3.10 o superior.
2. Abre una terminal en la carpeta del proyecto y ejecuta `pip install -r requirements.txt`.
3. Inicia la herramienta con `streamlit run app.py`.
4. Antes del primer acceso, crea las cuentas del laboratorio con `python crear_usuarios.py`. Cada encargado inicia sesión con su propio usuario y contraseña.
5. En el panel lateral izquierdo, selecciona el área de laboratorio (Microbiología, Hematología y coagulación, Bioquímica, Hormonas e inmunidad o Derivados).
6. En el cuadro modal del área, selecciona un examen existente o utiliza los controles para añadirlo o quitarlo de la lista.
7. Selecciona el laboratorio de procesamiento, carga el PDF y pulsa **Analizar documento**.
8. Descarga el Excel actualizado y revisa la información contra el documento original antes de utilizarla.

## Cambio de área: limpieza automática de la sesión

Al seleccionar un área distinta en la barra lateral, la aplicación **descarta el trabajo
no guardado de esa sesión**: examen activo, selección de la Cartera, campos operacionales,
resultados y ediciones de extracción, libro Excel generado, cola de procesamiento,
archivos PDF cargados y archivo Excel de la pestaña de conversión. Al volver al área
anterior, los selectores y cargadores estarán vacíos.

Esto **no elimina** los exámenes, plantillas ni fuentes almacenadas en SQLite, ni modifica
las cuentas o el historial de trazabilidad. Tampoco borra una ficha que ya se haya
confirmado y guardado en la base de datos. Pulsar nuevamente el área que ya está
activa no descarta el trabajo.

La limpieza de PDF usa una nueva clave del widget `st.file_uploader` para evitar
modificaciones prohibidas de su estado en Streamlit. Los cambios de área se gestionan
mediante el callback del botón antes de que se construyan los widgets.

## Estado de la vinculación exacta en Cartera de Exámenes del CRS

La aplicación calcula el emparejamiento del examen activo con la Cartera del CRS
**antes de dibujar el selector** y lo recalcula cuando cambia el examen, el área
o la revisión de la Cartera almacenada en SQLite. Si un selector queda vacío por
un estado de sesión anterior, se recupera una coincidencia exacta no ambigua;
el usuario puede elegir otra prestación o desmarcarla sin que un rerun deshaga
su elección. Cuando se vincula «FIBRINOGENO», la variante «FIBRINOGENO – HOSMET»
permanece disponible para seleccionarla manualmente.

Las sugerencias muestran «Posibles coincidencias:» subrayado y una línea
separadora entre filas; se aplica a ambas Carteras. La puntuación de similitud
no equivale a una vinculación automática de nombres difusos.

## Reinicio flotante y emparejamiento asistido

Al seleccionar un área de laboratorio aparece un botón circular en la **esquina inferior derecha**:

- En reposo: **flecha circular negra sobre fondo gris muy claro**, sin texto ni emojis.
- Primer clic: la **misma flecha gira durante 3 segundos** a la espera de confirmación.
- Segundo clic dentro del plazo: descarta solo el trabajo temporal de la sesión,
  conservando el área, los usuarios, la base SQLite y el historial. Se muestra un
  destello de **verde suave durante 1,5 segundos** para confirmar.
- Si no hay segundo clic dentro de 3 segundos: se cancela silenciosamente y la
  flecha vuelve a su estado inicial. No hay X, color rojo ni pérdida de información.

Los plazos se inician cuando se presenta el estado en pantalla, para evitar que el
procesamiento de documentos grandes acorte la confirmación o el destello visual.

**Emparejamiento con Cartera de Exámenes del CRS:** se vincula automáticamente el
nombre exacto compatible con el área. En el caso especial en que hay un único
nombre idéntico y otra prestación cuyo nombre lo contiene completo (por ejemplo,
«FIBRINOGENO» y «FIBRINOGENO – HOSMET»), se selecciona el exacto incluso si
falta el dato de área y **la variante permanece visible** como alternativa manual.
Nunca se autovincula una mera similitud parcial. Bajo el selector aparece
«Vinculada: ... · exacta» (o «manual»), y las alternativas van bajo el título
«Posibles coincidencias», sin un aviso azul adicional. La Similitud mostrada
es la similitud real de los nombres; el bonus por área no puede elevarla al 100%.
El selector es opcional. No se descargan propuestas CSV.

**Emparejamiento con Cartera de acreditación:** cuando hay coincidencias
posibles, el **nombre completo del examen es seleccionable** y muestra una gama
del rojo al verde según el porcentaje «Similitud». El botón «Seleccionar» toma
el mismo color y sigue disponible, incluso con una sola sugerencia. Al pasar
el mouse sobre el nombre se muestra texto blanco. El programa no presupone
que un nombre parecido sea el mismo examen.

Los avisos informativos y de advertencia se ajustan al tamaño del texto sin
sobrepasar el espacio disponible, incluso en pantallas pequeñas.

En «Referencias» se mantienen las dos vistas previas APA 7 (solo el PDF técnico
versus todas las fuentes efectivamente utilizadas). Se puede elegir la alternativa
antes de incorporarla y revisar los datos bibliográficos incompletos.

## Actualizar sin perder cuentas ni cambios (IMPORTANTE)

**`Testeo.zip` contiene la carpeta `Testeo/` con todo el código, pero NO contiene `base_de_datos.db`.** Se excluye
intencionalmente para proteger las cuentas `ADMIN`, los otros encargados, sus
contraseñas, documentos, fichas e historial de trazabilidad ya guardados.

1. Detén Streamlit con `Ctrl + C` y respalda tu `base_de_datos.db` actual con la
   aplicación cerrada (y la carpeta `respaldos/`, si existe).
2. Extrae `Testeo.zip` **en una carpeta temporal** (crea `Testeo/`).
3. Copia los archivos de código y recursos de `Testeo/` sobre **la carpeta de la
   instalación que ya funciona y tiene tus cuentas**. No ejecutes otra carpeta
   recién descomprimida si allí no está tu base. **No borres ni reemplaces
   `base_de_datos.db`.** Si usas `HERRAMIENTA_DB_PATH`, conserva
   esa ruta persistente sin cambios.
4. Abre **CMD** en la carpeta donde ya está tu base y ejecuta, uno por uno:
   `python -m pip install -r requirements.txt`, `python verificar_usuarios_sqlite.py`
   y `streamlit run app.py`.
5. Accede con tus usuarios existentes; no es necesario recrearlos.

**No copies una base SQLite mientras Streamlit esté en ejecución.** Para respaldos
con la aplicación activa, utiliza la función de copia consistente integrada.

Tampoco se incluyen plantillas Excel sueltas: las cinco están guardadas en tu SQLite.

## Versión multiusuario (3 encargados) y SQLite persistente

**Esta versión no utiliza la contraseña compartida antigua.** Las tres cuentas se crean por
el responsable técnico en un terminal del servidor. Las contraseñas no se incluyen en el ZIP,
no se guardan en texto legible y no deben subirse a GitHub.

### Prueba en Windows, sin WordPress

1. Descomprime el ZIP en una carpeta del computador; respalda la SQLite anterior si ya actualizaste documentos.
2. Abre PowerShell en la carpeta principal del programa. Para que la información
   permanezca separada de los archivos de código, configura una ruta de datos local:

   ```powershell
   $env:HERRAMIENTA_DB_PATH = "$PWD\datos\base_de_datos.db"
   $env:HERRAMIENTA_BACKUP_DIR = "$PWD\datos\respaldos"
   ```

   Si cambias la ruta de SQLite, primero traslada con la aplicación cerrada
   tu base actual a esa ubicación; este paquete no contiene una base inicial.
   **No cambies la ruta si estás probando con los usuarios que ya creaste.**
3. Instala las dependencias según «Ejecución en Windows» y crea las tres cuentas:

   ```powershell
   python crear_usuarios.py --usuarios 3
   streamlit run app.py
   ```

4. Abre `http://localhost:8501` en tres ventanas de navegador **independientes**
   (puedes usar navegadores distintos o ventanas privadas). Inicia sesión con cada cuenta.
5. Carga un documento en una fuente de SQLite. Antes de guardarlo, verás un diálogo con
   la fecha y el usuario de la última modificación, además de una **casilla de confirmación**.
6. Para simular una colisión: abre la confirmación de la misma fuente en dos sesiones,
   guarda desde la primera y después intenta confirmar desde la segunda. La segunda
   no sobrescribirá la primera; deberá revisar y volver a iniciar la actualización.

No es necesario instalar tres veces el programa: los tres navegadores se conectan
al mismo proceso de Streamlit y a la misma SQLite centralizada.

### Instalación institucional con disco persistente

- Configurar `HERRAMIENTA_DB_PATH` con una **ruta de disco local y persistente del servidor**,
  fuera del árbol de código Git/GitHub (por ejemplo `/datos/laboratorio/base_de_datos.db`).
- Configurar opcionalmente `HERRAMIENTA_BACKUP_DIR` en otro directorio persistente
  del servidor, idealmente en almacenamiento respaldado por TI. El programa genera
  un respaldo antes de confirmar un reemplazo y otro después de cada modificación.
  Conserva un máximo de **20 copias locales recientes** (ajustable en el código).
- La base utiliza **WAL y transacciones SQLite**, con **una escritura a la vez** y
  lecturas concurrentes; la confirmación compara una revisión por recurso.
- Si dos encargados modifican **la misma área**, incluso exámenes diferentes dentro
  de una plantilla común, el segundo guardado se bloquea para evitar sobrescribir
  información. Debe revisar nuevamente la versión actual.
- **No utilizar SQLite/WAL sobre una carpeta de red compartida.** Utilizar un volumen
  local al proceso que ejecuta Streamlit, con respaldo a un almacenamiento externo
  mediante herramientas institucionales. No compartir el archivo `.db` por SMB/NFS.
- Las cuentas tienen acceso equivalente para modificar fuentes, pero **toda modificación
  desde la interfaz necesita confirmación explícita y casilla marcada**.
- La sección lateral **Historial de modificaciones** identifica usuario, fecha,
  fuente, operación, revisión y documento; **Actividad de usuarios** registra
  inicios/cierres de sesión y generación de fichas, sin almacenar los PDF de pacientes en estos registros.
- Los registros heredados pueden mostrar «No registrado (versión anterior)», porque no
  existía historial de autores cuando se creó la base; no se inventan usuarios.
- Para recuperar una copia de seguridad, detener la aplicación y solicitar a TI que
  restaure una copia de SQLite; no copiar el `.db` de una aplicación activa como respaldo.

La persistencia **no queda garantizada** si se ejecuta en Streamlit Community Cloud
con un disco temporal. El servidor institucional debe ofrecer un volumen persistente
verdadero, permisos restringidos, HTTPS, cortafuegos y política de copias externas.
Este proyecto no incluye todavía integración de autenticación con WordPress o SSO institucional;
la cuenta de usuario se controla dentro de SQLite. La instalación definitiva debe
revisarse con el equipo de TI y seguridad del CRS antes de permitir acceso de usuarios reales
u operar con documentos que contengan información de pacientes.

### Administración de usuarios en la aplicación (nuevo)

Desde la sección lateral **Acceso**, la cuenta ya existente cuyo nombre de inicio de sesión
sea `ADMIN` (sin distinguir mayúsculas) tendrá el rol **Administrador** automáticamente al
abrir esta nueva versión. La migración no modifica el hash ni la contraseña de ninguna
cuenta anterior. Los otros usuarios conservan el rol **Encargado** y pueden seguir
actualizando las fuentes con su confirmación correspondiente.

Desde la pantalla **Acceso a herramienta (Login)** también hay una opción
**Crear nuevo usuario**. Al seleccionarla, se abre una ventana emergente donde
se solicitan **usuario y contraseña del administrador**. Tras verificarlos,
el mismo diálogo permite crear una cuenta con nombre de acceso, nombre visible
y contraseña repetida. **Las cuentas creadas desde el Login siempre tienen rol
Encargado**. La autorización dura tres minutos y se invalida si la cuenta
administradora cambia sus credenciales o permisos. Esta opción no funciona si
no existe aún un administrador; la cuenta ADMIN inicial se crea en el servidor
mediante `python crear_usuarios.py`.

La creación realizada desde Login reutiliza el control de versiones, la auditoría
y los respaldos SQLite existentes. Las contraseñas del administrador no se
almacenan en el estado de la aplicación. Para cambiar roles o contraseñas,
utiliza el panel interno **Administrar usuarios**.

El botón flotante de reinicio muestra únicamente el consejo
**«Pulsa dos veces para reiniciar»**. Acepta dos clics rápidos con la misma
clave de control y sin una recarga de página al primer clic. La flecha gira
hasta 3 segundos; tras confirmar, el verde se muestra durante 1,5 segundos.
Si no se pulsa por segunda vez, cancela sin limpiar los datos de la sesión.

El Administrador verá **Administrar usuarios**, con cuatro opciones:

- **Crear usuario**: usuario de acceso, nombre visible, rol y contraseña.
- **Editar usuario**: nombre visible o rol.
- **Contraseña**: establecer una contraseña nueva (la antigua no se puede consultar).
- **Activar/desactivar**: revocar o recuperar acceso sin perder auditoría histórica.

Cada operación exige marcar una confirmación y se registra con fecha, responsable,
acción y revisión en **Trazabilidad > Historial de modificaciones**. Se impide
que el último administrador activo pierda el rol o quede desactivado. Si otro
administrador modifica las cuentas mientras el diálogo está abierto, se debe
actualizar la lista antes de guardar. El cambio de contraseña o rol invalida
las sesiones anteriores de la cuenta afectada.

**Contraseñas:** se almacenan como una derivación criptográfica
**PBKDF2-HMAC-SHA256** con sal aleatoria individual y 500.000 iteraciones,
no en texto claro. Los administradores no pueden visualizarlas, solo cambiarlas.

**Actualización sobre la instalación existente:** cierra Streamlit, respalda
`base_de_datos.db` y copia únicamente los archivos Python del paquete de
actualización sobre el programa instalado. **No copies ni reemplaces el .db**:
así conservarás las tres cuentas, fichas, documentos y modificaciones actuales.
Al iniciar la herramienta, SQLite migrará automáticamente los roles. El ZIP
incremental no incluye base de datos ni plantillas `.xlsx`.

**Si utilizas CMD**, basta con `streamlit run app.py` desde la carpeta del
programa para una prueba local; no necesitas ejecutar comandos de PowerShell
`$env:...`. En el servidor institucional, configura la ruta persistente mediante
sus variables de entorno antes de iniciar Streamlit.

### Administración local de usuarios

```powershell
# Crear los 3 encargados iniciales:
python crear_usuarios.py --usuarios 3

# Añadir o cambiar la contraseña de una sola cuenta:
python crear_usuarios.py --usuarios 1

# Desactivar una cuenta (conservando su historial):
python crear_usuarios.py --desactivar nombre_usuario
```

**Resguardos:** protege el acceso al directorio de datos y sus respaldos: incluyen
plantillas, documentos originales y hashes de contraseñas. No subas la SQLite modificada,
la carpeta `datos/` ni los respaldos a un repositorio público.

### Puesta en marcha mediante Docker (opcional)

El paquete incluye `Dockerfile` y `compose.yaml` como punto de partida para un servidor
institucional. La configuración expone Streamlit solo por `127.0.0.1:8501` del servidor,
para conectarlo más adelante a un proxy HTTPS administrado por TI:

```bash
docker compose build
docker compose run --rm herramienta python crear_usuarios.py --usuarios 3
docker compose up -d
```

El volumen `laboratorio_datos` de Docker retiene SQLite y los respaldos cuando se reinicia
el contenedor. **No equivale por sí solo a un respaldo fuera del servidor**. Antes del despliegue
real deben revisarse permisos del volumen, proxy, certificados TLS y restauración de respaldos.

## Comportamiento de los exámenes

- El nombre del examen se elige desde el menú; no se intenta deducirlo interpretando el texto del PDF.
- El nombre completo seleccionado se conserva en `Nombre LIS` y en el selector del área.
- Excel limita los títulos de hoja a 31 caracteres y prohíbe `\ / ? * : [ ]`. Para prestaciones con nombres largos o estos signos, la herramienta utiliza un nombre **interno de hoja compatible** y la identifica por su `Nombre LIS` completo, sin duplicar fichas al actualizarla.
- Las cinco áreas tienen plantillas Excel almacenadas en SQLite; no es necesario mantener archivos `.xlsx` sueltos en la carpeta.
- Las plantillas nuevas se generaron usando la misma estructura de 24 campos de Hematología y coagulación, con `Código` y `Nombre LIS` tomados de la Cartera del 13/10/2026.
- **Derivados**: 205 exámenes cuya columna `LABORATORIO` es exactamente `DERIVADO EN CONVENIO`. Se excluyen los 17 de `DERIVADO ESPECIAL`.
- **Hormonas e inmunidad**: 14 exámenes cuya columna `LABORATORIO` es exactamente `HORMONAS E INMUNOLOGÍA`.
- Cada plantilla conserva inicialmente en el campo `Laboratorio de procesamiento` el valor correspondiente de `LABORATORIO` de la Cartera. Al procesar, se utiliza el laboratorio que el usuario seleccione.
- Las listas de exámenes, laboratorios, encargados y documentos institucionales se guardan en SQLite (`base_de_datos.db`). Si se ejecuta en un alojamiento temporal, configura almacenamiento persistente para la base de datos.
- Los selectores de exámenes se muestran en orden alfabético sin distinguir tildes ni mayúsculas. Así, `ÍNDICE ANDROGENICO LIBRE` aparece bajo I en Derivados, sin cambiar su nombre original ni los registros de SQLite.

## Formato del texto y Excel

- Se normalizan espacios innecesarios y se conservan saltos de línea útiles, como los de las viñetas.
- Se convierten cifras pegadas a letras en notación con subíndice (por ejemplo, `CO2` a `CO₂`), sin convertir números separados por espacios como volúmenes o tiempos.
- El ancho de las columnas se ajusta automáticamente mediante una aproximación compatible con Python y Streamlit Cloud. `Cells.EntireColumn.AutoFit` es una función de automatización de Excel/VBA y no está disponible directamente en `openpyxl` en el servidor Linux.

## Limitaciones

La extracción depende del diseño y calidad del PDF. Encabezados, pies, logotipos, tablas complejas y documentos escaneados pueden producir errores. Los resultados son una propuesta de extracción y deben ser revisados por una persona antes de incorporarlos a procedimientos del laboratorio.

La herramienta no debe completar información técnica faltante por suposición. La traducción automática técnica del inglés al español y la validación clínica automática no están implementadas en esta versión.

## Ajustes de precisión de esta versión

- Se conserva la interfaz original, los campos de búsqueda, las áreas, los Excel y la base de datos SQLite.
- La extracción PDF se evalúa **página por página**: si una página tiene muy poco texto digital, se intenta OCR solo en esa página. Esto mejora el tratamiento de documentos mixtos (texto y escaneo) sin reemplazar texto digital válido.
- Los errores al abrir un PDF se informan de forma clara; si una página falla durante OCR, el texto legible de las demás se conserva.
- El reconocimiento de etiquetas admite separadores `:`, `-`, `–`, `—` y `：` con espacios opcionales, respetando los alias y campos existentes.
- No se agregan datos clínicos por inferencia ni se alteran registros de la base de datos.

**Importante:** Los valores extraídos siguen necesitando revisión contra el PDF original, especialmente rangos de referencia, unidades, tipo de muestra y estabilidad. La extracción de tablas complejas puede requerir ajustes con ejemplos reales anonimizados.

### Ejecución en Windows

En la carpeta de la herramienta, abre PowerShell y ejecuta:

```powershell
py -m venv .venv
.\\.venv\\Scripts\\Activate.ps1
python -m pip install -r requirements.txt
streamlit run app.py
```

Si PowerShell bloquea la activación, utiliza directamente `.venv\\Scripts\\python.exe -m streamlit run app.py` después de instalar las dependencias con ese mismo Python.


## Catálogos en SQLite
Los laboratorios y encargados se almacenan únicamente en `base_de_datos.db`. Dentro de «Administrar opciones», la pestaña **Encargados** es el único lugar para añadir o desactivar los nombres que alimentan las dos selecciones de la ficha: «Actualizado por» y «Autorizado por». Las selecciones de cada ficha siguen siendo independientes, pero los nombres provienen de un único catálogo en SQLite. Se migraron los laboratorios existentes y los encargados predeterminados. No se requieren archivos JSON. Antes de reemplazar una base de datos existente, realiza una copia de seguridad.

## Dos modalidades de procesamiento PDF

En la pestaña de extracción, selecciona **Un PDF por examen (cola actual)** para conservar el flujo anterior, o **Varios PDF para un mismo examen** para complementar una sola ficha con hasta cinco PDF. En este último modo, selecciona primero el examen y el laboratorio, carga los documentos y pulsa **Analizar todos los PDF del examen**.

Cada PDF se extrae por separado. Los campos encontrados en documentos posteriores completan los campos vacíos del primero; si hay valores diferentes para un mismo campo, se muestran en la revisión para seleccionar el correcto y aplicar el cambio a la ficha Excel. La vista previa corresponde al primer archivo cargado. Los valores de la Cartera institucional siguen aplicándose con sus reglas existentes. Antes de actualizar la plantilla en la herramienta, se deben resolver las discrepancias entre PDF. Revisa siempre el Excel final con los documentos originales.


## Fuentes documentales institucionales — actualización

Esta versión conserva la **Cartera de Exámenes** original de SQLite (393 registros del proyecto base) y mantiene **tres fuentes independientes**:

1. **Normas Generales de Toma de Muestras**: contiene los textos predeterminados de **Preparación del paciente**, **Procedimiento de toma de muestras**, **Rotulación** y **Criterios de rechazo**. El original incorporado es `documentos institucionales.docx.pdf`, la copia PDF disponible del documento Word. Se puede actualizar desde el menú lateral con `.docx` o `.pdf` que incluya los cuatro apartados.
2. **Cartera de Acreditación**: contiene **374 prestaciones** del documento `Cartera_Acreditacion_CRS_v0.2.3.xlsx`, incluida su columna **Prestación**, área y los seis pares de texto/estado. Se puede reemplazar desde el menú lateral por otra versión válida que conserve la hoja `Cartera_Completa` y los encabezados correspondientes.

3. **Manual de toma de muestras**: incluye el PDF original `MANUAL DE TOMA DE MUESTRAS LABORATORIO CLINICO DEL  CRSCO.pdf` y dos textos extraídos y depurados: **Recepción de muestras** y **Procesamiento de muestras**. El primero empieza en «Las muestras deben ser recepcionadas por un técnico paramédico o profesional de la salud» y termina en «cumplen con lo estipulado por el Laboratorio del CRSCO»; el segundo comienza después del título «2. MUESTRAS PROCEDENTES DE LA TOMA DE MUESTRAS DEL LABORATORIO DEL CRSCO» y termina en «secretaría de laboratorio». Las cabeceras, pies, códigos y folios se excluyen según posición en la página y frases delimitadoras. Si el manual actualizado no contiene ambos fragmentos completos, **la actualización se rechaza y el contenido anterior no se modifica**. Solo acepta PDF con texto seleccionable. El nombre del original y la fecha de actualización se actualizan dinámicamente al reemplazarlo.


Las fuentes documentales muestran la **fecha y hora de actualización** en formato día/mes/año, permiten **descargar el archivo original guardado en SQLite** y **actualizarlo**. Las cargas nuevas se validan antes de reemplazar registros; si falla la importación, se conserva el archivo previo.

En el menú lateral **Base de datos (SQLite)**, las fuentes se presentan en orden alfabético mediante **Cartera de exámenes y acreditación** (pestañas **Acreditación** y **Exámenes**) y **Manual de toma de muestras** (pestañas **Manual** y **Normas generales**). Cada pestaña conserva por separado los controles de carga, fecha, nombre dinámico del documento y descarga. Los botones de descarga utilizan solo la primera letra mayúscula; los nombres propios de las áreas se mantienen sin cambios.


### Prioridad de llenado y procedencia de los ocho campos

| Campo de la ficha | Primera prioridad | Respaldo en orden |
|---|---|---|
| Preparación del paciente | **Cartera de Exámenes del CRS**, columna `INDICACIONES` del examen emparejado | **Cartera de acreditación** (`Preparación del paciente · texto`) y luego **Normas generales** |
| Procedimiento de toma de muestras | Cartera de acreditación | Normas generales |
| Rotulación | Cartera de acreditación | Normas generales |
| Criterios de rechazo | Cartera de acreditación | Normas generales |
| Muestra requerida y recipiente | Cartera de acreditación | Sin respaldo genérico |
| Condiciones de envío al Laboratorio | Cartera de acreditación | Sin respaldo genérico |
| Recepción de muestras | Manual de toma de muestras | Sin respaldo alternativo |
| Procesamiento de muestras | Manual de toma de muestras | Sin respaldo alternativo |

**Caso particular de Preparación:** en **Acreditación**, «No aplica» y las frases pendientes (incluyendo la instrucción de consultar manuales del laboratorio externo) NO bloquean el respaldo de Normas. En el primer documento del CRS, el contenido no vacío y válido de `INDICACIONES` prevalece, incluso si dice «No aplica». Si `INDICACIONES` del CRS es válida, tiene prioridad absoluta. Para los otros campos, «No aplica» puede seguir siendo una respuesta institucional válida. Los PDF técnicos analizados no sustituyen estos ocho campos; las ediciones manuales confirmadas prevalecen cuando se vuelve a analizar el mismo examen.

La sección **Procedencia de campos documentales** usa un aspecto de cuaderno con líneas tenues y nombres de campos en azul marino; siempre muestra la fuente realmente seleccionada.

**Emparejamientos:** ambas Carteras muestran sugerencias verticales con nombre en negrita y cursiva, área en texto normal entre paréntesis y porcentaje de Similitud. Las coincidencias exactas se autovinculan; para la Cartera CRS se contempla además el caso particular de un examen exacto único con una variante cuyo nombre contiene íntegramente el examen original. El examen exacto se vincula y la variante permanece elegible. Nombres ambiguos no se autovinculan. Las alternativas parecidas se vinculan mediante «Seleccionar» o desde el menú. La Cartera del CRS usa su columna `LABORATORIO`; Acreditación usa `Área`. La estructura de SQLite se amplía automáticamente para guardar `laboratorio` en `cartera_examenes`, recuperándolo del Excel original ya almacenado en `documentos_fuente` cuando está disponible. Si la base antigua no conserva ese Excel, vuelve a cargar la Cartera desde el panel lateral para completar los laboratorios. No se inventan áreas.

### Vinculación con «Prestación»

La herramienta busca el nombre del examen activo en la columna `Prestación` de la Cartera de Acreditación, normalizando mayúsculas, tildes y puntuación. Cuando hay varias coincidencias iguales, diferencia por área. **Nunca elige una coincidencia aproximada de manera automática**: si no encuentra un nombre exacto, el operador puede vincular la prestación correcta desde el selector. Esto evita introducir datos de otro examen por semejanza de nombres.

### Pruebas y respaldo

Para comprobar los datos preinstalados y las reglas de prioridad:

```bash
python -m unittest discover -s tests -v
```

Antes de actualizar SQLite con documentos nuevos, conserva una copia de seguridad de `base_de_datos.db`. **Este ZIP no incluye SQLite**: conserva siempre la base existente con tus cuentas y tus documentos. Si carece del Excel original del CRS, cárgalo desde el panel lateral para completar los laboratorios; no reemplaces la base.


## Correcciones de interfaz y trazabilidad documental

- Los selectores desplegables usan `index=None` y un texto de ayuda en español, por lo que es posible dejarlos vacíos mediante su control de limpieza **×**. Esto incluye la «Prestación vinculada en la Cartera de Acreditación» y el selector de la vista previa de PDF; dejar la prestación en blanco activa la vinculación automática **exacta**. Si no se selecciona un PDF de vista previa, se utiliza el documento activo.
- Los cargadores de archivos utilizan la apariencia nativa de Streamlit (sin sobreescribir mediante CSS el texto inglés del botón **Upload**).
- El recuento de la Cartera original incluye los exámenes cuyos nombres empiezan por dígitos, como `17-HIDROXIPROGESTERONA` y `17-BETA-ESTRADIOL`: **393** exámenes, no 391.
- En las secciones laterales se muestra el nombre del documento almacenado. Al reemplazar Normas Generales o Acreditación, el nombre se actualiza automáticamente en SQLite. La Cartera de Exámenes almacena el nuevo Excel original (con su nombre) a partir de la próxima actualización. Las plantillas por área también registran el nombre del XLSX cargado o del archivo generado por la herramienta.
- La migración de `plantillas_area` añade `nombre_archivo` sin borrar las fichas. Para los archivos cargados en versiones antiguas, se identifica su nombre únicamente cuando coincide byte a byte con la plantilla externa disponible. Si no se conoce la fuente original, se muestra **«Documento de origen: no registrado...»**, en vez de inventar un nombre; se corregirá automáticamente al cargar una nueva versión.
- Para la Cartera de Exámenes cargada en versiones anteriores, no existe copia del Excel original en la base: hasta que se vuelva a actualizar se ofrece la exportación de los registros SQLite, no una falsa copia del original.

Pruebas de regresión: `python -m unittest discover -s tests -v` (sin necesidad de ejecutar Streamlit). La apariencia final del botón de subida y la × deben comprobarse en la versión de Streamlit instalada en el equipo.

## Ajustes de interfaz — avance de cola y carga de Excel

- El botón **Continuar al siguiente examen (n/total)** aparece en azul cuando quedan documentos PDF en la cola. El numerador es la posición del documento **actual**, no la posición del siguiente (ej.: `1/4` al revisar el primero de cuatro documentos); al hacer clic, avanza un examen.
- Se retiró del panel lateral la carga opcional **Cargar Excel del área (opcional)**, por lo que el análisis siempre usa la plantilla del área almacenada en SQLite (o el libro consolidado de la cola en curso).
- Para conservar **Exportar a otro formato**, esta segunda pestaña dispone de su propio control **Cargar Excel para exportar**, independiente de las plantillas técnicas de SQLite.

## Emparejamiento institucional y referencias (versión actual)

- **Emparejamiento con Cartera de Exámenes del CRS**: usa un flujo equivalente a Acreditación, pero consulta la Cartera institucional del CRS. Selecciona automáticamente un examen exacto no ambiguo y deja visible la variante más larga que contiene el nombre (FIBRINOGENO vs. FIBRINOGENO – HOSMET). Las sugerencias no exactas nunca se asignan sin confirmación. El vínculo es opcional; ya no existe «Continuar sin vinculación» ni «Preparar incorporación».
- **Referencias**: en la revisión de una ficha, el apartado desplegable «Referencias bibliográficas (APA 7) — comparar y aplicar» muestra **dos vistas previas**: «Solo PDF técnico» y «Todas las fuentes utilizadas». El usuario elige la alternativa y pulsa «Aplicar opción seleccionada a Referencias». No se añadieron diálogos emergentes.
- Las referencias se construyen a partir de los **documentos que aportaron información** y sus metadatos verificables. Cuando faltan autor/año/título se indica expresamente en la interfaz y se utiliza «s. f.» si corresponde. La fecha de carga a SQLite **no se considera fecha de publicación**. Una revisión manual de datos incompletos es necesaria antes de dar una referencia por definitiva.
- Las referencias generadas se ordenan alfabéticamente y evitan duplicados. Al guardarlas en Excel, el título de cada documento se escribe en cursiva. Si el campo ya fue editado por el usuario, la herramienta exige una autorización explícita para reemplazarlo.
- **Instalación/actualización**: distribuimos el **código completo sin `base_de_datos.db`** para proteger las tres cuentas, las plantillas institucionales y el historial de cambios. Copia el código nuevo **sobre tu instalación existente, sin borrar la carpeta ni reemplazar SQLite**, y conserva `respaldos/`. Si usas una carpeta nueva, antes de iniciar copia tu `base_de_datos.db` anterior y verifica la ruta mediante `python verificar_usuarios_sqlite.py`.

### Edición de textos extensos y formato de Normas generales

La columna «Información extraída» utiliza el editor de texto amplio de AG Grid, con
capacidad de hasta **32.767 caracteres por celda**, que es el límite técnico del
formato Excel (se ha eliminado el límite predeterminado de 200 caracteres del
editor). Para editar, haz doble clic en la celda y confirma el cambio.

En la fuente «Normas generales de toma de muestras», la herramienta recompone
las viñetas y une los renglones cortados artificialmente por el PDF, conserva
los cuatro apartados y evita la duplicación detectada en la versión original
del procedimiento de extracción venosa. En el Excel de salida, los subtítulos
reconocidos de esta fuente se guardan en **negrita** y se mantienen las viñetas
y los saltos de párrafo. Esta limpieza se aplica también al consultar Normas
ya almacenadas, sin modificar el archivo original ni la base SQLite.

Para conservar el contenido clínico, cualquier edición manual tiene prioridad;
no se eliminan instrucciones médicas por heurísticas amplias ni se mezcla la
preparación de fuentes con menor prioridad.


### Compatibilidad de exportación con Microsoft Excel

- Los campos de las fichas se escriben como **texto literal**, incluso si comienzan por `=`. Esto evita interpretar indicaciones clínicas como fórmulas de Excel.
- Antes de habilitar la descarga, el programa revisa la integridad del paquete `.xlsx`, sus documentos XML y que el libro contenga hojas. Si falla la validación, muestra un aviso en vez de descargar un archivo dañado.
- Cuando no hubo cambios en la tabla, el libro existente se conserva sin serializarlo otra vez, manteniendo el formato ya incorporado.
- Las negritas de Normas generales y cursivas APA se mantienen. La validación interna no equivale a haber abierto el libro en Microsoft Excel; si Excel todavía solicita reparar un archivo, conviene conservar y examinar **el `.xlsx` descargado** y el registro de reparación de Excel.

### Plantilla reconstruida de Hematología y coagulación

La actualización de la plantilla reconstruida se realiza cargando el `.xlsx`
proporcionado por separado desde **Base de datos (SQLite) → Área: Hematología y
coagulación**. Al confirmar esta importación se conservan los datos que ya
estén completos en las fichas anteriores y los exámenes creados después. Las
ediciones confirmadas de fichas dentro de la aplicación siguen guardándose
normalmente, sin recuperación automática que las sobrescriba. Los textos con
viñetas documentales se escriben en el Excel con saltos de línea y subtítulos
en negrita; los datos de la tabla no editados no vuelven a serializar esas
celdas. Mantener respaldos de SQLite antes de actualizar.


### Excel: cadenas enriquecidas y autoajuste de columnas

La descarga `.xlsx` **ya no crea ni modifica manualmente** `xl/sharedStrings.xml`.
Una versión anterior añadía fragmentos vacíos (`<t />`) dentro de cadenas
con negrita y cursiva; Microsoft Excel podía solicitar repararlas. Ahora la
herramienta elimina únicamente los fragmentos vacíos utilizando el escritor
de Excel de `openpyxl` y conserva los saltos de línea, las viñetas y el texto
enriquecido. Si un libro anterior contenía cadenas compartidas, se reescribe
una sola vez sin utilizar conversiones XML manuales. El resto de los libros
se descargan sin reescrituras innecesarias.

Esta corrección no cambia las prioridades de «Preparación del paciente».
El aviso de Microsoft Excel solo puede descartarse definitivamente tras abrir
el archivo descargado en el programa de escritorio; la prueba XML no equivale
a esa apertura.

La herramienta calcula automáticamente el ancho de las columnas de las
plantillas sin depender de Excel. Los anchos se limitan para evitar columnas
desmesuradas con indicaciones clínicas extensas.

**VBA opcional en Microsoft Excel para ajuste exacto:** se incluye
`AutoajustarColumnas.bas`, cuyo procedimiento contiene la instrucción
`Cells.EntireColumn.AutoFit`. Excel no ejecuta macros dentro de archivos
`.xlsx`. Para utilizarla, abre una copia del libro en Excel de escritorio,
guárdala como `.xlsm`, importa el archivo `.bas` desde el editor de VBA y
ejecuta `AutoajustarColumnas`. Esta macro es **opcional** y nunca se ejecuta
automáticamente; solo importa archivos y macros de confianza.

### Presentación de procedencia, referencias y compatibilidad Excel (noviembre 2026)

- **Procedencia de campos documentales:** muestra, bajo cada uno de los ocho campos documentales, el orden de consulta institucional (de mayor a menor prioridad). Se destaca la fuente efectivamente aplicada, aunque sea una fuente de respaldo. Es una visualización: **no cambia las reglas de selección de datos**.
- **Referencias bibliográficas (histórico, antes de 09-11-2026):** el panel se cerraba al confirmar. Este comportamiento **se revirtió** en Test.zip: ahora conserva su estado abierto/cerrado.
- **Excel:** los fragmentos contiguos sin formato se agrupan y se evitan runs separados para cada salto de línea. Se conservan las viñetas, los saltos y los estilos de negrita/cursiva. Este intento de compatibilidad responde al error de Microsoft Excel en la hoja FIBRINOGENO. Las pruebas de Python verifican integridad y conservación del contenido, pero **solo Microsoft Excel puede confirmar que desaparezca el aviso de reparación**; si reaparece, se desactivará el formato enriquecido de las celdas afectadas, manteniendo su texto.
- **SQLite:** ninguna de estas mejoras requiere sustituir `base_de_datos.db` ni modificar los usuarios.

### Interacciones suaves y posición de lectura

- «Referencias bibliográficas» **ya no se cierra automáticamente** al aplicar la opción. Se conserva la identidad del panel y la posición de lectura; la animación de cierre implementada antes se revirtió.
- «Descargar Excel consolidado» y la descarga del Excel importado emplean `on_click="ignore"` para evitar un rerun innecesario.
- «Actualizar en base de datos» ya no provoca un segundo `st.rerun()` redundante al abrir su confirmación.
- `interfaz_fluida.py` incorpora un auxiliar visual que restaura la posición del desplazamiento en las acciones previstas. No accede a datos clínicos ni altera SQLite. En navegadores que impidan el acceso de un iframe local a la página, las acciones seguirán funcionando, aunque el ajuste visual de desplazamiento puede no estar disponible.
- La flecha del botón flotante mantiene la animación CSS pausada en su última orientación al terminar el giro y reanuda desde ese punto cuando se vuelve a activar.

### Optimización de navegación y referencias (8 de noviembre de 2026)

**Histórico:** en la actualización del 09-11-2026 se revirtió el cierre
al confirmar referencias descrito más abajo. Consultar «Actualización Test».

Esta versión conserva los procesos de lectura y extracción de PDF, OCR,
cruce de fuentes documentales, formato APA y exportación a Excel. Los cambios
se limitan al comportamiento visual y a evitar trabajo innecesario:

- Históricamente, el panel principal «Referencias bibliográficas» se
  cerraba tras aplicar una selección (revertido el 09-11-2026). Se eliminó la combinación de animaciones
  CSS + JavaScript que podía provocar el efecto de doble desplegado. Ahora se
  utiliza un único cierre suave (560 ms); el desplazamiento se restaura una
  sola vez tras montarse el panel actualizado.
- Se eliminó el subdesplegable «Datos bibliográficos que conviene verificar».
  **Los avisos bibliográficos no se eliminan**: se presentan como texto visible
  dentro de «Referencias bibliográficas» para mantener la revisión humana.
- Al salir de Bienvenida hacia un área, el navegador oculta temporalmente la
  vista antigua hasta que la nueva vista esté lista, evitando el *ghosting*.
  La ocultación se limita a ese tránsito y posee un respaldo automático.
- El control flotante de reinicio actualiza su fragmento cada 1,0 segundo en
  vez de 0,4 segundos. Siguen vigentes los dos clics y el plazo de 3 segundos.
- Se almacenan solo en la sesión los **metadatos APA ya calculados de PDF**,
  hasta 12 archivos/combinaciones. Una huella SHA-256 obliga a recalcular al
  reemplazar un PDF aunque conserve el mismo nombre. La extracción, OCR y
  los datos de la ficha no están sujetos a esta caché.
- Las descargas de plantillas y fuentes en el panel lateral no ejecutan un
  recálculo completo de Streamlit (`on_click="ignore"`).

**Actualización y reversión.** Mantener el respaldo de la versión anterior del
código antes de publicar la nueva. **No reemplazar ni subir a GitHub la base de
datos SQLite activa**, los directorios `datos/` y `respaldos/`, ni credenciales.
Si la interfaz no se comporta según lo previsto, restaurar desde el ZIP
original los archivos de código y reiniciar Streamlit; no restaurar una base
de datos antigua, ya que podría borrar el trabajo de otros operadores.

**Verificación.** Pruebas estáticas y unitarias con `python -m pytest -q`
(requiere las dependencias en `requirements.txt`). Parte de las pruebas
históricas utiliza una base de datos de demostración con plantillas,
393 exámenes y otros documentos, que **no forma parte de este paquete**;
por ello no es posible ejecutar esa parte del conjunto sin la base de datos de
pruebas correspondiente. También es necesario comprobar manualmente la
transición en la versión de Streamlit instalada en el servidor.


## Actualización Test (9 de noviembre de 2026)

El archivo **Test.zip** parte de la versión optimizada del 8 de noviembre y
**no contiene base_de_datos.db**. Para actualizar una instalación existente,
reemplazar los archivos de código conservando la base SQLite, usuarios,
documentos, respaldos y secretos de la instalación.

- **Excel:** la validación de descarga inspecciona cadenas enriquecidas
  (APA y Normas), reserializa fragmentos con propiedades de texto vacías
  y comprueba el límite de 32.767 unidades UTF-16 por celda. Si una celda
  lo supera, muestra su referencia y evita descargar un XLSX defectuoso,
  sin truncar contenido clínico. Para confirmar el aviso «sheet11.xml»
  se necesita probar en Microsoft Excel el archivo original que falló.
- **Ambas carteras:** cada sugerencia tiene un color acorde a su similitud,
  entre rojo (55 %) y verde (100 %). Pulsar el nombre del examen o el botón
  «Seleccionar» produce la misma vinculación, siempre manual en candidatos
  difusos. El texto se torna blanco al posar el cursor sobre el nombre.
- **PDF:** dos flechas laterales cambian de página y se sincronizan con el
  slider. Los PDF de un lote conservan sus páginas por separado.
- **Referencias:** al pulsar «Aplicar opción seleccionada a Referencias»,
  no se cierra ni se comprime automáticamente el desplegable.

La lógica de extracción, OCR, comparaciones y edición de los datos clínicos
no se modificó. Detalles y verificaciones en `CAMBIOS_TEST_2026-11-09.txt`.
