# Herramienta de Extracción de Datos de Fichas Técnicas

Aplicación de prueba desarrollada con Python y Streamlit para extraer información de documentos PDF y completar fichas técnicas en Excel.

## Uso

1. Instala Python 3.10 o superior.
2. Abre una terminal en la carpeta del proyecto y ejecuta `pip install -r requirements.txt`.
3. Inicia la herramienta con `streamlit run app.py`.
4. Inicia sesión con las credenciales de prueba indicadas en la pantalla, salvo que se hayan configurado otras mediante `st.secrets`.
5. En el panel lateral izquierdo, selecciona el área de laboratorio (Microbiología, Hematología y coagulación, Bioquímica, Hormonas e inmunidad o Derivados).
6. En el cuadro modal del área, selecciona un examen existente o utiliza los controles para añadirlo o quitarlo de la lista.
7. Selecciona el laboratorio de procesamiento, carga el PDF y pulsa **Analizar documento**.
8. Descarga el Excel actualizado y revisa la información contra el documento original antes de utilizarla.

## Comportamiento de los exámenes

- El nombre del examen se elige desde el menú; no se intenta deducirlo interpretando el texto del PDF.
- El nombre seleccionado se utiliza como nombre de la hoja y como valor de `Nombre LIS`.
- Para que ambos nombres coincidan exactamente, Excel limita el nombre de las hojas a 31 caracteres y prohíbe `\\ / ? * : [ ]`.
- Los exámenes de Microbiología, Hematología y Bioquímica parten de las plantillas Excel almacenadas en SQLite; no es necesario mantener archivos `.xlsx` sueltos en la carpeta.
- Hormonas e inmunidad comienza sin una plantilla de exámenes; al seleccionar un examen nuevo, la aplicación crea su ficha con los campos estándar.
- Las listas de exámenes, laboratorios, encargados y documentos institucionales se guardan en SQLite (`base_de_datos.db`). Si se ejecuta en un alojamiento temporal, configura almacenamiento persistente para la base de datos.

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

En la carpeta `Testeo`, abre PowerShell y ejecuta:

```powershell
py -m venv .venv
.\\.venv\\Scripts\\Activate.ps1
python -m pip install -r requirements.txt
streamlit run app.py
```

Si PowerShell bloquea la activación, utiliza directamente `.venv\\Scripts\\python.exe -m streamlit run app.py` después de instalar las dependencias con ese mismo Python.


## Catálogos en SQLite
Los laboratorios y encargados se almacenan únicamente en `base_de_datos.db`. Las listas «Actualizado por» y «Autorizado por» comparten los mismos nombres, pero las selecciones de cada ficha son independientes. Añadir o eliminar (desactivar) desde cualquiera de las dos pestañas actualiza ambas listas. Se migraron los laboratorios existentes y los encargados predeterminados. No se requieren archivos JSON. Antes de reemplazar una base de datos existente, realiza una copia de seguridad.

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


### Prioridad de llenado de los ocho campos

| Campo de la ficha | Columna de Cartera de Acreditación | Respaldo |
| --- | --- | --- |
| Preparación del paciente | Preparación del paciente · texto | Normas Generales |
| Muestra requerida y recipiente | Muestra y recipiente · texto | Sin respaldo institucional |
| Procedimiento de toma de muestras | Procedimiento de toma · texto | Normas Generales |
| Rotulación | Rotulación · texto | Normas Generales |
| Condiciones de envío al Laboratorio | Traslado · texto | Sin respaldo institucional |
| Criterios de rechazo | Criterios de rechazo · texto | Normas Generales |
| Recepción de muestras | — | Manual de toma de muestras |
| Procesamiento de muestras | — | Manual de toma de muestras |

En los seis primeros campos, la **Cartera de Acreditación tiene prioridad** cuando el campo contiene texto válido y su estado es `Específico` o `Heredado`. El estado `Pendiente` no se toma como instrucción, incluso si la celda tiene texto. `No aplica` se considera un valor válido cuando está documentado. Si el campo de Acreditación no sirve, se usa el respaldo de Normas en los cuatro campos previstos; en los otros dos queda vacío.

**Importante**: Estos ocho campos **no se completan desde los PDF cargados para extracción** ni desde la Cartera tradicional. Los demás campos siguen usando el procesamiento de PDF y las reglas de la Cartera de Exámenes original. Las ediciones manuales confirmadas tienen prioridad cuando se reanaliza el mismo examen en la sesión. Los otros dos campos se obtienen exclusivamente del Manual de toma de muestras, sin consultar la Cartera de Acreditación ni los PDF que el usuario analiza. La tabla conserva los estados y colores actuales. Se añadió un desplegable para consultar la procedencia de los ocho campos, sin cambiar el sistema de estados.

### Vinculación con «Prestación»

Testeo busca el nombre del examen activo en la columna `Prestación` de la Cartera de Acreditación, normalizando mayúsculas, tildes y puntuación. Cuando hay varias coincidencias iguales, diferencia por área. **Nunca elige una coincidencia aproximada de manera automática**: si no encuentra un nombre exacto, el operador puede vincular la prestación correcta desde el selector. Esto evita introducir datos de otro examen por semejanza de nombres.

### Pruebas y respaldo

Para comprobar los datos preinstalados y las reglas de prioridad:

```bash
python -m unittest discover -s tests -v
```

Antes de actualizar SQLite con documentos nuevos, conserva una copia de seguridad de `base_de_datos.db`. Para una instalación existente, usa el archivo SQLite suministrado con esta versión o importa ambas fuentes desde el menú lateral; no reemplaces una base personalizada sin respaldarla.


## Correcciones de interfaz y trazabilidad documental

- Los selectores desplegables usan `index=None` y un texto de ayuda en español, por lo que es posible dejarlos vacíos mediante su control de limpieza **×**. Esto incluye la «Prestación vinculada en la Cartera de Acreditación» y el selector de la vista previa de PDF; dejar la prestación en blanco activa la vinculación automática **exacta**. Si no se selecciona un PDF de vista previa, se utiliza el documento activo.
- Los cargadores de archivos utilizan la apariencia nativa de Streamlit (sin sobreescribir mediante CSS el texto inglés del botón **Upload**).
- El recuento de la Cartera original incluye los exámenes cuyos nombres empiezan por dígitos, como `17-HIDROXIPROGESTERONA` y `17-BETA-ESTRADIOL`: **393** exámenes, no 391.
- En las secciones laterales se muestra el nombre del documento almacenado. Al reemplazar Normas Generales o Acreditación, el nombre se actualiza automáticamente en SQLite. La Cartera de Exámenes almacena el nuevo Excel original (con su nombre) a partir de la próxima actualización. Las plantillas por área también registran el nombre del XLSX cargado o del archivo generado por Testeo.
- La migración de `plantillas_area` añade `nombre_archivo` sin borrar las fichas. Para los archivos cargados en versiones antiguas, se identifica su nombre únicamente cuando coincide byte a byte con la plantilla externa disponible. Si no se conoce la fuente original, se muestra **«Documento de origen: no registrado...»**, en vez de inventar un nombre; se corregirá automáticamente al cargar una nueva versión.
- Para la Cartera de Exámenes cargada en versiones anteriores, no existe copia del Excel original en la base: hasta que se vuelva a actualizar se ofrece la exportación de los registros SQLite, no una falsa copia del original.

Pruebas de regresión: `python -m unittest discover -s tests -v` (sin necesidad de ejecutar Streamlit). La apariencia final del botón de subida y la × deben comprobarse en la versión de Streamlit instalada en el equipo.
