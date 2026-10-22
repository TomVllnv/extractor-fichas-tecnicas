# Herramienta de Extracción de Datos de Fichas Técnicas

Aplicación de prueba desarrollada con Python y Streamlit para extraer información de documentos PDF y completar fichas técnicas en Excel.

## Uso

1. Instala Python 3.10 o superior.
2. Abre una terminal en la carpeta del proyecto y ejecuta `pip install -r requirements.txt`.
3. Inicia la herramienta con `streamlit run app.py`.
4. Inicia sesión con las credenciales de prueba indicadas en la pantalla, salvo que se hayan configurado otras mediante `st.secrets`.
5. En el panel lateral izquierdo, selecciona una de las cuatro áreas: Microbiología, Hematología, Bioquímica o Hormonas e inmunidad.
6. En el cuadro modal del área, selecciona un examen existente o utiliza los controles para añadirlo o quitarlo de la lista.
7. Selecciona el laboratorio de procesamiento, carga el PDF y pulsa **Analizar documento**.
8. Descarga el Excel actualizado y revisa la información contra el documento original antes de utilizarla.

## Comportamiento de los exámenes

- El nombre del examen se elige desde el menú; no se intenta deducirlo interpretando el texto del PDF.
- El nombre seleccionado se utiliza como nombre de la hoja y como valor de `Nombre LIS`.
- Para que ambos nombres coincidan exactamente, Excel limita el nombre de las hojas a 31 caracteres y prohíbe `\\ / ? * : [ ]`.
- Los exámenes de Microbiología, Hematología y Bioquímica parten de las plantillas Excel incluidas en el proyecto.
- Hormonas e inmunidad comienza sin una plantilla de exámenes; al seleccionar un examen nuevo, la aplicación crea su ficha con los campos estándar.
- Las listas añadidas/eliminadas se guardan en `examenes_por_area.json`. En servicios de alojamiento que no conservan los archivos locales entre reinicios, esta lista puede necesitar almacenamiento persistente externo.

## Formato del texto y Excel

- Se normalizan espacios innecesarios y se conservan saltos de línea útiles, como los de las viñetas.
- Se convierten cifras pegadas a letras en notación con subíndice (por ejemplo, `CO2` a `CO₂`), sin convertir números separados por espacios como volúmenes o tiempos.
- El ancho de las columnas se ajusta automáticamente mediante una aproximación compatible con Python y Streamlit Cloud. `Cells.EntireColumn.AutoFit` es una función de automatización de Excel/VBA y no está disponible directamente en `openpyxl` en el servidor Linux.

## Limitaciones

La extracción depende del diseño y calidad del PDF. Encabezados, pies, logotipos, tablas complejas y documentos escaneados pueden producir errores. Los resultados son una propuesta de extracción y deben ser revisados por una persona antes de incorporarlos a procedimientos del laboratorio.

La herramienta no debe completar información técnica faltante por suposición. La traducción automática técnica del inglés al español y una base de datos central persistente no están implementadas en esta versión.
