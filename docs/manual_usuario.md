# Manual de Usuario

Este manual guía el uso de la aplicación "Vista de 2 filas de Excel + QR".

## Requisitos

- Sistema operativo: Windows 10/11 o Linux.
- Python 3.8+ (solo si va a ejecutar desde código). Para el .exe en Windows no es necesario.
- Archivo Excel con datos en la hoja activa. Columnas usadas:
  - Columna A: título (para nombrar la imagen del QR)
  - Columna B: URL o texto que codificará el QR

## Instalación

- Desde ejecutable (Windows):
  1. Descarga `ExcelQR-windows` desde los artefactos del build (GitHub Actions) o compílalo localmente.
  2. Abre la carpeta `dist/ExcelQR/` y ejecuta `ExcelQR.exe`.

- Desde código (Linux/Windows):
  1. Instala dependencias: `pip install -r requirements.txt`
  2. Ejecuta `python app.py`

## Uso paso a paso

1. Abrir la aplicación.
2. Presionar el botón "Cargar Excel (.xlsx)".
3. Seleccionar el archivo Excel. La app mostrará todas las filas de las columnas A y B en la tabla.
4. (Opcional) Dejar marcada "Omitir primera fila (encabezado)" si tu primera fila son títulos de columnas.
5. Presionar "Generar QR".
6. La aplicación creará imágenes PNG de códigos QR por cada fila válida.

## ¿Dónde se guardan los QR?

- Ejecutable (.exe): en la carpeta `img` junto al `ExcelQR.exe`.
- Desde código: en la carpeta `img` junto al archivo `app.py`.

Si ya existe un archivo con el mismo nombre, se agregará un sufijo `_1`, `_2`, etc.

## Formato del archivo Excel

- Hoja activa: se toma la hoja activa del libro.
- Columnas leídas: solo A (título) y B (URL/texto).
- Filas incompletas:
  - Si A o B están vacías, la fila se omite al generar QR y se cuenta como "omitida".

## Mensajes comunes

- "Sin archivo": primero carga un archivo Excel.
- "Sin datos": no se encontraron filas en la hoja activa.
- "No se pudieron generar los QR": revisa que A y B tengan datos y que el archivo no esté bloqueado.

## Consejos y buenas prácticas

- No incluyas caracteres especiales prohibidos en el título de la columna A; la app los reemplazará por `_` para generar nombres de archivo válidos.
- Usa URLs completas en la columna B (por ejemplo, `https://...`).
- Verifica que la hoja correcta esté marcada como activa en Excel.

## Solución de problemas

- En Linux, si aparece un error de Tkinter al iniciar:
  - Instala `python3-tk` desde tu gestor de paquetes.
- Si `openpyxl` no abre tu archivo:
  - Asegúrate de que sea `.xlsx` o `.xlsm`. `.xls` (muy antiguos) no está soportado.
- Si no se generan algunos QR:
  - Revisa si hay filas con A o B vacías.

## Atajos y navegación

- Botón "Limpiar": limpia la tabla para cargar otro archivo.
- Botón "Salir": cierra la aplicación.

## Privacidad

- La aplicación no envía datos a internet. Trabaja localmente con tus archivos y genera imágenes de QR en tu equipo.

## Contacto

- Reporta incidencias o solicitudes desde Issues en el repositorio.
