# Vista rápida de 2 filas de Excel (Tkinter)

Pequeña app de escritorio en Python (Tkinter) que permite seleccionar un archivo Excel (`.xlsx`) y muestra visualmente las primeras dos filas en una tabla.

## Requisitos

- Python 3.8+
- Tkinter (en Linux puede requerir instalar `python3-tk` desde tu gestor de paquetes)

## Instalación de dependencias

Instala las dependencias del proyecto:

```
pip install -r requirements.txt
```

En algunas distros Linux, si Tkinter no está disponible:

```
sudo apt-get update
sudo apt-get install -y python3-tk
```

## Ejecutar

```
python app.py
```

Luego, presiona "Cargar Excel (.xlsx)", elige tu archivo y verás sus primeras dos filas en la tabla.

## Notas

- Se usa `openpyxl` para leer archivos `.xlsx`. Para `.xls` antiguos no está soportado.
- La app toma la hoja activa del libro.
- Celdas vacías se muestran como vacías.

## Documentación

- Manual de Usuario: `docs/manual_usuario.md`
- Diagramas de flujo: `docs/diagramas/flujo.md`

## Construir ejecutable para Windows (.exe)

Hay dos formas:

1) Local en Windows (recomendado)

- Instala Python 3.11 (o superior) y agrega "Add python to PATH" en el instalador.
- En una terminal PowerShell o CMD dentro del proyecto:

```
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --noconfirm --clean --name "ExcelQR" --windowed --add-data "img;img" app.py
```

El ejecutable quedará en `dist/ExcelQR/ExcelQR.exe`. La app guarda los QR en la carpeta `img` junto al `.exe`.

Como alternativa, puedes usar el spec incluido:

```
pyinstaller pyinstaller.spec
```

2) Vía GitHub Actions

- Este repo incluye `.github/workflows/windows-build.yml`.
- Al crear un tag `vX.Y.Z` se dispara el build y podrás descargar el artefacto `ExcelQR-windows` desde la sección de Actions.

### Observaciones

- Si vas a ejecutar el .exe en Windows con archivos `.xlsm`, `openpyxl` no carga macros (pero sí lee celdas). Asegúrate de que los datos estén en la hoja activa.
- Tkinter ya viene con Python oficial para Windows; no requiere dependencias extras.
