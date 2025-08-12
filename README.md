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
