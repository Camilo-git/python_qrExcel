# Diagrama de flujo (Mermaid)

```mermaid
flowchart TD
    A[Inicio app] --> B[Interfaz Tkinter]
    B --> C[Botón Cargar Excel]
    C -->|Selecciona archivo| D{openpyxl lee hoja activa}
    D -->|OK| E[Mostrar columnas A y B en tabla]
    D -->|Error| X[Mensaje de error]

    E --> F{Omitir primera fila?}
    F -->|Sí| G[Preparar datos desde fila 2]
    F -->|No| H[Preparar datos desde fila 1]

    G --> I[Botón Generar QR]
    H --> I[Botón Generar QR]

    I --> J{A y B con datos?}
    J -->|No| K[Omitir fila y contar]
    J -->|Sí| L[Crear QR con qrcode]

    L --> M[Guardar PNG en carpeta img]
    K --> N[Continuar siguiente fila]
    M --> N[Continuar siguiente fila]

    N --> O{Quedan filas?}
    O -->|Sí| J
    O -->|No| P[Mostrar resumen (generados/omitidos)]

    P --> Q[Fin]
```

## Notas de implementación

- Cuando se ejecuta como .exe (PyInstaller), los archivos se guardan junto al ejecutable.
- Los nombres de archivo se sanea sustituyendo caracteres inválidos por `_`.
