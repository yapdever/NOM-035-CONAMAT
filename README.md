# Evaluador de Riesgos Psicosociales

Este programa permite evaluar riesgos psicosociales a partir de un archivo Excel con respuestas, generando reportes individuales y un reporte general en formato Excel. Incluye una interfaz gráfica amigable para el usuario.

## Características

- Interfaz gráfica sencilla (Tkinter)
- Selección de archivo Excel de respuestas
- Selección de carpeta de destino para los reportes
- Generación de reporte general y reportes individuales por trabajador
- Recomendaciones automáticas según el nivel de riesgo

## Requisitos

- Python 3.8 o superior
- Paquetes: `pandas`, `openpyxl`

Instala los requisitos con:

```sh
pip install -r requirements.txt
```

## Uso

1. Ejecuta el programa:

   ```sh
   python main.py
   ```

2. Selecciona el archivo Excel con las respuestas (debe tener una hoja llamada `Respuestas de formulario 1`).
3. Selecciona la carpeta donde se guardarán los reportes.
4. Haz clic en "Procesar y generar reportes".
5. Los archivos generados estarán en la carpeta seleccionada.

## Generar ejecutable para Windows

Puedes crear un `.exe` usando PyInstaller:

```sh
pip install pyinstaller
pyinstaller --onefile --windowed main.py
```

El ejecutable estará en la carpeta `dist`.

## Estructura del proyecto

```
Interfaz/
├── main.py
├── requirements.txt
└── README.md
```

## Créditos

Desarrollado por [Hecotr Andres Herrera Ramos (Yap)].

---

¿Dudas o sugerencias? ¡Contáctanos!
correo yapdever@gmail.com
