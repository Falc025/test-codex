# Generador Documental Local (MVP)

Aplicación de escritorio en Python para leer datos desde Excel y generar documentos Word localmente.

## Estructura

```text
.
├── main.py
├── models/
│   └── document_data.py
├── output/
├── scripts/
│   └── create_sample_files.py
├── services/
│   ├── document_builder.py
│   ├── excel_reader.py
│   └── validator.py
├── templates/
├── ui/
│   └── main_window.py
├── utils/
│   └── file_utils.py
└── requirements.txt
```

## Instalación

```bash
python -m venv .venv
# Windows PowerShell:
.venv\Scripts\Activate.ps1
# Linux/macOS:
# source .venv/bin/activate

pip install -r requirements.txt
```

## Ejemplo funcional mínimo

1. Crear archivos de ejemplo (Excel + plantilla):

```bash
python scripts/create_sample_files.py
```

2. Ejecutar la app:

```bash
python main.py
```

3. En la interfaz:
   - Seleccionar `sample_input.xlsx`
   - Verificar plantilla `templates/plantilla_base.docx`
   - Seleccionar carpeta `output/`
   - Clic en **Cargar Excel**
   - Clic en **Generar documento**

## Supuestos de datos del Excel (MVP)

- Hoja requerida: `Datos`
- Estructura clave/valor en columnas A:B
- Campos obligatorios:
  - `expediente`
  - `fecha`
  - `administrado`

## Marcadores de plantilla Word

- `{{expediente}}`
- `{{fecha}}`
- `{{administrado}}`

## Empaquetado base con PyInstaller

```bash
pyinstaller --noconfirm --windowed --name GeneradorDocumental \
  --add-data "templates;templates" \
  --add-data "output;output" \
  main.py
```

> En PowerShell puede usarse el comando en una sola línea para evitar problemas con `\`.

## Buenas prácticas aplicadas

- Arquitectura por capas (UI, servicios, modelo, utilidades)
- Validaciones aisladas y reutilizables
- Modelo de datos con `dataclass`
- Manejo explícito de errores para feedback al usuario
- Preparación para escalabilidad (múltiples plantillas y tablas)
- Resolución de rutas compatible con entorno PyInstaller
