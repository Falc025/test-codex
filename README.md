# Generador Documental Local (MVP)

Aplicación de escritorio en Python para leer datos tabulares desde Excel y generar documentos Word localmente.

## Flujo principal implementado

- La hoja `Datos` se interpreta como tabla (1 registro por fila).
- Se genera **un documento por cada registro**.
- Se usan 3 plantillas según el valor del campo `total`:
  - `total == 0` → plantilla de cero
  - `total > 0` → plantilla de positivo
  - `total < 0` → plantilla de negativo

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
   - Seleccionar `templates/plantilla_total_cero.docx`
   - Seleccionar `templates/plantilla_total_positivo.docx`
   - Seleccionar `templates/plantilla_total_negativo.docx`
   - Seleccionar carpeta `output/`
   - Clic en **Cargar Excel**
   - Clic en **Generar documentos**

## Formato de Excel esperado (hoja `Datos`)

La fila 1 debe tener encabezados. Columnas mínimas requeridas:

- `ruc`
- `razon_social`
- `domicilio`
- `periodo`
- `total`

## Marcadores soportados en Word

La app reemplaza ambos formatos de marcador:

- `{{clave}}`
- `{clave}`

Ejemplos de claves: `{{n° siged}}`, `{{ruc}}`, `{{razon_social}}`, `{{periodo}}`, `{{total}}`, `{{sector}}`.

## Empaquetado base con PyInstaller

```bash
pyinstaller --noconfirm --windowed --name GeneradorDocumental \
  --add-data "templates;templates" \
  --add-data "output;output" \
  main.py
```
