# Generador Masivo de Documentos Word desde Excel

Aplicación de escritorio en **PySide6** para uso interno. Permite cargar un Excel, tres plantillas DOCX y generar documentos masivos conservando formato Word.

## Características

- Selección de Excel, plantillas `cero`, `negativo`, `positivo` y carpeta de salida desde GUI.
- Detección automática de placeholders `{{campo}}` por plantilla.
- Lectura robusta de Excel con normalización de encabezados (minúsculas, espacios, tildes/símbolos).
- Selección de plantilla por regla de negocio sobre `apr_omitido`:
  - `== 0` → `plantilla_total_cero`
  - `< 0` → `plantilla_total_negativo`
  - `> 0` → `plantilla_total_positivo`
- Conversión robusta de `apr_omitido` (número, texto con comas/puntos, espacios, formatos mixtos).
- Reemplazo exclusivo de textos `{{campo}}` en párrafos, tablas, encabezados y pies de página.
- El texto fijo fuera de `{{ }}` **no se modifica**.
- Preservación de formato general de Word (estructura, tablas, estilos de runs).
- Vista previa de datos (primeras filas) y columna calculada de plantilla asignada.
- Vista de placeholders resueltos por fila seleccionada.
- Validación previa obligatoria antes de habilitar generación.
- Tolerancia a errores por fila (si falla una, continúa con las demás).
- Generación de:
  - `log_generacion.txt`
  - `reporte_generacion.xlsx`
- Configuración persistente en `config.json` (últimas rutas, hoja seleccionada).

## Estructura

```text
app/
 ├── main.py
 ├── ui/
 │    └── main_window.py
 ├── core/
 │    ├── excel_reader.py
 │    ├── template_engine.py
 │    ├── validator.py
 │    ├── generator.py
 ├── services/
 │    ├── config_service.py
 │    ├── logging_service.py
 │    ├── report_service.py
 ├── utils/
 │    ├── file_utils.py
 │    ├── text_utils.py
 ├── requirements.txt
 ├── README.md
 └── config.json
```

## Instalación

```bash
python -m venv .venv
# Windows PowerShell
.venv\Scripts\Activate.ps1
# Linux/macOS
# source .venv/bin/activate

pip install -r app/requirements.txt
```

## Ejecución

```bash
python -m app.main
```

## Uso

1. Seleccionar Excel.
2. Seleccionar las 3 plantillas `.docx`.
3. Elegir carpeta de salida.
4. (Opcional) elegir hoja del Excel.
5. Clic en **Validar**.
6. Revisar columnas, placeholders, mapeo/faltantes y vista previa.
7. Clic en **Generar documentos**.
8. Revisar resumen, `log_generacion.txt` y `reporte_generacion.xlsx`.

## Reglas de reemplazo

- Solo se reemplazan tokens con patrón exacto `{{nombre_campo}}`.
- Si un placeholder no tiene columna en Excel, se reemplaza por cadena vacía y se registra advertencia.
- Texto legal fijo, firmas, nombres/cargos y cualquier texto fuera de `{{ }}` no se modifica.

## Build .exe (Windows)

### Opción 1: PyInstaller

```bash
pyinstaller --noconfirm --windowed --name GeneradorMasivoDOCX app/main.py
```

### Opción 2: pyside6-deploy (opcional)

```bash
pyside6-deploy app/main.py
```

## Validación funcional explícita

- ✅ La selección de plantilla por `apr_omitido` funciona.
- ✅ Solo se reemplazan textos `{{campo}}`.
- ✅ No se altera texto fijo fuera de `{{ }}`.
- ✅ El usuario puede cambiar luego Excel/plantillas desde la GUI.
- ✅ Se preserva formato general del documento Word.
- ✅ Se generan logs y reporte final.
