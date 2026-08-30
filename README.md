# Sistema de Digitalización de Planillas de Notas

Aplicación de escritorio para convertir tus planillas de notas en papel a un
Excel con las notas definitivas ya calculadas.

Está pensada para docentes sin conocimientos técnicos: escaneás las planillas
de una misma asignatura (pueden ser de varios cursos), las cargás en la app, y
la app lee las notas escritas a mano y genera un Excel con una hoja por curso.

## Qué hace

- Lee un PDF con varias planillas escaneadas de una misma asignatura.
- Detecta automáticamente el curso y el periodo de cada planilla.
- Lee las notas manuscritas (con el modelo de visión de Google Gemini).
- Calcula la definitiva de cada estudiante con fórmulas reales de Excel.
- Si la planilla es del periodo 4, agrega la "Definitiva Anual".
- Marca en amarillo cualquier nota dudosa para que la verifiques contra el papel.
- Detecta estudiantes retirados (filas con `****`) y los excluye de los cálculos.

## Cómo instalarlo (para desarrollo)

1. Crear el entorno e instalar dependencias:

   ```
   python -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   ```

2. Ejecutar la app:

   ```
   python app.py
   ```

3. La primera vez la app te pedirá tu clave de Google AI (se obtiene gratis en
   Google AI Studio). La clave queda guardada en tu computador en
   `%APPDATA%\SistemaNotas\config.json`.

   ⚠️ **IMPORTANTE:** esa clave se guarda en **texto plano** en
   `%APPDATA%\SistemaNotas\config.json` y **no debe compartirse**: cualquiera
   que la tenga puede usar tu cuenta de Google AI (y consumir tu cuota). Si
   creés que se comprometió, revocala en Google AI Studio.

## Cómo usarla

1. Al abrir, pega tu clave de Google AI y presiona "Guardar y continuar".
2. En la pantalla principal presiona "Cargar PDF de planillas".
3. Elegí el PDF con las planillas escaneadas de una asignatura.
4. Esperá el progreso (la app está leyendo las planillas).
5. Revisá la pantalla de revisión: las celdas dudosas aparecen en amarillo.
   Corregí un valor tocándolo y escribiendo la nota correcta si hace falta.
6. Presioná "Generar Excel", indicá dónde guardarlo y listo.
7. En la pantalla final podés "Abrir archivo" o "Abrir carpeta".

## Estructura del proyecto

```
app.py                  # punto de entrada (ventana principal)
config/                 # API key y preferencias (%APPDATA%/SistemaNotas)
pdf_processing/         # carga de PDF y conversión a imágenes (PyMuPDF)
vision/                 # lectura de planillas con Gemini
excel/                  # generación del Excel (openpyxl)
gui/                    # interfaz (CustomTkinter): pantallas y proceso en segundo plano
tests/                  # test del generador de Excel
requirements.txt
build_exe.md            # instrucciones para crear el .exe
```

## Registro de errores

Si algo sale mal, la app guarda un registro en
`%APPDATA%\SistemaNotas\log.txt` para ayudar a diagnosticar el problema.
