# Sistema de Digitalización de Planillas de Notas

Aplicación de escritorio para docentes: escaneás las planillas de notas en
papel de una asignatura, la app lee lo manuscrito con visión por IA y te
genera un Excel con una hoja por curso y las definitivas ya calculadas.

## Cómo se usa

1. **Primera vez, pegá tu clave de Google AI** (gratis en Google AI Studio) y
   presioná "Guardar y continuar".
2. **Elegí el PDF** con las planillas escaneadas de una asignatura (puede tener
   varios cursos: la app los ordena solos).
3. **Revisá lo que se leyó**: las notas dudosas aparecen en amarillo. Tocá una
   celda y corregí el valor si hace falta.
4. **Generá el Excel**: indicá dónde guardarlo y listo. En la pantalla final
   podés abrir el archivo o la carpeta.

> Detalles útiles: los estudiantes duplicados por un escaneo solapado se
> descartan automáticamente; las filas retiradas (`****`) se excluyen de los
> cálculos; en el periodo 4 se agrega la "Definitiva Anual".

## Dónde quedan los archivos

| Qué | Dónde |
|-----|-------|
| Excel generado | La carpeta que elijas al guardar (recuerda la última usada) |
| Clave de Google AI | `%APPDATA%\SistemaNotas\config.json` |
| Registro de errores | `%APPDATA%\SistemaNotas\log.txt` |

## Cómo correrla desde el código (desarrollo)

1. Creá el entorno e instalá las dependencias:

   ```
   python -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   ```

2. Ejecutá la app:

   ```
   python app.py
   ```

> ¿Usuaria final? No necesitás nada de esto: existe un `.exe` ya armado.
> Consultá `build_exe.md` para conocer su ubicación y cómo reconstruirlo si
> cambia el código.

## Solución de problemas

- **"Clave inválida" o no responde la app**: revisá que la clave de Google AI
  sea correcta y que tengas conexión a internet. Podés cambiarla desde el
  botón "Cambiar la clave" de la pantalla principal.
- **Una página no se leyó**: la app genera las que sí se leyeron y te avisa
  cuáles fallaron. Verificá que esa planilla esté bien escaneada (sin manchas
  ni cortes) y escaneala de nuevo.
- **Ninguna página se leyó**: revisá la calidad del escaneo y la conexión.
- **El .exe abre y se cierra de inmediato**: revisá
  `%APPDATA%\SistemaNotas\log.txt` para ver el error.

## Seguridad

⚠️ **Tu clave de Google AI se guarda en texto plano** en
`%APPDATA%\SistemaNotas\config.json` y **no debe compartirse**: cualquiera que
la tenga puede usar tu cuenta de Google AI (y consumir tu cuota). Si creés que
se comprometió, revocala en Google AI Studio y pegá una nueva.

## Estructura del proyecto

```
app.py                  # punto de entrada (ventana principal)
config/                 # API key y preferencias (%APPDATA%\SistemaNotas)
pdf_processing/         # carga de PDF y render perezoso de páginas (PyMuPDF)
vision/                 # lectura de planillas con Gemini
excel/                  # generación del Excel (openpyxl) y agrupación compartida
gui/                    # interfaz (CustomTkinter): pantallas y worker en segundo plano
tests/                  # tests del pipeline completo y de cada capa
requirements.txt
build_exe.md            # instrucciones para crear el .exe
```