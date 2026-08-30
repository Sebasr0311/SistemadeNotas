# Cómo generar el .exe (Windows) con PyInstaller

Estás instrucciones crean un único ejecutable `.exe` de la app para Windows,
sin necesidad de tener Python instalado en la máquina de la usuaria.

## Requisitos

- Windows, con Python 3.13 instalado (u otra versión 3.11+).
- Acceso a internet la primera vez para generar el .exe.

## Pasos

1. Abrí una terminal (PowerShell) en la carpeta del proyecto:

   ```
   cd C:\Users\JUAN\SistemadeNotas
   ```

2. Creá y activá un entorno virtual (opcional pero recomendado):

   ```
   python -m venv .venv
   .venv\Scripts\activate
   ```

3. Instalá las dependencias de la app:

   ```
   pip install -r requirements.txt
   ```

4. Instalá PyInstaller:

   ```
   pip install pyinstaller
   ```

5. Generá el ejecutable (una sola ventana de aplicación, sin consola):

   ```
   pyinstaller --onefile --windowed --name SistemaNotas app.py
   ```

   Opciones:
   - `--onefile`: empaqueta todo en un único `.exe`.
   - `--windowed`: ventana de aplicación, sin consola de terminal.
   - `--name SistemaNotas`: nombre del ejecutable.

6. El ejecutable queda en:

   ```
   C:\Users\JUAN\SistemadeNotas\dist\SistemaNotas.exe
   ```

   Podés copiarlo a donde quieras (por ejemplo, al escritorio) y usarlo sin
   instalar nada más.

## Prueba final

1. Ejecutá `dist\SistemaNotas.exe`.
2. La app debe abrir la primera pantalla pidiendo la clave de Google AI.
3. Pegá una clave válida, presioná "Guardar y continuar" y completá un ciclo
   de carga de un PDF de planillas hasta obtener el Excel.

## Notas

- La primera vez que se abre, el .exe puede tardar unos segundos más (desempaqueta
  los archivos necesarios). Es normal.
- Si Windows muestra un aviso de "protección" al ejecutar el .exe, hay que
  presionar "Más información" y luego "Ejecutar de todas formas". Suele pasar
  con programas sin firma digital; podés firmarlo o agregarlo como excepción.

## Solución de problemas comunes

- **"No module named ..." al compilar**: asegurate de haber instalado las
  dependencias en el mismo entorno desde el que corrés PyInstaller.
- **El .exe abre y se cierra de inmediato**: revisá `%APPDATA%\SistemaNotas\log.txt`
  para ver el error registrado por la app.
