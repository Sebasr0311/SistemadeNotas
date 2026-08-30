# Cómo generar el .exe (Windows) con PyInstaller

> ⚠️ **Si cambiás el código, RE-CONSTRUÍ el .exe: el `dist\SistemaNotas.exe`
> vale solo para el commit en que se construyó. No se actualiza solo.**

Estas instrucciones crean un único ejecutable `.exe` de la app para Windows,
sin necesidad de tener Python instalado en la máquina de la usuaria.

> **Usá SIEMPRE el PyInstaller del entorno virtual**
> (`.venv\Scripts\pyinstaller.exe`): el Python del sistema no tiene las
> dependencias de la app instaladas, y compilar contra el entorno equivocado
> produce un .exe que falla al abrir.

## Requisitos

- Windows, con Python 3.13 instalado (u otra versión 3.11+).
- Acceso a internet la primera vez para generar el .exe.

## Pasos

1. Abrí una terminal (PowerShell) en la carpeta del proyecto:

   ```
   cd C:\Users\JUAN\SistemadeNotas
   ```

2. Si todavía no existe, creá y activá el entorno virtual e instalá las
   dependencias de la app y PyInstaller:

   ```
   python -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   pip install pyinstaller
   ```

3. Generá el ejecutable (una sola ventana de aplicación, sin consola) con el
   PyInstaller del venv:

   ```
   .venv\Scripts\pyinstaller --onefile --windowed --name SistemaNotas app.py
   ```

   Opciones:
   - `--onefile`: empaqueta todo en un único `.exe`.
   - `--windowed`: ventana de aplicación, sin consola de terminal.
   - `--name SistemaNotas`: nombre del ejecutable.

4. El ejecutable queda en:

   ```
   C:\Users\JUAN\SistemadeNotas\dist\SistemaNotas.exe
   ```

   Podés copiarlo a donde quieras (por ejemplo, al escritorio) y usarlo sin
   instalar nada más.

## Prueba final (posición de nunca jamás)

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

- **"No module named ..." al compilar**: asegurate de estar corriendo el
  PyInstaller del venv (`.venv\Scripts\pyinstaller.exe`), no el del sistema.
- **El .exe abre y se cierra de inmediato**: revisá `%APPDATA%\SistemaNotas\log.txt`
  para ver el error registrado por la app.