"""
Configuración de la aplicación.

Se encarga de:
- Saber dónde vive la carpeta de configuración del usuario (%APPDATA%/SistemaNotas).
- Cargar y guardar la API key de Gemini y las preferencias en un archivo JSON.
- Proveer rutas útiles (carpeta de logs, carpeta de configuración).

Nunca se debe hardcodear la API key en el código fuente: sólo vive en este
archivo de configuración local del usuario.
"""

import json
import os

# Nombre de la carpeta de configuración dentro de %APPDATA%.
_CONFIG_DIR_NAME = "SistemaNotas"
_CONFIG_FILE_NAME = "config.json"
_LOG_FILE_NAME = "log.txt"


def _appdata_dir() -> str:
    """Devuelve la carpeta %APPDATA% del usuario o un respaldo si no existe."""
    base = os.environ.get("APPDATA")
    if not base:
        # Respaldo para entornos sin %APPDATA% (rara vez pasa en Windows).
        base = os.path.expanduser("~")
    return base


def config_dir() -> str:
    """Carpeta de configuración de la app (se crea si no existe)."""
    path = os.path.join(_appdata_dir(), _CONFIG_DIR_NAME)
    os.makedirs(path, exist_ok=True)
    return path


def config_path() -> str:
    """Ruta completa del archivo de configuración config.json."""
    return os.path.join(config_dir(), _CONFIG_FILE_NAME)


def log_path() -> str:
    """Ruta completa del archivo de log de errores log.txt."""
    return os.path.join(config_dir(), _LOG_FILE_NAME)


def logs_dir() -> str:
    """Carpeta de logs (coincide con la carpeta de configuración por ahora)."""
    return config_dir()


def _config_defaults() -> dict:
    return {
        "api_key": "",
        "modelo_vision": "gemini-2.5-flash-lite",
        "preferencias": {
            "dpi_pdf": 250,
            "carpeta_salida": "",
        },
    }


def load_config() -> dict:
    """
    Carga la configuración desde config.json.
    Si el archivo no existe o está dañado, devuelve la configuración por defecto
    (sin lanzar errores que asusten a la usuaria).
    """
    defaults = _config_defaults()
    path = config_path()
    if not os.path.exists(path):
        save_config(defaults)
        return defaults
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError) as e:
        escribir_log(f"No se pudo leer la configuración ({e}). Se usan valores por defecto.")
        return defaults
    # Combinar con los defaults para no romper si faltan claves nuevas.
    for k, v in defaults.items():
        data.setdefault(k, v)
    return data


def save_config(config: dict) -> None:
    """Guarda la configuración (la API key incluida) en config.json."""
    path = config_path()
    try:
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except OSError as e:
        escribir_log(f"No se pudo guardar la configuración: {e}")
        raise


def get_api_key() -> str:
    """Devuelve la API key guardada, o cadena vacía si no hay."""
    return (load_config().get("api_key") or "").strip()


def set_api_key(api_key: str) -> None:
    """Guarda una API key nueva sin tocar el resto de la configuración."""
    config = load_config()
    config["api_key"] = (api_key or "").strip()
    save_config(config)


def has_api_key() -> bool:
    """Verdadero si ya hay una API key guardada."""
    return bool(get_api_key())


def get_modelo_vision() -> str:
    """Nombre del modelo de visión configurado (por defecto un modelo balanceado)."""
    return load_config().get("modelo_vision") or _config_defaults()["modelo_vision"]


def set_modelo_vision(modelo: str) -> None:
    config = load_config()
    config["modelo_vision"] = modelo
    save_config(config)


# --- Ruta de salida para los Excel generados ---


def default_output_dir() -> str:
    """Carpeta de salida por defecto para los Excel."""
    pref = load_config().get("preferencias", {}).get("carpeta_salida")
    if pref and os.path.isdir(pref):
        return pref
    return os.path.expanduser("~\\Documents")


def set_output_dir(ruta: str) -> None:
    """Guarda la última carpeta usada para guardar los Excel."""
    config = load_config()
    config.setdefault("preferencias", {})["carpeta_salida"] = ruta
    save_config(config)


# --- Log de errores ---


def escribir_log(mensaje: str) -> None:
    """
    Escribe una línea en el archivo de log de errores (log.txt).
    Nunca lanza excepciones: si el log falla, simplemente se ignora.
    """
    try:
        with open(log_path(), "a", encoding="utf-8") as f:
            f.write(mensaje + "\n")
    except OSError:
        pass
