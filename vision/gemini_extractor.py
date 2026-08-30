"""
Extracción de datos de las planillas usando Gemini (visión).

Por cada página del PDF se envía la imagen al modelo de Gemini con un prompt
en español que pide devolver la planilla en formato JSON. Luego se valida y
normaliza el resultado para conectarlo directo con el generador de Excel.

Salida esperada por página (la MISMA forma que espera `generar_excel_asignatura`):
    {
        "encabezado": {
            "institucion": str, "sede": str, "año_lectivo": str, "jornada": str,
            "grupo": "0GNN" (hasta "05XX"), "asignatura": str, "docente": str,
            "periodo": int (1-4),
        },
        "estudiantes": [
            {
                "no": int, "nombre": str,
                "ev_anteriores": [float|None, ...],
                "area_trabajo": [float|None, float|None],
                "retirado": bool,
                "revisar": [bool, bool],
            }, ...
        ],
    }

Reglas aplicadas:
- Celda en blanco -> None (nunca 0).
- Tachón -> se devuelve el valor vigente (el reescrito).
- Valor dudoso -> se devuelve igual pero con su flag de revisión en True.
- Rango razonable de notas 0-100; fuera de rango -> revisión True.
- Filas "****" (retirado) -> retirado=True, sin datos.
- Se ignoran las columnas "Min" y "Fls.".
"""

import json
import re
import time

from google import genai
from google.genai import types

from config import app_config

# El mejor modelo multimodal con buena relación calidad/precio hoy.
MODELO_POR_DEFECTO = "gemini-3.7-flash"
INTENTOS_MAX = 3
ESPERA_BASE_SEG = 2.0


class VisionError(Exception):
    """Error de visión con mensaje amigable para la usuaria."""


def _cliente(api_key: str) -> genai.Client:
    if not api_key:
        raise VisionError("Falta la clave de Google AI. Cerra la app y configurala de nuevo.")
    return genai.Client(api_key=api_key)


def _rango_grupo_valido(grupo: str) -> bool:
    """
    Formato esperado: 4 dígitos, primero 0, segundo 1-5 (grados 0 a 5),
    ej. "0302", "0401", "0501". Hasta quinto grado ("05XX"), nunca más alto.
    """
    if not grupo or not re.fullmatch(r"\d{4}", str(grupo)):
        return False
    grado = int(str(grupo)[0:2])
    return 0 <= grado <= 5


def _numero_plausible(valor):
    """Devuelve el número si parece una nota razonable, o None/flag según el caso."""
    # Los modelos a veces devuelven cadenas tipo "4 5", "45 " o "4,5".
    if isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return valor
    if isinstance(valor, str):
        limpio = valor.strip().replace(",", ".")
        limpio = re.sub(r"\s+", "", limpio)
        try:
            return float(limpio)
        except ValueError:
            return None
    return None


def _normalizar_area(valores, revisar_flags):
    """
    Devuelve (area_trabajo, revisar) con exactamente 2 celdas.
    - Celdas en blanco -> None.
    - Fuera de rango (no entre 0 y 100) -> marcar revisión True y dejar el valor.
    """
    area = [None, None]
    revisar = [bool(revisar_flags[0]), bool(revisar_flags[1]) if len(revisar_flags) > 1 else False]
    if not valores:
        return area, revisar
    for i in range(2):
        v = valores[i] if i < len(valores) else None
        num = _numero_plausible(v)
        area[i] = num
        if num is not None and not (0 <= num <= 100):
            revisar[i] = True
    return area, revisar


def _normalizar_ev(valores, n_esperadas):
    """Ev. Anteriores: lista de n_esperadas celdas (blanco -> None)."""
    ev = []
    for i in range(n_esperadas):
        v = valores[i] if valores and i < len(valores) else None
        num = _numero_plausible(v)
        ev.append(num)
    return ev


def _normalizar_planilla(datos: dict) -> dict:
    """
    Convierte el JSON crudo del modelo en la forma EXACTA que espera el
    generador de Excel, aplicando todas las reglas de negocio.
    """
    enc_raw = datos.get("encabezado") or {}

    # Clave "año_lectivo" (el generador la usa así).
    periodo = int(enc_raw.get("periodo") or 0)
    if periodo < 1 or periodo > 4:
        periodo = 1

    grupo = str(enc_raw.get("grupo") or "").strip()
    # Si el grupo no cumple el formato esperado, se marca en el encabezado
    # como 'grupo_erroneo' para que la pantalla de revisión lo avise, pero se
    # mantiene el valor leído para no perder información.
    grupo_ok = _rango_grupo_valido(grupo)

    encabezado = {
        "institucion": str(enc_raw.get("institucion") or "").strip(),
        "sede": str(enc_raw.get("sede") or "").strip(),
        "año_lectivo": str(enc_raw.get("año_lectivo") or "").strip(),
        "jornada": str(enc_raw.get("jornada") or "").strip(),
        "grupo": grupo,
        "asignatura": str(enc_raw.get("asignatura") or "").strip(),
        "docente": str(enc_raw.get("docente") or "").strip(),
        "periodo": periodo,
    }
    if not grupo_ok:
        encabezado["grupo_erroneo"] = True

    n_ev_anteriores = periodo - 1
    estudiantes = []
    for est in (datos.get("estudiantes") or []):
        if not isinstance(est, dict):
            continue
        retirado = bool(est.get("retirado")) or _fila_retirada(est)
        nombre = str(est.get("nombre") or "").strip()

        if retirado:
            estudiantes.append({
                "no": int(est.get("no") or 0),
                "nombre": nombre,
                "ev_anteriores": [],
                "area_trabajo": None,
                "retirado": True,
                "revisar": [False, False],
            })
            continue

        rev = est.get("revisar")
        rev_flags = [False, False]
        if isinstance(rev, list):
            rev_flags = [bool(rev[0]), bool(rev[1]) if len(rev) > 1 else False]
        elif isinstance(rev, (int, float)) and not isinstance(rev, bool):
            rev_flags = [rev == 1, False]

        area, revisar = _normalizar_area(est.get("area_trabajo"), rev_flags)
        ev = _normalizar_ev(est.get("ev_anteriores"), n_ev_anteriores)

        estudiantes.append({
            "no": int(est.get("no") or 0),
            "nombre": nombre,
            "ev_anteriores": ev,
            "area_trabajo": area,
            "retirado": False,
            "revisar": revisar,
        })

    return {"encabezado": encabezado, "estudiantes": estudiantes}


def _fila_retirada(est: dict) -> bool:
    """Detecta filas marcadas como retirado por usar asteriscos en el nombre."""
    nombre = str(est.get("nombre") or "")
    return "****" in nombre


def _extraer_json(texto: str) -> dict:
    """Extrae el JSON de la respuesta del modelo, tolerando texto alrededor."""
    texto = (texto or "").strip()
    # Quitar cercos de código si el modelo los usa.
    texto = re.sub(r"^```(?:json)?\s*", "", texto)
    texto = re.sub(r"\s*```$", "", texto)
    try:
        return json.loads(texto)
    except json.JSONDecodeError:
        # Intentar encontrar el primer objeto JSON en el texto.
        m = re.search(r"\{.*\}", texto, re.DOTALL)
        if m:
            return json.loads(m.group(0))
        raise


def extraer_planilla_pagina(
    imagen,
    api_key: str,
    modelo: str = None,
    pagina_label: str = "",
    progreso_cb=None,
):
    """
    Envía una sola página (imagen PIL) al modelo y devuelve la planilla
    normalizada en el formato del generador de Excel.

    Args:
        imagen: PIL.Image de la página.
        api_key: clave de Google AI.
        modelo: nombre del modelo de visión (por defecto el configurado).
        pagina_label: texto para los mensajes de progreso (ej. "Curso 0401").
        progreso_cb: callback opcional que recibe un mensaje de estado.
    """
    if not modelo:
        modelo = app_config.get_modelo_vision() or MODELO_POR_DEFECTO

    client = _cliente(api_key)
    prompt = _PROMPT_PLANILLA

    def _notify(msg):
        if progreso_cb:
            try:
                progreso_cb(msg)
            except Exception:
                pass

    ultimo_error = None
    for intento in range(1, INTENTOS_MAX + 1):
        try:
            if intento > 1 and _notify:
                _notify(f"Reintentando lectura de {pagina_label or 'esta planilla'} (intento {intento})...")
            response = client.models.generate_content(
                model=modelo,
                contents=[prompt, imagen],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.0,
                ),
            )
            texto = response.text
            if not texto:
                raise VisionError("El modelo no devolvió respuesta.")
            datos = _extraer_json(texto)
            return _normalizar_planilla(datos)
        except VisionError as e:
            ultimo_error = e
        except Exception as e:
            # Errores de red/API: reintentar con espera.
            ultimo_error = e
        if intento < INTENTOS_MAX:
            time.sleep(ESPERA_BASE_SEG * intento)

    app_config.escribir_log(f"Error de visión en {pagina_label}: {ultimo_error}")
    raise VisionError(
        "No se pudo leer la planilla. Revisá tu conexión a internet y que la clave "
        "de Google AI sea correcta. Si el problema continúa, podés reintentar."
    ) from ultimo_error


def extraer_planilla_pdf(
    paginas,
    api_key: str,
    modelo: str = None,
    progreso_cb=None,
    por_pagina_cb=None,
):
    """
    Procesa todas las páginas del PDF y devuelve la lista de planillas
    (una por página) en el formato del generador de Excel.

    Args:
        paginas: lista de dicts de `pdf_loader.cargar_paginas`.
        api_key: clave de Google AI.
        modelo: modelo de visión (opcional).
        progreso_cb: callback de estado (mensaje).
        por_pagina_cb: callback opcional que recibe (indice, planilla) tras cada
                       página exitosa (útil para cancelar/saltar).
    """
    planillas = []
    total = len(paginas)
    for i, pag in enumerate(paginas, start=1):
        if progreso_cb:
            progreso_cb(f"Leyendo planilla {i} de {total}...")
        planilla = extraer_planilla_pagina(
            pag["imagen"],
            api_key=api_key,
            modelo=modelo,
            pagina_label=f"planilla {i} de {total}",
            progreso_cb=progreso_cb,
        )
        planillas.append(planilla)
        if por_pagina_cb:
            por_pagina_cb(i, planilla)
    return planillas


# Prompt en español, claro y simple, que le pide al modelo devolver JSON con la
# estructura exacta que espera el generador de Excel.
_PROMPT_PLANILLA = """
Eres un asistente de lectura de planillas de notas escolares en papel.
Te voy a pasar la foto de UNA página de una planilla de notas manuscrita con lapicero.

Debés leer el encabezado y la tabla de estudiantes, y responder SOLAMENTE con
JSON válido (sin texto adicional, sin marcas de código), con esta estructura:

{
  "encabezado": {
    "institucion": "texto",
    "sede": "texto",
    "año_lectivo": "texto",
    "jornada": "texto",
    "grupo": "texto de 4 dígitos, ej. 0302",
    "asignatura": "texto",
    "docente": "texto",
    "periodo": 3
  },
  "estudiantes": [
    {
      "no": 1,
      "nombre": "APELLIDO NOMBRE",
      "ev_anteriores": [45, 45],
      "area_trabajo": [40, 50],
      "retirado": false,
      "revisar": [false, false]
    }
  ]
}

REGLAS IMPORTANTES:

1. ENCABEZADO:
   - "periodo" es un número entero del 1 al 4 (el periodo académico de la planilla).
   - "grupo" es un número de 4 dígitos como "0302" (grado y grupo). Nunca inventes
     un grupo que no esté escrito.
   - "año_lectivo" suele ser un año (ej. "2026").

2. TABLA DE ESTUDIANTES:
   - "ev_anteriores" son las notas definitivas de periodos anteriores (una por cada
     columna que aparezca; si hay 2 columnas van 2 valores, si hay 3 van 3).
   - "area_trabajo" son EXACTAMENTE 2 valores: las dos notas manuscritas del periodo
     actual. Si una celda está en blanco, pon null en esa posición.
   - IGNORA por completo las columnas tituladas "Min" y "Fls.": no las leas.
   - Si una fila está marcada con asteriscos (****) o dice "retirado", pon
     "retirado": true, "area_trabajo": null y "ev_anteriores": [].
   - Si una celda está en blanco o tachada sin valor claro, devolvé null (nunca 0).
   - Si una nota fue corregida (tachada y reescrita), devolvé el valor VIGENTE
     (el reescrito), no el tachado.
   - "revisar" es un arreglo de 2 booleanos, uno por cada celda de "area_trabajo".
     Pon true en la posición de una celda cuyo valor NO estés seguro de haber leído
     bien (letra ambigua, tachón difícil, mancha, número raro). Si estás seguro,
     pon false. Nunca inventes un valor dudoso para evitar la revisión.
   - Las notas suelen ser números enteros o decimales (ej. 45, 40, 4.5, 50). Devolvé
     el número tal cual aparece en la planilla, sin cambiar su escala.

3. Devolvé SOLO el JSON. No agregues explicaciones.
""".strip()
