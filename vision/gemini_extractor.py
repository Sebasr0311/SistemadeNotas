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
                "area_trabajo": [float|None, ...],  # 1 a MAX_N_AREAS notas
                "retirado": bool,
                "revisar": [bool, ...],  # alineado a len(area_trabajo)
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

# Modelo por defecto: gemini-2.5-flash-lite tiene el nivel gratuito más
# generoso y confiable de Gemini (sin tarjeta de crédito, ~1000-1500
# solicitudes/día). "gemini-3.7-flash" es más potente pero NO tiene cuota
# gratuita garantizada; se deja como opción para quien sí quiera pagar.
MODELO_POR_DEFECTO = "gemini-2.5-flash-lite"
MODELO_RESPALDO_PAGO = "gemini-3.7-flash"
INTENTOS_MAX = 3
ESPERA_BASE_SEG = 2.0

# Rango razonable para una nota (0 a 100). Fuera de rango -> revisión.
RANGO_NOTA = (0, 100)

# Sanidad de la cantidad de alumnos por planilla (spec v2): el spec pedía
# 40-50; se usan límites generosos para no dar falsas alarmas con cursos
# chicos o planillas con más alumnos de lo habitual.
MIN_ESTUDIANTES_SANOS = 10
MAX_ESTUDIANTES_SANOS = 60

# Cantidad máxima de notas de "área de trabajo" por alumno (spec v2): el rango
# válido es de 1 a 16 notas por alumno. Debajo se trunca a este tope de
# seguridad; la validación de planilla es la que se encarga de chequear la
# consistencia con n_area_trabajo.
MAX_N_AREAS = 16


class VisionError(Exception):
    """Error de visión con mensaje amigable para la usuaria."""


class _PlanillaParseError(VisionError):
    """Error de PARSEO del contenido devuelto por el modelo (no de red/API).
    No se reintenta: el problema es que no se pudo interpretar la planilla."""


def _cliente(api_key: str) -> genai.Client:
    if not api_key:
        raise VisionError("Falta la clave de Google AI. Cerra la app y configurala de nuevo.")
    # Timeout de 120 s (120000 ms): el log muestra "The read operation timed
    # out" en planillas reales a 60 s incluso con flash-lite + dpi 200; 120 s
    # da margen real sin dejar el hilo colgado (una conexión muerta igualmente
    # se corta y Cancelar sigue siendo útil). (Verificado por introspección:
    # google-genai 2.20.0 acepta
    # Client(..., http_options=types.HttpOptions(timeout=<ms>)).)
    return genai.Client(
        api_key=api_key,
        http_options=types.HttpOptions(timeout=120000),
    )


def _imagen_de(pag):
    """
    Devuelve la imagen PIL de una página, tolerando los dos formatos.

    Las páginas nuevas de `pdf_loader.cargar_paginas` (S7) NO traen la imagen
    materializada: exponen `_render`, un callable que rasteriza ESA página
    recién cuando se lo llama. Los fixtures/tests legacy traen la imagen ya
    materializada en la clave "imagen". El PDF nunca contiene "imagen" desde
    pdf_loader: se verifica acá que el render lazy sea el camino normal.
    """
    render = pag.get("_render") if isinstance(pag, dict) else None
    if callable(render):
        return render()
    return pag["imagen"]


def _rango_grupo_valido(grupo: str) -> bool:
    """
    Formato esperado: 4 dígitos, primero 0, segundo 1-5 (grados 0 a 5),
    ej. "0302", "0401", "0501". Hasta quinto grado ("05XX"), nunca más alto.
    """
    if not grupo or not re.fullmatch(r"\d{4}", str(grupo)):
        return False
    grado = int(str(grupo)[0:2])
    return 0 <= grado <= 5


def _parsear_periodo(valor):
    """
    Parsea el periodo de forma tolerante (int, "3", "3.0", " 3 ").
    Devuelve (periodo, ok). ok=False si no se puede convertir o queda fuera del
    rango 1-4; en ese caso se devuelve el valor leído (entero si se puede) para
    que la pantalla de revisión lo avise, sin corregirlo silenciosamente.
    """
    if isinstance(valor, bool):
        return 0, False
    if isinstance(valor, (int, float)):
        p = int(valor)
    elif isinstance(valor, str):
        try:
            p = int(float(valor.strip()))
        except (ValueError, TypeError):
            return 0, False
    else:
        return 0, False
    return p, (1 <= p <= 4)


def _coercion_bool(valor):
    """
    Coerción explícita de un valor booleano leído del modelo (los modelos suelen
    devolver strings tipo "false"/"true"; bool("false") es True, un bug que esta
    función corrige).
    Devuelve (booleano, confiable). confiable=False indica un valor no
    reconocido: se conserva y se marca para revisión, nunca se adivina.
    true/1/sí/si/yes (case-insensitive) -> True
    false/0/no / vacío/None -> False
    """
    if valor is None:
        return False, True
    if isinstance(valor, bool):
        return valor, True
    if isinstance(valor, (int, float)):
        return valor != 0, True
    s = str(valor).strip().lower()
    if s in ("true", "1", "1.0", "sí", "si", "yes"):
        return True, True
    if s in ("false", "0", "0.0", "no", ""):
        return False, True
    return bool(valor), False


def _numero_plausible(valor):
    """
    Devuelve (valor, confiable).

    confiable=False => el valor crudo es ambiguo (dígitos muy separados tipo
    "4 5") o no es un número: se devuelve None para que la pantalla de revisión
    lo corrija. NUNCA se fusionan dígitos separados (S6): "4 5" no es 45.
    Los bools no son notas y no son ambiguos: (None, True).
    """
    if isinstance(valor, bool):
        return None, True
    if isinstance(valor, (int, float)):
        return valor, True
    if isinstance(valor, str):
        s = valor.strip().replace(",", ".")
        # Decimales españoles con espacios sueltos alrededor del punto:
        # "4 ,5" -> "4.5" es confiable. Fuera de eso, cualquier espacio
        # interno restante (ej. "4 5", "12 3") es un número partido: no se
        # fusiona jamás.
        s = re.sub(r"\s*\.\s*", ".", s)
        if re.search(r"\s", s):
            return None, False
        try:
            return float(s), True
        except ValueError:
            return None, False
    return None, True


def _normalizar_area(valores, revisar_flags):
    """
    Devuelve (area_trabajo, revisar) con TANTAS celdas como valores traiga el
    modelo (hasta MAX_N_AREAS = 16, el rango válido del spec v2: de 1 a 16 notas
    por alumno). Antes estaba hardcodeado a 2 celdas y TODA planilla salía con
    exactamente 2 notas aunque el modelo leyera 5, 10 o 16.
    - Celdas en blanco -> None (sin marcar revisión).
    - Valor ambiguo (dígitos separados) o no numérico -> None + marcar revisión.
    - Fuera de rango (no entre 0 y 100) -> marcar revisión True y dejar el valor.
    """
    if not valores:
        return [], []
    # Tope de seguridad: nunca más de MAX_N_AREAS celdas. La validación de la
    # planilla (ver _normalizar_planilla) es la que cubre la consistencia con
    # n_area_trabajo; acá sólo se acota por sanidad.
    valores = valores[:MAX_N_AREAS]
    revisar = [
        bool(revisar_flags[i]) if i < len(revisar_flags) else False
        for i in range(len(valores))
    ]
    area = []
    for i, v in enumerate(valores):
        num, confiable = _numero_plausible(v)
        area.append(num)
        if not confiable and v not in (None, ""):
            revisar[i] = True
        if num is not None and not (RANGO_NOTA[0] <= num <= RANGO_NOTA[1]):
            revisar[i] = True
    return area, revisar


def _normalizar_ev(valores, n_esperadas):
    """Ev. Anteriores: lista de n_esperadas celdas (blanco -> None).

    Devuelve (ev, revisar_ev): revisar_ev=True si alguna celda dio un valor
    ambiguo (dígitos separados) o quedó fuera del rango 0-100. Celda en
    blanco/None no marca revisión.
    """
    ev = []
    revisar_ev = False
    for i in range(n_esperadas):
        v = valores[i] if valores and i < len(valores) else None
        num, confiable = _numero_plausible(v)
        ev.append(num)
        if not confiable and v not in (None, ""):
            revisar_ev = True
        if num is not None and not (RANGO_NOTA[0] <= num <= RANGO_NOTA[1]):
            revisar_ev = True
    return ev, revisar_ev


def _normalizar_planilla(datos: dict) -> dict:
    """
    Convierte el JSON crudo del modelo en la forma EXACTA que espera el
    generador de Excel, aplicando todas las reglas de negocio.
    """
    enc_raw = datos.get("encabezado") or {}

    # n_area_trabajo declarado en el encabezado (spec v2): cuántas columnas de
    # área de trabajo dice la planilla que hay. Si viene y no coincide con la
    # cantidad de notas leídas, la planilla completa se marca para revisión
    # manual (NUNCA se descarta). Un valor que no se puede interpretar (o un
    # bool, que no es un conteo) se ignora: el flag queda en manos de la
    # heurística de cantidad de alumnos.
    n_at_raw = enc_raw.get("n_area_trabajo")
    n_area_declarado = None
    if isinstance(n_at_raw, bool):
        n_area_declarado = None
    elif isinstance(n_at_raw, (int, float)):
        n_area_declarado = int(n_at_raw)
    elif isinstance(n_at_raw, str):
        try:
            n_area_declarado = int(float(n_at_raw.strip()))
        except (ValueError, TypeError):
            n_area_declarado = None

    # Periodo parseado de forma tolerante (W2) y SIN corregir silenciosamente:
    # un periodo fuera de rango (ej. 9) ya no se transforma a 1 (W1).
    periodo, periodo_ok = _parsear_periodo(enc_raw.get("periodo"))

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
        # S13: el declarado se CONSERVA en el encabezado (antes se usaba sólo
        # para validar y se descartaba). Así calcular_n_areas de la GUI y del
        # generador pueden preferir el conteo de columnas del encabezado sobre
        # lo observado, y mostrar columnas con celdas vacías como corresponde.
        "n_area_trabajo": n_area_declarado,
    }
    if not grupo_ok:
        encabezado["grupo_erroneo"] = True
    if not periodo_ok:
        encabezado["periodo_erroneo"] = True

    n_ev_anteriores = periodo - 1
    estudiantes = []
    for est in (datos.get("estudiantes") or []):
        if not isinstance(est, dict):
            continue
        # retirado con coerción explícita (W6): bool("false") ya no da True.
        retirado_b, retirado_conf = _coercion_bool(est.get("retirado"))
        if retirado_conf:
            retirado = retirado_b or _fila_retirada(est)
        else:
            # Valor no reconocido: no marcar retirado a ciegas (evita perder un
            # estudiante del cálculo); se usa la marca de fila como respaldo.
            retirado = _fila_retirada(est)
        nombre = str(est.get("nombre") or "").strip()

        if retirado:
            estudiantes.append({
                "no": int(est.get("no") or 0),
                "nombre": nombre,
                "ev_anteriores": [],
                "area_trabajo": None,
                "retirado": True,
                "revisar": [],
                "revisar_ev": False,
            })
            continue

        rev = est.get("revisar")
        # LONGITUD DINÁMICA: los flags del modelo se alinean a la cantidad de
        # áreas que trae el modelo en area_trabajo (hasta MAX_N_AREAS), no a un
        # fijo de 2. Si el modelo no trae "revisar", quedan todos False y sólo
        # actúa la heurística de _normalizar_area.
        at_modelo = est.get("area_trabajo")
        n_at = len(at_modelo) if isinstance(at_modelo, list) else 0
        rev_flags = [False] * n_at
        if isinstance(rev, list):
            for i in range(n_at):
                val = rev[i] if i < len(rev) else None
                r, conf = _coercion_bool(val)
                rev_flags[i] = r if conf else True
        elif isinstance(rev, (int, float)) and not isinstance(rev, bool):
            # Flag escalar (valor único): se aplica a la primera celda, el resto
            # queda sin revisar, con la MISMA longitud dinámica.
            rev_flags = ([rev == 1] + [False] * (n_at - 1)) if n_at > 0 else []

        area, revisar = _normalizar_area(est.get("area_trabajo"), rev_flags)
        ev, revisar_ev = _normalizar_ev(est.get("ev_anteriores"), n_ev_anteriores)

        estudiantes.append({
            "no": int(est.get("no") or 0),
            "nombre": nombre,
            "ev_anteriores": ev,
            "area_trabajo": area,
            "retirado": False,
            "revisar": revisar,
            "revisar_ev": revisar_ev,
        })

    # Validación fuerte de la planilla (spec v2): si algo no cierra, la
    # planilla se MARCA para revisión manual pero sigue en `planillas` — nunca
    # se descarta ni pasa a `fallidas`. La pantalla de revisión avisa y la
    # usuaria verifica contra la planilla física.
    revisar_planilla = False
    # 1) Cantidad de notas: si el encabezado declara n_area_trabajo y algún
    #    estudiante no retirado trae otra cantidad de notas -> revisión manual.
    if n_area_declarado is not None:
        for e in estudiantes:
            if e.get("retirado"):
                continue
            if len(e.get("area_trabajo") or []) != n_area_declarado:
                revisar_planilla = True
                break
    # 2) Cantidad de alumnos no retirados fuera del rango sano -> revisión.
    activos = [e for e in estudiantes if not e.get("retirado")]
    if not (MIN_ESTUDIANTES_SANOS <= len(activos) <= MAX_ESTUDIANTES_SANOS):
        revisar_planilla = True

    return {
        "encabezado": encabezado,
        "estudiantes": estudiantes,
        "revisar_planilla": revisar_planilla,
    }


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
        # Chequear cancelación en cada intento: el callback del worker lanza si
        # la usuaria canceló, abortando la extracción sin esperar reintentos
        # (W3). El valor None es sólo un chequeo, no un mensaje visible.
        if progreso_cb:
            progreso_cb(None)
        if intento > 1 and _notify:
            _notify(f"Reintentando lectura de {pagina_label or 'esta planilla'} (intento {intento})...")
        try:
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
            try:
                datos = _extraer_json(texto)
                return _normalizar_planilla(datos)
            except Exception as e:
                # El modelo devolvió algo que no se pudo interpretar: es un
                # error de PARSEO (W2), no de red/API; no se reintenta.
                # W-C: dejar rastro en el log para diagnóstico (antes el log
                # quedaba mudo ante este tipo de falla).
                app_config.escribir_log(
                    f"Error de parseo en {pagina_label or 'planilla'}: {e!r}"
                )
                raise _PlanillaParseError(
                    "No se pudo interpretar el contenido de la planilla."
                ) from e
        except _PlanillaParseError:
            raise
        except VisionError as e:
            ultimo_error = e
        except Exception as e:
            # Errores de red/API: reintentar con espera (backoff actual).
            ultimo_error = e
        if intento < INTENTOS_MAX:
            time.sleep(ESPERA_BASE_SEG * intento)

    # S12: si la usuaria canceló DURANTE la última llamada (o durante la espera
    # del último reintento), el chequeo del inicio del loop ya no volverá a
    # correr y el loop terminaría elevando un VisionError en vez de la
    # cancelación limpia. Este chequeo final propaga el _Cancelado del worker
    # (que NO es VisionError ni _PlanillaParseError), y la GUI muestra la
    # cancelación en vez de "No se pudo leer...".
    if progreso_cb:
        progreso_cb(None)

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
    Procesa todas las páginas del PDF y devuelve (planillas, fallidas).

    Una página que falla (error de visión tras los reintentos o error de
    parseo) NO tira el lote (S2): se anota en `fallidas` con su número y tipo
    y se continúa con las demás. Cualquier OTRA excepción (incluida la
    cancelación del worker) se propaga sin atrapar: no se traga nada que no
    sea un fallo propio de lectura de una planilla.

    Args:
        paginas: lista de dicts de `pdf_loader.cargar_paginas`.
        api_key: clave de Google AI.
        modelo: modelo de visión (opcional).
        progreso_cb: callback de estado (mensaje).
        por_pagina_cb: callback opcional que recibe (indice, planilla) tras cada
                       página exitosa (útil para cancelar/saltar).

    Returns:
        (planillas, fallidas):
        - planillas: lista de planillas normalizadas (una por página OK).
        - fallidas: lista de dicts {"pagina": int, "total": int,
          "tipo": "vision"|"parseo"} por cada página que no se pudo leer.
    """
    planillas = []
    fallidas = []
    total = len(paginas)
    for i, pag in enumerate(paginas, start=1):
        if progreso_cb:
            # Valor = i/total (fracción 0..1) para actualizar la barra real
            # (S1): cada página procesada mueve el progreso, no sólo la primera.
            progreso_cb(f"Leyendo planilla {i} de {total}...", i / total)
        # S7: la imagen de la página se rasteriza acá (carga perezosa) y se
        # cierra apenas se termina de procesar, liberando el buffer por página.
        imagen = _imagen_de(pag)
        try:
            planilla = extraer_planilla_pagina(
                imagen,
                api_key=api_key,
                modelo=modelo,
                pagina_label=f"planilla {i} de {total}",
                progreso_cb=progreso_cb,
            )
        except (VisionError, _PlanillaParseError) as e:
            # S2: página ilegible -> se registra y se sigue con las demás.
            tipo = "parseo" if isinstance(e, _PlanillaParseError) else "vision"
            app_config.escribir_log(f"No se pudo leer la planilla {i} de {total}: {e!r}")
            fallidas.append({"pagina": i, "total": total, "tipo": tipo})
            continue
        finally:
            # Liberar la imagen de ESTA página (el render lazy crea una nueva
            # por llamada). Tolerante a imagenes legacy sin close().
            cerrar = getattr(imagen, "close", None)
            if callable(cerrar):
                cerrar()
        planillas.append(planilla)
        if por_pagina_cb:
            por_pagina_cb(i, planilla)
    return planillas, fallidas


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
    "periodo": 3,
    "n_area_trabajo": 4
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

NOTA: el ejemplo muestra 2 áreas, pero el número REAL de valores de
"area_trabajo" (y de "revisar") varía con la planilla: pueden ser de 1 a 16.
Lee SIEMPRE la cantidad real de columnas, nunca fijes el tamaño.

REGLAS IMPORTANTES:

1. ENCABEZADO:
   - "periodo" es un número entero del 1 al 4 (el periodo académico de la planilla).
   - "grupo" es un número de 4 dígitos como "0302" (grado y grupo). Nunca inventes
     un grupo que no esté escrito.
   - "año_lectivo" suele ser un año (ej. "2026").
   - El texto fijo "LISTA AUXILIAR DE CLASE" es el membrete impreso de la
     planilla, NO un dato: ignoralo (no lo extraigas como institución, sede ni
     en ningún otro campo).
   - Los escaneos (por ejemplo de CamScanner) suelen cortar o inclinar el
     margen izquierdo: etiquetas como "ASIGNATURA", "DOCENTE" o "AÑO LECTIVO"
     pueden aparecer truncadas ("GNATURA", "CENTE", "O LECTIVO"). Inferí el
     campo por su POSICIÓN en la tabla del encabezado (la fila/columna donde
     siempre aparece), no solo por el texto literal. No dejes el campo vacío ni
     falles por un rótulo parcialmente cortado: el valor de la celda a la
     derecha del rótulo suele estar legible.
   - "n_area_trabajo" es la cantidad de columnas "ÁREA DE TRABAJO"/"AREA DE
     TRABAJO" impresas en el ENCABEZADO de la tabla (de 1 a 16). Contá las
     columnas del encabezado, NO las celdas llenas: si la fila de títulos dice
     ÁREA DE TRABAJO 1, 2, 3 y 4, devolvé 4 aunque en varios alumnos la última
     celda esté en blanco (las planillas reales suelen dejar celdas vacías).

2. TABLA DE ESTUDIANTES:
   - "ev_anteriores" son las notas definitivas de periodos anteriores (una por cada
     columna que aparezca; si hay 2 columnas van 2 valores, si hay 3 van 3).
   - "area_trabajo" son las notas manuscritas del periodo actual, UNA por cada
     columna de área que aparezca en la planilla (de 1 a 16 valores, NUNCA fijes
     la cantidad: leé exactamente cuántas columnas hay). Si una celda está en
     blanco, pon null en esa posición. ATENCIÓN: cada alumno debe traer
     EXACTAMENTE n_area_trabajo valores: si el encabezado tiene 4 columnas, el
     arreglo de cada alumno tiene 4 posiciones (las vacías van como null), aunque
     ese alumno no tenga todas las notas.
   - Nunca asumas que dos planillas del mismo curso y periodo tienen la misma
     cantidad de columnas de Área de Trabajo. Cada planilla se cuenta de forma
     independiente, aunque sea del mismo curso y periodo que otra que ya
     procesaste.
   - IGNORA por completo las columnas tituladas "Min" y "Fls.": no las leas ni
     las guardes. Aunque traigan números reales (minutos de tardanza y cantidad
     de faltas), esos datos no son parte del sistema.
   - Si una fila está marcada con asteriscos (****) o dice "retirado", pon
     "retirado": true, "area_trabajo": null y "ev_anteriores": [].
   - La fila retirado suele venir con "********" en TODAS sus celdas (leyenda,
     Ev. Anteriores, Min, Fls., cada columna de Área de Trabajo y Fallas). Son
     una sola marca, no datos: poné "retirado": true y no intentes leer los
     asteriscos como números ni como notas ("area_trabajo" va en null).
   - Si una celda está en blanco o tachada sin valor claro, devolvé null (nunca 0).
   - Si una nota fue corregida (tachada y reescrita), devolvé el valor VIGENTE
     (el reescrito), no el tachado.
   - "revisar" es un arreglo de booleanos, uno por cada celda de "area_trabajo"
     (la MISMA cantidad de valores).
     Si NO estás completamente seguro de un dígito (letra ambigua, tachón difícil,
     mancha), poné true en esa posición de "revisar" (y null en el valor si no podés
     leerlo). Si estás seguro, pon false. Nunca inventes un valor dudoso para evitar
     la revisión.
   - Las notas suelen ser números enteros o decimales (ej. 45, 40, 4.5, 50). Devolvé
     el número tal cual aparece en la planilla, sin cambiar su escala.
   - Si un número está escrito con los dígitos muy separados (ej. "4 5") NO lo
     escribas como uno solo: devolvé null y marcá revisar en true.
   - Ignorá por completo cualquier texto manuscrito que aparezca DEBAJO de la
     última fila de la tabla de estudiantes o debajo de la línea "Firma Docente"
     (fechas de clase, temas, tareas, firmas). Esos apuntes no son notas de
     ningún estudiante.

3. Devolvé SOLO el JSON. No agregues explicaciones.
""".strip()
