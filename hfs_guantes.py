"""
=============================================================================
PROGRAMACIÓN DE UN HYBRID FLOW SHOP (HFS) CON LOT STREAMING DE SUBLOTES
CONSISTENTES Y ALGORITMO GENÉTICO — PRODUCCIÓN DE GUANTES INDUSTRIALES
=============================================================================

DATOS
  Todos los datos del problema se leen del libro de Excel `BD_HFS.xlsx`; el
  código no contiene trabajos, cantidades, fechas ni tiempos.
    • "Tiempo de procesamiento": Codigo producto | Maquina | Etapa |
                                 min x docena | h x docena
      (tiempo de cada producto en cada máquina de cada etapa)
    • "Trabajos": Orden de trabajo | Codigo Producto | Cantidad (pares) |
                  Cantidad (docenas) | Fecha de liberación | Fechas de entrega |
                  Horas disponibles (fecha de entrega en horas laborables)
    • "Disponibilidad": Fecha | Dia semana | Horas disponibles | Horas acumuladas
      (calendario laboral: convierte fechas ↔ horas laborables)
  Las etapas se ordenan según la numeración de sus máquinas (M1–M4 Troquelado,
  M5–M24 Ensamble, …). Si un producto no tiene tiempo en una máquina, esa
  máquina no puede procesarlo (capacidad de máquina).

LOT STREAMING — SUBLOTES CONSISTENTES
  Se sigue la estrategia 2 / escenario S2 de Gürsoy Yılmaz y Yılmaz (2022),
  "Lot streaming in hybrid flowshop scheduling problem by considering equal
  and consistent sublots…", Computers & Industrial Engineering 173, 108745:
    • Cada lote (orden de trabajo) se divide en B_j sublotes.
    • Los tamaños de los sublotes de un lote pueden ser distintos entre sí,
      pero son LOS MISMOS EN TODAS LAS ETAPAS (sublotes consistentes).
    • Un sublote pasa a la etapa s+1 apenas termina en la etapa s, aunque el
      resto del lote siga en proceso (traslape entre etapas).
    • La asignación de las posiciones de los sublotes en el cromosoma es
      aleatoria (random-based sublot assignment).

  Cromosoma:
    nivel 1 'tamanos'  : proporciones de los sublotes de cada trabajo
                         (una sola lista por trabajo → igual en todas las etapas)
    nivel 2 'secuencia': posiciones de los sublotes (permutación con repetición:
                         la n-ésima aparición del trabajo j es su sublote n)
    'k'                : número de sublotes B_j de cada trabajo
  Decodificación: la etapa 1 procesa los sublotes en el orden del cromosoma;
  en las etapas siguientes se atienden en orden de llegada (desempate por la
  posición en el cromosoma). Cada sublote se asigna a la máquina capaz que lo
  TERMINA más temprano (las máquinas no son idénticas).

  Función objetivo: minimizar la tardanza total ΣTj (desempate: Cmax).
  Las horas son horas laborables según la hoja "Disponibilidad".

Validación con LEKIN:
  LEKIN no hace lot streaming y admite máximo 10 trabajos. Por eso este
  programa incluye un simulador de reglas de despacho (EDD, SPT, LPT, FIFO,
  LIFO, MS, CR) sin lot streaming, y genera la hoja `LEKIN_datos` en minutos
  enteros para cargar en LEKIN. Si el simulador está bien, cada regla debe
  dar el mismo ΣTj en Python y en LEKIN.
"""

import copy
import datetime as dt
import itertools
import math
import os
import random
import re
import sys
import unicodedata
from typing import Dict, List, Optional, Tuple

import pandas as pd

# NO se llama a matplotlib.use(): Colab y Jupyter configuran su propio backend.
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker as ticker

RUTA_EXCEL = "BD_HFS.xlsx"
RUTA_RESULTADOS = "resultados_programacion.xlsx"
RUTA_GANTT = "gantt.png"

EPS = 1e-6

# ---------------------------------------------------------------------------
# 1.  PARÁMETROS DEL MÉTODO (no son datos del problema)
# ---------------------------------------------------------------------------

MAX_SUBLOTES = 4              # máximo de sublotes por lote (B_j)
MIN_PARES_SUBLOTE = 60        # tamaño mínimo de un sublote [pares] (5 docenas)
ENUMERACION_MAX_N = 8         # óptimo exacto sin LS solo si n ≤ este valor
MAX_TRABAJOS_LEKIN = 10

GA_PARAMS = {
    'poblacion'    : 100,
    'generaciones' : 300,
    'prob_cruce'   : 0.90,
    'prob_mutacion': 0.30,
    'prob_mut_seq' : 0.50,    # intercambio de dos posiciones de sublotes
    'prob_mut_tam' : 0.20,    # por trabajo: perturba los tamaños de sus sublotes
    'prob_mut_k'   : 0.15,    # por trabajo: suma o quita un sublote
    'tam_torneo'   : 3,
    'elitismo'     : True,
    'semilla'      : 42,
}

# ---------------------------------------------------------------------------
# 2.  DATOS DEL PROBLEMA (se llenan desde el Excel con cargar_datos)
# ---------------------------------------------------------------------------

TRABAJOS: List[str] = []                  # órdenes de trabajo, en el orden del Excel
PRODUCTO: Dict[str, str] = {}             # trabajo → código de producto
DEMANDA: Dict[str, int] = {}              # [pares]
DUE_DATE: Dict[str, float] = {}           # [horas laborables]
RELEASE: Dict[str, float] = {}            # [horas laborables]
FECHA_ENTREGA: Dict[str, Optional[dt.date]] = {}
ETAPAS: List[int] = []                    # 1..S en orden de flujo
NOMBRE_ETAPA: Dict[int, str] = {}
MAQUINAS: Dict[int, List[str]] = {}       # etapa → ['M1', 'M2', …]
MAQUINAS_POR_ETAPA: Dict[int, int] = {}
TIEMPO_MIN_DOCENA: Dict[str, Dict[str, float]] = {}   # producto → máquina → min/docena
CALENDARIO: List[Tuple[dt.date, float, float]] = []   # (fecha, horas, horas acumuladas)

# Derivado: RAPIDEZ[j][s] = [(índice de máquina capaz, horas por par), …]
RAPIDEZ: Dict[str, Dict[int, List[Tuple[int, float]]]] = {}

# ---------------------------------------------------------------------------
# 3.  LECTURA DEL EXCEL
# ---------------------------------------------------------------------------

def _norm(t) -> str:
    """Texto en minúsculas, sin tildes ni espacios repetidos (para comparar nombres)."""
    t = unicodedata.normalize("NFKD", str(t)).encode("ascii", "ignore").decode()
    return " ".join(t.lower().replace("_", " ").split())


def _vacio(v) -> bool:
    return v is None or (not isinstance(v, str) and pd.isna(v)) or str(v).strip() == ""


def _buscar_hoja(xls: pd.ExcelFile, clave: str, obligatoria: bool = True) -> Optional[str]:
    for h in xls.sheet_names:
        if clave in _norm(h):
            return h
    if obligatoria:
        raise ValueError(f"El Excel no tiene una hoja '{clave}…'. Hojas: {xls.sheet_names}")
    return None


def _columna(df: pd.DataFrame, hoja: str, *candidatos: str,
             obligatoria: bool = True) -> Optional[str]:
    """Nombre real de la columna que coincide con alguno de los candidatos."""
    cols = {_norm(c): c for c in df.columns}
    for c in candidatos:
        if _norm(c) in cols:
            return cols[_norm(c)]
    for c in candidatos:
        for k, real in cols.items():
            if _norm(c) in k:
                return real
    if obligatoria:
        raise ValueError(f"Hoja '{hoja}': falta la columna '{candidatos[0]}'. "
                         f"Columnas encontradas: {list(df.columns)}")
    return None


def _numero(v, hoja: str, campo: str, minimo: float = 0.0, estricto: bool = False) -> float:
    try:
        x = float(v)
    except (TypeError, ValueError):
        raise ValueError(f"Hoja '{hoja}': '{campo}' debe ser numérico (valor leído: {v!r})")
    if math.isnan(x) or x < minimo or (estricto and x <= minimo):
        signo = ">" if estricto else "≥"
        raise ValueError(f"Hoja '{hoja}': '{campo}' debe ser {signo} {minimo} (valor leído: {v!r})")
    return x


def _fecha(v, hoja: str, campo: str) -> dt.date:
    try:
        f = pd.Timestamp(v)
    except (TypeError, ValueError):
        f = pd.NaT
    if pd.isna(f):
        raise ValueError(f"Hoja '{hoja}': '{campo}' debe ser una fecha (valor leído: {v!r})")
    return f.date()


def _num_maquina(m: str) -> Tuple[int, str]:
    d = re.sub(r"\D", "", m)
    return (int(d) if d else 0, m)


def _leer_calendario(xls: pd.ExcelFile):
    CALENDARIO.clear()
    hoja = _buscar_hoja(xls, "disponib", obligatoria=False)
    if hoja is None:
        return
    df = xls.parse(hoja)
    c_fecha = _columna(df, hoja, "Fecha")
    c_horas = _columna(df, hoja, "Horas disponibles")
    # Respaldo si 'Horas disponibles' es una fórmula sin valor calculado:
    # tabla auxiliar Dia ("1 Lunes", …) → Horas laborables
    c_dia = _columna(df, hoja, "Dia", obligatoria=False)
    c_hlab = _columna(df, hoja, "Horas laborables", obligatoria=False)
    horas_dia = {}
    if c_dia is not None and c_hlab is not None:
        for d, hl in zip(df[c_dia], df[c_hlab]):
            num = re.match(r"\s*(\d)", str(d))
            if num and not _vacio(hl):
                horas_dia[int(num.group(1))] = float(hl)
    filas = []
    for _, f in df.iterrows():
        if _vacio(f[c_fecha]):
            continue
        fecha = _fecha(f[c_fecha], hoja, "Fecha")
        if _vacio(f[c_horas]) and fecha.isoweekday() in horas_dia:
            horas = horas_dia[fecha.isoweekday()]
        else:
            horas = _numero(f[c_horas], hoja, f"Horas disponibles ({fecha})")
        filas.append((fecha, horas))
    filas.sort()
    acum = 0.0
    for fecha, horas in filas:
        acum += horas
        CALENDARIO.append((fecha, horas, acum))


def fecha_a_horas(fecha: dt.date, fin_del_dia: bool = True) -> float:
    """Horas laborables acumuladas al final (o al inicio) de 'fecha'."""
    if not CALENDARIO:
        raise ValueError("No hay hoja 'Disponibilidad' para convertir fechas a horas")
    if fecha < CALENDARIO[0][0]:
        return 0.0
    for f, horas, acum in CALENDARIO:
        if f >= fecha:
            return acum if (fin_del_dia and f == fecha) else acum - horas
    raise ValueError(f"La fecha {fecha} es posterior al calendario de la hoja 'Disponibilidad'")


def horas_a_fecha(h: float) -> Optional[dt.date]:
    """Día laborable en que se alcanza la hora laborable 'h' (None si sale del calendario)."""
    for f, horas, acum in CALENDARIO:
        if horas > 0 and acum >= h - EPS:
            return f
    return None


def cargar_datos(ruta: str = RUTA_EXCEL):
    """Lee el Excel y llena las estructuras globales del modelo."""
    xls = pd.ExcelFile(ruta)
    _leer_calendario(xls)

    # ── Tiempos de procesamiento (definen etapas y máquinas) ──
    hoja = _buscar_hoja(xls, "tiempo")
    df = xls.parse(hoja)
    c_prod = _columna(df, hoja, "Codigo producto", "Producto")
    c_maq = _columna(df, hoja, "Maquina")
    c_etapa = _columna(df, hoja, "Etapa")
    c_min = _columna(df, hoja, "min x docena", obligatoria=False)
    c_h = _columna(df, hoja, "h x docena", obligatoria=False)
    if c_min is None and c_h is None:
        raise ValueError(f"Hoja '{hoja}': falta la columna 'min x docena' (o 'h x docena')")

    TIEMPO_MIN_DOCENA.clear()
    etapa_de_maquina: Dict[str, str] = {}
    nombre_etapa: Dict[str, str] = {}             # nombre normalizado → nombre original
    for _, f in df.iterrows():
        if _vacio(f[c_prod]) and _vacio(f[c_maq]):
            continue
        prod, maq, etapa = str(f[c_prod]).strip(), str(f[c_maq]).strip(), str(f[c_etapa]).strip()
        if "" in (prod, maq, etapa):
            raise ValueError(f"Hoja '{hoja}': fila incompleta ({prod!r}, {maq!r}, {etapa!r})")
        nombre_etapa.setdefault(_norm(etapa), etapa)
        if etapa_de_maquina.setdefault(maq, _norm(etapa)) != _norm(etapa):
            raise ValueError(f"Hoja '{hoja}': la máquina {maq} aparece en dos etapas")
        if c_min is not None and not _vacio(f[c_min]):
            t = _numero(f[c_min], hoja, f"min x docena ({prod}, {maq})", estricto=True)
        elif c_h is not None and not _vacio(f[c_h]):
            t = 60.0 * _numero(f[c_h], hoja, f"h x docena ({prod}, {maq})", estricto=True)
        else:
            continue          # sin tiempo → la máquina no procesa ese producto
        if maq in TIEMPO_MIN_DOCENA.setdefault(prod, {}):
            raise ValueError(f"Hoja '{hoja}': el tiempo de {prod} en {maq} está repetido")
        TIEMPO_MIN_DOCENA[prod][maq] = t

    maq_por_etapa: Dict[str, List[str]] = {}
    for maq, e in etapa_de_maquina.items():
        maq_por_etapa.setdefault(e, []).append(maq)
    for lista in maq_por_etapa.values():
        lista.sort(key=_num_maquina)
    # Orden de flujo: según la numeración de las máquinas de cada etapa
    orden = sorted(maq_por_etapa, key=lambda e: _num_maquina(maq_por_etapa[e][0]))
    ETAPAS[:] = list(range(1, len(orden) + 1))
    NOMBRE_ETAPA.clear(); MAQUINAS.clear(); MAQUINAS_POR_ETAPA.clear()
    for s, e in zip(ETAPAS, orden):
        NOMBRE_ETAPA[s] = nombre_etapa[e]
        MAQUINAS[s] = maq_por_etapa[e]
        MAQUINAS_POR_ETAPA[s] = len(maq_por_etapa[e])

    # ── Trabajos ──
    hoja = _buscar_hoja(xls, "trabajo")
    df = xls.parse(hoja)
    c_ot = _columna(df, hoja, "Orden de trabajo", "Trabajo")
    c_prod = _columna(df, hoja, "Codigo Producto", "Producto")
    c_pares = _columna(df, hoja, "Cantidad (pares)", "pares", obligatoria=False)
    c_doc = _columna(df, hoja, "Cantidad (docenas)", "docenas", obligatoria=False)
    c_lib = _columna(df, hoja, "Fecha de liberacion", "liberacion", obligatoria=False)
    c_ent = _columna(df, hoja, "Fechas de entrega", "Fecha de entrega", obligatoria=False)
    c_hdisp = _columna(df, hoja, "Horas disponibles", obligatoria=False)
    if c_pares is None and c_doc is None:
        raise ValueError(f"Hoja '{hoja}': falta la columna 'Cantidad (pares)'")

    TRABAJOS.clear(); PRODUCTO.clear(); DEMANDA.clear(); DUE_DATE.clear()
    RELEASE.clear(); FECHA_ENTREGA.clear()
    for _, f in df.iterrows():
        if _vacio(f[c_ot]):
            continue
        j = str(f[c_ot]).strip()
        if j in DEMANDA:
            raise ValueError(f"Hoja '{hoja}': la orden de trabajo {j} está repetida")
        prod = str(f[c_prod]).strip()
        if prod not in TIEMPO_MIN_DOCENA:
            raise ValueError(f"Hoja '{hoja}': el producto {prod} de {j} no tiene tiempos de "
                             "procesamiento")
        # Cantidad en pares (si solo hay docenas: pares = docenas × 12)
        if c_pares is not None and not _vacio(f[c_pares]):
            pares = _numero(f[c_pares], hoja, f"Cantidad (pares) ({j})", estricto=True)
        else:
            pares = 12 * _numero(f[c_doc], hoja, f"Cantidad (docenas) ({j})", estricto=True)
        if abs(pares - round(pares)) > 1e-6:
            raise ValueError(f"Hoja '{hoja}': la cantidad de {j} no es un número entero de pares")

        # Fecha de entrega en horas laborables: 'Horas disponibles' o, si la
        # celda está vacía (fórmula sin calcular), desde el calendario
        fecha_ent = None
        if c_ent is not None and not _vacio(f[c_ent]):
            fecha_ent = _fecha(f[c_ent], hoja, f"Fechas de entrega ({j})")
        if c_hdisp is not None and not _vacio(f[c_hdisp]):
            due = _numero(f[c_hdisp], hoja, f"Horas disponibles ({j})")
        elif fecha_ent is not None:
            due = fecha_a_horas(fecha_ent)
        else:
            raise ValueError(f"Hoja '{hoja}': {j} no tiene 'Horas disponibles' ni fecha de entrega")

        lib = 0.0
        if c_lib is not None and not _vacio(f[c_lib]) and CALENDARIO:
            lib = fecha_a_horas(_fecha(f[c_lib], hoja, f"Fecha de liberación ({j})"),
                                fin_del_dia=False)

        TRABAJOS.append(j)
        PRODUCTO[j] = prod
        DEMANDA[j] = int(round(pares))
        DUE_DATE[j] = due
        RELEASE[j] = lib
        FECHA_ENTREGA[j] = fecha_ent if fecha_ent is not None else horas_a_fecha(due)
    if len(TRABAJOS) < 2:
        raise ValueError(f"Hoja '{hoja}': se necesitan al menos 2 trabajos")

    # ── Rapidez por máquina capaz [h/par] ──
    RAPIDEZ.clear()
    for j in TRABAJOS:
        RAPIDEZ[j] = {}
        for s in ETAPAS:
            tiempos = TIEMPO_MIN_DOCENA[PRODUCTO[j]]
            RAPIDEZ[j][s] = [(m, tiempos[maq] / 12.0 / 60.0)
                             for m, maq in enumerate(MAQUINAS[s]) if maq in tiempos]
            if not RAPIDEZ[j][s]:
                raise ValueError(f"Ninguna máquina de {NOMBRE_ETAPA[s]} tiene tiempo para el "
                                 f"producto {PRODUCTO[j]} ({j})")


def obtener_ruta_excel(ruta: str = RUTA_EXCEL) -> str:
    """Devuelve la ruta del Excel; en Colab, si no existe, pide subirlo."""
    if os.path.exists(ruta):
        return ruta
    try:
        from google.colab import files  # type: ignore
    except ImportError:
        raise FileNotFoundError(
            f"No se encontró '{ruta}'. Póngalo en la misma carpeta del script "
            f"o cambie la variable RUTA_EXCEL.")
    print(f"No se encontró '{ruta}'. Seleccione el archivo Excel para subirlo:")
    subidos = files.upload()
    if not subidos:
        raise FileNotFoundError("No se subió ningún archivo.")
    return next(iter(subidos))

# ---------------------------------------------------------------------------
# 4.  TIEMPOS
# ---------------------------------------------------------------------------

def docenas(pares: float) -> float:
    """Convierte pares de guantes a docenas (1 docena = 12 pares)."""
    return pares / 12.0


def tiempo_horas(trabajo: str, etapa: int, m: int, pares: float) -> float:
    """Tiempo [h] = min/docena del producto en la máquina m × docenas / 60."""
    maq = MAQUINAS[etapa][m]
    return TIEMPO_MIN_DOCENA[PRODUCTO[trabajo]][maq] * docenas(pares) / 60.0


def tiempos_lote_completo() -> Dict[str, Dict[int, List[Tuple[int, float]]]]:
    """p[j][s] = [(máquina, horas del lote completo en esa máquina), …]."""
    return {j: {s: [(m, h * DEMANDA[j]) for m, h in RAPIDEZ[j][s]] for s in ETAPAS}
            for j in TRABAJOS}


def tiempo_promedio(p_js: List[Tuple[int, float]]) -> float:
    return sum(t for _, t in p_js) / len(p_js)

# ---------------------------------------------------------------------------
# 5.  REPRESENTACIÓN DEL CROMOSOMA (sublotes consistentes)
# ---------------------------------------------------------------------------

def max_sublotes(trabajo: str) -> int:
    """B_j máximo: limitado por MAX_SUBLOTES y por el tamaño mínimo de sublote."""
    return max(1, min(MAX_SUBLOTES, DEMANDA[trabajo] // MIN_PARES_SUBLOTE))


def _dirichlet(k: int) -> List[float]:
    """k proporciones aleatorias positivas que suman 1."""
    muestras = [random.expovariate(1.0) + 1e-9 for _ in range(k)]
    total = sum(muestras)
    return [x / total for x in muestras]


def tamanos_pares(trabajo: str, props: List[float]) -> List[int]:
    """
    Convierte las proporciones del cromosoma en tamaños ENTEROS de sublote
    [pares]: cada sublote tiene al menos MIN_PARES_SUBLOTE y la suma es
    exactamente la demanda (método del mayor residuo). Estos tamaños se usan
    en TODAS las etapas (sublotes consistentes).
    """
    total = DEMANDA[trabajo]
    k = len(props)
    minimo = min(MIN_PARES_SUBLOTE, total // k)
    libre = total - k * minimo
    suma = sum(props)
    ideal = [libre * p / suma for p in props]
    base = [int(x) for x in ideal]
    por_repartir = libre - sum(base)
    for i in sorted(range(k), key=lambda i: ideal[i] - base[i], reverse=True)[:por_repartir]:
        base[i] += 1
    return [minimo + b for b in base]


def crear_individuo(orden: Optional[List[str]] = None, k_fijo: Optional[int] = None,
                    iguales: bool = False) -> dict:
    """
    orden   : si se da, los sublotes de cada trabajo van seguidos en ese orden;
              si no, las posiciones de los sublotes son aleatorias (escenario S2)
    k_fijo  : número de sublotes para todos los trabajos (None → aleatorio)
    iguales : sublotes de igual tamaño (si no, tamaños aleatorios)
    """
    k = {j: min(k_fijo, max_sublotes(j)) if k_fijo else random.randint(1, max_sublotes(j))
         for j in TRABAJOS}
    tamanos = {j: [1.0 / k[j]] * k[j] if iguales else _dirichlet(k[j]) for j in TRABAJOS}
    if orden is not None:
        secuencia = [j for j in orden for _ in range(k[j])]
    else:
        secuencia = [j for j in TRABAJOS for _ in range(k[j])]
        random.shuffle(secuencia)
    return {'secuencia': secuencia, 'k': k, 'tamanos': tamanos}

# ---------------------------------------------------------------------------
# 6.  SIMULADOR HFS CON SUBLOTES CONSISTENTES (decodificador del AG)
# ---------------------------------------------------------------------------

def simular_hfs(individuo: dict) -> Tuple[Dict[str, float], Dict[int, list]]:
    """
    Simula el HFS con sublotes consistentes.

    - Etapa 1: los sublotes se procesan en el orden del cromosoma.
    - Etapas siguientes: en orden de llegada (fin en la etapa anterior);
      empates → posición en el cromosoma.
    - Un sublote empieza en la etapa s cuando ese MISMO sublote terminó en la
      etapa s−1 (el tamaño del sublote no cambia entre etapas).
    - Cada sublote va a la máquina capaz que lo TERMINA más temprano.

    Retorna
    -------
    C     : {trabajo: terminación de su último sublote en la última etapa}
    gantt : {etapa: [(trabajo, sublote, maquina_idx, inicio, fin, pares), ...]}
    """
    tam = {j: tamanos_pares(j, individuo['tamanos'][j]) for j in TRABAJOS}
    orden, n = [], {j: 0 for j in TRABAJOS}
    for pos, j in enumerate(individuo['secuencia']):
        n[j] += 1
        orden.append((j, n[j], pos))
    listo = {(j, b): RELEASE[j] for j, b, _ in orden}
    gantt: Dict[int, list] = {s: [] for s in ETAPAS}

    for s in ETAPAS:
        libre = [0.0] * MAQUINAS_POR_ETAPA[s]
        if s != ETAPAS[0]:
            orden.sort(key=lambda x: (listo[(x[0], x[1])], x[2]))
        for j, b, _ in orden:
            q = tam[j][b - 1]
            t0 = listo[(j, b)]
            mejor_m, mejor_ini, mejor_fin = -1, 0.0, math.inf
            for m, h in RAPIDEZ[j][s]:
                ini = t0 if t0 > libre[m] else libre[m]
                fin = ini + h * q
                if fin < mejor_fin - EPS:
                    mejor_m, mejor_ini, mejor_fin = m, ini, fin
            libre[mejor_m] = mejor_fin
            listo[(j, b)] = mejor_fin
            gantt[s].append((j, b, mejor_m, mejor_ini, mejor_fin, q))

    C = {j: max(listo[(j, b)] for b in range(1, len(tam[j]) + 1)) for j in TRABAJOS}
    return C, gantt


def simular_sin_ls(orden: List[str]) -> Tuple[Dict[str, float], Dict[int, list]]:
    """HFS de permutación sin lot streaming (un solo sublote por lote)."""
    return simular_hfs(crear_individuo(orden, k_fijo=1))


def verificar_programa(gantt: Dict[int, list]) -> List[str]:
    """
    Revisa un programa de forma independiente del simulador. Comprueba:
      1. Ninguna máquina procesa dos sublotes a la vez.
      2. En cada etapa se procesa exactamente la demanda de cada trabajo.
      3. La máquina puede procesar el producto y la duración es
         min/docena × docenas / 60.
      4. Ningún sublote empieza antes de la liberación del trabajo.
      5. Sublotes consistentes: cada sublote tiene el mismo tamaño en todas
         las etapas y no empieza en la etapa s antes de terminar en la s−1.
    Retorna la lista de errores (vacía = programa factible).
    """
    errores = []
    fin_anterior: Dict[Tuple[str, int], float] = {}
    tam_anterior: Dict[Tuple[str, int], float] = {}
    for s in ETAPAS:
        por_maquina: Dict[int, list] = {}
        volumen = {j: 0.0 for j in TRABAJOS}
        fin_actual, tam_actual = {}, {}
        for (j, b, m, ini, fin, q) in gantt[s]:
            etq = f"{j}-S{b} en {NOMBRE_ETAPA[s]}"
            por_maquina.setdefault(m, []).append((ini, fin, j, b))
            volumen[j] += q
            if MAQUINAS[s][m] not in TIEMPO_MIN_DOCENA[PRODUCTO[j]]:
                errores.append(f"{etq}: la máquina {MAQUINAS[s][m]} no procesa {PRODUCTO[j]}")
            elif abs((fin - ini) - tiempo_horas(j, s, m, q)) > 1e-6:
                errores.append(f"Duración incorrecta: {etq}")
            if ini < RELEASE[j] - EPS:
                errores.append(f"{etq} inicia antes de su liberación")
            if s != ETAPAS[0]:
                if (j, b) not in fin_anterior:
                    errores.append(f"{etq} no pasó por la etapa anterior")
                else:
                    if abs(tam_anterior[(j, b)] - q) > 1e-6:
                        errores.append(f"{etq}: el tamaño cambia entre etapas (no consistente)")
                    if ini < fin_anterior[(j, b)] - EPS:
                        errores.append(f"{etq} inicia antes de terminar en la etapa anterior")
            fin_actual[(j, b)], tam_actual[(j, b)] = fin, q
        for m, ops in por_maquina.items():
            ops.sort()
            for a, c in zip(ops, ops[1:]):
                if c[0] < a[1] - EPS:
                    errores.append(f"Traslape en {NOMBRE_ETAPA[s]} {MAQUINAS[s][m]}: "
                                   f"{a[2]}-S{a[3]} y {c[2]}-S{c[3]}")
        for j in TRABAJOS:
            if abs(volumen[j] - DEMANDA[j]) > 1e-6:
                errores.append(f"{j}: el volumen en {NOMBRE_ETAPA[s]} no suma la demanda")
        fin_anterior, tam_anterior = fin_actual, tam_actual
    return errores

# ---------------------------------------------------------------------------
# 7.  SIMULADOR DE REGLAS DE DESPACHO (sin lot streaming, como LEKIN)
# ---------------------------------------------------------------------------

REGLAS = ["EDD", "SPT", "LPT", "FIFO", "LIFO", "MS", "CR"]


def simular_despacho(regla: str,
                     trabajos: Optional[List[str]] = None,
                     p_maq: Optional[Dict[str, Dict[int, List[Tuple[int, float]]]]] = None,
                     due: Optional[Dict[str, float]] = None,
                     rel: Optional[Dict[str, float]] = None
                     ) -> Tuple[Dict[str, float], Dict[int, list]]:
    """
    Programación no-delay por reglas de despacho, lote completo (sin LS).

    Cada vez que una máquina queda libre y hay trabajos esperando en su
    etapa, se elige el trabajo de mayor prioridad según la regla y se asigna
    a la máquina capaz que lo termina primero. Prioridades (con el tiempo
    promedio de la operación en las máquinas capaces de la etapa):
      EDD  : menor fecha de entrega
      SPT  : menor tiempo de la operación en la etapa
      LPT  : mayor tiempo de la operación en la etapa
      FIFO : el que llegó primero a la cola de la etapa
      LIFO : el que llegó último a la cola de la etapa
      MS   : menor holgura  (d_j − t − trabajo restante)
      CR   : menor razón crítica (d_j − t) / trabajo restante
    Empates → orden de la hoja Trabajos (LIFO: el último).

    p_maq[j][s] = [(máquina, horas del lote completo), …]; los argumentos
    opcionales permiten usar otros datos (p. ej. los minutos enteros de LEKIN).
    """
    trabajos = list(trabajos or TRABAJOS)
    p_maq = p_maq or tiempos_lote_completo()
    due = due or DUE_DATE
    rel = rel or RELEASE
    p = {j: {s: tiempo_promedio(p_maq[j][s]) for s in ETAPAS} for j in trabajos}

    orden = {j: i for i, j in enumerate(trabajos)}
    maq_libre = {s: [0.0] * MAQUINAS_POR_ETAPA[s] for s in ETAPAS}
    idx_etapa = {j: 0 for j in trabajos}            # índice de la próxima etapa
    listo = {j: rel[j] for j in trabajos}           # llegada a la cola actual
    C: Dict[str, float] = {}
    gantt: Dict[int, list] = {s: [] for s in ETAPAS}

    def restante(j, s):
        return sum(p[j][e] for e in ETAPAS[ETAPAS.index(s):])

    def prioridad(j, s, t):
        if regla == "EDD":
            return (due[j], orden[j])
        if regla == "SPT":
            return (p[j][s], orden[j])
        if regla == "LPT":
            return (-p[j][s], orden[j])
        if regla == "FIFO":
            return (listo[j], orden[j])
        if regla == "LIFO":
            return (-listo[j], -orden[j])
        if regla == "MS":
            return (due[j] - t - restante(j, s), orden[j])
        if regla == "CR":
            r = restante(j, s)
            return ((due[j] - t) / r if r > 0 else -math.inf, orden[j])
        raise ValueError(f"Regla desconocida: {regla}")

    pendientes = set(trabajos)
    while pendientes:
        # Instante más temprano en que alguna operación puede iniciar
        t_min, s_min = math.inf, None
        for j in pendientes:
            s = ETAPAS[idx_etapa[j]]
            t = max(listo[j], min(maq_libre[s][m] for m, _ in p_maq[j][s]))
            if t < t_min - EPS or (abs(t - t_min) <= EPS and s < s_min):
                t_min, s_min = t, s
        s = s_min
        cola = [j for j in pendientes
                if ETAPAS[idx_etapa[j]] == s and listo[j] <= t_min + EPS
                and min(maq_libre[s][m] for m, _ in p_maq[j][s]) <= t_min + EPS]
        j = min(cola, key=lambda x: prioridad(x, s, t_min))
        # Máquina capaz que TERMINA primero la operación (con velocidades
        # distintas puede convenir esperar un poco a una máquina rápida)
        m, dur = min(p_maq[j][s], key=lambda x: (max(t_min, maq_libre[s][x[0]]) + x[1], x[0]))
        ini = max(t_min, maq_libre[s][m])
        fin = ini + dur
        maq_libre[s][m] = fin
        gantt[s].append((j, 1, m, ini, fin, DEMANDA[j]))

        idx_etapa[j] += 1
        listo[j] = fin
        if idx_etapa[j] == len(ETAPAS):
            C[j] = fin
            pendientes.remove(j)

    return C, gantt

# ---------------------------------------------------------------------------
# 8.  FUNCIÓN OBJETIVO E INDICADORES
# ---------------------------------------------------------------------------

def calcular_indicadores(C: Dict[str, float],
                         due: Optional[Dict[str, float]] = None,
                         rel: Optional[Dict[str, float]] = None) -> Dict:
    """Indicadores estándar de scheduling para los trabajos de C."""
    due = due or DUE_DATE
    rel = rel or RELEASE
    ind_j = {}
    for j in C:
        ci, di, ri = C[j], due[j], rel[j]
        li = ci - di
        ti = max(0.0, li)
        ind_j[j] = {'Ci': ci, 'ri': ri, 'di': di, 'Fi': ci - ri, 'Li': li, 'Ti': ti,
                    'Ei': max(0.0, -li), 'delta_i': 1 if ti > EPS else 0}
    n = len(C)
    return {
        'por_trabajo': ind_j,
        'Cmax'     : max(C.values()),
        'suma_Ci'  : sum(C.values()),
        'Lmax'     : max(v['Li'] for v in ind_j.values()),
        'Tmax'     : max(v['Ti'] for v in ind_j.values()),
        'suma_Ti'  : sum(v['Ti'] for v in ind_j.values()),
        'suma_Ei'  : sum(v['Ei'] for v in ind_j.values()),
        'suma_Fi'  : sum(v['Fi'] for v in ind_j.values()),
        'media_Fi' : sum(v['Fi'] for v in ind_j.values()) / n,
        'num_tardy': sum(v['delta_i'] for v in ind_j.values()),
    }


def fitness(individuo: dict) -> Tuple[float, float]:
    """
    (ΣTi, Cmax): se minimiza ΣTi y, en empate, Cmax. El desempate hace que
    el AG siga mejorando el programa cuando ya encontró ΣTi = 0.
    Se guarda en el individuo para no recalcular.
    """
    if '_fitness' not in individuo:
        C, _ = simular_hfs(individuo)
        ind = calcular_indicadores(C)
        individuo['_fitness'] = (round(ind['suma_Ti'], 9), ind['Cmax'])
    return individuo['_fitness']

# ---------------------------------------------------------------------------
# 9.  OPERADORES GENÉTICOS
# ---------------------------------------------------------------------------

def seleccion_torneo(poblacion: List[dict], tam: int) -> dict:
    return min(random.sample(poblacion, tam), key=fitness)


def _cruce_un_hijo(padre: dict, madre: dict, de_padre: set) -> dict:
    """
    Cruce por trabajos (tipo POX) para la permutación con repetición: los
    sublotes de los trabajos de 'de_padre' conservan sus posiciones relativas
    del padre; los demás se intercalan en el orden de la madre. Cada trabajo
    hereda k y tamaños del MISMO progenitor que le da sus posiciones, así el
    hijo siempre tiene tantas posiciones como sublotes.
    """
    lp, lm = len(padre['secuencia']), len(madre['secuencia'])
    genes = [((i + 0.5) / lp, 0, j) for i, j in enumerate(padre['secuencia']) if j in de_padre]
    genes += [((i + 0.5) / lm, 1, j) for i, j in enumerate(madre['secuencia'])
              if j not in de_padre]
    genes.sort()
    origen = {j: (padre if j in de_padre else madre) for j in TRABAJOS}
    return {'secuencia': [j for _, _, j in genes],
            'k': {j: origen[j]['k'][j] for j in TRABAJOS},
            'tamanos': {j: list(origen[j]['tamanos'][j]) for j in TRABAJOS}}


def cruce(p1: dict, p2: dict) -> Tuple[dict, dict]:
    """Cruce por subconjunto aleatorio de trabajos; produce dos hijos simétricos."""
    n = len(TRABAJOS)
    conjunto = set(random.sample(TRABAJOS, random.randint(1, n - 1)))
    return _cruce_un_hijo(p1, p2, conjunto), _cruce_un_hijo(p2, p1, conjunto)


def mutar(individuo: dict) -> dict:
    """
    Mutación:
      - posiciones: intercambio (swap) de dos sublotes de trabajos distintos
      - tamaños   : perturbación gaussiana de una proporción (igual en todas
                    las etapas, el sublote sigue siendo consistente)
      - k         : se agrega o quita un sublote (y su posición)
    """
    ind = copy.deepcopy(individuo)
    ind.pop('_fitness', None)
    seq = ind['secuencia']

    if random.random() < GA_PARAMS['prob_mut_seq'] and len(set(seq)) > 1:
        while True:
            a, b = random.sample(range(len(seq)), 2)
            if seq[a] != seq[b]:
                break
        seq[a], seq[b] = seq[b], seq[a]

    for j in TRABAJOS:
        props = ind['tamanos'][j]
        if random.random() < GA_PARAMS['prob_mut_tam'] and len(props) > 1:
            i = random.randrange(len(props))
            props[i] = max(1e-9, props[i] + random.gauss(0, 0.15))
            t = sum(props)
            ind['tamanos'][j] = [x / t for x in props]

        if random.random() < GA_PARAMS['prob_mut_k']:
            k_act = ind['k'][j]
            k_nuevo = max(1, min(max_sublotes(j), k_act + random.choice([-1, 1])))
            if k_nuevo > k_act:
                # El sublote más grande se parte en dos; el nuevo va a una posición al azar
                props = ind['tamanos'][j]
                i = max(range(k_act), key=lambda x: props[x])
                props[i] /= 2.0
                props.append(props[i])
                seq.insert(random.randint(0, len(seq)), j)
            elif k_nuevo < k_act:
                # Un sublote al azar se une a otro; se quita una posición de j
                props = ind['tamanos'][j]
                quitado = props.pop(random.randrange(k_act))
                props[random.randrange(len(props))] += quitado
                posiciones = [p for p, x in enumerate(seq) if x == j]
                del seq[random.choice(posiciones)]
            ind['k'][j] = k_nuevo
    return ind

# ---------------------------------------------------------------------------
# 10. ALGORITMO GENÉTICO PRINCIPAL
# ---------------------------------------------------------------------------

def algoritmo_genetico(semillas: Optional[List[List[str]]] = None,
                       verbose: bool = True) -> Tuple[dict, List[float]]:
    """
    Ejecuta el AG y retorna (mejor_individuo, historial del mejor ΣTi).

    La población inicial es aleatoria (escenario S2: tamaños de sublote y
    posiciones aleatorios). 'semillas' son órdenes de trabajos (p. ej. de las
    reglas de despacho) que se insertan sin lot streaming y con sublotes
    iguales; así el AG nunca termina peor que esas secuencias.
    """
    if GA_PARAMS['semilla'] is not None:
        random.seed(GA_PARAMS['semilla'])

    poblacion = []
    for orden in semillas or []:
        poblacion.append(crear_individuo(orden, k_fijo=1))
        poblacion.append(crear_individuo(orden, k_fijo=MAX_SUBLOTES, iguales=True))
    poblacion = poblacion[:GA_PARAMS['poblacion']]
    while len(poblacion) < GA_PARAMS['poblacion']:
        poblacion.append(crear_individuo())
    for ind in poblacion:
        fitness(ind)

    mejor_global = copy.deepcopy(min(poblacion, key=fitness))
    historial = []
    tam_torneo = min(GA_PARAMS['tam_torneo'], GA_PARAMS['poblacion'])

    if verbose:
        print(f"\n{'=' * 72}")
        print("  ALGORITMO GENÉTICO — HFS CON SUBLOTES CONSISTENTES (escenario S2)")
        print(f"  Población: {GA_PARAMS['poblacion']}  |  Generaciones: "
              f"{GA_PARAMS['generaciones']}  |  Máx. sublotes: {MAX_SUBLOTES}  |  "
              f"Mín. sublote: {MIN_PARES_SUBLOTE} pares")
        print("  Máquinas: " + " | ".join(f"{NOMBRE_ETAPA[s]}={MAQUINAS_POR_ETAPA[s]}"
                                          for s in ETAPAS))
        print(f"{'=' * 72}")
        print(f"  {'Gen':>5} | {'Mejor ΣTi':>12} | {'Su Cmax':>10} | {'Media ΣTi':>12} | "
              f"{'Peor ΣTi':>12}")
        print(f"  {'-' * 65}")

    for gen in range(GA_PARAMS['generaciones']):
        nueva = []
        if GA_PARAMS['elitismo']:
            nueva.append(copy.deepcopy(min(poblacion, key=fitness)))

        while len(nueva) < GA_PARAMS['poblacion']:
            p1 = seleccion_torneo(poblacion, tam_torneo)
            p2 = seleccion_torneo(poblacion, tam_torneo)
            if random.random() < GA_PARAMS['prob_cruce']:
                h1, h2 = cruce(p1, p2)
            else:
                h1, h2 = copy.deepcopy(p1), copy.deepcopy(p2)
            if random.random() < GA_PARAMS['prob_mutacion']:
                h1 = mutar(h1)
            if random.random() < GA_PARAMS['prob_mutacion']:
                h2 = mutar(h2)
            nueva.extend([h1, h2])

        poblacion = nueva[:GA_PARAMS['poblacion']]
        f_vals = [fitness(i)[0] for i in poblacion]

        mejor_gen = min(poblacion, key=fitness)
        if fitness(mejor_gen) < fitness(mejor_global):
            mejor_global = copy.deepcopy(mejor_gen)
        historial.append(fitness(mejor_global)[0])

        if verbose and (gen % 50 == 0 or gen == GA_PARAMS['generaciones'] - 1):
            print(f"  {gen + 1:>5} | {fitness(mejor_global)[0]:>12.4f} | "
                  f"{fitness(mejor_global)[1]:>10.4f} | "
                  f"{sum(f_vals) / len(f_vals):>12.4f} | {max(f_vals):>12.4f}")

    if verbose:
        print(f"{'=' * 72}\n")
    return mejor_global, historial

# ---------------------------------------------------------------------------
# 11. REFERENCIAS: REGLAS DE DESPACHO Y ENUMERACIÓN EXACTA
# ---------------------------------------------------------------------------

def evaluar_reglas() -> List[dict]:
    """Evalúa las reglas de despacho sin lot streaming con los datos reales."""
    resultados = []
    for regla in REGLAS:
        C, gantt = simular_despacho(regla)
        # Orden de entrada a la primera etapa (sirve de semilla para el AG)
        seq = [j for (j, *_r) in sorted(gantt[ETAPAS[0]], key=lambda x: (x[3], x[2]))]
        resultados.append({'regla': regla, 'secuencia': seq, **calcular_indicadores(C)})
    return resultados


def enumeracion_sin_ls() -> Optional[dict]:
    """
    Si n ≤ ENUMERACION_MAX_N, evalúa TODAS las secuencias de permutación sin
    lot streaming. Da el óptimo exacto de ese modelo como referencia.
    """
    if len(TRABAJOS) > ENUMERACION_MAX_N:
        return None
    mejor = None
    for seq in itertools.permutations(TRABAJOS):
        C, _ = simular_sin_ls(list(seq))
        ind = calcular_indicadores(C)
        if mejor is None or ind['suma_Ti'] < mejor['suma_Ti'] - EPS:
            mejor = {'secuencia': list(seq), **ind}
    mejor['evaluadas'] = math.factorial(len(TRABAJOS))
    return mejor

# ---------------------------------------------------------------------------
# 12. VALIDACIÓN CON LEKIN
# ---------------------------------------------------------------------------

def datos_lekin() -> Tuple[List[str], pd.DataFrame, Dict]:
    """
    Datos para cargar en LEKIN (Flexible Flow Shop), en MINUTOS ENTEROS:
    tiempo del lote completo en cada etapa (workcenter) con el PROMEDIO de
    min/docena de las máquinas capaces (máquinas idénticas), fecha de
    liberación, fecha de entrega y peso = 1.
    """
    trabajos = TRABAJOS[:MAX_TRABAJOS_LEKIN]
    p_min = {}
    for j in trabajos:
        tiempos = TIEMPO_MIN_DOCENA[PRODUCTO[j]]
        p_min[j] = {s: int(round(docenas(DEMANDA[j]) *
                                 tiempo_promedio([(m, tiempos[MAQUINAS[s][m]])
                                                  for m, _ in RAPIDEZ[j][s]])))
                    for s in ETAPAS}
    due_min = {j: int(round(DUE_DATE[j] * 60)) for j in trabajos}
    rel_min = {j: int(round(RELEASE[j] * 60)) for j in trabajos}
    filas = []
    for j in trabajos:
        fila = {'Job': j, 'Producto': PRODUCTO[j],
                'Release (min)': rel_min[j], 'Due (min)': due_min[j], 'Weight': 1}
        for s in ETAPAS:
            fila[f"W{s} {NOMBRE_ETAPA[s]} ({MAQUINAS_POR_ETAPA[s]} maq) [min]"] = p_min[j][s]
        filas.append(fila)
    return trabajos, pd.DataFrame(filas), {'p': p_min, 'due': due_min, 'rel': rel_min}


def validacion_lekin() -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Corre las reglas de despacho con EXACTAMENTE los datos de la hoja
    LEKIN_datos (minutos enteros, máquinas idénticas). Los resultados deben
    coincidir con los que muestra LEKIN para la misma regla.
    """
    trabajos, tabla, d = datos_lekin()
    p_maq = {j: {s: [(m, d['p'][j][s] / 60.0) for m, _ in RAPIDEZ[j][s]] for s in ETAPAS}
             for j in trabajos}
    filas = []
    for regla in REGLAS:
        C_h, _ = simular_despacho(regla, trabajos, p_maq=p_maq,
                                  due={j: d['due'][j] / 60.0 for j in trabajos},
                                  rel={j: d['rel'][j] / 60.0 for j in trabajos})
        C = {j: round(C_h[j] * 60) for j in trabajos}
        ind = calcular_indicadores(C, due=d['due'], rel=d['rel'])
        filas.append({'Regla': regla, 'Cmax': ind['Cmax'], 'Lmax': ind['Lmax'],
                      'ΣUj (tardíos)': ind['num_tardy'], 'ΣCj': ind['suma_Ci'],
                      'ΣTj': ind['suma_Ti'], 'ΣwjCj': ind['suma_Ci'], 'ΣwjTj': ind['suma_Ti']})
    return tabla, pd.DataFrame(filas)

# ---------------------------------------------------------------------------
# 13. REPORTES EN CONSOLA
# ---------------------------------------------------------------------------

SEP = "=" * 76
SEP2 = "-" * 76


def _fmt_fecha(h: float) -> str:
    f = horas_a_fecha(h)
    return f.isoformat() if f else ("fuera de calendario" if CALENDARIO else "")


def imprimir_datos():
    print(SEP)
    print("  DATOS LEÍDOS DEL EXCEL")
    print(SEP)
    for s in ETAPAS:
        maq = MAQUINAS[s]
        print(f"  Etapa {s} {NOMBRE_ETAPA[s]:<12} → {len(maq)} máquinas ({maq[0]}…{maq[-1]})")
    if CALENDARIO:
        print(f"  Calendario: {CALENDARIO[0][0]} a {CALENDARIO[-1][0]} "
              f"({CALENDARIO[-1][2]:.0f} h laborables)")
    print()
    print(SEP)
    print("  TIEMPOS POR PRODUCTO [min/docena] (mín–máx entre las máquinas de la etapa)")
    print(SEP)
    print(f"  {'Prod':>5} | " + " | ".join(f"{NOMBRE_ETAPA[s][:11]:>11}" for s in ETAPAS))
    print(SEP2)
    for prod in sorted({PRODUCTO[j] for j in TRABAJOS}):
        celdas = []
        for s in ETAPAS:
            t = [TIEMPO_MIN_DOCENA[prod][m] for m in MAQUINAS[s] if m in TIEMPO_MIN_DOCENA[prod]]
            celdas.append(f"{min(t):>5.2f}–{max(t):<5.2f}" if min(t) != max(t)
                          else f"{min(t):>11.2f}")
        print(f"  {prod:>5} | " + " | ".join(celdas))
    print()
    print(SEP)
    print("  TRABAJOS")
    print(SEP)
    print(f"  {'OT':>4} | {'Prod':>5} | {'Pares':>6} | {'Docenas':>7} | {'ri(h)':>6} | "
          f"{'di(h)':>6} | {'Entrega':>10}")
    print(SEP2)
    for j in TRABAJOS:
        fe = FECHA_ENTREGA[j].isoformat() if FECHA_ENTREGA[j] else ""
        print(f"  {j:>4} | {PRODUCTO[j]:>5} | {DEMANDA[j]:>6} | {docenas(DEMANDA[j]):>7.2f} | "
              f"{RELEASE[j]:>6.1f} | {DUE_DATE[j]:>6.1f} | {fe:>10}")
    print()


def imprimir_reglas(resultados: List[dict], exacto: Optional[dict]):
    print(SEP)
    print("  REGLAS DE DESPACHO — SIN LOT STREAMING (referencia)")
    print(SEP)
    print(f"  {'Regla':>6} | {'ΣTi':>9} | {'Tmax':>9} | {'Cmax':>9} | {'#tard':>5} | "
          f"Orden en etapa 1")
    print(SEP2)
    for r in sorted(resultados, key=lambda x: x['suma_Ti']):
        print(f"  {r['regla']:>6} | {r['suma_Ti']:>9.4f} | {r['Tmax']:>9.4f} | "
              f"{r['Cmax']:>9.4f} | {r['num_tardy']:>5} | {' '.join(r['secuencia'])}")
    if exacto:
        print(SEP2)
        print(f"  {'ÓPTIMO':>6} | {exacto['suma_Ti']:>9.4f} | {exacto['Tmax']:>9.4f} | "
              f"{exacto['Cmax']:>9.4f} | {exacto['num_tardy']:>5} | "
              f"{' '.join(exacto['secuencia'])}")
        print(f"  (óptimo exacto de permutación sin LS: {exacto['evaluadas']} secuencias evaluadas)")
    print()


def tabla_sublotes(individuo: dict) -> pd.DataFrame:
    filas = []
    for j in TRABAJOS:
        tam = tamanos_pares(j, individuo['tamanos'][j])
        for b, q in enumerate(tam, start=1):
            filas.append({'Trabajo': j, 'Producto': PRODUCTO[j], 'k_sublotes': len(tam),
                          'Sublote': b, 'Pares': q, 'Docenas': round(docenas(q), 3)})
    return pd.DataFrame(filas)


def tabla_programa(gantt: Dict[int, list]) -> pd.DataFrame:
    filas = []
    for s in ETAPAS:
        for (j, b, m, ini, fin, q) in sorted(gantt[s], key=lambda x: (x[3], x[2])):
            filas.append({'Etapa': NOMBRE_ETAPA[s], 'Maquina': MAQUINAS[s][m],
                          'Trabajo': j, 'Producto': PRODUCTO[j], 'Sublote': b,
                          'Pares': q, 'Docenas': round(docenas(q), 3),
                          'min_x_docena': TIEMPO_MIN_DOCENA[PRODUCTO[j]][MAQUINAS[s][m]],
                          'Inicio_h': round(ini, 4), 'Fin_h': round(fin, 4),
                          'Duracion_h': round(fin - ini, 4),
                          'Fecha_inicio': _fmt_fecha(ini), 'Fecha_fin': _fmt_fecha(fin)})
    return pd.DataFrame(filas)


def tabla_indicadores(ind: dict) -> pd.DataFrame:
    filas = []
    for j, vals in ind['por_trabajo'].items():
        fe = FECHA_ENTREGA[j].isoformat() if FECHA_ENTREGA[j] else ""
        filas.append({'Trabajo': j, 'Producto': PRODUCTO[j], 'Pares': DEMANDA[j],
                      **{k: (round(v, 4) if isinstance(v, float) else v) for k, v in vals.items()},
                      'Fecha_entrega': fe, 'Fecha_terminacion': _fmt_fecha(vals['Ci'])})
    return pd.DataFrame(filas)


def imprimir_solucion(titulo: str, individuo: dict):
    C, gantt = simular_hfs(individuo)
    ind = calcular_indicadores(C)

    print(SEP)
    print(f"  {titulo}")
    print("  Posiciones de los sublotes (etapa 1): " + " ".join(individuo['secuencia']))
    print("\n  SUBLOTES CONSISTENTES (mismo tamaño en todas las etapas):")
    for j in TRABAJOS:
        tam = tamanos_pares(j, individuo['tamanos'][j])
        print(f"    {j} ({PRODUCTO[j]}, {DEMANDA[j]} pares): k = {len(tam)} → "
              + " + ".join(f"{q}" for q in tam) + " pares")

    print(f"\n  {'OT':>4} | {'di(h)':>7} | {'Ci(h)':>8} | {'Fi(h)':>8} | {'Li(h)':>8} | "
          f"{'Ti(h)':>8} | {'Ei(h)':>8} | {'δi':>3} | Terminación")
    print("  " + "-" * 84)
    for j in TRABAJOS:
        v = ind['por_trabajo'][j]
        print(f"  {j:>4} | {v['di']:>7.1f} | {v['Ci']:>8.4f} | {v['Fi']:>8.4f} | "
              f"{v['Li']:>8.4f} | {v['Ti']:>8.4f} | {v['Ei']:>8.4f} | {v['delta_i']:>3} | "
              f"{_fmt_fecha(v['Ci'])}")

    print("\n  INDICADORES GLOBALES")
    print(f"    Cmax      (makespan)              = {ind['Cmax']:>10.4f} h")
    print(f"    Lmax      (retraso máximo)        = {ind['Lmax']:>10.4f} h")
    print(f"    Tmax      (tardanza máxima)       = {ind['Tmax']:>10.4f} h")
    print(f"    ΣTi       (tardanza total) [FO]   = {ind['suma_Ti']:>10.4f} h  ← FO")
    print(f"    ΣEi       (adelanto total)        = {ind['suma_Ei']:>10.4f} h")
    print(f"    ΣFi       (flujo total)           = {ind['suma_Fi']:>10.4f} h")
    print(f"    F̄         (flujo medio)           = {ind['media_Fi']:>10.4f} h")
    print(f"    # tardíos (Σδi)                   = {ind['num_tardy']:>10}")
    print()


def imprimir_comparacion(ind_sin: dict, ind_con: dict, nombre_sin: str):
    print(SEP)
    print(f"  COMPARACIÓN: SIN LOT STREAMING ({nombre_sin}) vs. SUBLOTES CONSISTENTES (AG)")
    print(SEP)
    print(f"  {'Indicador':>25} | {'Sin LS':>12} | {'Con LS':>12} | {'Δ%':>9}")
    print("  " + "-" * 66)
    for nombre, clave in [('ΣTi (tardanza total)', 'suma_Ti'), ('Cmax (makespan)', 'Cmax'),
                          ('Tmax (tardanza máx.)', 'Tmax'), ('# tardíos (Σδi)', 'num_tardy')]:
        v_sin, v_con = ind_sin[clave], ind_con[clave]
        delta = (v_con - v_sin) / v_sin * 100 if v_sin > 0 else 0.0
        signo = "▼" if delta < 0 else ("▲" if delta > 0 else "─")
        print(f"  {nombre:>25} | {v_sin:>12.4f} | {v_con:>12.4f} |   {signo}{abs(delta):>5.1f}%")
    print()


def imprimir_lekin(tabla: pd.DataFrame, res: pd.DataFrame):
    print(SEP)
    print("  VALIDACIÓN CON LEKIN — datos a cargar (minutos enteros, máquinas idénticas)")
    print(SEP)
    if len(TRABAJOS) > MAX_TRABAJOS_LEKIN:
        print(f"  ⚠ Hay {len(TRABAJOS)} trabajos; LEKIN admite {MAX_TRABAJOS_LEKIN}. "
              f"Se usan los primeros {MAX_TRABAJOS_LEKIN} de la hoja Trabajos.")
    print("  " + tabla.to_string(index=False).replace("\n", "\n  "))
    print("\n  Resultados que deben coincidir con LEKIN (mismos datos y regla) [min]:")
    print("  " + res.to_string(index=False).replace("\n", "\n  "))
    print()

# ---------------------------------------------------------------------------
# 14. DIAGRAMA DE GANTT
# ---------------------------------------------------------------------------

def graficar_gantt(individuo: dict,
                   titulo: str = "Diagrama de Gantt — HFS con sublotes consistentes",
                   guardar_como: Optional[str] = None):
    C, gantt = simular_hfs(individuo)
    ind = calcular_indicadores(C)
    cmax = ind['Cmax']

    colores_base = ['#4E79A7', '#F28E2B', '#59A14F', '#E15759', '#B07AA1',
                    '#76B7B2', '#EDC948', '#FF9DA7', '#9C755F', '#BAB0AC']
    color_trabajo = {j: colores_base[i % len(colores_base)] for i, j in enumerate(TRABAJOS)}

    alturas = [MAQUINAS_POR_ETAPA[s] for s in ETAPAS]
    fig, axes = plt.subplots(len(ETAPAS), 1,
                             figsize=(20, max(8, sum(alturas) * 0.35 + 3)),
                             gridspec_kw={'height_ratios': alturas}, sharex=True)
    if len(ETAPAS) == 1:
        axes = [axes]
    fig.suptitle(titulo, fontsize=14, fontweight='bold', y=1.0)
    umbral_texto = (cmax if cmax > 0 else 1) * 0.02

    for ax, s in zip(axes, ETAPAS):
        n_maq = MAQUINAS_POR_ETAPA[s]
        for (j, b, m, ini, fin, _q) in gantt[s]:
            dur = fin - ini
            ax.barh(m, dur, left=ini, height=0.72, color=color_trabajo[j],
                    edgecolor='white', linewidth=0.6, alpha=0.92)
            if dur >= umbral_texto:
                ax.text(ini + dur / 2, m, f"{j}-S{b}", ha='center', va='center',
                        fontsize=6.5, fontweight='bold', color='white', clip_on=True)
        ax.axvline(cmax, color='#333333', linewidth=1.2, linestyle='--', alpha=0.7)
        ax.set_yticks(range(n_maq))
        ax.set_yticklabels(MAQUINAS[s], fontsize=7)
        ax.set_ylim(n_maq - 0.4, -0.6)   # primera máquina arriba
        ax.set_ylabel(f"Etapa {s}\n{NOMBRE_ETAPA[s]}", fontsize=9, fontweight='bold', labelpad=8)
        ax.xaxis.set_minor_locator(ticker.AutoMinorLocator(4))
        ax.grid(axis='x', which='major', linestyle='--', linewidth=0.5, alpha=0.5, color='#bbbbbb')
        ax.grid(axis='x', which='minor', linestyle=':', linewidth=0.3, alpha=0.35, color='#cccccc')
        ax.set_facecolor('#F8F9FA')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

    axes[-1].set_xlabel("Tiempo (horas laborables)", fontsize=10)
    axes[-1].xaxis.set_major_locator(ticker.MaxNLocator(nbins=15))
    axes[-1].annotate(f"Cmax = {cmax:.2f} h", xy=(cmax, 1.0), xycoords=('data', 'axes fraction'),
                      xytext=(-4, -4), textcoords='offset points', fontsize=8, ha='right',
                      va='top', color='#333333')

    fig.legend(handles=[mpatches.Patch(color=color_trabajo[j], label=f"{j} ({PRODUCTO[j]})")
                        for j in TRABAJOS],
               loc='lower center', ncol=min(len(TRABAJOS), 10), fontsize=9, frameon=True,
               framealpha=0.9, edgecolor='#cccccc', bbox_to_anchor=(0.5, -0.03))
    fig.text(0.5, 1.002, f"ΣTi = {ind['suma_Ti']:.4f} h  |  Cmax = {cmax:.4f} h  |  "
                         f"# tardíos = {ind['num_tardy']}",
             ha='center', va='bottom', fontsize=8.5, color='#444444', transform=fig.transFigure)
    plt.tight_layout(rect=[0, 0.04, 1, 0.99])

    # Se guarda ANTES de plt.show(): en algunos entornos show() vacía la figura
    if guardar_como:
        fig.savefig(guardar_como, dpi=150, bbox_inches='tight')
        print(f"  Gantt guardado en: {guardar_como}")
    plt.show()
    return fig

# ---------------------------------------------------------------------------
# 15. EXPORTAR RESULTADOS A EXCEL
# ---------------------------------------------------------------------------

def exportar_resultados(ruta: str, resumen: List[dict], reglas: List[dict],
                        exacto: Optional[dict], mejor: dict, historial: List[float],
                        lekin_tabla: pd.DataFrame, lekin_res: pd.DataFrame):
    C, gantt = simular_hfs(mejor)
    ind = calcular_indicadores(C)
    df_reglas = pd.DataFrame([{
        'Regla': r['regla'], 'Orden_etapa1': " ".join(r['secuencia']),
        'ΣTi_h': round(r['suma_Ti'], 4), 'Tmax_h': round(r['Tmax'], 4),
        'Cmax_h': round(r['Cmax'], 4), 'Tardios': r['num_tardy']} for r in reglas])
    if exacto:
        df_reglas = pd.concat([df_reglas, pd.DataFrame([{
            'Regla': 'ÓPTIMO permutación sin LS', 'Orden_etapa1': " ".join(exacto['secuencia']),
            'ΣTi_h': round(exacto['suma_Ti'], 4), 'Tmax_h': round(exacto['Tmax'], 4),
            'Cmax_h': round(exacto['Cmax'], 4), 'Tardios': exacto['num_tardy']}])])

    with pd.ExcelWriter(ruta, engine="openpyxl") as w:
        pd.DataFrame(resumen).to_excel(w, sheet_name="Resumen", index=False)
        df_reglas.to_excel(w, sheet_name="Reglas_sin_LS", index=False)
        tabla_indicadores(ind).to_excel(w, sheet_name="Indicadores_AG", index=False)
        tabla_sublotes(mejor).to_excel(w, sheet_name="Sublotes_AG", index=False)
        tabla_programa(gantt).to_excel(w, sheet_name="Programa_AG", index=False)
        pd.DataFrame({'Generacion': range(1, len(historial) + 1), 'Mejor_ΣTi_h': historial}) \
            .to_excel(w, sheet_name="Convergencia", index=False)
        lekin_tabla.to_excel(w, sheet_name="LEKIN_datos", index=False)
        lekin_res.to_excel(w, sheet_name="LEKIN_esperado", index=False)
        for hoja in w.sheets.values():
            for col in hoja.columns:
                ancho = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                hoja.column_dimensions[col[0].column_letter].width = min(max(10, ancho + 2), 60)
    print(f"  Resultados guardados en: {ruta}")

# ---------------------------------------------------------------------------
# 16. PROGRAMA PRINCIPAL
# ---------------------------------------------------------------------------

def main(ruta_excel: str = RUTA_EXCEL, mostrar_gantt: bool = True):
    cargar_datos(obtener_ruta_excel(ruta_excel))

    print()
    print(SEP)
    print("  HFS CON LOT STREAMING — SUBLOTES CONSISTENTES — GUANTES INDUSTRIALES")
    print("  Escenario S2 (Gürsoy Yılmaz & Yılmaz, 2022) | AG | Mín. tardanza total")
    print(SEP)
    print()
    imprimir_datos()

    # ── Referencias sin lot streaming ──
    reglas = evaluar_reglas()
    exacto = enumeracion_sin_ls()
    imprimir_reglas(reglas, exacto)
    mejor_regla = min(reglas, key=lambda x: x['suma_Ti'])
    print(f"  ✓ Mejor regla sin LS: {mejor_regla['regla']} → ΣTi = {mejor_regla['suma_Ti']:.4f} h\n")

    # ── Validación LEKIN ──
    lekin_tabla, lekin_res = validacion_lekin()
    imprimir_lekin(lekin_tabla, lekin_res)

    # ── AG con sublotes consistentes (población inicial sembrada con las reglas) ──
    semillas = [r['secuencia'] for r in reglas]
    if exacto:
        semillas.append(exacto['secuencia'])
    mejor, historial = algoritmo_genetico(semillas=semillas, verbose=True)
    C_con, gantt_con = simular_hfs(mejor)
    ind_con = calcular_indicadores(C_con)

    imprimir_solucion("MEJOR SOLUCIÓN — AG CON SUBLOTES CONSISTENTES", mejor)
    errores = verificar_programa(gantt_con)
    for r in reglas:
        errores += verificar_programa(simular_despacho(r['regla'])[1])
    if errores:
        print("  ✗ SE ENCONTRARON ERRORES DE FACTIBILIDAD:")
        for e in errores:
            print(f"    - {e}")
    else:
        print("  ✓ Verificación: los programas del AG y de las reglas son factibles "
              "(sin traslapes, volúmenes completos, sublotes consistentes y "
              "precedencias respetadas)\n")
    imprimir_comparacion(mejor_regla, ind_con, f"regla {mejor_regla['regla']}")

    if mostrar_gantt:
        print("  Generando diagrama de Gantt...")
        graficar_gantt(mejor, "Diagrama de Gantt — HFS con sublotes consistentes (AG)",
                       guardar_como=RUTA_GANTT)

    resumen = [
        {'Metodo': f"Mejor regla sin LS ({mejor_regla['regla']})",
         'ΣTi_h': mejor_regla['suma_Ti'], 'Cmax_h': mejor_regla['Cmax'],
         'Tmax_h': mejor_regla['Tmax'], 'Tardios': mejor_regla['num_tardy']},
    ]
    if exacto:
        resumen.append({'Metodo': 'Óptimo permutación sin LS (enumeración)',
                        'ΣTi_h': exacto['suma_Ti'], 'Cmax_h': exacto['Cmax'],
                        'Tmax_h': exacto['Tmax'], 'Tardios': exacto['num_tardy']})
    resumen.append({'Metodo': 'AG con sublotes consistentes', 'ΣTi_h': ind_con['suma_Ti'],
                    'Cmax_h': ind_con['Cmax'], 'Tmax_h': ind_con['Tmax'],
                    'Tardios': ind_con['num_tardy']})
    resumen.append({'Metodo': "Posiciones de sublotes AG: " + " ".join(mejor['secuencia'])})
    exportar_resultados(RUTA_RESULTADOS, resumen, reglas, exacto, mejor, historial,
                        lekin_tabla, lekin_res)

    print(SEP)
    print("  RESUMEN EJECUTIVO")
    print(SEP)
    for r in resumen[:-1]:
        print(f"  {r['Metodo']:<42} ΣTi={r['ΣTi_h']:>9.4f} h | Cmax={r['Cmax_h']:>9.4f} h | "
              f"#tard={r['Tardios']}")
    if mejor_regla['suma_Ti'] > 0:
        mej = (mejor_regla['suma_Ti'] - ind_con['suma_Ti']) / mejor_regla['suma_Ti'] * 100
        print(f"\n  Mejora de ΣTi del AG frente a la mejor regla: {mej:.1f}%")
    print(SEP)
    print()

    # En Colab, ofrecer la descarga de los resultados
    try:
        from google.colab import files  # type: ignore
        files.download(RUTA_RESULTADOS)
        if mostrar_gantt:
            files.download(RUTA_GANTT)
    except ImportError:
        pass

    return mejor


if __name__ == "__main__":
    # Permite: python hfs_guantes.py otro_archivo.xlsx
    args = [a for a in sys.argv[1:] if a.lower().endswith(".xlsx")]
    main(args[0] if args else RUTA_EXCEL)
