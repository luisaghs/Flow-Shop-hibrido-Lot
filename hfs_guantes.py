"""
=============================================================================
PROGRAMACIÓN DE UN HYBRID FLOW SHOP (HFS) CON LOT STREAMING POR ETAPA
Y ALGORITMO GENÉTICO — PRODUCCIÓN DE GUANTES INDUSTRIALES
=============================================================================

Etapas (definidas en el Excel): Troquelado → Ensamblaje → Costura → Voteado.
Ensamblaje tiene máquinas paralelas NO idénticas (uniformes): cada grupo de
máquinas tiene un factor de tiempo según su velocidad por docena.

Los datos se leen de `datos_guantes.xlsx` (ver hoja "Instrucciones").
Los resultados se guardan en `resultados_programacion.xlsx` y `gantt.png`.

Validación con LEKIN:
  LEKIN no hace lot streaming y admite máximo 10 trabajos. Por eso este
  programa incluye un simulador de reglas de despacho (EDD, SPT, LPT, FIFO,
  LIFO, MS, CR) para el modelo SIN lot streaming, y genera la hoja
  `LEKIN_datos` con los tiempos en minutos enteros para cargar en LEKIN.
  Si el simulador está bien, cada regla debe dar el mismo ΣTj en Python y en
  LEKIN. Después se compara el AG contra la mejor regla de LEKIN.
"""

import copy
import itertools
import math
import os
import random
import sys
from typing import Dict, List, Optional, Tuple

import pandas as pd

# NO se llama a matplotlib.use(): Colab y Jupyter configuran su propio backend.
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker as ticker

RUTA_EXCEL = "datos_guantes.xlsx"
RUTA_RESULTADOS = "resultados_programacion.xlsx"
RUTA_GANTT = "gantt.png"

EPS = 1e-6

# ---------------------------------------------------------------------------
# 1.  DATOS DEL PROBLEMA (se llenan desde el Excel con cargar_datos)
# ---------------------------------------------------------------------------

TRABAJOS: List[int] = []
DESCRIPCION: Dict[int, str] = {}
DEMANDA: Dict[int, float] = {}            # [pares]
DUE_DATE: Dict[int, float] = {}           # [horas]
RELEASE: Dict[int, float] = {}            # [horas]
ETAPAS: List[int] = []                    # 1..S en orden de flujo
NOMBRE_ETAPA: Dict[int, str] = {}
TIEMPO_MIN_DOCENA: Dict[int, Dict[int, float]] = {}   # [min/docena] estándar
MAQUINAS: Dict[int, List[dict]] = {}      # etapa → [{'id', 'grupo', 'factor'}]
MAQUINAS_POR_ETAPA: Dict[int, int] = {}

MAX_SPLIT = 4
MIN_DOCENAS_POR_SUBLOTE = 0.5
USAR_FACTORES = True
ENUMERACION_MAX_N = 8
MAX_TRABAJOS_LEKIN = 10

GA_PARAMS = {
    'poblacion'    : 100,
    'generaciones' : 300,
    'prob_cruce'   : 0.85,
    'prob_mutacion': 0.25,
    'prob_mut_seq' : 0.30,
    'prob_mut_prop': 0.40,
    'prob_mut_k'   : 0.30,
    'tam_torneo'   : 3,
    'elitismo'     : True,
    'semilla'      : 42,
}

# ---------------------------------------------------------------------------
# 2.  LECTURA DEL EXCEL
# ---------------------------------------------------------------------------

def _texto(v) -> str:
    return "" if pd.isna(v) else str(v).strip()


def _si_no(v) -> bool:
    return _texto(v).lower() in ("sí", "si", "s", "yes", "true", "1", "1.0", "verdadero")


def _leer_hoja(xls: pd.ExcelFile, hoja: str, columnas: List[str]) -> pd.DataFrame:
    if hoja not in xls.sheet_names:
        raise ValueError(f"El Excel no tiene la hoja '{hoja}'. Hojas encontradas: {xls.sheet_names}")
    df = xls.parse(hoja)
    df.columns = [str(c).strip() for c in df.columns]
    faltantes = [c for c in columnas if c not in df.columns]
    if faltantes:
        raise ValueError(f"Hoja '{hoja}': faltan las columnas {faltantes}")
    # Se descartan filas totalmente vacías (p. ej. formato aplicado sin datos)
    return df.dropna(how="all", subset=columnas).reset_index(drop=True)


def _numero(v, hoja: str, campo: str, minimo: float = 0.0, estricto: bool = False) -> float:
    try:
        x = float(v)
    except (TypeError, ValueError):
        raise ValueError(f"Hoja '{hoja}': '{campo}' debe ser numérico (valor leído: {v!r})")
    if math.isnan(x) or x < minimo or (estricto and x <= minimo):
        signo = ">" if estricto else "≥"
        raise ValueError(f"Hoja '{hoja}': '{campo}' debe ser {signo} {minimo} (valor leído: {v!r})")
    return x


def cargar_datos(ruta: str = RUTA_EXCEL):
    """Lee el Excel y llena las estructuras globales del modelo."""
    global MAX_SPLIT, MIN_DOCENAS_POR_SUBLOTE, USAR_FACTORES
    global ENUMERACION_MAX_N, MAX_TRABAJOS_LEKIN

    xls = pd.ExcelFile(ruta)

    # ── Etapas ──
    df = _leer_hoja(xls, "Etapas", ["Orden", "Etapa"])
    df = df.sort_values("Orden")
    nombres = [_texto(n) for n in df["Etapa"]]
    if len(set(nombres)) != len(nombres) or "" in nombres:
        raise ValueError("Hoja 'Etapas': los nombres de etapa deben ser únicos y no vacíos")
    ETAPAS[:] = list(range(1, len(nombres) + 1))
    NOMBRE_ETAPA.clear()
    NOMBRE_ETAPA.update({s: n for s, n in zip(ETAPAS, nombres)})
    etapa_por_nombre = {n.lower(): s for s, n in NOMBRE_ETAPA.items()}

    # ── Parámetros ──
    df = _leer_hoja(xls, "Parametros", ["Parametro", "Valor"])
    par = {_texto(p).lower(): v for p, v in zip(df["Parametro"], df["Valor"])}

    def p(nombre, defecto):
        v = par.get(nombre.lower())
        return defecto if v is None or _texto(v) == "" else v

    GA_PARAMS.update({
        'poblacion'    : int(p("Poblacion", 100)),
        'generaciones' : int(p("Generaciones", 300)),
        'prob_cruce'   : float(p("Prob_cruce", 0.85)),
        'prob_mutacion': float(p("Prob_mutacion", 0.25)),
        'prob_mut_seq' : float(p("Prob_mut_secuencia", 0.30)),
        'prob_mut_prop': float(p("Prob_mut_proporciones", 0.40)),
        'prob_mut_k'   : float(p("Prob_mut_k", 0.30)),
        'tam_torneo'   : int(p("Tam_torneo", 3)),
        'elitismo'     : _si_no(p("Elitismo", "Sí")),
        'semilla'      : None if _texto(par.get("semilla")) == "" else int(float(par["semilla"])),
    })
    MAX_SPLIT = int(p("Max_split", 4))
    MIN_DOCENAS_POR_SUBLOTE = float(p("Min_docenas_sublote", 0.5))
    USAR_FACTORES = _si_no(p("Usar_factores_velocidad", "Sí"))
    ENUMERACION_MAX_N = int(p("Enumeracion_max_trabajos", 8))
    MAX_TRABAJOS_LEKIN = int(p("Max_trabajos_LEKIN", 10))
    if GA_PARAMS['poblacion'] < 2 or GA_PARAMS['tam_torneo'] < 1 or MAX_SPLIT < 1:
        raise ValueError("Hoja 'Parametros': Poblacion ≥ 2, Tam_torneo ≥ 1 y Max_split ≥ 1")
    GA_PARAMS['tam_torneo'] = min(GA_PARAMS['tam_torneo'], GA_PARAMS['poblacion'])

    # ── Máquinas ──
    df = _leer_hoja(xls, "Maquinas", ["Etapa", "Grupo", "Cantidad_maquinas", "Factor_tiempo"])
    MAQUINAS.clear()
    MAQUINAS.update({s: [] for s in ETAPAS})
    for _, f in df.iterrows():
        nombre = _texto(f["Etapa"])
        s = etapa_por_nombre.get(nombre.lower())
        if s is None:
            raise ValueError(f"Hoja 'Maquinas': la etapa '{nombre}' no existe en la hoja 'Etapas'")
        grupo = _texto(f["Grupo"]) or "G1"
        cant = _numero(f["Cantidad_maquinas"], "Maquinas", f"Cantidad_maquinas ({nombre} {grupo})")
        if cant != int(cant):
            raise ValueError(f"Hoja 'Maquinas': Cantidad_maquinas de {nombre} {grupo} debe ser entero")
        factor = _numero(f["Factor_tiempo"], "Maquinas", f"Factor_tiempo ({nombre} {grupo})",
                         estricto=True)
        for _ in range(int(cant)):
            MAQUINAS[s].append({'grupo': grupo, 'factor': factor})
    for s in ETAPAS:
        if not MAQUINAS[s]:
            raise ValueError(f"Hoja 'Maquinas': la etapa '{NOMBRE_ETAPA[s]}' no tiene máquinas")
        # Más rápidas primero: así M1, M2… son las de menor factor
        MAQUINAS[s].sort(key=lambda m: m['factor'])
        for i, m in enumerate(MAQUINAS[s], start=1):
            m['id'] = f"M{i}" if len({x['grupo'] for x in MAQUINAS[s]}) == 1 \
                else f"M{i}-{m['grupo']}"
    MAQUINAS_POR_ETAPA.clear()
    MAQUINAS_POR_ETAPA.update({s: len(MAQUINAS[s]) for s in ETAPAS})

    # ── Trabajos ──
    df = _leer_hoja(xls, "Trabajos", ["Trabajo", "Demanda_pares", "Fecha_entrega_h"])
    TRABAJOS.clear(); DESCRIPCION.clear(); DEMANDA.clear(); DUE_DATE.clear(); RELEASE.clear()
    for _, f in df.iterrows():
        j = int(_numero(f["Trabajo"], "Trabajos", "Trabajo"))
        if j in DEMANDA:
            raise ValueError(f"Hoja 'Trabajos': el trabajo {j} está repetido")
        TRABAJOS.append(j)
        DESCRIPCION[j] = _texto(f.get("Descripcion", "")) or f"Trabajo {j}"
        DEMANDA[j] = _numero(f["Demanda_pares"], "Trabajos", f"Demanda_pares (trabajo {j})",
                             estricto=True)
        DUE_DATE[j] = _numero(f["Fecha_entrega_h"], "Trabajos", f"Fecha_entrega_h (trabajo {j})")
        lib = f.get("Liberacion_h", 0)
        RELEASE[j] = 0.0 if _texto(lib) == "" else _numero(lib, "Trabajos",
                                                            f"Liberacion_h (trabajo {j})")
    if len(TRABAJOS) < 2:
        raise ValueError("Hoja 'Trabajos': se necesitan al menos 2 trabajos")

    # ── Tiempos ──
    df = _leer_hoja(xls, "Tiempos", ["Trabajo"])
    col_por_nombre = {c.lower(): c for c in df.columns}
    faltantes = [NOMBRE_ETAPA[s] for s in ETAPAS if NOMBRE_ETAPA[s].lower() not in col_por_nombre]
    if faltantes:
        raise ValueError(f"Hoja 'Tiempos': faltan las columnas de las etapas {faltantes} "
                         "(deben llamarse igual que en la hoja 'Etapas')")
    TIEMPO_MIN_DOCENA.clear()
    for _, f in df.iterrows():
        j = int(_numero(f["Trabajo"], "Tiempos", "Trabajo"))
        TIEMPO_MIN_DOCENA[j] = {
            s: _numero(f[col_por_nombre[NOMBRE_ETAPA[s].lower()]], "Tiempos",
                       f"{NOMBRE_ETAPA[s]} (trabajo {j})")
            for s in ETAPAS
        }
    sin_tiempos = [j for j in TRABAJOS if j not in TIEMPO_MIN_DOCENA]
    if sin_tiempos:
        raise ValueError(f"Hoja 'Tiempos': no hay tiempos para los trabajos {sin_tiempos}")


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
# 3.  CONVERSIÓN DE UNIDADES
# ---------------------------------------------------------------------------

def docenas(pares: float) -> float:
    """Convierte pares de guantes a docenas (1 docena = 12 pares)."""
    return pares / 12.0


def factor_maquina(s: int, m: int) -> float:
    return MAQUINAS[s][m]['factor'] if USAR_FACTORES else 1.0


def tiempo_horas(trabajo: int, etapa: int, cant_docenas: float, factor: float = 1.0) -> float:
    """Tiempo [h] = (min/docena estándar) × factor de la máquina × docenas / 60."""
    return TIEMPO_MIN_DOCENA[trabajo][etapa] * factor * cant_docenas / 60.0


def calcular_tiempos_completos() -> Dict[int, Dict[int, float]]:
    """p[trabajo][etapa] = tiempo [h] del lote completo en una máquina estándar."""
    return {j: {s: tiempo_horas(j, s, docenas(DEMANDA[j])) for s in ETAPAS} for j in TRABAJOS}

# ---------------------------------------------------------------------------
# 4.  REPRESENTACIÓN DEL CROMOSOMA
# ---------------------------------------------------------------------------

def _max_k_valido(trabajo: int, etapa: int) -> int:
    """Máximo de sublotes: limitado por máquinas, MAX_SPLIT y tamaño mínimo de sublote."""
    por_volumen = int(docenas(DEMANDA[trabajo]) / MIN_DOCENAS_POR_SUBLOTE + EPS)
    return max(1, min(MAQUINAS_POR_ETAPA[etapa], MAX_SPLIT, por_volumen))


def _dirichlet(k: int) -> List[float]:
    """k proporciones aleatorias positivas que suman 1."""
    muestras = [random.expovariate(1.0) + 1e-9 for _ in range(k)]
    total = sum(muestras)
    return [m / total for m in muestras]


def props_a_docenas(trabajo: int, props: List[float]) -> List[float]:
    """
    Convierte proporciones a docenas, garantizando que cada sublote tenga al
    menos MIN_DOCENAS_POR_SUBLOTE (reparación en la decodificación) y que la
    suma sea exactamente la demanda.
    """
    total = docenas(DEMANDA[trabajo])
    k = len(props)
    minimo = min(MIN_DOCENAS_POR_SUBLOTE, total / k)
    libre = total - k * minimo
    suma = sum(props)
    ds = [minimo + libre * p / suma for p in props]
    ds[-1] += total - sum(ds)   # corrección de redondeo flotante
    return ds


def crear_individuo(secuencia: Optional[List[int]] = None, sin_split: bool = False) -> dict:
    """
    secuencia     : orden de entrada de los trabajos a la primera etapa
    k_por_etapa   : k_por_etapa[j][s] = número de sublotes del trabajo j en la etapa s
    proporciones  : proporciones[j][s] = fracción del lote en cada sublote
    """
    if secuencia is None:
        secuencia = list(TRABAJOS)
        random.shuffle(secuencia)
    k_por_etapa, proporciones = {}, {}
    for j in TRABAJOS:
        k_por_etapa[j], proporciones[j] = {}, {}
        for s in ETAPAS:
            k = 1 if sin_split else random.randint(1, _max_k_valido(j, s))
            k_por_etapa[j][s] = k
            proporciones[j][s] = _dirichlet(k)
    return {'secuencia': list(secuencia), 'k_por_etapa': k_por_etapa,
            'proporciones': proporciones}

# ---------------------------------------------------------------------------
# 5.  SIMULADOR HFS CON LOT STREAMING POR ETAPA (decodificador del AG)
# ---------------------------------------------------------------------------

def _tiempo_disponible(curva: List[Tuple[float, float]], volumen: float) -> float:
    """
    curva: [(instante, docenas que quedan disponibles en ese instante), ...]
    Devuelve el primer instante en que el volumen acumulado disponible
    alcanza 'volumen'.
    """
    acumulado = 0.0
    for t, d in curva:
        acumulado += d
        if acumulado >= volumen - EPS:
            return t
    return curva[-1][0]


def simular_hfs(individuo: dict) -> Tuple[Dict[int, float], Dict[int, list]]:
    """
    Simula el HFS con sublotes independientes por etapa.

    - Los trabajos entran en el orden de 'secuencia' y en cada etapa sus
      sublotes se programan en ese mismo orden.
    - Un sublote de la etapa s puede empezar cuando la etapa anterior ya
      entregó todo el volumen acumulado que necesita (la etapa anterior
      entrega material al terminar cada uno de sus sublotes).
    - Cada sublote se asigna a la máquina que lo TERMINA más temprano
      (considera la disponibilidad y la velocidad de cada máquina).

    Retorna
    -------
    C     : {trabajo: tiempo de terminación en la última etapa}
    gantt : {etapa: [(trabajo, sublote, maquina_idx, inicio, fin, docenas), ...]}
    """
    maq_libre = {s: [0.0] * MAQUINAS_POR_ETAPA[s] for s in ETAPAS}
    C: Dict[int, float] = {}
    gantt: Dict[int, list] = {s: [] for s in ETAPAS}

    for j in individuo['secuencia']:
        # Al inicio todo el lote está disponible en la fecha de liberación
        curva = [(RELEASE[j], docenas(DEMANDA[j]))]

        for s in ETAPAS:
            docs = props_a_docenas(j, individuo['proporciones'][j][s])
            curva.sort()
            entregas = []
            vol_acum = 0.0
            for b, d in enumerate(docs):
                vol_acum += d
                t_listo = _tiempo_disponible(curva, vol_acum)

                mejor_m, mejor_ini, mejor_fin = None, 0.0, math.inf
                for m in range(MAQUINAS_POR_ETAPA[s]):
                    ini = max(t_listo, maq_libre[s][m])
                    fin = ini + tiempo_horas(j, s, d, factor_maquina(s, m))
                    if fin < mejor_fin - EPS:
                        mejor_m, mejor_ini, mejor_fin = m, ini, fin

                maq_libre[s][mejor_m] = mejor_fin
                entregas.append((mejor_fin, d))
                gantt[s].append((j, b + 1, mejor_m, mejor_ini, mejor_fin, d))
            curva = entregas

        C[j] = max(t for t, _ in curva)

    return C, gantt


def simular_sin_ls(secuencia: List[int]) -> Tuple[Dict[int, float], Dict[int, list]]:
    """HFS de permutación sin lot streaming (k = 1 en todas las etapas)."""
    return simular_hfs(crear_individuo(secuencia, sin_split=True))


def verificar_programa(gantt: Dict[int, list], usar_factores: Optional[bool] = None) -> List[str]:
    """
    Revisa un programa de forma independiente del simulador. Comprueba:
      1. Ninguna máquina procesa dos sublotes a la vez.
      2. En cada etapa se procesa exactamente la demanda de cada trabajo.
      3. Duración = min/docena × factor de la máquina × docenas / 60.
      4. Ningún sublote empieza antes de la liberación del trabajo ni antes
         de que la etapa anterior haya terminado el volumen que necesita.
    Retorna la lista de errores (vacía = programa factible).
    """
    usar_factores = USAR_FACTORES if usar_factores is None else usar_factores
    errores = []
    for s in ETAPAS:
        por_maquina: Dict[int, list] = {}
        for (j, b, m, ini, fin, d) in gantt[s]:
            por_maquina.setdefault(m, []).append((ini, fin, j, b))
            f = MAQUINAS[s][m]['factor'] if usar_factores else 1.0
            if abs((fin - ini) - tiempo_horas(j, s, d, f)) > 1e-6:
                errores.append(f"Duración incorrecta: T{j}-S{b} en {NOMBRE_ETAPA[s]}")
            if ini < RELEASE[j] - EPS:
                errores.append(f"T{j}-S{b} inicia antes de su liberación en {NOMBRE_ETAPA[s]}")
        for m, ops in por_maquina.items():
            ops.sort()
            for a, b in zip(ops, ops[1:]):
                if b[0] < a[1] - EPS:
                    errores.append(f"Traslape en {NOMBRE_ETAPA[s]} {MAQUINAS[s][m]['id']}: "
                                   f"T{a[2]}-S{a[3]} y T{b[2]}-S{b[3]}")
        for j in TRABAJOS:
            subs = sorted((x for x in gantt[s] if x[0] == j), key=lambda x: x[1])
            if abs(sum(x[5] for x in subs) - docenas(DEMANDA[j])) > 1e-6:
                errores.append(f"T{j}: el volumen en {NOMBRE_ETAPA[s]} no suma la demanda")
            if s == ETAPAS[0]:
                continue
            previos = [x for x in gantt[ETAPAS[ETAPAS.index(s) - 1]] if x[0] == j]
            acumulado = 0.0
            for x in subs:
                acumulado += x[5]
                listo = sum(y[5] for y in previos if y[4] <= x[3] + EPS)
                if listo < acumulado - 1e-6:
                    errores.append(f"T{j}-S{x[1]} en {NOMBRE_ETAPA[s]} inicia sin material "
                                   f"suficiente de la etapa anterior")
    return errores

# ---------------------------------------------------------------------------
# 6.  SIMULADOR DE REGLAS DE DESPACHO (sin lot streaming, como LEKIN)
# ---------------------------------------------------------------------------

REGLAS = ["EDD", "SPT", "LPT", "FIFO", "LIFO", "MS", "CR"]


def simular_despacho(regla: str,
                     trabajos: Optional[List[int]] = None,
                     p: Optional[Dict[int, Dict[int, float]]] = None,
                     due: Optional[Dict[int, float]] = None,
                     rel: Optional[Dict[int, float]] = None,
                     usar_factores: Optional[bool] = None
                     ) -> Tuple[Dict[int, float], Dict[int, list]]:
    """
    Programación no-delay por reglas de despacho, lote completo (sin LS).

    Cada vez que una máquina queda libre y hay trabajos esperando en su
    etapa, se elige el trabajo de mayor prioridad según la regla y se asigna
    a la máquina que lo termina primero. Prioridades:
      EDD  : menor fecha de entrega
      SPT  : menor tiempo de la operación en la etapa
      LPT  : mayor tiempo de la operación en la etapa
      FIFO : el que llegó primero a la cola de la etapa
      LIFO : el que llegó último a la cola de la etapa
      MS   : menor holgura  (d_j − t − trabajo restante)
      CR   : menor razón crítica (d_j − t) / trabajo restante
    Empates → menor número de trabajo (LIFO: mayor).

    Los argumentos opcionales permiten usar otros datos (p. ej. los minutos
    enteros de LEKIN). p está en horas en una máquina de factor 1.
    """
    trabajos = list(trabajos or TRABAJOS)
    p = p or calcular_tiempos_completos()
    due = due or DUE_DATE
    rel = rel or RELEASE
    usar_factores = USAR_FACTORES if usar_factores is None else usar_factores
    fac = {s: [MAQUINAS[s][m]['factor'] if usar_factores else 1.0
               for m in range(MAQUINAS_POR_ETAPA[s])] for s in ETAPAS}

    orden = {j: i for i, j in enumerate(trabajos)}
    maq_libre = {s: [0.0] * MAQUINAS_POR_ETAPA[s] for s in ETAPAS}
    idx_etapa = {j: 0 for j in trabajos}            # índice de la próxima etapa
    listo = {j: rel[j] for j in trabajos}           # llegada a la cola actual
    C: Dict[int, float] = {}
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
            t = max(listo[j], min(maq_libre[s]))
            if t < t_min - EPS or (abs(t - t_min) <= EPS and s < s_min):
                t_min, s_min = t, s
        s = s_min
        cola = [j for j in pendientes
                if ETAPAS[idx_etapa[j]] == s and listo[j] <= t_min + EPS]
        j = min(cola, key=lambda x: prioridad(x, s, t_min))
        # Máquina que TERMINA primero la operación (con máquinas idénticas es
        # la que está libre en t_min; con velocidades distintas puede convenir
        # esperar un poco a una máquina rápida)
        m = min(range(MAQUINAS_POR_ETAPA[s]),
                key=lambda x: (max(t_min, maq_libre[s][x]) + p[j][s] * fac[s][x], x))
        ini = max(t_min, maq_libre[s][m])
        fin = ini + p[j][s] * fac[s][m]
        maq_libre[s][m] = fin
        gantt[s].append((j, 1, m, ini, fin, docenas(DEMANDA[j])))

        idx_etapa[j] += 1
        listo[j] = fin
        if idx_etapa[j] == len(ETAPAS):
            C[j] = fin
            pendientes.remove(j)

    return C, gantt

# ---------------------------------------------------------------------------
# 7.  FUNCIÓN OBJETIVO E INDICADORES
# ---------------------------------------------------------------------------

def calcular_indicadores(C: Dict[int, float],
                         due: Optional[Dict[int, float]] = None,
                         rel: Optional[Dict[int, float]] = None) -> Dict:
    """Indicadores estándar de scheduling para los trabajos de C."""
    due = due or DUE_DATE
    rel = rel or RELEASE
    trabajos = list(C)
    ind_j = {}
    for j in trabajos:
        ci, di, ri = C[j], due[j], rel[j]
        li = ci - di
        ti = max(0.0, li)
        ind_j[j] = {'Ci': ci, 'ri': ri, 'di': di, 'Fi': ci - ri, 'Li': li, 'Ti': ti,
                    'Ei': max(0.0, -li), 'delta_i': 1 if ti > EPS else 0}
    n = len(trabajos)
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
# 8.  OPERADORES GENÉTICOS
# ---------------------------------------------------------------------------

def seleccion_torneo(poblacion: List[dict], tam: int) -> dict:
    return min(random.sample(poblacion, tam), key=fitness)


def cruce_ox(seq1: List[int], seq2: List[int]) -> List[int]:
    """Cruce de orden OX: conserva una permutación válida."""
    n = len(seq1)
    p1, p2 = sorted(random.sample(range(n), 2))
    hijo = [None] * n
    hijo[p1:p2 + 1] = seq1[p1:p2 + 1]
    usados = set(hijo[p1:p2 + 1])
    pos = (p2 + 1) % n
    for gen in seq2[p2 + 1:] + seq2[:p2 + 1]:
        if gen not in usados:
            hijo[pos] = gen
            usados.add(gen)
            pos = (pos + 1) % n
    return hijo


def _ajustar_props(ref: List[float], otro: List[float], k: int) -> List[float]:
    """k proporciones obtenidas al promediar 'ref' y 'otro' adaptados a longitud k."""
    def adaptar(lst):
        r = lst[:k] if len(lst) >= k else lst + [lst[-1]] * (k - len(lst))
        t = sum(r)
        return [x / t for x in r] if t > 0 else _dirichlet(k)

    mezcla = [(a + b) / 2.0 for a, b in zip(adaptar(ref), adaptar(otro))]
    t = sum(mezcla)
    return [x / t for x in mezcla] if t > 0 else _dirichlet(k)


def cruce(p1: dict, p2: dict) -> Tuple[dict, dict]:
    """Cruce de secuencia (OX) y cruce uniforme de k y proporciones por (j, s)."""
    h1, h2 = copy.deepcopy(p1), copy.deepcopy(p2)
    h1.pop('_fitness', None)
    h2.pop('_fitness', None)

    h1['secuencia'] = cruce_ox(p1['secuencia'], p2['secuencia'])
    h2['secuencia'] = cruce_ox(p2['secuencia'], p1['secuencia'])

    for j in TRABAJOS:
        for s in ETAPAS:
            if random.random() < 0.5:
                h1['k_por_etapa'][j][s] = p2['k_por_etapa'][j][s]
                h2['k_por_etapa'][j][s] = p1['k_por_etapa'][j][s]
            k1, k2 = h1['k_por_etapa'][j][s], h2['k_por_etapa'][j][s]
            h1['proporciones'][j][s] = _ajustar_props(
                p1['proporciones'][j][s], p2['proporciones'][j][s], k1)
            h2['proporciones'][j][s] = _ajustar_props(
                p2['proporciones'][j][s], p1['proporciones'][j][s], k2)
    return h1, h2


def mutar(individuo: dict) -> dict:
    """Mutación de secuencia (swap), proporciones (gaussiana) y k (±1)."""
    ind = copy.deepcopy(individuo)
    ind.pop('_fitness', None)

    if random.random() < GA_PARAMS['prob_mut_seq']:
        a, b = random.sample(range(len(ind['secuencia'])), 2)
        ind['secuencia'][a], ind['secuencia'][b] = ind['secuencia'][b], ind['secuencia'][a]

    for j in TRABAJOS:
        for s in ETAPAS:
            if random.random() < GA_PARAMS['prob_mut_prop']:
                props = ind['proporciones'][j][s][:]
                idx = random.randrange(len(props))
                props[idx] = max(1e-9, props[idx] + random.gauss(0, 0.15))
                t = sum(props)
                ind['proporciones'][j][s] = [x / t for x in props]

            if random.random() < GA_PARAMS['prob_mut_k']:
                k_actual = ind['k_por_etapa'][j][s]
                k_nuevo = max(1, min(_max_k_valido(j, s), k_actual + random.choice([-1, 1])))
                if k_nuevo != k_actual:
                    ind['k_por_etapa'][j][s] = k_nuevo
                    ind['proporciones'][j][s] = _ajustar_props(
                        ind['proporciones'][j][s], _dirichlet(k_nuevo), k_nuevo)
    return ind

# ---------------------------------------------------------------------------
# 9.  ALGORITMO GENÉTICO PRINCIPAL
# ---------------------------------------------------------------------------

def algoritmo_genetico(semillas: Optional[List[List[int]]] = None,
                       verbose: bool = True) -> Tuple[dict, List[float]]:
    """
    Ejecuta el AG y retorna (mejor_individuo, historial del mejor ΣTi).

    semillas: secuencias que se insertan en la población inicial sin lot
    streaming (p. ej. las de las reglas de despacho). Así el AG nunca
    termina peor que esas secuencias.
    """
    if GA_PARAMS['semilla'] is not None:
        random.seed(GA_PARAMS['semilla'])

    poblacion = [crear_individuo(seq, sin_split=True) for seq in (semillas or [])]
    poblacion = poblacion[:GA_PARAMS['poblacion']]
    while len(poblacion) < GA_PARAMS['poblacion']:
        poblacion.append(crear_individuo())
    for ind in poblacion:
        fitness(ind)

    mejor_global = copy.deepcopy(min(poblacion, key=fitness))
    historial = []

    if verbose:
        print(f"\n{'=' * 72}")
        print("  ALGORITMO GENÉTICO — HFS CON LOT STREAMING POR ETAPA")
        print(f"  Población: {GA_PARAMS['poblacion']}  |  "
              f"Generaciones: {GA_PARAMS['generaciones']}  |  MAX_SPLIT: {MAX_SPLIT}")
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
            p1 = seleccion_torneo(poblacion, GA_PARAMS['tam_torneo'])
            p2 = seleccion_torneo(poblacion, GA_PARAMS['tam_torneo'])
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
# 10. REFERENCIAS: REGLAS DE DESPACHO Y ENUMERACIÓN EXACTA
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
# 11. VALIDACIÓN CON LEKIN
# ---------------------------------------------------------------------------

def datos_lekin() -> Tuple[List[int], pd.DataFrame, Dict]:
    """
    Datos para cargar en LEKIN (Flexible Flow Shop), en MINUTOS ENTEROS:
    tiempo de procesamiento del lote completo en cada etapa (workcenter),
    fecha de liberación, fecha de entrega y peso = 1.
    """
    trabajos = TRABAJOS[:MAX_TRABAJOS_LEKIN]
    p_min = {j: {s: int(round(TIEMPO_MIN_DOCENA[j][s] * docenas(DEMANDA[j]))) for s in ETAPAS}
             for j in trabajos}
    due_min = {j: int(round(DUE_DATE[j] * 60)) for j in trabajos}
    rel_min = {j: int(round(RELEASE[j] * 60)) for j in trabajos}
    filas = []
    for j in trabajos:
        fila = {'Job': f"J{j:02d}", 'Descripcion': DESCRIPCION[j],
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
    p_h = {j: {s: d['p'][j][s] / 60.0 for s in ETAPAS} for j in trabajos}
    filas = []
    for regla in REGLAS:
        C_h, _ = simular_despacho(regla, trabajos, p=p_h,
                                  due={j: d['due'][j] / 60.0 for j in trabajos},
                                  rel={j: d['rel'][j] / 60.0 for j in trabajos},
                                  usar_factores=False)
        C = {j: round(C_h[j] * 60) for j in trabajos}
        ind = calcular_indicadores(C, due=d['due'], rel=d['rel'])
        filas.append({'Regla': regla, 'Cmax': ind['Cmax'], 'Lmax': ind['Lmax'],
                      'ΣUj (tardíos)': ind['num_tardy'], 'ΣCj': ind['suma_Ci'],
                      'ΣTj': ind['suma_Ti'], 'ΣwjCj': ind['suma_Ci'], 'ΣwjTj': ind['suma_Ti']})
    return tabla, pd.DataFrame(filas)

# ---------------------------------------------------------------------------
# 12. REPORTES EN CONSOLA
# ---------------------------------------------------------------------------

SEP = "=" * 76
SEP2 = "-" * 76


def imprimir_datos():
    print(SEP)
    print("  DATOS LEÍDOS DEL EXCEL")
    print(SEP)
    for s in ETAPAS:
        grupos = {}
        for m in MAQUINAS[s]:
            grupos.setdefault(m['grupo'], [0, m['factor']])[0] += 1
        detalle = ", ".join(f"{g}: {n} maq × factor {f:.2f}" for g, (n, f) in grupos.items())
        print(f"  Etapa {s} {NOMBRE_ETAPA[s]:<12} → {detalle}")
    if not USAR_FACTORES:
        print("  (Usar_factores_velocidad = No → todas las máquinas se tratan como idénticas)")
    print()
    print(SEP)
    print("  TIEMPOS DE PROCESAMIENTO — LOTE COMPLETO EN MÁQUINA ESTÁNDAR [horas]")
    print(SEP)
    enc = " | ".join(f"{NOMBRE_ETAPA[s][:10]:>10}" for s in ETAPAS)
    print(f"  {'Trab':>4} | {'Docenas':>8} | {'di(h)':>6} | {enc} | {'Total':>8}")
    print(SEP2)
    p = calcular_tiempos_completos()
    for j in TRABAJOS:
        e = " | ".join(f"{p[j][s]:>10.3f}" for s in ETAPAS)
        print(f"  {j:>4} | {docenas(DEMANDA[j]):>8.2f} | {DUE_DATE[j]:>6.1f} | {e} | "
              f"{sum(p[j].values()):>8.3f}")
    print()


def imprimir_reglas(resultados: List[dict], exacto: Optional[dict]):
    print(SEP)
    print("  REGLAS DE DESPACHO — SIN LOT STREAMING (referencia)")
    print(SEP)
    print(f"  {'Regla':>6} | {'Orden en etapa 1':>24} | {'ΣTi':>9} | {'Tmax':>9} | "
          f"{'Cmax':>9} | {'#tard':>5}")
    print(SEP2)
    for r in sorted(resultados, key=lambda x: x['suma_Ti']):
        print(f"  {r['regla']:>6} | {str(r['secuencia']):>24} | {r['suma_Ti']:>9.4f} | "
              f"{r['Tmax']:>9.4f} | {r['Cmax']:>9.4f} | {r['num_tardy']:>5}")
    if exacto:
        print(SEP2)
        print(f"  {'ÓPTIMO':>6} | {str(exacto['secuencia']):>24} | {exacto['suma_Ti']:>9.4f} | "
              f"{exacto['Tmax']:>9.4f} | {exacto['Cmax']:>9.4f} | {exacto['num_tardy']:>5}")
        print(f"  (óptimo exacto de permutación sin LS: {exacto['evaluadas']} secuencias evaluadas)")
    print()


def tabla_sublotes(individuo: dict) -> pd.DataFrame:
    filas = []
    for j in TRABAJOS:
        for s in ETAPAS:
            docs = props_a_docenas(j, individuo['proporciones'][j][s])
            for b, d in enumerate(docs, start=1):
                filas.append({'Trabajo': j, 'Etapa': NOMBRE_ETAPA[s],
                              'Maquinas_etapa': MAQUINAS_POR_ETAPA[s],
                              'k_sublotes': len(docs), 'Sublote': b,
                              'Docenas': round(d, 3), 'Pares': round(d * 12, 1)})
    return pd.DataFrame(filas)


def tabla_programa(gantt: Dict[int, list]) -> pd.DataFrame:
    filas = []
    for s in ETAPAS:
        for (j, b, m, ini, fin, d) in sorted(gantt[s], key=lambda x: (x[3], x[2])):
            filas.append({'Etapa': NOMBRE_ETAPA[s], 'Maquina': MAQUINAS[s][m]['id'],
                          'Grupo': MAQUINAS[s][m]['grupo'],
                          'Factor': factor_maquina(s, m), 'Trabajo': j, 'Sublote': b,
                          'Docenas': round(d, 3), 'Inicio_h': round(ini, 4),
                          'Fin_h': round(fin, 4), 'Duracion_h': round(fin - ini, 4)})
    return pd.DataFrame(filas)


def tabla_indicadores(ind: dict) -> pd.DataFrame:
    return pd.DataFrame([{'Trabajo': j, 'Descripcion': DESCRIPCION[j], **{
        k: (round(v, 4) if isinstance(v, float) else v) for k, v in vals.items()}}
        for j, vals in ind['por_trabajo'].items()])


def imprimir_solucion(titulo: str, individuo: dict):
    C, gantt = simular_hfs(individuo)
    ind = calcular_indicadores(C)

    print(SEP)
    print(f"  {titulo}")
    print(f"  Secuencia: {individuo['secuencia']}")
    print("\n  SUBLOTES ELEGIDOS POR EL AG (k independiente por etapa):")
    print("  " + tabla_sublotes(individuo).to_string(index=False).replace("\n", "\n  "))

    print(f"\n  {'Tj':>4} | {'di(h)':>7} | {'Ci(h)':>8} | {'Fi(h)':>8} | {'Li(h)':>8} | "
          f"{'Ti(h)':>8} | {'Ei(h)':>8} | {'δi':>3}")
    print("  " + "-" * 70)
    for j in TRABAJOS:
        v = ind['por_trabajo'][j]
        print(f"  {j:>4} | {v['di']:>7.1f} | {v['Ci']:>8.4f} | {v['Fi']:>8.4f} | "
              f"{v['Li']:>8.4f} | {v['Ti']:>8.4f} | {v['Ei']:>8.4f} | {v['delta_i']:>3}")

    print("\n  INDICADORES GLOBALES")
    print(f"    Cmax      (makespan)              = {ind['Cmax']:>10.4f} h")
    print(f"    Lmax      (retraso máximo)        = {ind['Lmax']:>10.4f} h")
    print(f"    Tmax      (tardanza máxima)       = {ind['Tmax']:>10.4f} h")
    print(f"    ΣTi       (tardanza total) [FO]   = {ind['suma_Ti']:>10.4f} h  ← FO")
    print(f"    ΣEi       (adelanto total)        = {ind['suma_Ei']:>10.4f} h")
    print(f"    ΣFi       (flujo total)           = {ind['suma_Fi']:>10.4f} h")
    print(f"    F̄         (flujo medio)           = {ind['media_Fi']:>10.4f} h")
    print(f"    # tardíos (Σδi)                   = {ind['num_tardy']:>10}")

    print("\n  PROGRAMA DETALLADO")
    print("  " + tabla_programa(gantt).to_string(index=False).replace("\n", "\n  "))
    print()


def imprimir_utilizacion(individuo: dict):
    print(SEP)
    print("  USO DE MÁQUINAS POR ETAPA (solución AG)")
    print(SEP)
    print(f"  {'Etapa':>12} | {'Disponibles':>11} | {'Máx sublotes/trabajo':>20} | "
          f"{'Prom sublotes/trabajo':>21}")
    print(SEP2)
    for s in ETAPAS:
        k_vals = [individuo['k_por_etapa'][j][s] for j in TRABAJOS]
        print(f"  {NOMBRE_ETAPA[s]:>12} | {MAQUINAS_POR_ETAPA[s]:>11} | {max(k_vals):>20} | "
              f"{sum(k_vals) / len(k_vals):>21.2f}")
    print()


def imprimir_comparacion(ind_sin: dict, ind_con: dict, nombre_sin: str):
    print(SEP)
    print(f"  COMPARACIÓN: SIN LOT STREAMING ({nombre_sin}) vs. CON LOT STREAMING (AG)")
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
    print("  " + tabla.drop(columns=['Descripcion']).to_string(index=False).replace("\n", "\n  "))
    print("\n  Resultados que deben coincidir con LEKIN (mismos datos y regla) [min]:")
    print("  " + res.to_string(index=False).replace("\n", "\n  "))
    print()

# ---------------------------------------------------------------------------
# 13. DIAGRAMA DE GANTT
# ---------------------------------------------------------------------------

def graficar_gantt(individuo: dict,
                   titulo: str = "Diagrama de Gantt — HFS con Lot Streaming",
                   guardar_como: Optional[str] = None):
    C, gantt = simular_hfs(individuo)
    ind = calcular_indicadores(C)
    cmax = ind['Cmax']

    colores_base = ['#4E79A7', '#F28E2B', '#59A14F', '#E15759', '#B07AA1',
                    '#76B7B2', '#EDC948', '#FF9DA7', '#9C755F', '#BAB0AC']
    color_trabajo = {j: colores_base[i % len(colores_base)] for i, j in enumerate(TRABAJOS)}

    alturas = [MAQUINAS_POR_ETAPA[s] for s in ETAPAS]
    fig, axes = plt.subplots(len(ETAPAS), 1,
                             figsize=(20, max(8, sum(alturas) * 0.45 + 3)),
                             gridspec_kw={'height_ratios': alturas}, sharex=True)
    if len(ETAPAS) == 1:
        axes = [axes]
    fig.suptitle(titulo, fontsize=14, fontweight='bold', y=1.0)
    umbral_texto = (cmax if cmax > 0 else 1) * 0.015

    for ax, s in zip(axes, ETAPAS):
        n_maq = MAQUINAS_POR_ETAPA[s]
        for (j, b, m, ini, fin, _d) in gantt[s]:
            dur = fin - ini
            ax.barh(m, dur, left=ini, height=0.72, color=color_trabajo[j],
                    edgecolor='white', linewidth=0.6, alpha=0.92)
            if dur >= umbral_texto:
                ax.text(ini + dur / 2, m, f"T{j}-S{b}", ha='center', va='center',
                        fontsize=6.5, fontweight='bold', color='white', clip_on=True)
        ax.axvline(cmax, color='#333333', linewidth=1.2, linestyle='--', alpha=0.7)
        ax.set_yticks(range(n_maq))
        ax.set_yticklabels([MAQUINAS[s][m]['id'] for m in range(n_maq)], fontsize=8)
        ax.set_ylim(n_maq - 0.4, -0.6)   # M1 arriba
        ax.set_ylabel(f"Etapa {s}\n{NOMBRE_ETAPA[s]}", fontsize=9, fontweight='bold', labelpad=8)
        ax.xaxis.set_minor_locator(ticker.AutoMinorLocator(4))
        ax.grid(axis='x', which='major', linestyle='--', linewidth=0.5, alpha=0.5, color='#bbbbbb')
        ax.grid(axis='x', which='minor', linestyle=':', linewidth=0.3, alpha=0.35, color='#cccccc')
        ax.set_facecolor('#F8F9FA')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

    axes[-1].set_xlabel("Tiempo (horas)", fontsize=10)
    axes[-1].xaxis.set_major_locator(ticker.MaxNLocator(nbins=15))
    axes[-1].annotate(f"Cmax = {cmax:.2f} h", xy=(cmax, 1.0), xycoords=('data', 'axes fraction'),
                      xytext=(-4, -4), textcoords='offset points', fontsize=8, ha='right',
                      va='top', color='#333333')

    fig.legend(handles=[mpatches.Patch(color=color_trabajo[j], label=f"Trabajo {j}")
                        for j in TRABAJOS],
               loc='lower center', ncol=min(len(TRABAJOS), 10), fontsize=9, frameon=True,
               framealpha=0.9, edgecolor='#cccccc', bbox_to_anchor=(0.5, -0.03))
    fig.text(0.5, 1.002, f"Secuencia: {individuo['secuencia']}  |  ΣTi = {ind['suma_Ti']:.4f} h"
                         f"  |  Cmax = {cmax:.4f} h  |  # tardíos = {ind['num_tardy']}",
             ha='center', va='bottom', fontsize=8.5, color='#444444', transform=fig.transFigure)
    plt.tight_layout(rect=[0, 0.04, 1, 0.99])

    # Se guarda ANTES de plt.show(): en algunos entornos show() vacía la figura
    if guardar_como:
        fig.savefig(guardar_como, dpi=150, bbox_inches='tight')
        print(f"  Gantt guardado en: {guardar_como}")
    plt.show()
    return fig

# ---------------------------------------------------------------------------
# 14. EXPORTAR RESULTADOS A EXCEL
# ---------------------------------------------------------------------------

def exportar_resultados(ruta: str, resumen: List[dict], reglas: List[dict],
                        exacto: Optional[dict], mejor: dict, historial: List[float],
                        lekin_tabla: pd.DataFrame, lekin_res: pd.DataFrame):
    C, gantt = simular_hfs(mejor)
    ind = calcular_indicadores(C)
    df_reglas = pd.DataFrame([{
        'Regla': r['regla'], 'Orden_etapa1': str(r['secuencia']),
        'ΣTi_h': round(r['suma_Ti'], 4), 'Tmax_h': round(r['Tmax'], 4),
        'Cmax_h': round(r['Cmax'], 4), 'Tardios': r['num_tardy']} for r in reglas])
    if exacto:
        df_reglas = pd.concat([df_reglas, pd.DataFrame([{
            'Regla': 'ÓPTIMO permutación sin LS', 'Orden_etapa1': str(exacto['secuencia']),
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
# 15. PROGRAMA PRINCIPAL
# ---------------------------------------------------------------------------

def main(ruta_excel: str = RUTA_EXCEL, mostrar_gantt: bool = True):
    cargar_datos(obtener_ruta_excel(ruta_excel))

    print()
    print(SEP)
    print("  HFS CON LOT STREAMING POR ETAPA — GUANTES INDUSTRIALES")
    print("  Sublotes independientes por etapa | AG | Mín. tardanza total")
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

    # ── AG con lot streaming (población inicial sembrada con las reglas) ──
    semillas = [r['secuencia'] for r in reglas]
    if exacto:
        semillas.append(exacto['secuencia'])
    mejor, historial = algoritmo_genetico(semillas=semillas, verbose=True)
    C_con, _ = simular_hfs(mejor)
    ind_con = calcular_indicadores(C_con)

    imprimir_solucion("MEJOR SOLUCIÓN — AG CON LOT STREAMING POR ETAPA", mejor)
    errores = verificar_programa(simular_hfs(mejor)[1])
    for r in reglas:
        errores += verificar_programa(simular_despacho(r['regla'])[1])
    if errores:
        print("  ✗ SE ENCONTRARON ERRORES DE FACTIBILIDAD:")
        for e in errores:
            print(f"    - {e}")
    else:
        print("  ✓ Verificación: los programas del AG y de las reglas son factibles "
              "(sin traslapes, volúmenes completos y precedencias respetadas)\n")
    imprimir_utilizacion(mejor)
    imprimir_comparacion(mejor_regla, ind_con, f"regla {mejor_regla['regla']}")

    print("  Generando diagrama de Gantt...")
    if mostrar_gantt:
        graficar_gantt(mejor, "Diagrama de Gantt — HFS con Lot Streaming por Etapa (AG)",
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
    resumen.append({'Metodo': 'AG con lot streaming', 'ΣTi_h': ind_con['suma_Ti'],
                    'Cmax_h': ind_con['Cmax'], 'Tmax_h': ind_con['Tmax'],
                    'Tardios': ind_con['num_tardy']})
    resumen.append({'Metodo': f"Secuencia AG: {mejor['secuencia']}"})
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
