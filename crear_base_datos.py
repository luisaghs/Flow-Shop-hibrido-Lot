"""
Genera la plantilla Excel `datos_guantes.xlsx` con los datos del problema
HFS de guantes industriales.

Uso (local o Colab):
    python crear_base_datos.py

Solo es necesario ejecutarlo si se quiere regenerar la plantilla desde cero;
para el uso normal basta con editar `datos_guantes.xlsx` directamente.
"""

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

RUTA_SALIDA = "datos_guantes.xlsx"

AZUL = PatternFill("solid", fgColor="1F4E78")
AMARILLO = PatternFill("solid", fgColor="FFF2CC")   # celdas con datos supuestos
GRIS = PatternFill("solid", fgColor="F2F2F2")
BLANCO_NEGRITA = Font(bold=True, color="FFFFFF")
BORDE = Border(*(Side(style="thin", color="BFBFBF"),) * 4)

# ---------------------------------------------------------------------------
# Datos
# ---------------------------------------------------------------------------

ETAPAS = [
    # Orden, Etapa, Descripción
    (1, "Troquelado", "Corte de piezas de carnaza / vaqueta"),
    (2, "Ensamblaje", "Máquinas NO idénticas: 3 grupos por velocidad"),
    (3, "Costura", "Costura de cierre"),
    (4, "Voteado", "Volteo del guante (etapa nueva)"),
]

# Etapa, Grupo, Descripción, Cantidad de máquinas, Factor de tiempo
# Factor = (min/docena del grupo) / (min/docena estándar de la hoja Tiempos).
# Factor < 1 → grupo más rápido; factor > 1 → grupo más lento.
MAQUINAS = [
    ("Troquelado", "G1", "Troqueladoras", 4, 1.00),
    ("Ensamblaje", "G1", "Grupo rápido", 3, 0.80),
    ("Ensamblaje", "G2", "Grupo estándar", 4, 1.00),
    ("Ensamblaje", "G3", "Grupo lento", 3, 1.25),
    ("Costura", "G1", "Máquinas de costura", 4, 1.00),
    ("Voteado", "G1", "Volteadoras", 8, 1.00),
]
# Filas de MAQUINAS cuyo dato es supuesto (se resaltan en amarillo)
MAQUINAS_SUPUESTAS = {1, 2, 3, 5}

# Trabajo, Descripción, Material, Refuerzo, Demanda [pares],
# Fecha de entrega [h], Liberación [h]
TRABAJOS = [
    (1, "Guante carnaza sin refuerzo", "Carnaza", "No", 1272, 200, 0),
    (2, "Guante carnaza con refuerzo", "Carnaza", "Sí", 628, 150, 0),
    (3, "Guante vaqueta sin refuerzo", "Vaqueta", "No", 629, 120, 0),
    (4, "Guante vaqueta con refuerzo", "Vaqueta", "Sí", 660, 160, 0),
    (5, "Guante carnaza con refuerzo", "Carnaza", "Sí", 423, 180, 0),
]

# Tiempos estándar [min/docena]: Troquelado, Ensamblaje, Costura, Voteado
TIEMPOS = {
    1: (30, 50, 50, 12),
    2: (25, 45, 50, 14),
    3: (33, 40, 46, 13),
    4: (28, 44, 52, 15),
    5: (36, 55, 52, 14),
}
COLUMNA_SUPUESTA_TIEMPOS = "Voteado"

PARAMETROS = [
    ("Poblacion", 100, "Tamaño de la población del AG"),
    ("Generaciones", 300, "Número de generaciones"),
    ("Prob_cruce", 0.85, "Probabilidad de cruce"),
    ("Prob_mutacion", 0.25, "Probabilidad de mutar a cada hijo"),
    ("Prob_mut_secuencia", 0.30, "Prob. de intercambiar dos trabajos de la secuencia"),
    ("Prob_mut_proporciones", 0.40, "Prob. de perturbar las proporciones de un (trabajo, etapa)"),
    ("Prob_mut_k", 0.30, "Prob. de cambiar ±1 el número de sublotes de un (trabajo, etapa)"),
    ("Tam_torneo", 3, "Tamaño del torneo de selección"),
    ("Elitismo", "Sí", "Conservar el mejor individuo entre generaciones (Sí/No)"),
    ("Semilla", 42, "Semilla aleatoria; dejar vacío para resultados distintos en cada corrida"),
    ("Max_split", 4, "Máximo de sublotes por (trabajo, etapa); nunca supera las máquinas de la etapa"),
    ("Min_docenas_sublote", 0.5, "Tamaño mínimo de un sublote [docenas]"),
    ("Usar_factores_velocidad", "Sí",
     "Sí = máquinas de Ensamblaje con velocidades distintas. "
     "No = todas idénticas (modelo equivalente al de LEKIN)"),
    ("Enumeracion_max_trabajos", 8,
     "Si hay ≤ este número de trabajos se enumeran todas las secuencias sin "
     "lot streaming (óptimo exacto de referencia)"),
    ("Max_trabajos_LEKIN", 10, "Límite de trabajos de LEKIN para la hoja de validación"),
]

INSTRUCCIONES = [
    "BASE DE DATOS — HFS CON LOT STREAMING PARA GUANTES INDUSTRIALES",
    "",
    "Cómo llenar el archivo:",
    "1. Etapas: una fila por etapa, en el orden en que la recorren los guantes (columna Orden).",
    "   El nombre de la etapa debe coincidir EXACTAMENTE con el encabezado de la hoja Tiempos",
    "   y con la columna Etapa de la hoja Maquinas.",
    "2. Maquinas: una fila por grupo de máquinas. Las etapas con máquinas idénticas tienen un",
    "   solo grupo con Factor_tiempo = 1. Ensamblaje tiene 3 grupos según su velocidad.",
    "   Factor_tiempo = (min/docena en ese grupo) / (min/docena estándar de la hoja Tiempos).",
    "   Ej.: si el estándar es 50 min/doc y el grupo rápido tarda 40 min/doc → factor 0.80.",
    "3. Trabajos: una fila por pedido. Demanda en PARES, fecha de entrega y liberación en HORAS",
    "   contadas desde el inicio del horizonte de programación (hora 0).",
    "4. Tiempos: min/docena estándar de cada trabajo en cada etapa.",
    "5. Parametros: parámetros del algoritmo genético y del lot streaming.",
    "",
    "Las celdas en AMARILLO son valores supuestos: reemplácelos por los datos reales de la planta",
    "(tiempos de Voteado y factores de velocidad de los grupos de Ensamblaje).",
    "",
    "No cambie el nombre de las hojas ni de los encabezados (fila 1).",
    "Puede agregar o quitar trabajos y grupos de máquinas libremente (sin dejar filas vacías en medio).",
]


# ---------------------------------------------------------------------------
# Utilidades de formato
# ---------------------------------------------------------------------------

def _hoja(wb, nombre, encabezados, filas, anchos):
    ws = wb.create_sheet(nombre)
    ws.append(encabezados)
    for fila in filas:
        ws.append(list(fila))
    for c, ancho in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(c)].width = ancho
        celda = ws.cell(row=1, column=c)
        celda.fill = AZUL
        celda.font = BLANCO_NEGRITA
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for fila in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=len(encabezados)):
        for celda in fila:
            celda.border = BORDE
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 30
    return ws


def crear_plantilla(ruta: str = RUTA_SALIDA) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "Instrucciones"
    for linea in INSTRUCCIONES:
        ws.append([linea])
    ws["A1"].font = Font(bold=True, size=14, color="1F4E78")
    ws["A16"].fill = AMARILLO
    ws.column_dimensions["A"].width = 100

    _hoja(wb, "Etapas", ["Orden", "Etapa", "Descripcion"], ETAPAS, [8, 16, 50])

    ws = _hoja(wb, "Maquinas",
               ["Etapa", "Grupo", "Descripcion", "Cantidad_maquinas", "Factor_tiempo"],
               MAQUINAS, [16, 8, 24, 18, 15])
    for i in MAQUINAS_SUPUESTAS:
        ws.cell(row=i + 2, column=5).fill = AMARILLO
    ws["E1"].comment = Comment(
        "Factor = tiempo del grupo / tiempo estándar.\n<1 más rápido, >1 más lento.", "Plantilla")

    ws = _hoja(wb, "Trabajos",
               ["Trabajo", "Descripcion", "Material", "Refuerzo", "Demanda_pares",
                "Fecha_entrega_h", "Liberacion_h"],
               TRABAJOS, [9, 32, 12, 10, 15, 16, 13])
    validacion = DataValidation(type="list", formula1='"Carnaza,Vaqueta"', allow_blank=True)
    ws.add_data_validation(validacion)
    validacion.add("C2:C200")
    validacion = DataValidation(type="list", formula1='"Sí,No"', allow_blank=True)
    ws.add_data_validation(validacion)
    validacion.add("D2:D200")

    nombres = [e[1] for e in ETAPAS]
    ws = _hoja(wb, "Tiempos", ["Trabajo"] + nombres,
               [(j, *t) for j, t in TIEMPOS.items()], [9] + [14] * len(nombres))
    col = nombres.index(COLUMNA_SUPUESTA_TIEMPOS) + 2
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=col).fill = AMARILLO

    ws = _hoja(wb, "Parametros", ["Parametro", "Valor", "Descripcion"], PARAMETROS, [26, 10, 95])
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=1).fill = GRIS

    wb.save(ruta)
    return ruta


if __name__ == "__main__":
    print(f"Plantilla creada: {crear_plantilla()}")
