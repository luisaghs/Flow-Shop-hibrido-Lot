# HFS con lot streaming (sublotes consistentes) para guantes industriales

Programación de un *Hybrid Flow Shop* (Troquelado → Ensamble → Costura → Volteado)
con *lot streaming* de **sublotes consistentes** y algoritmo genético, minimizando la
tardanza total ΣTj (desempate: Cmax).

| Archivo | Contenido |
|---|---|
| `BD_HFS.xlsx` | Base de datos del problema (ver abajo). El código no trae datos escritos. |
| `hfs_guantes.py` | Modelo: lectura del Excel, simulador, AG, reglas de despacho, validación LEKIN, Gantt y exportación. |

## Base de datos `BD_HFS.xlsx`

| Hoja | Columnas | Uso |
|---|---|---|
| `Tiempo de procesamiento` | Codigo producto, Maquina, Etapa, min x docena, h x docena | Tiempo de cada producto en cada máquina. Las etapas y máquinas salen de aquí; el orden de las etapas sigue la numeración de las máquinas (M1–M4, M5–M24, …). Si falta la fila de un producto en una máquina, esa máquina no puede procesarlo. |
| `Trabajos` | Orden de trabajo, Codigo Producto, Cantidad (pares), Cantidad (docenas), Fecha de liberación, Fechas de entrega, Horas disponibles | Una fila por orden. `Horas disponibles` es la fecha de entrega en horas laborables (si está vacía se calcula con el calendario). |
| `Disponibilidad` | Fecha, Dia semana, Horas disponibles, Horas acumuladas | Calendario laboral: convierte fechas ↔ horas laborables (liberación y fechas de terminación). |

Para otro escenario basta con editar el Excel: agregar órdenes, cambiar cantidades,
fechas, tiempos o máquinas. Los parámetros del método (tamaño de población,
generaciones, `MAX_SUBLOTES`, `MIN_PARES_SUBLOTE`) están al inicio de `hfs_guantes.py`.

## Sublotes consistentes (escenario S2)

Se sigue la estrategia 2 / escenario S2 de Gürsoy Yılmaz y Yılmaz (2022), *Lot streaming
in hybrid flowshop scheduling problem by considering equal and consistent sublots…*,
Computers & Industrial Engineering 173, 108745:

- Cada orden se divide en B_j sublotes (1 ≤ B_j ≤ `MAX_SUBLOTES`); los tamaños pueden ser
  distintos entre sí, pero **son los mismos en todas las etapas**.
- Un sublote pasa a la etapa siguiente apenas termina, aunque el resto del lote siga en
  proceso.
- Cromosoma: tamaños de sublote (uno por trabajo, comunes a todas las etapas), posiciones
  de los sublotes (aleatorias en la población inicial) y número de sublotes.
- Decodificación: la etapa 1 sigue el orden del cromosoma; en las siguientes los sublotes
  se atienden por orden de llegada; cada sublote va a la máquina capaz que lo termina
  primero (las máquinas de Ensamble tienen tiempos distintos).
- Los tamaños son números enteros de pares, con un mínimo de `MIN_PARES_SUBLOTE`.

Salidas: `resultados_programacion.xlsx` (resumen, reglas, indicadores con fechas de
terminación, sublotes, programa, convergencia, datos y resultados esperados de LEKIN)
y `gantt.png`.

## Ejecutar en Google Colab

**Opción A — subir los archivos (rápido)**

1. Abra <https://colab.research.google.com> → *Nuevo notebook*.
2. Panel izquierdo → ícono de carpeta 📁 → *Subir* → suba `hfs_guantes.py` y `BD_HFS.xlsx`.
3. En una celda:
   ```python
   import hfs_guantes
   mejor = hfs_guantes.main("BD_HFS.xlsx")
   ```
   Si el Excel no está, el programa abre un selector para subirlo.

Los archivos subidos así se borran cuando se cierra la sesión de Colab.

**Opción B — Google Drive (persistente)**

1. Cree en Drive la carpeta `HFS_guantes` y ponga ahí los dos archivos.
2. En Colab:
   ```python
   from google.colab import drive
   drive.mount('/content/drive')
   %cd /content/drive/MyDrive/HFS_guantes
   import hfs_guantes
   mejor = hfs_guantes.main("BD_HFS.xlsx")
   ```

Si modifica `hfs_guantes.py` después de importarlo, reinicie el entorno
(*Entorno de ejecución → Reiniciar*) o use `import importlib; importlib.reload(hfs_guantes)`.

## Validación con LEKIN

LEKIN no hace lot streaming, admite máximo 10 trabajos y usa tiempos enteros. Por eso:

1. La hoja `LEKIN_datos` trae, por trabajo y etapa, el tiempo del lote completo en minutos
   enteros con el **promedio** de min/docena de las máquinas de la etapa (en LEKIN las
   máquinas de un *workcenter* se cargan como idénticas).
2. Cargue en LEKIN (*Flexible Flow Shop*) un *workcenter* por etapa con su número de
   máquinas y esos trabajos (release y due en minutos, peso 1).
3. Ejecute cada regla en LEKIN (EDD, SPT, LPT, FCFS, MS, CR…) y compare ΣTj, Cmax, Lmax, ΣUj
   con la hoja `LEKIN_esperado`. Si coinciden, el simulador está validado.
   Pueden aparecer diferencias pequeñas por desempates entre trabajos con igual prioridad.
4. Compare el AG con sublotes consistentes (hoja `Resumen`) contra la mejor regla.

Además, el programa verifica cada solución por separado (sin traslapes, volúmenes
completos, mismo tamaño de cada sublote en todas las etapas y precedencia de cada sublote
entre etapas) y, con ≤ 8 trabajos, calcula el óptimo exacto de permutación sin lot
streaming por enumeración completa.
