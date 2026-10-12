# HFS con lot streaming para guantes industriales

Programación de un *Hybrid Flow Shop* (Troquelado → Ensamblaje → Costura → Voteado)
con *lot streaming* por etapa y algoritmo genético, minimizando la tardanza total ΣTj.

| Archivo | Contenido |
|---|---|
| `datos_guantes.xlsx` | Base de datos (Etapas, Maquinas, Trabajos, Tiempos, Parametros). Las celdas amarillas son valores supuestos. |
| `hfs_guantes.py` | Modelo: lectura del Excel, simulador, AG, reglas de despacho, validación LEKIN, Gantt y exportación. |
| `crear_base_datos.py` | Regenera la plantilla Excel desde cero (opcional). |

Salidas: `resultados_programacion.xlsx` (resumen, reglas, sublotes, programa, convergencia,
datos y resultados esperados de LEKIN) y `gantt.png`.

## Ejecutar en Google Colab

**Opción A — subir los archivos (rápido)**

1. Abra <https://colab.research.google.com> → *Nuevo notebook*.
2. Panel izquierdo → ícono de carpeta 📁 → *Subir* → suba `hfs_guantes.py` y `datos_guantes.xlsx`.
3. En una celda:
   ```python
   import hfs_guantes
   mejor = hfs_guantes.main("datos_guantes.xlsx")
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
   mejor = hfs_guantes.main("datos_guantes.xlsx")
   ```
   Los resultados quedan guardados en la misma carpeta de Drive. Puede editar el Excel en
   Drive (abrir con Google Sheets y *Archivo → Descargar → .xlsx* o editar con Excel) y
   volver a ejecutar.

Si modifica `hfs_guantes.py` después de importarlo, reinicie el entorno
(*Entorno de ejecución → Reiniciar*) o use `import importlib; importlib.reload(hfs_guantes)`.

## Validación con LEKIN

LEKIN no hace lot streaming, admite máximo 10 trabajos y usa tiempos enteros. Por eso:

1. Ponga `Usar_factores_velocidad = No` si su versión de LEKIN no maneja velocidades
   distintas por máquina (así ambos modelos son idénticos).
2. Cargue en LEKIN (*Flexible Flow Shop*) un *workcenter* por etapa con su número de
   máquinas y los trabajos de la hoja `LEKIN_datos` (tiempos, release y due en minutos, peso 1).
3. Ejecute cada regla en LEKIN (EDD, SPT, LPT, FCFS, MS, CR…) y compare ΣTj, Cmax, Lmax, ΣUj
   con la hoja `LEKIN_esperado`. Si coinciden, el simulador está validado.
   Pueden aparecer diferencias pequeñas por desempates entre trabajos con igual prioridad.
4. Compare el AG con lot streaming (hoja `Resumen`) contra la mejor regla de LEKIN.

Además, el programa verifica cada solución por separado (sin traslapes, volúmenes
completos, precedencias de material entre etapas) y, con ≤ 8 trabajos, calcula el óptimo
exacto de permutación sin lot streaming por enumeración completa.
