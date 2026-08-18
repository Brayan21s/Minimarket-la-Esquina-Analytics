# Pipeline de Power Query

Este documento resume la arquitectura de las consultas de Power Query construidas
para este proyecto, y sirve como referencia del código M final. **El código de tu
workbook es la fuente de verdad** — este archivo documenta la lógica y el porqué
de cada decisión, no reemplaza el Editor Avanzado de tu Excel.

## Arquitectura general

```
Carpeta dataset_minimarket/
  (244 archivos CSV: 61 días × 4 tiendas)
        │
        ▼
┌─────────────────────┐
│ Import_POS_Diario    │  Power Query "Desde carpeta" + Combinar archivos
│ (solo consolida)      │  Tipos de dato corregidos (Cantidad/Precio: errores→null)
└─────────────────────┘
        │  (referencia, no duplica datos)
        ▼
┌─────────────────────┐
│ Staging_Limpio        │  Limpieza, recuperación y validación
└─────────────────────┘
        │  Cargada como tabla en hoja Staging
        ▼
Tablas Dinámicas + Segmentadores → Dashboard
```

**Por qué dos consultas separadas en vez de una sola:** `Import_POS_Diario` solo
sabe "traer datos de una carpeta" — es la pieza que cualquier proyecto similar
reutilizaría tal cual. `Staging_Limpio` contiene toda la lógica de negocio
específica de este dataset (qué tiendas son válidas, cómo recuperar un precio
faltante, etc.). Separarlas hace que cada una se pueda probar y modificar de
forma independiente.

## Reglas de limpieza y recuperación

| Campo | Problema | Estrategia |
|---|---|---|
| Tienda | Espacios/mayúsculas inconsistentes | `Text.Trim` + `Text.Proper` |
| Tienda | Typos reales (ej. "Nrote") | Corrección explícita por valor conocido, o Fuzzy Merge contra la lista oficial |
| Producto | Espacios/mayúsculas inconsistentes | `Text.Trim` + `Text.Clean` + `Text.Proper` |
| Cantidad | Negativa | Valor absoluto (`Number.Abs`) |
| Cantidad | Vacía | `Total ÷ Precio_Unitario` (Total siempre es confiable en este dataset) |
| Precio_Unitario | Vacío | `Total ÷ Cantidad` (ya corregida) |
| Fecha | Fuera del rango operativo (jun-jul 2026) | Se descarta — no se puede inferir la fecha real |
| ID_Transaccion | Duplicado | `Table.Distinct` — se conserva la primera aparición |

**Principio de diseño:** un error solo se "resuelve" cuando hay otra columna que
permite inferirlo matemáticamente sin ambigüedad (ej. Cantidad y Precio se
recuperan mutuamente porque `Total` es un ancla confiable). Cuando no existe esa
ancla (fecha corrupta, fila duplicada), el dato se excluye explícitamente en vez
de adivinar — así el pipeline nunca inventa información.

## Columnas derivadas para análisis

- `Hora_Exacta`: hora del día sin minutos/segundos (`Time.Hour`), para el
  histograma de horas pico.
- `Semanas`: número de semana relativo al inicio del dataset.
- `Semana_Inicio` / `Semana_Fin` / `Semana_Etiqueta`: fechas de inicio/fin de
  cada semana y una etiqueta legible tipo "Semana 1 (1-7 jun)", usada como eje
  X en el gráfico de tendencia semanal.

## Panel de calidad de datos

```
Filas cargadas (Import_POS_Diario)   =COUNTA(Import_POS_Diario[ID_Transaccion])
Filas finales (Staging_Limpio)        =COUNTA(Staging_Limpio[ID_Transaccion])
% de datos aprovechados                =Filas_finales / Filas_cargadas
```

Este indicador no mide "ventas perdidas" — mide la salud del proceso de
ingesta: qué proporción de los datos crudos el sistema pudo usar de forma
confiable, sea porque llegaron limpios o porque se lograron recuperar.

## Notas de mantenimiento

- Si vuelves a generar el dataset con `data/generar_dataset.py` y cambias el
  rango de fechas, actualiza el paso `Fecha_Valida` en `Staging_Limpio` para
  que coincida con el nuevo rango operativo.
- El reemplazo de typos de tienda (`Table.ReplaceValue`) está escrito contra
  los valores específicos que genera el script actual. Si regeneras el dataset
  con una semilla distinta o agregas nuevas variantes de error, esos typos
  puntuales no se van a atrapar automáticamente — considera migrar a Fuzzy
  Merge si esto se vuelve recurrente.
