# Analizador de Asistencia y Turnos (demo 6)

**El problema real:** el cierre de nómina de asistencia toma horas: cruzar el reloj con los turnos a mano, sumar tardanzas en una calculadora y rezar que no se escape nadie. Un empleado que llegó 7 minutos tarde en un turno con 5 de tolerancia son 2 minutos de descuento; uno que se quedó 40 minutos después de su salida son 40 de extra, no 55. Con 30 empleados y 20 días hábiles, ese cruce manual son 600 filas miradas una por una.

**La solución:** este programa cruza el reloj del mes con los turnos esperados y deja el cierre listo en segundos: `resumen_asistencia_2026-11.xlsx`, con el detalle auditable de cada tardanza, cada extra y cada día bajo dotación. Nadie suma a mano; el respaldo queda en el mismo archivo.

## Qué hace

Genera un Excel con 4 hojas:

1. **Resumen por Persona** — título con el periodo ("Resumen de Asistencia — Noviembre 2026") y, por empleado: turno, días registrados, llegadas tarde, minutos de tardanza acumulados, salidas antes, HH extra (horas) del mes, ausencias y olvidos de marcaje. La columna de minutos de tardanza lleva escala de color (blanco a rojo) y la de HH extra otra (blanco a verde). Cierra con fila de TOTALES y una nota metodológica al pie con las reglas de tolerancia y la gracia de 15 minutos para las extras.
2. **Tardanzas del Mes** — una fila por tardanza efectiva: empleado, fecha, turno, hora real, entrada ideal y minutos sobre la tolerancia. Ordenada por empleado y fecha. Solo tardanzas mayores que cero.
3. **Extras** — horas extra y días con extra por persona, ordenado de mayor a menor: quien más se quedó aparece arriba. Cierra con TOTALES.
4. **Cobertura** — los 20 días hábiles con su dotación presente, el mínimo exigido y estado `OK` o `DEFICIT`. Los días con déficit quedan al inicio y pintados de rojo; el resto, de verde.

## Reglas de cálculo

- **Tardanza:** `max(0, entrada_real - entrada_ideal - tolerancia)`. La tolerancia es por turno (Mañana 5 min; Día y Tarde 10 min). Llegar dentro de la tolerancia no genera descuento.
- **Horas extra:** solo cuentan si la salida supera la salida ideal por más de 15 minutos de gracia. Pasada la gracia se abona el total de la diferencia (no se descuentan los 15).
- **Salida antes:** se cuenta cuando el empleado sale más de la tolerancia del turno antes de su hora ideal. No genera minutos extra negativos.
- **Ausencia:** no genera tardanza ni extra.
- **Olvido de marcaje:** se cuenta aparte; no es tardanza ni ausencia, y para cobertura cuenta como presente (estuvo, solo no marcó).
- **Turnos que cruzan medianoche:** `minutos_trabajados` suma 24 horas cuando la salida es menor o igual que la entrada.
- **Cobertura:** un día queda en `DEFICIT` si los presentes (asistencias más olvidos) son menos que `--minimo`.

## Cómo ejecutarlo

```bash
# con los datos de ejemplo del directorio
python3 asistencia.py

# con tus propios archivos y otro mínimo de dotación
python3 asistencia.py --reloj reloj.csv --turnos turnos.csv \
    --mes 2026-11 --salida ./entrega --minimo 27

# tests
python3 -m pytest test_asistencia.py -v
```

- `--reloj`: CSV con columnas `empleado, fecha, turno, estado, hora_entrada, hora_salida`.
- `--turnos`: CSV con columnas `turno, hora_entrada_ideal, hora_salida_ideal, tolerancia_min`.
- `--mes`: mes a analizar en formato `AAAA-MM` (por defecto, `2026-11`). Si el mes no tiene registros, termina con error claro y código 2.
- `--salida`: directorio donde guardar el XLSX (por defecto, `.`).
- `--minimo`: dotación mínima diaria para marcar déficit (por defecto, 25).

Las horas se leen en formato `HH:MM` y las fechas en `AAAA-MM-DD`. Una fila malformada detiene el proceso con un `ValueError` que indica el número de fila y el valor culpable.

## Datos de ejemplo

- `reloj_asistencia.csv`: 600 registros ficticios de 30 empleados (`EMP-001` a `EMP-030`) en los 20 días hábiles de noviembre de 2026. Turnos fijos: Mañana (EMP-001 a EMP-010), Día (EMP-011 a EMP-022) y Tarde (EMP-023 a EMP-030).
- `turnos_esperados.csv`: los tres turnos con sus horarios ideales y tolerancias.
- `generar_datos.py`: genera ambos CSV de forma determinista (semilla fija). Los casos están declarados: 18 tardanzas, 6 salidas anticipadas, 10 empleados con extra, 19 ausencias (seis el 13 y seis el 24 de noviembre para provocar déficit) y 4 olvidos de marcaje.
- `capturas/`: capturas a 150 dpi del resumen generado (una por hoja).

Ningún dato corresponde a una persona real.

## Tests

21 casos de pytest en 11 funciones (3,3 s): tardanza con tolerancia y casos borde, extras con y sin gracia, jornada que cruza medianoche, ausencia detectada, olvido contado aparte y presente en cobertura, columna faltante, hora inválida con fila y valor, totales que cuadran entre resumen y detalle, formato del XLSX (4 hojas, encabezado azul, escala de color, congelado, autofiltro) y cobertura con déficit los días 13 y 24.

## Stack

Python 3 · pandas · openpyxl (formato condicional, anchos fijos, paneles congelados) · pytest.
