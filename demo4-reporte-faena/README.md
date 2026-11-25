# Generador de Reporte Diario de Faena (demo 4 · flagship)

**El problema real:** en toda faena chilena, el jefe de faena va dictando la noche por WhatsApp — "cuadrilla 3 no salio, colacion", "permiso PT-4471 aprobado para izaje", "casi incidente: herramienta cayo 3 m, sin lesionados", "avance PLANTA-NORTE 82%". Cada mañana, alguien tiene que transformar ese desorden de mensajes en el reporte formal que el cliente exige a las 07:00: **una o dos horas tecleando**, copiando HH a mano, buscando qué permiso se aprobó y cuál quedó esperando firma, y rezando que no se escape un casi incidente (que después lo cobra el cliente o la supervisión).

**La solución:** este programa lee el mismo WhatsApp del jefe (texto libre, con sus errores de tipeo: *izo*, *calamon*, *colacion*), lo cruza con el Excel de horas-hombre de las cuadrillas y deja el reporte formal listo: `reporte_faena_2026-11-23.xlsx` + PDF con formato de industria, listo para enviar al cliente. **Nadie teclea nada; el reporte sale solo.**

## Qué hace

- **Parsea el texto libre** del jefe de faena con reglas robustas (tildes y mayúsculas opcionales):
  - Incidentes y **casi incidentes** con clasificación automática de severidad (Alta / Media / Baja).
  - **Permisos de trabajo** (PT-####) con estado: Aprobado, Rechazado, Pendiente.
  - **Avance por frente** (PLANTA-NORTE, PLANTA-SUR, TUNEL-1, PATIO-1).
  - **Cuadrillas ausentes** con motivo ("no salio, colacion").
  - Turno del reporte (noche/día).
- **Cruza con el Excel de HH** (`horas_cuadrillas.xlsx`): HH normales, extras y dotación por cuadrilla y por frente, filtrando por la fecha del reporte.
- **Genera** `reporte_faena_<fecha>.xlsx` con 4 hojas y exportación **PDF** (A4 horizontal, una hoja por página):
  1. **Resumen del Día** — encabezado con faena "PROYECTO CORDILLERA 2026", fecha y turno; avance por frente; HH normales, extras, totales y dotación.
  2. **Incidentes y Casi Incidentes** — hora, tipo, severidad con código de color y descripción.
  3. **Permisos de Trabajo** — PT, hora, estado con código de color y detalle.
  4. **Detalle de Cuadrillas** — cuadrilla, frente, líder, trabajadores, HH y observaciones (ej: cuadrilla 3 "No salió (colacion)").

## Por qué este demo es distinto

Ningún vendedor de automatización genérico conoce la operación real de una faena: cuadrillas, frentes, permisos de trabajo, casi incidentes que hay que reportar sí o sí, HH extras que el cliente audita. Esto no es una tabla de "Producto A / Producto B": es el flujo real de reportabilidad diaria de una faena chilena, con el lenguaje y los errores de tipeo reales del terreno.

## Cómo ejecutarlo

```bash
# correr el demo con los datos de ejemplo
python3 reporte_faena.py --fecha 2026-11-23 --salida .

# con tus propios archivos
python3 reporte_faena.py --reporte notas_jefe.txt --horas horas.xlsx \
    --fecha 2026-11-23 --salida ./entrega --turno Noche

# tests
python3 -m pytest test_reporte_faena.py -v
```

- `--reporte`: txt con las notas del jefe (estilo export de WhatsApp).
- `--horas`: Excel con columnas `cuadrilla, frente, lider, HH_normales, HH_extras, n_trabajadores, fecha`.
- `--fecha`: fecha del reporte (AAAA-MM-DD); valida el formato y falla con mensaje claro si es inválida.
- `--salida`: directorio donde se guardan XLSX y PDF.
- `--turno`: opcional; si se omite se detecta del texto ("turno noche").
- `--clima` y `--estado-faena`: opcionales para el encabezado del documento controlado (ej. `--clima "Nublado, lluvia en la madrugada" --estado-faena "Operativa"`).

El reporte sale con formato de **documento controlado**: banner a todo el ancho, bloque de identificación (Mandante, Contratista, N° de Reporte RDF-AAAA-NNN, código de documento RDF-01 Rev. 2, clima y estado de faena) y pie con tres líneas de firma (Elaborado por / Revisado por / Aprobado por).

## Datos de ejemplo

- `reporte_jefe.txt`: notas reales de tres días (21–23 nov 2026) estilo WhatsApp de faena, con los errores de tipeo habituales del terreno.
- `horas_cuadrillas.xlsx`: HH de 6 cuadrillas en los frentes PLANTA-NORTE, PLANTA-SUR, TUNEL-1 y PATIO-1 por 3 días.
- `capturas/`: capturas a 150 dpi del reporte generado en PDF.

## Tests

8 tests de pytest (2,5 s): parseo de frases clave, clasificación de severidad, suma de HH, fecha inválida, día sin incidentes no rompe, generación y existencia del PDF, y detección de turno.

## Stack

Python 3 · pandas · openpyxl (formato condicional, anchos fijos, PDF vía LibreOffice headless) · pytest.
