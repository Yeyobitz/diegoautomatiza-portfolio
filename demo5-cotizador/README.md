# Generador de cotizaciones PDF con folio e IVA (demo 5)

**El problema real:** si vendes servicios (ingeniería, montaje, mantención, lo que sea), cada cotización es la misma tortura: abrir la cotización anterior, borrar los ítems del cliente anterior, revisar que no quede ningún dato heredado, recalcular el IVA a mano en la calculadora, rezar que el total cuadre y formatear la tabla para que se vea presentable. **20 minutos por cotización** para un documento donde un cero de más o un cliente equivocado te cuesta la credibilidad — y el folio, si lo llevas a mano en una planilla, se salta o se repite tarde o temprano.

**La solución:** este programa arma la cotización en **30 segundos**: le pasas el CSV de ítems y el cliente, y sale `cotizacion_COT-2026-0001.xlsx` + PDF con tu marca, folio consecutivo que se lleva solo, IVA 19% calculado sin errores y validez de la oferta. Copiar de la anterior quedó en el pasado — y el folio nunca se repite.

## Qué hace

- **Lee un CSV de ítems** (`descripcion, unidad, cantidad, precio_unitario`) con validación estricta y mensajes claros: columnas faltantes, CSV malformado, cantidad negativa o no numérica, precio negativo — nada entra al documento.
- **Calcula en pesos, con redondeo decidido** (ROUND_HALF_UP, no el redondeo bancario de Python): importe por línea, **SUBTOTAL**, **IVA (19%)** y **TOTAL**, en formato chileno `$#,##0`.
- **Folio consecutivo** `COT-2026-0001`, `COT-2026-0002`…: escanea las cotizaciones ya emitidas en el directorio de salida y continúa la numeración, aunque las corridas sean de días distintos.
- **Documento presentable, no tabla cruda**: banner de marca con RUT y contacto, bloque "Señores / RUT / Atención", tabla con encabezado en azul y líneas alternadas, totales destacados, **validez de la oferta (15 días por defecto, con fecha de vencimiento explícita)**, notas comerciales y **dos líneas de firma** (empresa y aceptación del cliente).
- **Salida Excel + PDF** (vía LibreOffice headless), A4 con ajuste a una página de ancho.

## Cómo ejecutarlo

```bash
# correr el demo con los datos de ejemplo
python3 cotizador.py --fecha 2026-11-25

# con tus propios ítems, cliente y validez
python3 cotizador.py --items mis_items.csv --cliente "Constructora Los Andes Ltda" \
    --rut "77.987.654-3" --contacto "Sr(a). Encargado(a) de Compras" \
    --validez 15 --salida ./entrega

# con tu propia marca (para revender esto tal cual)
python3 cotizador.py --items mis_items.csv --empresa "Mi Empresa SpA" \
    --rut-empresa "78.543.210-9" --contacto-empresa "Nombre · Cargo · email · fono"

# tests
python3 -m pytest test_cotizador.py -v
```

- `--items`: CSV con `descripcion, unidad, cantidad, precio_unitario` (por defecto `items.csv` junto al script).
- `--cliente`, `--rut`, `--contacto`: datos del cliente.
- `--empresa`, `--rut-empresa`, `--contacto-empresa`: tu marca en el encabezado.
- `--validez`: días de validez de la oferta (por defecto 15; aparece en el documento con la fecha de vencimiento).
- `--fecha`: fecha de emisión AAAA-MM-DD (por defecto hoy).
- `--salida`: directorio donde se guardan XLSX y PDF (y donde se continúa el folio).
- `--pdf`: `auto` (sí si LibreOffice está disponible), `si` o `no`.

## Datos de ejemplo

- `items.csv`: 6 ítems típicos de un servicio de montaje (horas-hombre, ml de barandas, viajes de retiro, grúa) — todo ficticio.
- `cotizacion_COT-2026-0001.xlsx` / `.pdf`: la salida generada (SUBTOTAL $7.648.000 · IVA $1.453.120 · TOTAL $9.101.120).
- `capturas/`: capturas a 150 dpi del PDF.

## Tests

15 tests de pytest (≈4,5 s): folio consecutivo entre corridas y continuación desde archivo previo, formato de folio, IVA 19% exacto, redondeo half-up (56,5 → 57, no el bancario de Python), total cuadra con la suma (en Python y en las celdas del xlsx), CSV malformado y columnas faltantes rechazados con mensaje claro, cantidad negativa/cero-numérica y precio negativo rechazados, PDF existe y no está vacío (con firma `%PDF-`), plantilla de marca presente (empresa, RUT, folio, relleno del encabezado), validez de 15 días con vencimiento calculado y líneas de firma.

## Stack

Python 3 · pandas (CSV y validación) · openpyxl (formato de celdas, impresión A4) · LibreOffice headless para el PDF · pytest.
