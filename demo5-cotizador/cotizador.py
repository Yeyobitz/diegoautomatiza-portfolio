"""Generador de cotizaciones PDF con folio e IVA.

Lee un CSV de ítems (descripcion, unidad, cantidad, precio_unitario) y arma
una cotización formal en Excel y PDF: encabezado de marca, folio consecutivo
COT-AAAA-NNNN, fecha, datos del cliente, tabla de ítems, SUBTOTAL, IVA 19%,
TOTAL, validez de la oferta (15 días por defecto), notas y línea de firma.

Uso::

    python cotizador.py --items items.csv \
        --cliente "Constructora Los Andes Ltda" --salida .

Sin argumentos usa ``items.csv`` junto a este script y los datos de ejemplo.
El folio continúa la numeración de las cotizaciones previas que existan en el
directorio de salida (cotizacion_COT-AAAA-NNNN.xlsx / .pdf).
"""

import argparse
import re
import shutil
import subprocess
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

BASE_DIR = Path(__file__).resolve().parent

# Marca ficticia del demo (empresa de ejemplo, datos inventados).
EMPRESA = "Ingeniería Austral SpA"
RUT_EMPRESA = "78.543.210-9"
CONTACTO_EMPRESA = "Marcelo Ibáñez · Gerente Comercial · cotizaciones@austral.cl · +56 9 6543 2109"
DIRECCION_EMPRESA = "Av. Providencia 1234, Of. 602, Santiago"

# Cliente de ejemplo (ficticio).
CLIENTE_DEFAULT = "Constructora Los Andes Ltda"
RUT_CLIENTE_DEFAULT = "77.987.654-3"
CONTACTO_CLIENTE_DEFAULT = "Sr(a). Encargado(a) de Compras"

COLUMNAS = ["descripcion", "unidad", "cantidad", "precio_unitario"]
IVA_TASA = Decimal("0.19")
VALIDEZ_DEFAULT = 15
FORMATO_MONEDA = "$#,##0"

NOTAS_DEFAULT = [
    "Los valores netos están expresados en pesos chilenos; el IVA (19%) se indica por separado.",
    "El plazo de ejecución se confirma al momento de recibir la orden de compra.",
    "No incluye arriendo de equipos mayores ni obras civiles mayores, salvo pacto expreso.",
    "Forma de pago: 50% contra orden de compra y 50% contra entrega.",
]

AZUL_OSCURO = "1F4E78"
RELLENO_MARCA = PatternFill("solid", fgColor=AZUL_OSCURO)
RELLENO_TOTAL = PatternFill("solid", fgColor="D9E2F3")
RELLENO_ZEBRA = PatternFill("solid", fgColor="F2F2F2")
BORDE_FINO = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
BORDE_TOTAL = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="medium"),
    bottom=Side(style="double"),
)
CENTRADO = Alignment(horizontal="center", vertical="center", wrap_text=True)
DERECHA = Alignment(horizontal="right", vertical="center")
IZQUIERDA = Alignment(horizontal="left", vertical="center", wrap_text=True)

RE_FOLIO = re.compile(r"COT-(\d{4})-(\d+)")
MESES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]


def redondear(valor):
    """Redondea a peso entero con ROUND_HALF_UP (decidido, no del float)."""
    return int(Decimal(str(valor)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def fecha_larga(fecha):
    """Devuelve la fecha como '25 de noviembre de 2026' sin depender del locale."""
    return f"{fecha.day} de {MESES[fecha.month - 1]} de {fecha.year}"


def _a_fecha(valor):
    """Convierte date, datetime o texto AAAA-MM-DD a date, con error claro."""
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    try:
        return date.fromisoformat(str(valor).strip())
    except ValueError as error:
        raise SystemExit(
            f"Fecha inválida: {valor!r}. Usa el formato AAAA-MM-DD (ej. 2026-11-25)."
        ) from error


def leer_items(ruta):
    """Lee el CSV de ítems; falla con mensaje claro si no es un CSV válido."""
    ruta = Path(ruta)
    if not ruta.is_file():
        raise SystemExit(
            f"El archivo de ítems no existe: '{ruta}'. "
            "Indica una ruta válida con --items."
        )
    try:
        return pd.read_csv(ruta, encoding="utf-8-sig")
    except pd.errors.EmptyDataError as error:
        raise SystemExit(
            f"El archivo '{ruta}' está vacío. Agrega el encabezado "
            f"({', '.join(COLUMNAS)}) y al menos una fila de ítems."
        ) from error
    except pd.errors.ParserError as error:
        raise SystemExit(
            f"El archivo '{ruta}' es un CSV malformado (filas con más o menos "
            f"columnas que el encabezado). Corrígelo e inténtalo de nuevo. "
            f"Detalle: {error}"
        ) from error
    except Exception as error:
        raise SystemExit(
            f"No se pudo leer '{ruta}' como CSV (separador coma). "
            f"Detalle: {error}"
        ) from error


def validar_items(df, ruta):
    """Valida las filas del CSV y devuelve la lista de ítems corregida."""
    faltantes = [columna for columna in COLUMNAS if columna not in df.columns]
    if faltantes:
        raise SystemExit(
            f"El archivo '{Path(ruta).name}' no tiene todas las columnas "
            f"requeridas. Faltan: {', '.join(faltantes)}. "
            f"Columnas requeridas: {', '.join(COLUMNAS)}."
        )
    if df.empty:
        raise SystemExit(
            f"El archivo '{Path(ruta).name}' no contiene filas de ítems. "
            "Agrega al menos una fila con descripcion, unidad, cantidad y "
            "precio_unitario."
        )

    items = []
    for numero, (_, fila) in enumerate(df.iterrows(), start=2):
        descripcion = fila["descripcion"]
        if pd.isna(descripcion) or not str(descripcion).strip():
            raise SystemExit(
                f"El archivo '{Path(ruta).name}' tiene 'descripcion' vacía en la "
                f"fila {numero}. Completa ese dato e inténtalo de nuevo."
            )
        unidad = fila["unidad"]
        if pd.isna(unidad) or not str(unidad).strip():
            raise SystemExit(
                f"El archivo '{Path(ruta).name}' tiene 'unidad' vacía en la fila "
                f"{numero} (ej.: hh, ml, viaje, día)."
            )

        cantidad = _a_numero(fila["cantidad"], "cantidad", numero, ruta)
        if cantidad <= 0:
            raise SystemExit(
                f"El archivo '{Path(ruta).name}' tiene una cantidad inválida en "
                f"la fila {numero}: {fila['cantidad']!r}. La cantidad debe ser "
                "mayor que cero."
            )

        precio = _a_numero(fila["precio_unitario"], "precio_unitario", numero, ruta)
        if precio < 0:
            raise SystemExit(
                f"El archivo '{Path(ruta).name}' tiene un precio_unitario negativo "
                f"en la fila {numero}: {precio}. Usa cero o un valor positivo."
            )

        items.append(
            {
                "descripcion": str(descripcion).strip(),
                "unidad": str(unidad).strip(),
                "cantidad": cantidad,
                "precio_unitario": precio,
            }
        )
    return items


def _a_numero(valor, columna, numero, ruta):
    """Convierte un valor CSV a float con mensaje de error claro."""
    if pd.isna(valor):
        raise SystemExit(
            f"El archivo '{Path(ruta).name}' tiene '{columna}' vacía en la fila "
            f"{numero}. Usa un número (ej.: 160)."
        )
    try:
        return float(str(valor).strip().replace(",", "."))
    except ValueError as error:
        raise SystemExit(
            f"El archivo '{Path(ruta).name}' tiene un valor no numérico en la "
            f"columna '{columna}', fila {numero}: {valor!r}. Usa un número "
            "(ej.: 160)."
        ) from error


def calcular_totales(items):
    """Calcula importe por línea, SUBTOTAL, IVA 19% y TOTAL (en pesos)."""
    lineas = []
    for item in items:
        importe = redondear(item["cantidad"] * item["precio_unitario"])
        lineas.append({**item, "importe": importe})
    subtotal = sum(linea["importe"] for linea in lineas)
    iva = redondear(Decimal(subtotal) * IVA_TASA)
    return {"lineas": lineas, "subtotal": subtotal, "iva": iva, "total": subtotal + iva}


def siguiente_folio(directorio, año=None):
    """Devuelve el próximo folio COT-AAAA-NNNN según lo ya emitido ahí."""
    año = año if año is not None else date.today().year
    maximo = 0
    directorio = Path(directorio)
    if directorio.is_dir():
        for archivo in directorio.iterdir():
            coincidencia = RE_FOLIO.search(archivo.name)
            if coincidencia and int(coincidencia.group(1)) == año:
                maximo = max(maximo, int(coincidencia.group(2)))
    return f"COT-{año}-{maximo + 1:04d}"


def _escribir_bloque(hoja, celda, texto, *, fill, font, alignment):
    hoja[celda] = texto
    hoja[celda].fill = fill
    hoja[celda].font = font
    hoja[celda].alignment = alignment


def construir_cotizacion(
    totales,
    folio,
    fecha,
    salida,
    validez_dias=VALIDEZ_DEFAULT,
    cliente=CLIENTE_DEFAULT,
    rut_cliente=RUT_CLIENTE_DEFAULT,
    contacto_cliente=CONTACTO_CLIENTE_DEFAULT,
    empresa=EMPRESA,
    rut_empresa=RUT_EMPRESA,
    contacto_empresa=CONTACTO_EMPRESA,
    direccion_empresa=DIRECCION_EMPRESA,
    notas=None,
):
    """Escribe la cotización formateada en xlsx y devuelve la ruta."""
    vencimiento = fecha + timedelta(days=validez_dias)
    notas = NOTAS_DEFAULT if notas is None else notas

    libro = Workbook()
    hoja = libro.active
    hoja.title = "Cotización"

    # --- Encabezado de marca + bloque de folio -----------------------------
    marca = (
        f"{empresa.upper()}\n"
        "Servicios de ingeniería, montaje y mantención\n"
        f"RUT: {rut_empresa}\n"
        f"{contacto_empresa}\n"
        f"{direccion_empresa}"
    )
    hoja.merge_cells("A1:D1")
    _escribir_bloque(
        hoja, "A1", marca.splitlines()[0],
        fill=RELLENO_MARCA, font=Font(name="Calibri", size=16, bold=True, color="FFFFFF"),
        alignment=Alignment(horizontal="left", vertical="center"),
    )
    hoja.merge_cells("A2:D4")
    _escribir_bloque(
        hoja, "A2", "\n".join(marca.splitlines()[1:]),
        fill=RELLENO_MARCA, font=Font(name="Calibri", size=10, color="FFFFFF"),
        alignment=Alignment(horizontal="left", vertical="top", wrap_text=True),
    )
    hoja.merge_cells("E1:F1")
    _escribir_bloque(
        hoja, "E1", "COTIZACIÓN",
        fill=RELLENO_MARCA, font=Font(name="Calibri", size=16, bold=True, color="FFFFFF"),
        alignment=CENTRADO,
    )
    hoja.merge_cells("E2:F2")
    _escribir_bloque(
        hoja, "E2", f"N° {folio}",
        fill=RELLENO_MARCA, font=Font(name="Calibri", size=12, bold=True, color="FFFFFF"),
        alignment=CENTRADO,
    )
    hoja.merge_cells("E3:F4")
    _escribir_bloque(
        hoja, "E3",
        f"Fecha de emisión:\n{fecha_larga(fecha)}",
        fill=RELLENO_MARCA, font=Font(name="Calibri", size=10, color="FFFFFF"),
        alignment=CENTRADO,
    )
    for fila, alto in ((1, 24), (2, 14), (3, 14), (4, 14)):
        hoja.row_dimensions[fila].height = alto

    # --- Datos del cliente --------------------------------------------------
    etiquetas = [("Señores:", cliente), ("RUT:", rut_cliente), ("Atención:", contacto_cliente)]
    for desplazamiento, (etiqueta, valor) in enumerate(etiquetas):
        fila = 6 + desplazamiento
        hoja[f"B{fila}"] = etiqueta
        hoja[f"B{fila}"].font = Font(bold=True, color="44546A")
        hoja[f"B{fila}"].alignment = Alignment(vertical="center")
        hoja.merge_cells(f"C{fila}:F{fila}")
        hoja[f"C{fila}"] = valor
        hoja[f"C{fila}"].font = Font(size=12, bold=(desplazamiento == 0))
        hoja[f"C{fila}"].alignment = Alignment(vertical="center")

    # --- Tabla de ítems -----------------------------------------------------
    fila_tabla = 10
    encabezados = ["N°", "Descripción", "Unidad", "Cantidad", "Precio Unitario", "Importe"]
    for columna, texto in enumerate(encabezados, start=1):
        celda = hoja.cell(fila_tabla, columna, texto)
        celda.fill = RELLENO_MARCA
        celda.font = Font(bold=True, color="FFFFFF")
        celda.alignment = CENTRADO
        celda.border = BORDE_FINO
    hoja.row_dimensions[fila_tabla].height = 20

    linea_actual = fila_tabla
    for numero, linea in enumerate(totales["lineas"], start=1):
        linea_actual += 1
        cantidad = (
            int(linea["cantidad"])
            if float(linea["cantidad"]).is_integer()
            else linea["cantidad"]
        )
        valores = [numero, linea["descripcion"], linea["unidad"], cantidad,
                   linea["precio_unitario"], linea["importe"]]
        for columna, valor in enumerate(valores, start=1):
            celda = hoja.cell(linea_actual, columna, valor)
            celda.border = BORDE_FINO
            if numero % 2 == 0:
                celda.fill = RELLENO_ZEBRA
            if columna == 2:
                celda.alignment = IZQUIERDA
            elif columna in (1, 3):
                celda.alignment = CENTRADO
            else:
                celda.alignment = DERECHA
            if columna >= 5:
                celda.number_format = FORMATO_MONEDA

    # --- Totales ------------------------------------------------------------
    fila_subtotal = linea_actual + 1
    totales_celdas = [
        ("SUBTOTAL", totales["subtotal"], Font(bold=True), None, BORDE_FINO),
        (f"IVA (19%)", totales["iva"], Font(bold=True), None, BORDE_FINO),
        ("TOTAL (pesos, IVA incluido)", totales["total"],
         Font(bold=True, size=12), RELLENO_TOTAL, BORDE_TOTAL),
    ]
    for desplazamiento, (etiqueta, valor, fuente, relleno, borde) in enumerate(totales_celdas):
        fila = fila_subtotal + desplazamiento
        hoja.merge_cells(f"A{fila}:E{fila}")
        hoja[f"A{fila}"] = etiqueta
        hoja[f"A{fila}"].font = fuente
        hoja[f"A{fila}"].alignment = DERECHA
        hoja[f"A{fila}"].border = BORDE_FINO
        celda_valor = hoja.cell(fila, 6, valor)
        celda_valor.font = fuente
        celda_valor.number_format = FORMATO_MONEDA
        celda_valor.alignment = DERECHA
        celda_valor.border = borde
        if relleno is not None:
            celda_valor.fill = relleno

    # --- Validez, notas y firmas ---------------------------------------------
    fila_validez = fila_subtotal + len(totales_celdas) + 1
    hoja.merge_cells(f"A{fila_validez}:F{fila_validez}")
    hoja[f"A{fila_validez}"] = (
        f"Validez de la oferta: {validez_dias} días desde la fecha de emisión "
        f"(vence el {fecha_larga(vencimiento)})."
    )
    hoja[f"A{fila_validez}"].font = Font(italic=True, bold=True)
    hoja[f"A{fila_validez}"].alignment = IZQUIERDA

    fila_notas = fila_validez + 1
    hoja[f"A{fila_notas}"] = "Notas:"
    hoja[f"A{fila_notas}"].font = Font(bold=True, color="44546A")
    for indice, nota in enumerate(notas, start=1):
        fila = fila_notas + indice
        hoja.merge_cells(f"A{fila}:F{fila}")
        hoja[f"A{fila}"] = f"• {nota}"
        hoja[f"A{fila}"].font = Font(size=9)
        hoja[f"A{fila}"].alignment = IZQUIERDA

    fila_firma = fila_notas + len(notas) + 2
    hoja.merge_cells(f"B{fila_firma}:D{fila_firma}")
    hoja.merge_cells(f"E{fila_firma}:F{fila_firma}")
    hoja[f"B{fila_firma}"] = "_______________________________"
    hoja[f"E{fila_firma}"] = "______________________"
    firmas = [
        (f"Por {empresa}", "Aceptación del cliente"),
        (contacto_empresa.split(" · ")[0], "Nombre y firma"),
        ("Gerente Comercial", "RUT:"),
    ]
    for desplazamiento, (texto_empresa, texto_cliente) in enumerate(firmas, start=1):
        fila = fila_firma + desplazamiento
        hoja.merge_cells(f"B{fila}:D{fila}")
        hoja[f"B{fila}"] = texto_empresa
        hoja[f"B{fila}"].font = Font(size=10, bold=(desplazamiento == 1))
        hoja[f"B{fila}"].alignment = CENTRADO
        hoja.merge_cells(f"E{fila}:F{fila}")
        hoja[f"E{fila}"] = texto_cliente
        hoja[f"E{fila}"].font = Font(size=10, bold=(desplazamiento == 1))
        hoja[f"E{fila}"].alignment = CENTRADO

    # --- Ajustes de página y columnas ----------------------------------------
    anchos = {"A": 5, "B": 40, "C": 11, "D": 12, "E": 16, "F": 16}
    for columna, ancho in anchos.items():
        hoja.column_dimensions[columna].width = ancho
    hoja.page_setup.paperSize = hoja.PAPERSIZE_A4
    hoja.page_setup.orientation = "portrait"
    hoja.page_setup.fitToWidth = 1
    hoja.page_setup.fitToHeight = 0
    hoja.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    hoja.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    hoja.print_area = f"A1:F{fila_firma + len(firmas)}"

    destino = Path(salida) / f"cotizacion_{folio}.xlsx"
    libro.save(destino)
    return destino


def convertir_a_pdf(ruta_xlsx):
    """Convierte el xlsx indicado a PDF con LibreOffice y comprueba el resultado."""
    if shutil.which("soffice") is None:
        raise SystemExit(
            "No se encontró LibreOffice (comando 'soffice'). Instálalo o agrega "
            "'soffice' al PATH para generar el PDF, o usa --pdf no."
        )
    try:
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf",
             "--outdir", str(ruta_xlsx.parent.resolve()), str(ruta_xlsx.resolve())],
            check=True,
            capture_output=True,
            timeout=120,
        )
    except subprocess.CalledProcessError as error:
        detalle = (error.stderr or b"").decode("utf-8", "replace").strip()
        raise SystemExit(
            f"No se pudo convertir '{ruta_xlsx.name}' a PDF. Detalle: {detalle}"
        ) from error

    ruta_pdf = ruta_xlsx.with_suffix(".pdf")
    if not ruta_pdf.is_file() or ruta_pdf.stat().st_size == 0:
        raise SystemExit(
            f"LibreOffice no generó el PDF de '{ruta_xlsx.name}' o quedó vacío."
        )
    return ruta_pdf


def main(argv=None):
    """Genera la cotización (xlsx y PDF) y devuelve un resumen de lo creado."""
    parser = argparse.ArgumentParser(
        description=(
            "Genera una cotización con folio consecutivo COT-AAAA-NNNN, IVA 19% "
            "y validez de la oferta, en Excel y PDF."
        ),
        epilog="Sin argumentos usa items.csv y los datos de ejemplo junto a este script.",
    )
    parser.add_argument("--items", type=Path, default=BASE_DIR / "items.csv",
                        metavar="CSV", help="CSV de ítems (descripcion, unidad, cantidad, precio_unitario).")
    parser.add_argument("--cliente", default=CLIENTE_DEFAULT,
                        help="Nombre del cliente (por defecto: cliente de ejemplo).")
    parser.add_argument("--rut", dest="rut_cliente", default=RUT_CLIENTE_DEFAULT,
                        help="RUT del cliente.")
    parser.add_argument("--contacto", default=CONTACTO_CLIENTE_DEFAULT,
                        help="Persona de contacto en el cliente.")
    parser.add_argument("--empresa", default=EMPRESA, help="Nombre de tu empresa (marca).")
    parser.add_argument("--rut-empresa", default=RUT_EMPRESA, help="RUT de tu empresa.")
    parser.add_argument("--contacto-empresa", default=CONTACTO_EMPRESA,
                        help="Contacto de tu empresa (nombre · cargo · email · fono).")
    parser.add_argument("--validez", type=int, default=VALIDEZ_DEFAULT,
                        help="Validez de la oferta en días (por defecto: 15).")
    parser.add_argument("--fecha", default=None, metavar="AAAA-MM-DD",
                        help="Fecha de emisión (por defecto: hoy).")
    parser.add_argument("--salida", type=Path, default=Path("."),
                        metavar="DIRECTORIO", help="Directorio de salida (por defecto: .).")
    parser.add_argument("--pdf", choices=["auto", "si", "no"], default="auto",
                        help="Generar PDF con LibreOffice: auto (sí si soffice existe), si o no.")
    args = parser.parse_args(argv)

    if args.validez < 1:
        raise SystemExit(f"La validez debe ser al menos 1 día (recibí {args.validez}).")

    fecha = date.today() if args.fecha is None else _a_fecha(args.fecha)
    items = validar_items(leer_items(args.items), args.items)
    totales = calcular_totales(items)

    salida = Path(args.salida)
    salida.mkdir(parents=True, exist_ok=True)
    folio = siguiente_folio(salida, año=fecha.year)

    ruta_xlsx = construir_cotizacion(
        totales, folio, fecha, salida,
        validez_dias=args.validez,
        cliente=args.cliente,
        rut_cliente=args.rut_cliente,
        contacto_cliente=args.contacto,
        empresa=args.empresa,
        rut_empresa=args.rut_empresa,
        contacto_empresa=args.contacto_empresa,
    )

    ruta_pdf = None
    if args.pdf == "si" or (args.pdf == "auto" and shutil.which("soffice") is not None):
        ruta_pdf = convertir_a_pdf(ruta_xlsx)
    else:
        print("LibreOffice no disponible (o --pdf no): cotización solo en Excel.")

    print(f"Cotización {folio}: SUBTOTAL ${totales['subtotal']:,} · "
          f"IVA (19%) ${totales['iva']:,} · TOTAL ${totales['total']:,}")
    print(f"XLSX: {ruta_xlsx.resolve()}")
    if ruta_pdf:
        print(f"PDF: {ruta_pdf.resolve()}")
    return {"folio": folio, "fecha": fecha, "xlsx": ruta_xlsx, "pdf": ruta_pdf,
            "totales": totales, "validez": args.validez}


if __name__ == "__main__":
    main()
