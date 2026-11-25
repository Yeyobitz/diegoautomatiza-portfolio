#!/usr/bin/env python3
"""Monitor de precios de catálogo con detección de cambios.

Recorre únicamente las páginas de listado de books.toscrape.com, guarda un
snapshot en Excel y, al comparar con un snapshot anterior, genera un informe
con los precios que subieron, bajaron, desaparecieron o son nuevos. También
mantiene un historial de precios entre corridas.
"""

import argparse
import re
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote, urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

URL_POR_DEFECTO = "http://books.toscrape.com/"
SNAPSHOT_POR_DEFECTO = "catalogo_snapshot.xlsx"
CAMBIOS_POR_DEFECTO = "cambios_precios.xlsx"
HISTORIAL_POR_DEFECTO = "historial_precios.xlsx"

RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

RELLENO_ENCABEZADO = PatternFill("solid", fgColor="1F4E78")
FUENTE_ENCABEZADO = Font(bold=True, color="FFFFFF", size=11)
RELLENO_PAR = PatternFill("solid", fgColor="F2F2F2")
FORMATO_PRECIO = '"£"#,##0.00'
FORMATO_VARIACION = "+0.0%;-0.0%;0.0%"
FORMATO_FECHA = "dd-mm-yyyy hh:mm"


def parse_price(texto):
    """Convierte un precio como «£51.77» en el número 51.77."""
    limpio = re.sub(r"[^0-9.]", "", texto or "")
    if not limpio:
        raise ValueError(f"No se pudo interpretar el precio: {texto!r}")
    return float(limpio)


def parse_rating(clases):
    """Convierte las clases CSS del rating («One»...«Five») en 1...5."""
    for clase in clases or []:
        if clase in RATINGS:
            return RATINGS[clase]
    raise ValueError(f"Rating desconocido: {clases!r}")


def parse_availability(texto):
    """Devuelve True si el texto indica disponibilidad."""
    texto = texto or ""
    if "In stock" in texto:
        return True
    if "Out of stock" in texto:
        return False
    return False


def parse_catalog_page(soup, base_url):
    """Extrae los productos de una página de listado, sin abrir fichas."""
    productos = []
    for articulo in soup.select("article.product_pod"):
        enlace = articulo.select_one("h3 a")
        if enlace is None:
            continue

        precio = articulo.select_one(".price_color")
        rating = articulo.select_one(".star-rating")
        disponibilidad = articulo.select_one(".instock.availability")

        productos.append(
            {
                "nombre": enlace.get("title", "").strip(),
                "precio": parse_price(
                    precio.get_text(strip=True) if precio else ""
                ),
                "rating": parse_rating(rating.get("class", []) if rating else []),
                "disponible": parse_availability(
                    disponibilidad.get_text(" ", strip=True)
                    if disponibilidad
                    else ""
                ),
                "url": urljoin(base_url, enlace.get("href", "")),
            }
        )
    return productos


def fetch_soup(url):
    """Devuelve el HTML parseado de una URL http(s) o de un archivo file://."""
    if url.startswith("file://"):
        ruta = Path(unquote(urlparse(url).path))
        try:
            html = ruta.read_text(encoding="utf-8")
        except OSError as error:
            raise SystemExit(
                f"No se pudo leer el archivo local {ruta}: {error}. "
                "Revisa que la ruta exista y que el archivo sea legible."
            ) from error
        return BeautifulSoup(html, "html.parser")

    if not url.startswith(("http://", "https://")):
        raise SystemExit(
            f"URL no admitida: {url}. Usa una dirección http, https o file://."
        )

    try:
        respuesta = requests.get(url, timeout=30)
        respuesta.raise_for_status()
    except requests.exceptions.Timeout as error:
        raise SystemExit(
            f"Tiempo de espera agotado al acceder a {url}. "
            "Revisa tu conexión e inténtalo de nuevo."
        ) from error
    except requests.exceptions.HTTPError as error:
        raise SystemExit(
            f"Error HTTP al acceder a {url}: {error}. "
            "Revisa la URL y si el sitio está disponible."
        ) from error
    except requests.exceptions.RequestException as error:
        raise SystemExit(
            f"Error de red al acceder a {url}: {error}. "
            "Revisa tu conexión e inténtalo de nuevo."
        ) from error

    return BeautifulSoup(respuesta.content, "html.parser", from_encoding="utf-8")


def scrape_catalog(url):
    """Recorre el catálogo completo siguiendo el enlace «next»."""
    productos = {}
    pagina = url

    while pagina:
        soup = fetch_soup(pagina)
        for producto in parse_catalog_page(soup, pagina):
            productos.setdefault(producto["url"], producto)

        siguiente = soup.select_one("li.next > a")
        pagina = urljoin(pagina, siguiente["href"]) if siguiente else None

    if not productos:
        raise SystemExit(
            "No se encontró ningún producto. Revisa la URL inicial, tu conexión "
            "y que el sitio mantenga la estructura del catálogo."
        )
    return list(productos.values())


def _aplicar_estilo_datos(ws, anchos):
    """Aplica encabezados, anchos, autofiltro, fila fija y filas alternas."""
    for indice, ancho in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(indice)].width = ancho

    for celda in ws[1][: len(anchos)]:
        celda.fill = RELLENO_ENCABEZADO
        celda.font = FUENTE_ENCABEZADO
        celda.alignment = Alignment(horizontal="center", vertical="center")

    total_filas = max(ws.max_row, 1)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = (
        f"A1:{get_column_letter(len(anchos))}{total_filas}"
    )

    # Las filas pares llevan el relleno de banda.
    for numero_fila in range(2, total_filas + 1):
        if numero_fila % 2 == 0:
            for indice in range(1, len(anchos) + 1):
                ws.cell(row=numero_fila, column=indice).fill = RELLENO_PAR

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_title_rows = "1:1"


def _escribir_hoja(ws, encabezados, filas, anchos):
    """Escribe una tabla de datos y le aplica el estilo común."""
    ws.append(encabezados)
    for fila in filas:
        ws.append(fila)
    _aplicar_estilo_datos(ws, anchos)


def write_snapshot(productos, ruta, fecha, fuente=None):
    """Escribe el snapshot del catálogo y su hoja de resumen."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Catálogo"

    encabezados = [
        "Producto",
        "Precio (£)",
        "Rating (1-5)",
        "Disponibilidad",
        "URL",
    ]
    filas = [
        [
            producto["nombre"],
            producto["precio"],
            producto["rating"],
            "En stock" if producto["disponible"] else "Agotado",
            producto["url"],
        ]
        for producto in productos
    ]
    _escribir_hoja(ws, encabezados, filas, [55, 12, 12, 15, 75])

    for numero_fila in range(2, ws.max_row + 1):
        ws.cell(row=numero_fila, column=2).number_format = FORMATO_PRECIO

    # La fuente es opcional para conservar la firma de tres argumentos.
    if fuente is None:
        fuente = productos[0]["url"] if productos else ""

    resumen = wb.create_sheet("Resumen")
    resumen.column_dimensions["A"].width = 24
    resumen.column_dimensions["B"].width = 40

    resumen["A1"] = "Monitor de precios de catálogo"
    resumen["A1"].font = Font(bold=True, size=14)

    precios = [producto["precio"] for producto in productos]
    promedio = sum(precios) / len(precios) if precios else 0
    minimo = min(precios) if precios else 0
    maximo = max(precios) if precios else 0

    datos_resumen = [
        ("Fecha de corrida", fecha, FORMATO_FECHA),
        ("Fuente", fuente, None),
        ("Productos", len(productos), None),
        ("Precio promedio", promedio, FORMATO_PRECIO),
        ("Precio mínimo", minimo, FORMATO_PRECIO),
        ("Precio máximo", maximo, FORMATO_PRECIO),
    ]
    for indice, (rotulo, valor, formato) in enumerate(datos_resumen, start=2):
        resumen.cell(row=indice, column=1, value=rotulo).font = Font(bold=True)
        celda = resumen.cell(row=indice, column=2, value=valor)
        if formato:
            celda.number_format = formato

    resumen["A9"] = "Distribución de precios"
    resumen["A9"].font = Font(bold=True, size=12)

    fila_encabezado = 10
    for columna, titulo in enumerate(["Rango de precio", "Productos"], start=1):
        celda = resumen.cell(row=fila_encabezado, column=columna, value=titulo)
        celda.fill = RELLENO_ENCABEZADO
        celda.font = FUENTE_ENCABEZADO
        celda.alignment = Alignment(horizontal="center", vertical="center")

    inicio = (int(minimo) // 10) * 10
    fin = (int(maximo) // 10) * 10
    if maximo > fin or fin == inicio:
        fin += 10

    fila = fila_encabezado + 1
    for limite in range(inicio, fin, 10):
        superior = limite + 10
        ultimo = superior >= fin
        cantidad = sum(
            1
            for precio in precios
            if limite <= precio < superior or (ultimo and precio == superior)
        )
        resumen.cell(row=fila, column=1, value=f"£{limite}–£{superior}")
        resumen.cell(row=fila, column=2, value=cantidad)
        fila += 1

    grafico = BarChart()
    grafico.title = "Distribución de precios"
    grafico.legend = None
    grafico.y_axis.scaling.min = 0
    datos = Reference(
        resumen,
        min_col=2,
        min_row=fila_encabezado,
        max_row=fila - 1,
    )
    categorias = Reference(
        resumen,
        min_col=1,
        min_row=fila_encabezado + 1,
        max_row=fila - 1,
    )
    grafico.add_data(datos, titles_from_data=True)
    grafico.set_categories(categorias)
    resumen.add_chart(grafico, "D9")

    resumen.page_setup.orientation = "landscape"
    resumen.page_setup.fitToWidth = 1
    resumen.page_setup.fitToHeight = 0
    resumen.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    resumen.print_title_rows = "1:1"

    wb.save(ruta)


def read_snapshot(ruta):
    """Lee la hoja «Catálogo» de un snapshot y devuelve sus productos."""
    try:
        wb = load_workbook(ruta, data_only=True)
    except FileNotFoundError as error:
        raise SystemExit(
            f"No se encontró el snapshot {ruta}. Revisa la ruta indicada."
        ) from error

    if "Catálogo" not in wb.sheetnames:
        raise SystemExit(
            f"El archivo {ruta} no contiene la hoja «Catálogo». "
            "Revisa que sea un snapshot generado por este programa."
        )

    ws = wb["Catálogo"]
    productos = []
    for fila in ws.iter_rows(min_row=2, values_only=True):
        if not fila or fila[0] is None:
            continue
        productos.append(
            {
                "nombre": fila[0],
                "precio": float(fila[1]),
                "rating": int(fila[2]),
                "disponible": str(fila[3]).strip() == "En stock",
                "url": fila[4],
            }
        )
    return productos


def detect_changes(anteriores, actuales):
    """Compara dos catálogos por URL y clasifica los cambios de precio."""
    previos = {producto["url"]: producto for producto in anteriores}
    nuevos_por_url = {producto["url"]: producto for producto in actuales}
    resultado = {"subieron": [], "bajaron": [], "desaparecidos": [], "nuevos": []}

    for url, producto in nuevos_por_url.items():
        anterior = previos.get(url)
        if anterior is None:
            resultado["nuevos"].append(
                {
                    "nombre": producto["nombre"],
                    "precio_nuevo": producto["precio"],
                    "url": url,
                }
            )
            continue

        precio_anterior = anterior["precio"]
        precio_nuevo = producto["precio"]
        if precio_nuevo == precio_anterior:
            continue
        if precio_anterior == 0:
            continue

        cambio = {
            "nombre": producto["nombre"],
            "precio_anterior": precio_anterior,
            "precio_nuevo": precio_nuevo,
            "variacion_pct": (precio_nuevo - precio_anterior) / precio_anterior,
            "url": url,
        }
        resultado["subieron" if precio_nuevo > precio_anterior else "bajaron"].append(
            cambio
        )

    for url, anterior in previos.items():
        if url not in nuevos_por_url:
            resultado["desaparecidos"].append(
                {
                    "nombre": anterior["nombre"],
                    "precio_anterior": anterior["precio"],
                    "url": url,
                }
            )

    for clave in ("subieron", "bajaron"):
        resultado[clave].sort(
            key=lambda cambio: abs(cambio["variacion_pct"]), reverse=True
        )
    return resultado


def _preparar_historial(wb):
    """Devuelve la hoja «Historial», creándola con encabezados si hace falta."""
    if "Historial" in wb.sheetnames:
        return wb["Historial"]
    ws = wb.create_sheet("Historial")
    ws.append(["Fecha", "Producto", "Precio"])
    return ws


def _formatear_historial(ws):
    """Aplica el estilo común a la hoja de historial."""
    _aplicar_estilo_datos(ws, [18, 55, 12])
    for numero_fila in range(2, ws.max_row + 1):
        ws.cell(row=numero_fila, column=1).number_format = FORMATO_FECHA
        ws.cell(row=numero_fila, column=3).number_format = FORMATO_PRECIO


def append_history(productos, ruta_historial, fecha):
    """Agrega una fila por producto al historial y devuelve todo el historial."""
    ruta = Path(ruta_historial)
    if ruta.exists():
        wb = load_workbook(ruta)
    else:
        wb = Workbook()
        wb.remove(wb.active)

    ws = _preparar_historial(wb)
    for producto in productos:
        ws.append([fecha, producto["nombre"], producto["precio"]])

    _formatear_historial(ws)
    wb.save(ruta)
    return read_history(ruta)


def read_history(ruta_historial):
    """Devuelve el historial completo como lista de diccionarios."""
    ruta = Path(ruta_historial)
    if not ruta.exists():
        return []

    wb = load_workbook(ruta, data_only=True)
    if "Historial" not in wb.sheetnames:
        return []

    ws = wb["Historial"]
    historial = []
    for fila in ws.iter_rows(min_row=2, values_only=True):
        if not fila or fila[0] is None:
            continue
        historial.append(
            {"fecha": fila[0], "producto": fila[1], "precio": float(fila[2])}
        )
    return historial


def write_changes(cambios, historial, ruta):
    """Escribe el informe de cambios con cinco hojas."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Subieron"

    _escribir_hoja(
        ws,
        ["Producto", "Precio anterior", "Precio nuevo", "Variación", "URL"],
        [
            [
                cambio["nombre"],
                cambio["precio_anterior"],
                cambio["precio_nuevo"],
                cambio["variacion_pct"],
                cambio["url"],
            ]
            for cambio in cambios["subieron"]
        ],
        [55, 15, 15, 12, 75],
    )

    ws_bajaron = wb.create_sheet("Bajaron")
    _escribir_hoja(
        ws_bajaron,
        ["Producto", "Precio anterior", "Precio nuevo", "Variación", "URL"],
        [
            [
                cambio["nombre"],
                cambio["precio_anterior"],
                cambio["precio_nuevo"],
                cambio["variacion_pct"],
                cambio["url"],
            ]
            for cambio in cambios["bajaron"]
        ],
        [55, 15, 15, 12, 75],
    )

    ws_desaparecidos = wb.create_sheet("Desaparecidos")
    _escribir_hoja(
        ws_desaparecidos,
        ["Producto", "Precio anterior", "URL"],
        [
            [cambio["nombre"], cambio["precio_anterior"], cambio["url"]]
            for cambio in cambios["desaparecidos"]
        ],
        [55, 15, 75],
    )

    ws_nuevos = wb.create_sheet("Nuevos")
    _escribir_hoja(
        ws_nuevos,
        ["Producto", "Precio nuevo", "URL"],
        [
            [cambio["nombre"], cambio["precio_nuevo"], cambio["url"]]
            for cambio in cambios["nuevos"]
        ],
        [55, 15, 75],
    )

    ws_historial = wb.create_sheet("Historial")
    _escribir_hoja(
        ws_historial,
        ["Fecha", "Producto", "Precio"],
        [
            [registro["fecha"], registro["producto"], registro["precio"]]
            for registro in historial
        ],
        [18, 55, 12],
    )

    for hoja, color in ((ws, "C00000"), (ws_bajaron, "2E7D32")):
        for numero_fila in range(2, hoja.max_row + 1):
            celda = hoja.cell(row=numero_fila, column=4)
            celda.font = Font(color=color)
            celda.number_format = FORMATO_VARIACION
        for numero_fila in range(2, hoja.max_row + 1):
            hoja.cell(row=numero_fila, column=2).number_format = FORMATO_PRECIO
            hoja.cell(row=numero_fila, column=3).number_format = FORMATO_PRECIO

    for numero_fila in range(2, ws_desaparecidos.max_row + 1):
        ws_desaparecidos.cell(
            row=numero_fila, column=2
        ).number_format = FORMATO_PRECIO
    for numero_fila in range(2, ws_nuevos.max_row + 1):
        ws_nuevos.cell(row=numero_fila, column=2).number_format = FORMATO_PRECIO

    for numero_fila in range(2, ws_historial.max_row + 1):
        ws_historial.cell(row=numero_fila, column=1).number_format = FORMATO_FECHA
        ws_historial.cell(row=numero_fila, column=3).number_format = FORMATO_PRECIO

    wb.save(ruta)


def parse_args(argv=None):
    """Define y analiza las opciones de línea de comandos."""
    parser = argparse.ArgumentParser(
        description=(
            "Monitorea los precios del catálogo de Books to Scrape, guarda un "
            "snapshot en Excel y detecta cambios frente a un snapshot anterior."
        )
    )
    parser.add_argument(
        "--url",
        default=URL_POR_DEFECTO,
        metavar="URL",
        help="URL inicial del catálogo (por defecto: %(default)s).",
    )
    parser.add_argument(
        "--output",
        default=None,
        metavar="ARCHIVO",
        help=(
            "Archivo de salida. Por defecto: catalogo_snapshot.xlsx, o "
            "cambios_precios.xlsx si se usa --compare."
        ),
    )
    parser.add_argument(
        "--compare",
        default=None,
        metavar="SNAPSHOT",
        help="Ruta de un snapshot anterior con el que comparar los precios.",
    )
    return parser.parse_args(argv)


def _ruta_historial(salida):
    """Coloca el historial en el mismo directorio que el archivo de salida."""
    return str(Path(salida).parent / HISTORIAL_POR_DEFECTO)


def main(argv=None):
    """Ejecuta el monitoreo y escribe los archivos correspondientes."""
    args = parse_args(argv)
    fecha = datetime.now()

    if args.compare:
        salida = args.output or CAMBIOS_POR_DEFECTO
        productos = scrape_catalog(args.url)
        anteriores = read_snapshot(args.compare)
        cambios = detect_changes(anteriores, productos)
        historial = append_history(productos, _ruta_historial(salida), fecha)
        write_changes(cambios, historial, salida)
        write_snapshot(productos, args.compare, fecha, args.url)
        print(
            "Cambios detectados: "
            f"{len(cambios['subieron'])} subieron, "
            f"{len(cambios['bajaron'])} bajaron, "
            f"{len(cambios['desaparecidos'])} desaparecidos, "
            f"{len(cambios['nuevos'])} nuevos. Archivo: {salida}"
        )
        return

    salida = args.output or SNAPSHOT_POR_DEFECTO
    productos = scrape_catalog(args.url)
    write_snapshot(productos, salida, fecha, args.url)
    append_history(productos, _ruta_historial(salida), fecha)

    precios = [producto["precio"] for producto in productos]
    promedio = sum(precios) / len(precios)
    print(
        f"Total de productos: {len(productos)}. "
        f"Precio promedio: £{promedio:.2f}. Archivo: {salida}"
    )


if __name__ == "__main__":
    main()
