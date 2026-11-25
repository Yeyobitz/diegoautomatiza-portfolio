"""Plantillas .xlsx de certificado con diseño de documento imprimible.

Uso::

    python plantillas.py            # crea ./plantillas/*.xlsx
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

MARCADORES = ["{{NOMBRE}}", "{{CURSO}}", "{{FECHA}}", "{{NOTA}}", "{{FOLIO}}"]

PALETAS = {
    "participacion": {
        "oscuro": (22, 48, 79),
        "claro": (96, 134, 176),
        "acento": "B08D3C",
        "tinta": "1A1A1A",
        "gris": "5A5A5A",
    },
    "distincion": {
        "oscuro": (96, 26, 48),
        "claro": (168, 102, 126),
        "acento": "B08D3C",
        "tinta": "1A1A1A",
        "gris": "5A5A5A",
    },
}

_TEXTOS = {
    "participacion": {
        "titulo": "C E R T I F I C A D O",
        "subtitulo": "de Participación",
        "intro": "Se otorga el presente certificado a",
        "logro": "por completar satisfactoriamente el curso",
        "nota": "obteniendo una nota final de {{NOTA}} (escala 1.0 – 7.0).",
    },
    "distincion": {
        "titulo": "D I S T I N C I Ó N",
        "subtitulo": "al Mejor Participante",
        "intro": "Se otorga la presente distinción a",
        "logro": "por alcanzar el mejor rendimiento del curso",
        "nota": "con nota final {{NOTA}}, la más alta de su generación.",
    },
}

_FIRMA = "Diego A. Quiñones"
_CARGO = "Director de Formación · DQ Academy"

# Texto de la banda superior con tracking. Los espacios finales compensan el
# centro de la celda M (12.5) contra el eje del documento (12.0); el valor se
# calibró midiendo el render hasta un desvío de ~1 px.
ESPACIOS_COMPENSACION = 12
TEXTO_BANDA = (
    "D Q   A C A D E M Y   —   C E N T R O   D E   F O R M A C I Ó N"
    + " " * ESPACIOS_COMPENSACION
)

FILA_TITULO = 17  # primera fila del título (merge hasta la 18)
FILA_BANDA_TEXTO = 13  # fila central de la banda (12-14)
FILA_FIRMA_NOMBRE = 31  # misma fila que el folio: línea base compartida

ASSET_MEDALLON = Path(__file__).resolve().parent / "assets" / "medallon-dq.png"

# (fila_inicio, columna_inicio, fila_fin, columna_fin)
_MERGES = (
    [
        (FILA_TITULO, 3, FILA_TITULO + 1, 22),  # título C17:V18
        (19, 3, 19, 22),  # subtítulo
        (21, 3, 21, 22),  # intro
        (22, 3, 23, 22),  # nombre
        (24, 3, 24, 22),  # logro
        (25, 3, 25, 22),  # curso
        (26, 3, 26, 22),  # nota
        (28, 3, 28, 22),  # ciudad y fecha
        (30, 14, 30, 20),  # línea de firma N30:T30
        (31, 14, 31, 20),  # nombre firmante N31:T31
        (32, 14, 32, 20),  # cargo N32:T32
        (31, 5, 31, 11),  # folio E31:K31
        (32, 5, 32, 11),  # aclaración bajo el folio E32:K32
    ]
)

_ALTURAS = {
    5: 30,
    6: 34,
    7: 34,
    8: 30,
    9: 12,
    10: 10,
    11: 3,
    12: 13,
    13: 13,
    14: 13,
    15: 3,
    16: 12,
    17: 24,
    18: 24,
    19: 18,
    20: 12,
    21: 16,
    22: 20,
    23: 20,
    24: 16,
    25: 18,
    26: 16,
    27: 10,
    28: 16,
    29: 10,
    30: 14,
    31: 18,
    32: 14,
}


def color_intermedio(c1: tuple, c2: tuple, t: float) -> str:
    """Interpola linealmente dos colores RGB y devuelve el hex sin '#'."""
    return "".join(
        f"{max(0, min(255, round(a + (b - a) * t))):02X}" for a, b in zip(c1, c2)
    )


def _hex(rgb: tuple) -> str:
    return "".join(f"{canal:02X}" for canal in rgb)


def _relleno(color: str) -> PatternFill:
    return PatternFill(fill_type="solid", start_color=color, end_color=color)


def _perimetro(bordes: dict, fila1: int, col1: int, fila2: int, col2: int, estilo: str, color: str) -> None:
    """Acumula el side que corresponde a cada celda del perímetro."""
    lado = Side(style=estilo, color=color)
    for col in range(col1, col2 + 1):
        bordes.setdefault((fila1, col), {})["top"] = lado
        bordes.setdefault((fila2, col), {})["bottom"] = lado
    for fila in range(fila1, fila2 + 1):
        bordes.setdefault((fila, col1), {})["left"] = lado
        bordes.setdefault((fila, col2), {})["right"] = lado


def _geometria(hoja) -> None:
    for col in range(1, 25):
        hoja.column_dimensions[get_column_letter(col)].width = 4.6
    for fila in range(1, 35):
        hoja.row_dimensions[fila].height = _ALTURAS.get(fila, 16)


def _estilos(hoja, paleta: dict, textos: dict) -> None:
    oscuro = _hex(paleta["oscuro"])
    acento = paleta["acento"]
    tinta = paleta["tinta"]
    gris = paleta["gris"]

    bordes = {}
    _perimetro(bordes, 1, 1, 34, 24, "thick", oscuro)
    _perimetro(bordes, 3, 3, 33, 22, "thin", acento)
    # Línea de firma: azul marino del diseño, grosor medium (más fina que el
    # marco exterior thick).
    linea = Side(style="medium", color=oscuro)
    for col in range(14, 21):  # N..T
        bordes.setdefault((30, col), {})["bottom"] = linea
    for (fila, col), lados in bordes.items():
        hoja.cell(fila, col).border = Border(**lados)

    # Esquinas en bloque + barra inferior que conecta las esquinas sin dejar
    # franja blanca entre medio.
    relleno_oscuro = _relleno(oscuro)
    for fila1, col1, fila2, col2 in (
        (1, 1, 2, 2),
        (1, 23, 2, 24),
        (33, 1, 34, 2),
        (33, 23, 34, 24),
        (34, 3, 34, 22),
    ):
        for fila in range(fila1, fila2 + 1):
            for col in range(col1, col2 + 1):
                hoja.cell(fila, col).fill = relleno_oscuro

    # Filetes dorados: filas finas (3 pt) rellenas de acento arriba y abajo
    # de la banda — más confiable que bordes sobre celdas fusionadas.
    relleno_acento = _relleno(acento)
    for fila in (11, 15):
        for col in range(4, 22):
            hoja.cell(fila, col).fill = relleno_acento

    # Medallón dorado con monograma DQ: imagen PNG nítida (asset generado con
    # PIL, ver .generar_medallon.py) incrustada sobre la cuadrícula. Anclada en
    # K5 mide 4 columnas (148 px) → su centro cae en el eje del documento
    # (columna 12.0) y su alto cabe en las filas 5-8.
    if ASSET_MEDALLON.is_file():
        medallon = XLImage(str(ASSET_MEDALLON))
        medallon.anchor = "K5"
        medallon.width = 148
        medallon.height = 160
        hoja.add_image(medallon)

    # Degradado simulado: mezcla 2D (columna dominante + fila) para suavizar
    # los escalones sin perder el efecto de rellenos escalonados.
    for col in range(4, 22):
        t_col = (col - 4) / 17 * 0.40
        for fila in (12, 13, 14):
            t_fila = (fila - 12) / 2 * 0.10
            relleno = _relleno(color_intermedio(paleta["oscuro"], paleta["claro"], t_col + t_fila))
            hoja.cell(fila, col).fill = relleno

    centrado = Alignment(horizontal="center", vertical="center")

    def escribir(celda, valor, fuente, alineacion=centrado):
        destino = hoja[celda]
        destino.value = valor
        destino.font = fuente
        destino.alignment = alineacion

    # Texto sobre la banda, blanco con tracking. Va en una sola celda con
    # centrado; los espacios finales compensan la media columna de desvío del
    # centro de M (12.5) respecto del eje del documento (12.0).
    fuente_banda = Font(name="Georgia", size=10, color="FFFFFF")
    escribir(f"M{FILA_BANDA_TEXTO}", TEXTO_BANDA, fuente_banda)

    escribir(
        f"C{FILA_TITULO}",
        textos["titulo"],
        Font(name="Georgia", size=30, bold=True, color=oscuro),
    )
    escribir("C19", textos["subtitulo"], Font(name="Georgia", size=14, italic=True, color=acento))
    escribir("C21", textos["intro"], Font(name="Georgia", size=11, italic=True, color=gris))
    escribir("C22", "{{NOMBRE}}", Font(name="Georgia", size=22, color=tinta))
    escribir("C24", textos["logro"], Font(name="Georgia", size=11, color=gris))
    escribir("C25", "{{CURSO}}", Font(name="Georgia", size=14, bold=True, color=oscuro))
    escribir("C26", textos["nota"], Font(name="Georgia", size=11, color=gris))
    escribir("C28", "Santiago, Chile · {{FECHA}}", Font(name="Georgia", size=10.5, color=gris))

    # Pie: folio a la izquierda y firma a la derecha comparten la línea base
    # (fila 31) y quedan a igual distancia de los filetes verticales dorados
    # (un respiro de 2 columnas por lado, bloques de 7 columnas c/u).
    escribir(f"N{FILA_FIRMA_NOMBRE}", _FIRMA, Font(name="Georgia", size=11, bold=True, color=tinta))
    escribir("N32", _CARGO, Font(name="Georgia", size=9, color=gris))
    escribir(
        f"E{FILA_FIRMA_NOMBRE}",
        "Folio {{FOLIO}}",
        Font(name="Georgia", size=9, color=gris),
        Alignment(horizontal="left", vertical="center"),
    )
    escribir(
        "E32",
        "DQ Academy · Registro de emisión",
        Font(name="Georgia", size=7.5, color=gris),
        Alignment(horizontal="left", vertical="top"),
    )


def _pagina(hoja) -> None:
    hoja.print_area = "A1:X34"
    hoja.page_setup.orientation = "landscape"
    hoja.page_setup.paperSize = 9
    hoja.page_setup.fitToWidth = 1
    hoja.page_setup.fitToHeight = 1
    hoja.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    hoja.page_margins = PageMargins(left=0.3, right=0.3, top=0.25, bottom=0.25)
    hoja.print_options.horizontalCentered = True
    hoja.print_options.verticalCentered = True


def construir_plantilla(tipo: str, ruta) -> None:
    """Genera la plantilla del tipo indicado en ``ruta``."""
    if tipo not in PALETAS:
        raise ValueError(f"Tipo de plantilla desconocido: {tipo!r}")

    libro = Workbook()
    hoja = libro.active
    hoja.title = "Certificado"
    hoja.sheet_view.showGridLines = False
    _geometria(hoja)
    for fila1, col1, fila2, col2 in _MERGES:
        hoja.merge_cells(start_row=fila1, start_column=col1, end_row=fila2, end_column=col2)
    _estilos(hoja, PALETAS[tipo], _TEXTOS[tipo])
    _pagina(hoja)
    libro.save(ruta)


def construir_todas(directorio):
    """Crea ambas plantillas en ``directorio`` y devuelve sus rutas."""
    directorio = Path(directorio)
    directorio.mkdir(parents=True, exist_ok=True)
    rutas = []
    for tipo, nombre in (
        ("participacion", "certificado_participacion.xlsx"),
        ("distincion", "certificado_distincion.xlsx"),
    ):
        ruta = directorio / nombre
        construir_plantilla(tipo, ruta)
        rutas.append(ruta)
    return rutas


if __name__ == "__main__":
    for ruta in construir_todas("./plantillas"):
        print(ruta)
