"""Generador de Reporte Diario de Faena.

Cruza las notas de WhatsApp del jefe de faena (reporte_jefe.txt) con el
registro de horas-hombre de las cuadrillas (horas_cuadrillas.xlsx) y genera
un reporte formal en Excel y PDF.
"""

import argparse
import re
import shutil
import subprocess
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

DIRECTORIO_BASE = Path(__file__).resolve().parent

# Formato de cada línea útil de un export de WhatsApp.
RE_LINEA = re.compile(
    r"^(\d{1,2})/(\d{1,2})/(\d{2,4}),\s*(\d{1,2}:\d{2})\s*-\s*(.*)$"
)
RE_PERMISO = re.compile(r"\bPT-(\d+)\b", flags=re.IGNORECASE)
RE_AVANCE = re.compile(
    r"avance\s+(?:en\s+)?([A-Za-z0-9]+(?:-[A-Za-z0-9]+)*)\s*(\d{1,3})\s*%", flags=re.IGNORECASE
)
RE_AUSENTE = re.compile(r"cuadrilla\s+(\d+)\s+no\s+(?:salio|asistio|partio)")

# Colores del reporte.
AZUL_OSCURO = "1F4E78"
RELLENO_ENCABEZADO = PatternFill("solid", fgColor=AZUL_OSCURO)
RELLENO_TOTALES = PatternFill("solid", fgColor="D9D9D9")
RELLENO_SEVERIDAD = {
    "Alta": PatternFill("solid", fgColor="FFC7CE"),
    "Baja": PatternFill("solid", fgColor="FFEB9C"),
    "Media": PatternFill("solid", fgColor="FCD5B4"),
}
RELLENO_ESTADO = {
    "Aprobado": PatternFill("solid", fgColor="C6EFCE"),
    "Rechazado": PatternFill("solid", fgColor="FFC7CE"),
    "Pendiente": PatternFill("solid", fgColor="FFEB9C"),
}
BORDE_FINO = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
ALINEACION_CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)


def normalizar(texto):
    """Pasa el texto a minúsculas y quita tildes para comparar."""
    sin_tildes = unicodedata.normalize("NFKD", str(texto))
    sin_tildes = "".join(c for c in sin_tildes if not unicodedata.combining(c))
    return sin_tildes.lower()


def _a_fecha(valor):
    """Convierte date, datetime o texto AAAA-MM-DD a date."""
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    return datetime.strptime(str(valor).strip(), "%Y-%m-%d").date()


def parsear_linea(linea):
    """Interpreta una línea de WhatsApp; devuelve None si no es un mensaje."""
    coincidencia = RE_LINEA.match(linea.strip())
    if not coincidencia:
        return None
    dia, mes, anio, hora, texto = coincidencia.groups()
    anio = int(anio)
    if anio < 100:
        anio += 2000
    try:
        fecha = date(anio, int(mes), int(dia))
    except ValueError:
        return None
    return {"fecha": fecha, "hora": hora, "texto": texto.strip()}


def clasificar_severidad(texto):
    """Clasifica la severidad: Alta si hay lesionados, Baja si es casi incidente."""
    limpio = normalizar(texto)
    if "lesionad" in limpio and "sin lesionad" not in limpio:
        return "Alta"
    if "casi incidente" in limpio:
        return "Baja"
    return "Media"


def _tipo_evento(limpio):
    """Devuelve el tipo de evento del mensaje normalizado, o None."""
    if "casi incidente" in limpio:
        return "Casi Incidente"
    if "accidente" in limpio:
        return "Accidente"
    if "incidente" in limpio:
        return "Incidente"
    return None


def _estado_permiso(limpio):
    """Determina el estado del permiso según el mensaje normalizado."""
    if "aprobado" in limpio:
        return "Aprobado"
    if "rechazado" in limpio:
        return "Rechazado"
    if "pendiente" in limpio or "espera firma" in limpio or "falta firma" in limpio:
        return "Pendiente"
    return "Pendiente"


def _detalle_permiso(texto, estado):
    """Quita el código y el estado para dejar el resto del mensaje como detalle."""
    resto = RE_PERMISO.sub("", texto)
    resto = re.sub(r"^\s*permiso\s*", "", resto, flags=re.IGNORECASE)
    patrones = {
        "Aprobado": r"aprobado",
        "Rechazado": r"rechazado",
        "Pendiente": r"(?:en\s+)?(?:espera|falta)\s+firma|pendiente(?:\s+firma)?",
    }
    resto = re.sub(patrones[estado], "", resto, count=1, flags=re.IGNORECASE)
    resto = resto.strip(" ,.;:-")
    return resto or texto.strip()


def _leer_origen(origen):
    """Lee el texto del reporte desde una ruta o lo recibe ya como contenido."""
    if isinstance(origen, Path):
        return origen.read_text(encoding="utf-8")
    try:
        ruta = Path(origen)
        if ruta.is_file():
            return ruta.read_text(encoding="utf-8")
    except OSError:
        pass
    return str(origen)


def parsear_reporte(origen, fecha):
    """Extrae eventos, permisos, avances, ausencias y turno de la fecha pedida."""
    fecha = _a_fecha(fecha)
    texto = _leer_origen(origen)
    datos = {
        "eventos": [],
        "permisos": [],
        "avance": {},
        "cuadrillas_ausentes": {},
        "turno": None,
    }
    for linea in texto.splitlines():
        mensaje = parsear_linea(linea)
        if not mensaje or mensaje["fecha"] != fecha:
            continue
        contenido = mensaje["texto"]
        hora = mensaje["hora"]
        limpio = normalizar(contenido)

        tipo = _tipo_evento(limpio)
        if tipo:
            datos["eventos"].append(
                {
                    "hora": hora,
                    "tipo": tipo,
                    "severidad": clasificar_severidad(contenido),
                    "descripcion": contenido,
                }
            )

        for numero in dict.fromkeys(RE_PERMISO.findall(contenido)):
            estado = _estado_permiso(limpio)
            datos["permisos"].append(
                {
                    "pt": f"PT-{numero}",
                    "hora": hora,
                    "estado": estado,
                    "detalle": _detalle_permiso(contenido, estado),
                }
            )

        for coincidencia in RE_AVANCE.finditer(contenido):
            datos["avance"][coincidencia.group(1).upper()] = int(
                coincidencia.group(2)
            )

        ausente = RE_AUSENTE.search(limpio)
        if ausente:
            partes = contenido.split(",")
            motivo = partes[-1].strip() if len(partes) > 1 else ""
            datos["cuadrillas_ausentes"][int(ausente.group(1))] = motivo

        if datos["turno"] is None:
            if "turno noche" in limpio:
                datos["turno"] = "Noche"
            elif "turno dia" in limpio:
                datos["turno"] = "Día"

    if not datos["turno"]:
        datos["turno"] = "Noche"
    return datos


def cargar_horas(ruta, fecha):
    """Carga las horas de las cuadrillas y filtra las filas de la fecha."""
    ruta = Path(ruta)
    if not ruta.is_file():
        raise FileNotFoundError(f"No se encontró el archivo de horas: {ruta}")
    horas = pd.read_excel(ruta)
    esperadas = {
        "cuadrilla",
        "frente",
        "lider",
        "HH_normales",
        "HH_extras",
        "n_trabajadores",
        "fecha",
    }
    faltantes = esperadas - set(horas.columns)
    if faltantes:
        raise ValueError(
            "El archivo de horas no tiene las columnas esperadas: "
            + ", ".join(sorted(faltantes))
        )
    fecha = _a_fecha(fecha)
    filtradas = horas[
        horas["fecha"].astype(str).str.strip() == fecha.isoformat()
    ].copy()
    if filtradas.empty:
        raise ValueError(
            f"No hay registros de horas para la fecha {fecha.isoformat()} "
            f"en {ruta.name}."
        )
    return filtradas.reset_index(drop=True)


def _estilar_encabezado(hoja, fila, columnas):
    """Escribe y estiliza la fila de encabezado de una tabla."""
    for indice, nombre in enumerate(columnas, start=1):
        celda = hoja.cell(row=fila, column=indice, value=nombre)
        celda.fill = RELLENO_ENCABEZADO
        celda.font = Font(bold=True, color="FFFFFF")
        celda.border = BORDE_FINO
        celda.alignment = ALINEACION_CENTRO


def _fijar_anchos(hoja, anchos):
    """Fija el ancho de cada columna uno por uno."""
    for indice, ancho in enumerate(anchos, start=1):
        hoja.column_dimensions[get_column_letter(indice)].width = ancho


def _bordear(fila_celdas):
    """Aplica borde fino a todas las celdas de una fila."""
    for celda in fila_celdas:
        celda.border = BORDE_FINO
        celda.alignment = ALINEACION_CENTRO


def _numero_reporte(fecha):
    """Número correlativo del reporte para el año de la fecha."""
    # shortcut: número fijo por año, actualizar cuando exista una fuente de numeración.
    return f"RDF-{fecha.year}-084"


def _hoja_resumen(libro, reporte, horas, fecha, turno, clima, estado_faena):
    """Construye la hoja Resumen del Día como documento controlado de faena."""
    hoja = libro.active
    hoja.title = "Resumen del Día"
    hoja.sheet_view.showGridLines = False

    columnas = ["Frente", "HH Normales", "HH Extras", "HH Total", "Trabajadores", "Avance"]
    ultima_columna = len(columnas)

    hoja.merge_cells(
        start_row=1, start_column=1, end_row=1, end_column=ultima_columna
    )
    titulo = hoja["A1"]
    titulo.value = "REPORTE DIARIO DE FAENA"
    titulo.font = Font(bold=True, size=16, color="FFFFFF")
    titulo.fill = RELLENO_ENCABEZADO
    titulo.alignment = ALINEACION_CENTRO
    hoja.row_dimensions[1].height = 36

    identificacion = [
        ("Mandante:", "Metalúrgica Andina S.A. (RUT 76.123.456-7)"),
        ("Contratista:", "Servicios Andamieros Austral Ltda. (RUT 77.234.567-8)"),
        ("Faena:", "PROYECTO CORDILLERA 2026"),
        (
            "Fecha:",
            f"{fecha:%d-%m-%Y}    |    Turno: {turno}    |    "
            f"N° de Reporte: {_numero_reporte(fecha)}",
        ),
        ("Código de documento:", "RDF-01 Rev. 2"),
        ("Clima del día:", clima),
        ("Estado de faena:", estado_faena),
    ]
    fila = 2
    for etiqueta, valor in identificacion:
        celda_etiqueta = hoja.cell(row=fila, column=1, value=etiqueta)
        celda_etiqueta.font = Font(bold=True)
        hoja.merge_cells(
            start_row=fila, start_column=2, end_row=fila, end_column=ultima_columna
        )
        celda_valor = hoja.cell(row=fila, column=2, value=valor)
        celda_valor.alignment = Alignment(horizontal="right", vertical="center")
        fila += 1

    fila += 1  # fila en blanco antes de la tabla
    _estilar_encabezado(hoja, fila, columnas)
    fila += 1

    por_frente = horas.groupby("frente", sort=False).agg(
        hh_normales=("HH_normales", "sum"),
        hh_extras=("HH_extras", "sum"),
        trabajadores=("n_trabajadores", "sum"),
    )

    for frente, datos in por_frente.iterrows():
        valores = [
            str(frente),
            int(datos["hh_normales"]),
            int(datos["hh_extras"]),
            int(datos["hh_normales"] + datos["hh_extras"]),
            int(datos["trabajadores"]),
        ]
        for columna, valor in enumerate(valores, start=1):
            hoja.cell(row=fila, column=columna, value=valor)
        avance = reporte["avance"].get(str(frente).upper())
        celda_avance = hoja.cell(row=fila, column=6)
        if avance is not None:
            celda_avance.value = avance / 100
            celda_avance.number_format = "0%"
        else:
            celda_avance.value = "—"
        _bordear(hoja[fila])
        fila += 1

    totales = [
        "TOTALES",
        int(horas["HH_normales"].sum()),
        int(horas["HH_extras"].sum()),
        int(horas["HH_normales"].sum() + horas["HH_extras"].sum()),
        int(horas["n_trabajadores"].sum()),
        "—",
    ]
    for columna, valor in enumerate(totales, start=1):
        celda = hoja.cell(row=fila, column=columna, value=valor)
        celda.font = Font(bold=True)
        celda.fill = RELLENO_TOTALES
    _bordear(hoja[fila])

    fila += 2  # fila en blanco de separación antes del pie
    firmas = [
        ("Elaborado por: " + "_" * 14, "Jefe de Faena"),
        ("Revisado por: " + "_" * 14, "Prevencionista de Riesgos"),
        ("Aprobado por: " + "_" * 14, "Administrador de Contrato"),
    ]
    for (inicio, fin), (linea, rol) in zip([(1, 2), (3, 4), (5, 6)], firmas):
        hoja.merge_cells(
            start_row=fila, start_column=inicio, end_row=fila, end_column=fin
        )
        celda_linea = hoja.cell(row=fila, column=inicio, value=linea)
        celda_linea.font = Font(size=9)
        celda_linea.alignment = ALINEACION_CENTRO
        hoja.merge_cells(
            start_row=fila + 1, start_column=inicio, end_row=fila + 1, end_column=fin
        )
        celda_rol = hoja.cell(row=fila + 1, column=inicio, value=rol)
        celda_rol.font = Font(size=8)
        celda_rol.alignment = ALINEACION_CENTRO
    fila += 2

    hoja.merge_cells(
        start_row=fila, start_column=1, end_row=fila, end_column=ultima_columna
    )
    nota = hoja.cell(
        row=fila,
        column=1,
        value="Documento generado automáticamente con reporte_faena.py",
    )
    nota.font = Font(size=8, italic=True, color="808080")
    nota.alignment = Alignment(horizontal="right")

    _fijar_anchos(hoja, [20, 13, 12, 12, 13, 12])


def _hoja_incidentes(libro, eventos):
    """Construye la hoja de incidentes y casi incidentes."""
    hoja = libro.create_sheet("Incidentes y Casi Incidentes")
    hoja.sheet_view.showGridLines = False
    _estilar_encabezado(hoja, 1, ["Hora", "Tipo", "Severidad", "Descripción"])

    if not eventos:
        hoja.merge_cells("A2:D2")
        hoja["A2"] = "Sin incidentes ni casi incidentes registrados."
    else:
        for fila, evento in enumerate(eventos, start=2):
            valores = [
                evento["hora"],
                evento["tipo"],
                evento["severidad"],
                evento["descripcion"],
            ]
            for columna, valor in enumerate(valores, start=1):
                hoja.cell(row=fila, column=columna, value=valor)
            hoja.cell(row=fila, column=4).alignment = Alignment(
                vertical="center", wrap_text=True
            )
            hoja.cell(
                row=fila, column=3
            ).fill = RELLENO_SEVERIDAD[evento["severidad"]]
            _bordear(hoja[fila])

    _fijar_anchos(hoja, [10, 16, 12, 70])


def _hoja_permisos(libro, permisos):
    """Construye la hoja de permisos de trabajo."""
    hoja = libro.create_sheet("Permisos de Trabajo")
    hoja.sheet_view.showGridLines = False
    _estilar_encabezado(hoja, 1, ["PT", "Hora", "Estado", "Detalle"])

    if not permisos:
        hoja.merge_cells("A2:D2")
        hoja["A2"] = "Sin permisos de trabajo registrados."
    else:
        for fila, permiso in enumerate(permisos, start=2):
            valores = [
                permiso["pt"],
                permiso["hora"],
                permiso["estado"],
                permiso["detalle"],
            ]
            for columna, valor in enumerate(valores, start=1):
                hoja.cell(row=fila, column=columna, value=valor)
            hoja.cell(row=fila, column=4).alignment = Alignment(
                vertical="center", wrap_text=True
            )
            hoja.cell(row=fila, column=3).fill = RELLENO_ESTADO[permiso["estado"]]
            _bordear(hoja[fila])

    _fijar_anchos(hoja, [12, 10, 12, 60])


def _hoja_cuadrillas(libro, horas, ausentes):
    """Construye la hoja de detalle de cuadrillas ordenada por número."""
    hoja = libro.create_sheet("Detalle de Cuadrillas")
    hoja.sheet_view.showGridLines = False
    _estilar_encabezado(
        hoja,
        1,
        [
            "Cuadrilla",
            "Frente",
            "Líder",
            "Trabajadores",
            "HH Normales",
            "HH Extras",
            "Observación",
        ],
    )

    ordenadas = horas.sort_values("cuadrilla")
    fila = 2
    for _, datos in ordenadas.iterrows():
        numero = int(datos["cuadrilla"])
        motivo = ausentes.get(numero, "").strip()
        observacion = f"No salió ({motivo})" if motivo else ("No salió" if numero in ausentes else "")
        valores = [
            numero,
            str(datos["frente"]),
            str(datos["lider"]),
            int(datos["n_trabajadores"]),
            int(datos["HH_normales"]),
            int(datos["HH_extras"]),
            observacion,
        ]
        for columna, valor in enumerate(valores, start=1):
            hoja.cell(row=fila, column=columna, value=valor)
        hoja.cell(row=fila, column=7).alignment = Alignment(
            vertical="center", wrap_text=True
        )
        _bordear(hoja[fila])
        fila += 1

    _fijar_anchos(hoja, [11, 10, 18, 13, 13, 11, 32])


def _construir_libro(
    reporte, horas, fecha, turno, clima="Despejado, viento moderado",
    estado_faena="Operativa",
):
    """Arma el libro Excel completo con las cuatro hojas en orden."""
    libro = Workbook()
    _hoja_resumen(libro, reporte, horas, fecha, turno, clima, estado_faena)
    _hoja_incidentes(libro, reporte["eventos"])
    _hoja_permisos(libro, reporte["permisos"])
    _hoja_cuadrillas(libro, horas, reporte["cuadrillas_ausentes"])
    _configurar_pagina(libro)
    return libro


def _configurar_pagina(libro):
    """Deja cada hoja lista para PDF: horizontal A4, ajustada a 1 página de ancho."""
    for hoja in libro.worksheets:
        hoja.page_setup.orientation = "landscape"
        hoja.page_setup.paperSize = hoja.PAPERSIZE_A4
        hoja.page_setup.fitToWidth = 1
        hoja.page_setup.fitToHeight = 0
        hoja.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        hoja.page_margins = PageMargins(
            left=0.5, right=0.5, top=0.6, bottom=0.6
        )


def _convertir_pdf(ruta_xlsx):
    """Convierte el XLSX a PDF con LibreOffice; si falla, solo advierte."""
    if not shutil.which("soffice"):
        print(
            "Advertencia: no se encontró soffice (LibreOffice); "
            "el PDF no fue generado.",
            file=sys.stderr,
        )
        return None
    comando = [
        "soffice",
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(ruta_xlsx.parent.resolve()),
        str(ruta_xlsx.resolve()),
    ]
    try:
        subprocess.run(
            comando,
            check=True,
            capture_output=True,
            timeout=120,
        )
    except subprocess.TimeoutExpired:
        print(
            "Advertencia: la conversión a PDF excedió 120 segundos; "
            "el XLSX quedó guardado.",
            file=sys.stderr,
        )
        return None
    except (subprocess.CalledProcessError, OSError) as error:
        print(
            f"Advertencia: no se pudo convertir el XLSX a PDF ({error}); "
            "el XLSX quedó guardado.",
            file=sys.stderr,
        )
        return None

    ruta_pdf = ruta_xlsx.with_suffix(".pdf")
    if not ruta_pdf.is_file():
        print(
            "Advertencia: LibreOffice no generó el PDF; el XLSX quedó guardado.",
            file=sys.stderr,
        )
        return None
    return ruta_pdf


def generar_reporte(
    ruta_txt, ruta_horas, fecha, salida=".", turno=None,
    clima="Despejado, viento moderado", estado_faena="Operativa",
):
    """Genera el XLSX y el PDF del reporte diario y devuelve sus rutas."""
    fecha = _a_fecha(fecha)
    ruta_txt = Path(ruta_txt)
    if not ruta_txt.is_file():
        raise FileNotFoundError(f"No se encontró el reporte del jefe: {ruta_txt}")

    reporte = parsear_reporte(ruta_txt, fecha)
    turno_final = turno or reporte["turno"]
    horas = cargar_horas(ruta_horas, fecha)

    directorio_salida = Path(salida)
    directorio_salida.mkdir(parents=True, exist_ok=True)
    ruta_xlsx = directorio_salida / f"reporte_faena_{fecha.isoformat()}.xlsx"

    libro = _construir_libro(
        reporte, horas, fecha, turno_final, clima, estado_faena
    )
    libro.save(ruta_xlsx)

    ruta_pdf = _convertir_pdf(ruta_xlsx)

    print(f"XLSX generado: {ruta_xlsx}")
    if ruta_pdf:
        print(f"PDF generado: {ruta_pdf}")

    return {"xlsx": ruta_xlsx, "pdf": ruta_pdf}


def main(argv=None):
    """Punto de entrada con validación de argumentos y códigos de salida."""
    analizador = argparse.ArgumentParser(
        description="Genera el reporte diario de faena en Excel y PDF."
    )
    analizador.add_argument(
        "--reporte",
        default=str(DIRECTORIO_BASE / "reporte_jefe.txt"),
        help="Ruta del txt con las notas del jefe de faena.",
    )
    analizador.add_argument(
        "--horas",
        default=str(DIRECTORIO_BASE / "horas_cuadrillas.xlsx"),
        help="Ruta del Excel con las horas-hombre de las cuadrillas.",
    )
    analizador.add_argument(
        "--fecha",
        required=True,
        help="Fecha del reporte en formato AAAA-MM-DD.",
    )
    analizador.add_argument(
        "--salida",
        default=".",
        help="Directorio donde guardar el reporte.",
    )
    analizador.add_argument(
        "--turno",
        default=None,
        help="Turno del reporte; si se omite se detecta del texto.",
    )
    analizador.add_argument(
        "--clima",
        default="Despejado, viento moderado",
        help="Clima del día que se anota en el documento.",
    )
    analizador.add_argument(
        "--estado-faena",
        default="Operativa",
        help="Estado de faena que se anota en el documento.",
    )
    argumentos = analizador.parse_args(argv)

    try:
        fecha = datetime.strptime(argumentos.fecha, "%Y-%m-%d").date()
    except ValueError:
        print(
            f"Error: fecha inválida '{argumentos.fecha}'. "
            "Use el formato AAAA-MM-DD (por ejemplo, 2026-11-23).",
            file=sys.stderr,
        )
        raise SystemExit(2)

    try:
        return generar_reporte(
            argumentos.reporte,
            argumentos.horas,
            fecha,
            argumentos.salida,
            argumentos.turno,
            argumentos.clima,
            argumentos.estado_faena,
        )
    except (FileNotFoundError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        raise SystemExit(2)


if __name__ == "__main__":
    main()
