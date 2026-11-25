"""Emite certificados individuales (Excel y PDF) con folio desde un CSV.

Uso::

    python generar_certificados.py --input participantes.csv --output salida

Cada certificado se genera a partir de la plantilla de su tipo, recibe un folio
consecutivo (CERT-<año>-NNNN o DIST-<año>-NNNN) y queda anotado en
``registro_emision.xlsx`` para poder seguir la numeración en próximas corridas.
"""

import argparse
from datetime import date, datetime
from pathlib import Path
import re
import shutil
import subprocess
import unicodedata
import warnings

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

from plantillas import MARCADORES, PALETAS, construir_todas


BASE_DIR = Path(__file__).resolve().parent
PARTICIPANTES_PATH = BASE_DIR / "participantes.csv"
PLANTILLAS_DIR = BASE_DIR / "plantillas"

COLUMNAS = ["nombre", "curso", "fecha", "nota", "tipo"]
COLUMNAS_REGISTRO = ["folio", "nombre", "curso", "fecha", "fecha_emision", "archivo"]
TIPOS = {"participacion": "CERT", "distincion": "DIST"}
ARCHIVOS_PLANTILLA = {
    "participacion": "certificado_participacion.xlsx",
    "distincion": "certificado_distincion.xlsx",
}
MESES = [
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
]


def leer_participantes(ruta):
    """Lee el CSV de participantes y devuelve un DataFrame sin modificar."""
    ruta = Path(ruta)
    if not ruta.is_file():
        raise SystemExit(
            f"El archivo de participantes no existe: '{ruta}'. "
            "Indica una ruta válida con --input."
        )
    try:
        return pd.read_csv(ruta, encoding="utf-8-sig")
    except pd.errors.EmptyDataError as error:
        raise SystemExit(
            f"El archivo '{ruta}' está vacío o no contiene filas de participantes. "
            "Agrega el encabezado y al menos una fila."
        ) from error
    except Exception as error:
        raise SystemExit(
            f"No se pudo leer el archivo de participantes '{ruta}'. "
            f"Comprueba que sea un CSV válido con separador coma. Detalle: {error}"
        ) from error


def _normalizar_tipo(valor):
    if pd.isna(valor):
        return None
    texto = unicodedata.normalize("NFKD", str(valor))
    texto = texto.encode("ascii", "ignore").decode("ascii").strip().lower()
    return texto if texto in TIPOS else None


def validar(df, ruta):
    """Valida las filas del CSV y devuelve una copia normalizada."""
    faltantes = [columna for columna in COLUMNAS if columna not in df.columns]
    if faltantes:
        raise SystemExit(
            f"El archivo '{ruta}' no tiene todas las columnas requeridas "
            f"(fila 1, encabezado). Faltan: {', '.join(faltantes)}. "
            f"Columnas requeridas: {', '.join(COLUMNAS)}."
        )
    if df.empty:
        raise SystemExit(
            f"El archivo '{ruta}' no contiene filas de participantes. "
            "Agrega al menos una fila con nombre, curso, fecha, nota y tipo."
        )

    filas = []
    for numero, (_, participante) in enumerate(df.iterrows(), start=2):
        nombre = participante["nombre"]
        curso = participante["curso"]
        if pd.isna(nombre) or not str(nombre).strip():
            raise SystemExit(
                f"El archivo '{ruta}' tiene el campo 'nombre' vacío en la fila "
                f"{numero}. Completa ese dato e inténtalo de nuevo."
            )
        if pd.isna(curso) or not str(curso).strip():
            raise SystemExit(
                f"El archivo '{ruta}' tiene el campo 'curso' vacío en la fila "
                f"{numero}. Completa ese dato e inténtalo de nuevo."
            )

        fecha_valor = participante["fecha"]
        if pd.isna(fecha_valor):
            raise SystemExit(
                f"El archivo '{ruta}' tiene la fecha vacía en la fila {numero} "
                "(columna 'fecha'). Usa un formato como 2026-03-15."
            )
        try:
            if isinstance(fecha_valor, str) and re.fullmatch(
                r"\d{4}-\d{2}-\d{2}", fecha_valor.strip()
            ):
                fecha = date.fromisoformat(fecha_valor.strip())
            else:
                with warnings.catch_warnings():
                    warnings.filterwarnings(
                        "ignore",
                        message="Parsing dates in %Y-%m-%d format when dayfirst=True",
                        category=UserWarning,
                    )
                    fecha = pd.to_datetime(
                        fecha_valor, dayfirst=True, errors="raise"
                    ).date()
        except (TypeError, ValueError) as error:
            raise SystemExit(
                f"El archivo '{ruta}' tiene una fecha inválida en la fila {numero} "
                f"(columna 'fecha'): {fecha_valor!r}. Usa un formato como 2026-03-15."
            ) from error

        nota_valor = participante["nota"]
        if pd.isna(nota_valor):
            raise SystemExit(
                f"El archivo '{ruta}' tiene la nota vacía o no numérica en la fila "
                f"{numero} (columna 'nota'). Usa un número como 6.5."
            )
        try:
            nota = float(nota_valor)
        except (TypeError, ValueError) as error:
            raise SystemExit(
                f"El archivo '{ruta}' tiene una nota no numérica en la fila {numero} "
                f"(columna 'nota'): {nota_valor!r}. Usa un número como 6.5."
            ) from error

        tipo = _normalizar_tipo(participante["tipo"])
        if tipo is None:
            raise SystemExit(
                f"El archivo '{ruta}' tiene un tipo inválido en la fila {numero} "
                f"(columna 'tipo'): {participante['tipo']!r}. "
                "Usa 'participacion' o 'distincion'."
            )

        filas.append(
            {
                "nombre": str(nombre).strip(),
                "curso": str(curso).strip(),
                "fecha": fecha,
                "nota": nota,
                "tipo": tipo,
            }
        )
    return pd.DataFrame(filas, columns=COLUMNAS)


def fecha_larga(fecha):
    """Devuelve la fecha como '15 de marzo de 2026' sin depender del locale."""
    return f"{fecha.day} de {MESES[fecha.month - 1]} de {fecha.year}"


def slug(nombre):
    """Convierte un nombre en un slug ascii seguro para el nombre de archivo."""
    texto = unicodedata.normalize("NFKD", str(nombre))
    texto = texto.encode("ascii", "ignore").decode("ascii").lower()
    return re.sub(r"[^a-z0-9]+", "_", texto).strip("_") or "participante"


def siguiente_folio(contadores, tipo, año):
    """Incrementa el contador del tipo y devuelve el folio correspondiente."""
    contadores[tipo] = contadores.get(tipo, 0) + 1
    return f"{TIPOS[tipo]}-{año}-{contadores[tipo]:04d}"


def preparar_plantillas(directorio=None):
    """Valida el directorio de plantillas (o lo crea si es el de por defecto)."""
    explicito = directorio is not None
    directorio = Path(directorio) if explicito else PLANTILLAS_DIR
    if not directorio.is_dir():
        if explicito:
            raise SystemExit(
                f"El directorio de plantillas no existe: '{directorio}'. "
                "Indica una carpeta válida con --plantilla."
            )
        construir_todas(directorio)

    faltantes = [
        nombre
        for nombre in ARCHIVOS_PLANTILLA.values()
        if not (directorio / nombre).is_file()
    ]
    if faltantes:
        raise SystemExit(
            f"El directorio de plantillas '{directorio}' no tiene los archivos "
            f"requeridos: {', '.join(faltantes)}. Genera las plantillas con "
            "plantillas.py."
        )
    return {
        tipo: directorio / nombre for tipo, nombre in ARCHIVOS_PLANTILLA.items()
    }


def emitir_certificado(plantilla, salida, nombre, curso, fecha, nota, folio):
    """Llena la plantilla con los datos y guarda el xlsx del certificado."""
    libro = load_workbook(plantilla)
    valores = {
        "{{NOMBRE}}": nombre,
        "{{CURSO}}": curso,
        "{{FECHA}}": fecha_larga(fecha),
        "{{NOTA}}": f"{nota:.1f}",
        "{{FOLIO}}": folio,
    }
    for hoja in libro.worksheets:
        for fila in hoja.iter_rows():
            for celda in fila:
                if not isinstance(celda.value, str):
                    continue
                texto = celda.value
                for marcador in MARCADORES:
                    if marcador in texto:
                        texto = texto.replace(marcador, valores[marcador])
                celda.value = texto

    destino = Path(salida) / f"{folio}_{slug(nombre)}.xlsx"
    libro.save(destino)
    return destino


def convertir_a_pdf(xlsx, salida):
    """Convierte el xlsx indicado a PDF con LibreOffice y comprueba el resultado."""
    if shutil.which("soffice") is None:
        raise SystemExit(
            "No se encontró LibreOffice (comando 'soffice'). Instálalo o agrega "
            "'soffice' al PATH para generar los PDF."
        )
    try:
        subprocess.run(
            [
                "soffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(salida),
                str(xlsx),
            ],
            check=True,
            capture_output=True,
        )
    except subprocess.CalledProcessError as error:
        detalle = (error.stderr or b"").decode("utf-8", "replace").strip()
        raise SystemExit(
            f"No se pudo convertir '{Path(xlsx).name}' a PDF. Detalle: {detalle}"
        ) from error

    pdf = Path(salida) / f"{Path(xlsx).stem}.pdf"
    if not pdf.is_file() or pdf.stat().st_size == 0:
        raise SystemExit(
            f"LibreOffice no generó el PDF de '{Path(xlsx).name}' o quedó vacío."
        )
    return pdf


def guardar_registro(registros, ruta):
    """Escribe y formatea el registro de emisión con una fila por certificado."""
    libro = Workbook()
    hoja = libro.active
    hoja.title = "Emisión"
    hoja.append(COLUMNAS_REGISTRO)
    for registro in registros:
        hoja.append([registro[columna] for columna in COLUMNAS_REGISTRO])

    relleno = PatternFill(
        fill_type="solid", start_color="16304F", end_color="16304F"
    )
    for celda in hoja[1]:
        celda.fill = relleno
        celda.font = Font(color="FFFFFF", bold=True)

    hoja.auto_filter.ref = hoja.dimensions
    hoja.freeze_panes = "A2"
    anchos = {"A": 16, "B": 26, "C": 36, "D": 14, "E": 16, "F": 34}
    for columna, ancho in anchos.items():
        hoja.column_dimensions[columna].width = ancho
    for fila in range(2, hoja.max_row + 1):
        hoja.cell(fila, 4).number_format = "dd-mm-yyyy"
        hoja.cell(fila, 5).number_format = "dd-mm-yyyy hh:mm"

    libro.save(ruta)


def leer_registro(ruta):
    """Carga el registro previo como lista de diccionarios (vacío si no existe)."""
    ruta = Path(ruta)
    if not ruta.is_file():
        return []
    try:
        libro = load_workbook(ruta, data_only=True)
    except Exception as error:
        raise SystemExit(
            f"No se pudo leer el registro previo '{ruta}'. Corrige o elimina ese "
            f"archivo e inténtalo de nuevo. Detalle: {error}"
        ) from error

    filas = libro.active.iter_rows(values_only=True)
    encabezados = next(filas, None)
    if not encabezados:
        return []
    return [
        dict(zip(encabezados, fila))
        for fila in filas
        if any(valor is not None for valor in fila)
    ]


def _contadores_desde(registros, año):
    """Obtiene el último número de folio de cada tipo para el año indicado."""
    contadores = {}
    for registro in registros:
        coincidencia = re.fullmatch(
            r"(CERT|DIST)-(\d{4})-(\d+)", str(registro.get("folio", ""))
        )
        if not coincidencia or int(coincidencia.group(2)) != año:
            continue
        tipo = next(
            tipo for tipo, prefijo in TIPOS.items() if prefijo == coincidencia.group(1)
        )
        contadores[tipo] = max(
            contadores.get(tipo, 0), int(coincidencia.group(3))
        )
    return contadores


def main(argv=None):
    """Ejecuta la emisión completa y devuelve un resumen de lo generado."""
    parser = argparse.ArgumentParser(
        description=(
            "Genera certificados individuales en Excel y PDF con folio, y un "
            "registro de emisión."
        ),
        epilog=(
            "Sin argumentos usa participantes.csv y ./plantillas junto a este script."
        ),
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=PARTICIPANTES_PATH,
        metavar="CSV",
        help="CSV de participantes (por defecto: participantes.csv junto al script).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("."),
        metavar="DIRECTORIO",
        help="Directorio de salida (por defecto: directorio actual).",
    )
    parser.add_argument(
        "--plantilla",
        type=Path,
        default=None,
        metavar="DIRECTORIO",
        help="Directorio con las plantillas (por defecto: ./plantillas junto al script).",
    )
    parser.add_argument(
        "--pdf",
        choices=["auto", "si", "no"],
        default="auto",
        help=(
            "Generar PDF con LibreOffice: auto (por defecto: sí si soffice "
            "existe, con aviso si no), si (exigir) o no (omitir)."
        ),
    )
    args = parser.parse_args(argv)

    participantes = validar(leer_participantes(args.input), args.input)
    plantillas = preparar_plantillas(args.plantilla)

    tiene_soffice = shutil.which("soffice") is not None
    if args.pdf == "si" and not tiene_soffice:
        raise SystemExit(
            "No se encontró LibreOffice (comando 'soffice'). Instálalo o agrega "
            "'soffice' al PATH para generar los PDF."
        )
    generar_pdf = args.pdf == "si" or (args.pdf == "auto" and tiene_soffice)
    if not generar_pdf:
        print(
            "LibreOffice no disponible (o --pdf no): se omiten los PDF; "
            "los certificados se emiten solo en Excel."
        )

    salida = Path(args.output)
    if salida.exists() and not salida.is_dir():
        raise SystemExit(
            f"La ruta de salida no es un directorio: '{salida}'. "
            "Indica una carpeta válida con --output."
        )
    try:
        salida.mkdir(parents=True, exist_ok=True)
    except OSError as error:
        raise SystemExit(
            f"No se pudo crear el directorio de salida '{salida}': {error}. "
            "Comprueba la ruta y los permisos de escritura."
        ) from error

    registro_ruta = salida / "registro_emision.xlsx"
    previos = leer_registro(registro_ruta)
    año = date.today().year
    contadores = _contadores_desde(previos, año)
    ahora = datetime.now()

    folios = []
    archivos = []
    registros = []
    total = len(participantes)
    for numero, participante in enumerate(
        participantes.itertuples(index=False), start=1
    ):
        folio = siguiente_folio(contadores, participante.tipo, año)
        xlsx = emitir_certificado(
            plantillas[participante.tipo],
            salida,
            participante.nombre,
            participante.curso,
            participante.fecha,
            participante.nota,
            folio,
        )
        if generar_pdf:
            convertir_a_pdf(xlsx, salida)

        folios.append(folio)
        archivos.append(xlsx)
        registros.append(
            {
                "folio": folio,
                "nombre": participante.nombre,
                "curso": participante.curso,
                "fecha": participante.fecha,
                "fecha_emision": ahora,
                "archivo": xlsx.name,
            }
        )
        print(f"certificado {numero}/{total}: {xlsx.name} ({folio})")

    guardar_registro(previos + registros, registro_ruta)
    print(f"Total de certificados emitidos: {total}")
    print(f"Registro de emisión: {registro_ruta}")
    print(f"Directorio de salida: {salida.resolve()}")
    return {
        "total": total,
        "folios": folios,
        "registro": registro_ruta,
        "archivos": archivos,
    }


if __name__ == "__main__":
    main()
