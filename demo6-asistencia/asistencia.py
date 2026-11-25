"""Analizador de asistencia y turnos.

Cruza el reloj de asistencia del mes (reloj_asistencia.csv) con los turnos
esperados (turnos_esperados.csv) y genera un resumen Excel auditable con
tardanzas, salidas antes, horas extra, ausencias, olvidos y cobertura diaria.
"""

import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DIRECTORIO_BASE = Path(__file__).resolve().parent

COLUMNAS_RELOJ = [
    "empleado",
    "fecha",
    "turno",
    "estado",
    "hora_entrada",
    "hora_salida",
]
COLUMNAS_TURNOS = [
    "turno",
    "hora_entrada_ideal",
    "hora_salida_ideal",
    "tolerancia_min",
]
ESTADOS = {"Asistió", "Ausencia", "Olvido marcaje"}
PRESENTES = {"Asistió", "Olvido marcaje"}
RE_MES = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
RE_HORA = re.compile(r"^\d{2}:\d{2}$")
RE_FECHA = re.compile(r"^\d{4}-\d{2}-\d{2}$")

AZUL_OSCURO = "1F4E78"
RELLENO_ENCABEZADO = PatternFill("solid", fgColor=AZUL_OSCURO)
RELLENO_TOTALES = PatternFill("solid", fgColor="D9D9D9")
RELLENO_ESTADO = {
    "OK": PatternFill("solid", fgColor="C6EFCE"),
    "DEFICIT": PatternFill("solid", fgColor="FFC7CE"),
}
BORDE_FINO = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
ALINEACION_CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _parsear_hora(valor, fila=None):
    """Convierte 'HH:MM' a minutos desde la medianoche y valida el formato."""
    texto = str(valor).strip()
    ubicacion = f"Fila {fila}: " if fila is not None else ""
    if not RE_HORA.fullmatch(texto):
        raise ValueError(f"{ubicacion}hora inválida '{valor}' (se espera HH:MM).")
    try:
        momento = datetime.strptime(texto, "%H:%M")
    except ValueError:
        raise ValueError(f"{ubicacion}hora inválida '{valor}' (se espera HH:MM).") from None
    return momento.hour * 60 + momento.minute


def _parsear_fecha(valor, fila=None):
    """Convierte 'AAAA-MM-DD' a date y valida el formato."""
    texto = str(valor).strip()
    ubicacion = f"Fila {fila}: " if fila is not None else ""
    if not RE_FECHA.fullmatch(texto):
        raise ValueError(
            f"{ubicacion}fecha inválida '{valor}' (se espera AAAA-MM-DD)."
        )
    try:
        return datetime.strptime(texto, "%Y-%m-%d").date()
    except ValueError:
        raise ValueError(
            f"{ubicacion}fecha inválida '{valor}' (se espera AAAA-MM-DD)."
        ) from None


def minutos_tardanza(hora_real, hora_ideal, tolerancia_min):
    """Minutos de atraso sobre la tolerancia del turno; 0 si llega a tiempo."""
    atraso = _parsear_hora(hora_real) - _parsear_hora(hora_ideal)
    return max(0, atraso - int(tolerancia_min))


def minutos_horas_extra(hora_salida_real, hora_salida_ideal, gracia_min=15):
    """Minutos extra solo si la salida supera la gracia; si no, 0."""
    # shortcut: asume que la salida extra cae el mismo día; si el turno cruza
    # medianoche, hay que incorporar la fecha de salida para desambiguar.
    diferencia = _parsear_hora(hora_salida_real) - _parsear_hora(hora_salida_ideal)
    return diferencia if diferencia > int(gracia_min) else 0


def minutos_salida_antes(hora_real, hora_salida_ideal, tolerancia_min):
    """Minutos de salida anticipada sobre la tolerancia; 0 si sale a tiempo."""
    adelanto = _parsear_hora(hora_salida_ideal) - _parsear_hora(hora_real)
    return max(0, adelanto - int(tolerancia_min))


def minutos_trabajados(hora_entrada, hora_salida):
    """Minutos trabajados; suma 24 h cuando el turno cruza la medianoche."""
    diferencia = _parsear_hora(hora_salida) - _parsear_hora(hora_entrada)
    if diferencia <= 0:
        diferencia += 24 * 60
    return diferencia


def _validar_columnas(datos, esperadas, nombre_archivo):
    """Falla con un mensaje claro si al CSV le falta alguna columna."""
    faltantes = [columna for columna in esperadas if columna not in datos.columns]
    if faltantes:
        raise ValueError(
            f"{nombre_archivo} no tiene las columnas esperadas: "
            + ", ".join(faltantes)
        )


def leer_turnos(ruta):
    """Lee los turnos esperados y devuelve un diccionario por nombre de turno."""
    ruta = Path(ruta)
    if not ruta.is_file():
        raise FileNotFoundError(f"No se encontró el archivo de turnos: {ruta}")

    datos = pd.read_csv(ruta, dtype=str, keep_default_na=False)
    _validar_columnas(datos, COLUMNAS_TURNOS, ruta.name)

    turnos = {}
    for numero, fila in enumerate(datos.itertuples(index=False), start=2):
        turno = str(fila.turno).strip()
        if not turno:
            raise ValueError(f"Fila {numero}: turno vacío en {ruta.name}.")
        _parsear_hora(fila.hora_entrada_ideal, numero)
        _parsear_hora(fila.hora_salida_ideal, numero)
        try:
            tolerancia = int(str(fila.tolerancia_min).strip())
        except ValueError:
            raise ValueError(
                f"Fila {numero}: tolerancia inválida "
                f"'{fila.tolerancia_min}' (se espera un entero)."
            ) from None
        if tolerancia < 0:
            raise ValueError(f"Fila {numero}: la tolerancia no puede ser negativa.")
        turnos[turno] = {
            "hora_entrada_ideal": str(fila.hora_entrada_ideal).strip(),
            "hora_salida_ideal": str(fila.hora_salida_ideal).strip(),
            "tolerancia_min": tolerancia,
        }
    return turnos


def leer_reloj(ruta, mes):
    """Lee y valida el reloj, y deja solo los registros del mes pedido."""
    ruta = Path(ruta)
    if not ruta.is_file():
        raise FileNotFoundError(f"No se encontró el archivo de reloj: {ruta}")
    if not RE_MES.fullmatch(str(mes)):
        raise ValueError(
            f"Mes inválido '{mes}'. Use el formato AAAA-MM (por ejemplo, 2026-11)."
        )

    datos = pd.read_csv(ruta, dtype=str, keep_default_na=False)
    _validar_columnas(datos, COLUMNAS_RELOJ, ruta.name)

    registros = []
    for numero, fila in enumerate(datos.itertuples(index=False), start=2):
        fecha = _parsear_fecha(fila.fecha, numero)
        estado = str(fila.estado).strip()
        if estado not in ESTADOS:
            raise ValueError(
                f"Fila {numero}: estado desconocido '{fila.estado}' "
                f"(se espera {', '.join(sorted(ESTADOS))})."
            )
        if fecha.strftime("%Y-%m") != mes:
            continue
        if estado == "Asistió":
            _parsear_hora(fila.hora_entrada, numero)
            _parsear_hora(fila.hora_salida, numero)
        else:
            if str(fila.hora_entrada).strip():
                _parsear_hora(fila.hora_entrada, numero)
            if str(fila.hora_salida).strip():
                _parsear_hora(fila.hora_salida, numero)
        registros.append(
            {
                "empleado": str(fila.empleado).strip(),
                "fecha": fecha,
                "turno": str(fila.turno).strip(),
                "estado": estado,
                "hora_entrada": str(fila.hora_entrada).strip(),
                "hora_salida": str(fila.hora_salida).strip(),
                "fila_csv": numero,
            }
        )

    if not registros:
        raise ValueError(f"No hay registros para el mes {mes} en {ruta.name}.")
    return pd.DataFrame(registros)


def calcular_registros(reloj, turnos):
    """Agrega a cada registro la tardanza, salida antes y horas extra."""
    registros = reloj.copy()
    tardanzas, salidas_antes, extras = [], [], []
    entradas_ideales = []

    for fila in registros.itertuples(index=False):
        turno = turnos.get(fila.turno)
        if turno is None:
            raise ValueError(
                f"Fila {fila.fila_csv}: turno desconocido '{fila.turno}' "
                f"para {fila.empleado}."
            )
        entradas_ideales.append(turno["hora_entrada_ideal"])
        if fila.estado != "Asistió":
            tardanzas.append(0)
            salidas_antes.append(0)
            extras.append(0)
            continue
        tardanzas.append(
            minutos_tardanza(
                fila.hora_entrada,
                turno["hora_entrada_ideal"],
                turno["tolerancia_min"],
            )
        )
        salidas_antes.append(
            minutos_salida_antes(
                fila.hora_salida,
                turno["hora_salida_ideal"],
                turno["tolerancia_min"],
            )
        )
        extras.append(
            minutos_horas_extra(fila.hora_salida, turno["hora_salida_ideal"])
        )

    registros["entrada_ideal"] = entradas_ideales
    registros["tardanza_min"] = tardanzas
    registros["salida_antes_min"] = salidas_antes
    registros["extra_min"] = extras
    return registros


def resumen_por_persona(registros):
    """Resume por empleado: tardanzas, extras, ausencias y olvidos."""
    resumen = (
        registros.groupby(["empleado", "turno"], sort=True)
        .agg(
            dias=("fecha", "count"),
            tardanzas=("tardanza_min", lambda serie: int((serie > 0).sum())),
            minutos_tardanza=("tardanza_min", "sum"),
            salidas_antes=("salida_antes_min", lambda serie: int((serie > 0).sum())),
            minutos_extra=("extra_min", "sum"),
            ausencias=("estado", lambda serie: int((serie == "Ausencia").sum())),
            olvidos=("estado", lambda serie: int((serie == "Olvido marcaje").sum())),
        )
        .reset_index()
    )
    resumen["horas_extra"] = (resumen["minutos_extra"] / 60).round(2)
    return resumen


def detalle_tardanzas(registros):
    """Una fila por tardanza efectiva, ordenada por empleado y fecha."""
    detalle = registros[registros["tardanza_min"] > 0].copy()
    detalle = detalle.sort_values(["empleado", "fecha"])
    return detalle[
        [
            "empleado",
            "fecha",
            "turno",
            "hora_entrada",
            "entrada_ideal",
            "tardanza_min",
        ]
    ]


def resumen_extras(registros):
    """Horas extra y días con extra por empleado, de mayor a menor."""
    con_extra = registros[registros["extra_min"] > 0]
    extras = (
        con_extra.groupby(["empleado", "turno"], sort=True)
        .agg(minutos=("extra_min", "sum"), dias=("fecha", "count"))
        .reset_index()
    )
    extras["horas"] = (extras["minutos"] / 60).round(2)
    return extras.sort_values(["horas", "empleado"], ascending=[False, True])


def calcular_cobertura(registros, minimo=25):
    """Dotación presente por día, con déficit primero y el resto por fecha."""
    if int(minimo) < 1:
        raise ValueError("El mínimo exigido debe ser un entero mayor o igual a 1.")
    cobertura = (
        registros.groupby("fecha", sort=True)
        .agg(presentes=("estado", lambda serie: int(serie.isin(PRESENTES).sum())))
        .reset_index()
    )
    cobertura["minimo"] = int(minimo)
    cobertura["estado"] = [
        "OK" if presentes >= int(minimo) else "DEFICIT"
        for presentes in cobertura["presentes"]
    ]
    cobertura["_orden"] = (cobertura["estado"] == "OK").astype(int)
    return cobertura.sort_values(["_orden", "fecha"]).drop(columns="_orden")


NOMBRES_MES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]


def _nombre_periodo(mes):
    """Convierte 'AAAA-MM' a texto legible, por ejemplo 'Noviembre 2026'."""
    anio, numero = mes.split("-")
    return f"{NOMBRES_MES[int(numero) - 1]} {anio}"


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
    """Aplica borde fino y centrado a todas las celdas de una fila."""
    for celda in fila_celdas:
        celda.border = BORDE_FINO
        celda.alignment = ALINEACION_CENTRO


NOTA_METODOLOGICA = (
    "Reglas de cálculo: la tardanza son los minutos sobre la hora de entrada "
    "ideal descontada la tolerancia del turno.",
    "Horas extra: se abonan solo si la salida supera la hora de salida ideal "
    "en más de 15 minutos de gracia.",
)


def _hoja_resumen(libro, resumen, mes):
    """Construye la hoja Resumen por Persona con título, escala de color y nota."""
    hoja = libro.active
    hoja.title = "Resumen por Persona"
    hoja.sheet_view.showGridLines = False

    columnas = [
        "Empleado",
        "Turno",
        "Días registrados",
        "Llegadas tarde",
        "Minutos de tardanza",
        "Salidas antes",
        "HH extra (horas)",
        "Ausencias",
        "Olvidos de marcaje",
    ]
    ultima_columna = len(columnas)

    hoja.merge_cells(
        start_row=1, start_column=1, end_row=1, end_column=ultima_columna
    )
    titulo = hoja["A1"]
    titulo.value = f"Resumen de Asistencia — {_nombre_periodo(mes)}"
    titulo.font = Font(bold=True, size=16, color="FFFFFF")
    titulo.fill = RELLENO_ENCABEZADO
    titulo.alignment = ALINEACION_CENTRO
    hoja.row_dimensions[1].height = 36

    _estilar_encabezado(hoja, 2, columnas)

    fila = 3
    for datos in resumen.itertuples(index=False):
        valores = [
            datos.empleado,
            datos.turno,
            int(datos.dias),
            int(datos.tardanzas),
            int(datos.minutos_tardanza),
            int(datos.salidas_antes),
            float(datos.horas_extra),
            int(datos.ausencias),
            int(datos.olvidos),
        ]
        for columna, valor in enumerate(valores, start=1):
            hoja.cell(row=fila, column=columna, value=valor)
        hoja.cell(row=fila, column=7).number_format = "0.00"
        _bordear(hoja[fila])
        fila += 1

    ultima_fila_datos = fila - 1
    totales = [
        "TOTALES",
        "",
        int(resumen["dias"].sum()),
        int(resumen["tardanzas"].sum()),
        int(resumen["minutos_tardanza"].sum()),
        int(resumen["salidas_antes"].sum()),
        round(resumen["horas_extra"].sum(), 2),
        int(resumen["ausencias"].sum()),
        int(resumen["olvidos"].sum()),
    ]
    for columna, valor in enumerate(totales, start=1):
        celda = hoja.cell(row=fila, column=columna, value=valor)
        celda.font = Font(bold=True)
        celda.fill = RELLENO_TOTALES
    hoja.cell(row=fila, column=7).number_format = "0.00"
    _bordear(hoja[fila])

    if ultima_fila_datos >= 3:
        hoja.conditional_formatting.add(
            f"E3:E{ultima_fila_datos}",
            ColorScaleRule(
                start_type="min",
                start_color="FFFFFF",
                end_type="max",
                end_color="F8696B",
            ),
        )
        hoja.conditional_formatting.add(
            f"G3:G{ultima_fila_datos}",
            ColorScaleRule(
                start_type="min",
                start_color="FFFFFF",
                end_type="max",
                end_color="63BE7B",
            ),
        )

    hoja.freeze_panes = "A3"
    hoja.auto_filter.ref = f"A2:I{ultima_fila_datos}"
    _fijar_anchos(hoja, [12, 10, 11, 11, 13, 11, 12, 10, 12])

    fila += 2  # fila en blanco antes de la nota metodológica
    for linea in NOTA_METODOLOGICA:
        hoja.merge_cells(
            start_row=fila, start_column=1, end_row=fila, end_column=ultima_columna
        )
        nota = hoja.cell(row=fila, column=1, value=linea)
        nota.font = Font(size=8, italic=True, color="808080")
        nota.alignment = Alignment(horizontal="left", vertical="center")
        fila += 1


def _hoja_tardanzas(libro, detalle):
    """Construye el detalle auditable de tardanzas del mes."""
    hoja = libro.create_sheet("Tardanzas del Mes")
    hoja.sheet_view.showGridLines = False
    columnas = [
        "Empleado",
        "Fecha",
        "Turno",
        "Hora de entrada real",
        "Entrada ideal",
        "Minutos de tardanza",
    ]
    _estilar_encabezado(hoja, 1, columnas)

    if detalle.empty:
        hoja.merge_cells("A2:F2")
        hoja["A2"] = "Sin tardanzas registradas en el mes."
        hoja["A2"].alignment = ALINEACION_CENTRO
        hoja.freeze_panes = "A2"
        _fijar_anchos(hoja, [12, 12, 10, 17, 12, 17])
        return

    fila = 2
    for datos in detalle.itertuples(index=False):
        valores = [
            datos.empleado,
            datos.fecha,
            datos.turno,
            datos.hora_entrada,
            datos.entrada_ideal,
            int(datos.tardanza_min),
        ]
        for columna, valor in enumerate(valores, start=1):
            hoja.cell(row=fila, column=columna, value=valor)
        hoja.cell(row=fila, column=2).number_format = "DD-MM-YYYY"
        _bordear(hoja[fila])
        fila += 1

    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = f"A1:F{fila - 1}"
    _fijar_anchos(hoja, [12, 12, 10, 17, 12, 17])


def _hoja_extras(libro, extras):
    """Construye la hoja de horas extra ordenada de mayor a menor."""
    hoja = libro.create_sheet("Extras")
    hoja.sheet_view.showGridLines = False
    columnas = ["Empleado", "Turno", "Horas extra del mes", "Días con extra"]
    _estilar_encabezado(hoja, 1, columnas)

    if extras.empty:
        hoja.merge_cells("A2:D2")
        hoja["A2"] = "Sin horas extra registradas en el mes."
        hoja["A2"].alignment = ALINEACION_CENTRO
        hoja.freeze_panes = "A2"
        _fijar_anchos(hoja, [12, 10, 15, 13])
        return

    fila = 2
    for datos in extras.itertuples(index=False):
        valores = [
            datos.empleado,
            datos.turno,
            float(datos.horas),
            int(datos.dias),
        ]
        for columna, valor in enumerate(valores, start=1):
            hoja.cell(row=fila, column=columna, value=valor)
        hoja.cell(row=fila, column=3).number_format = "0.00"
        _bordear(hoja[fila])
        fila += 1

    totales = ["TOTALES", "", round(extras["horas"].sum(), 2), int(extras["dias"].sum())]
    for columna, valor in enumerate(totales, start=1):
        celda = hoja.cell(row=fila, column=columna, value=valor)
        celda.font = Font(bold=True)
        celda.fill = RELLENO_TOTALES
    hoja.cell(row=fila, column=3).number_format = "0.00"
    _bordear(hoja[fila])

    hoja.freeze_panes = "A2"
    _fijar_anchos(hoja, [12, 10, 15, 13])


def _hoja_cobertura(libro, cobertura):
    """Construye la cobertura diaria con déficit destacado al inicio."""
    hoja = libro.create_sheet("Cobertura")
    hoja.sheet_view.showGridLines = False
    columnas = ["Fecha", "Presentes", "Mínimo exigido", "Estado"]
    _estilar_encabezado(hoja, 1, columnas)

    fila = 2
    for datos in cobertura.itertuples(index=False):
        valores = [datos.fecha, int(datos.presentes), int(datos.minimo), datos.estado]
        for columna, valor in enumerate(valores, start=1):
            hoja.cell(row=fila, column=columna, value=valor)
        hoja.cell(row=fila, column=1).number_format = "DD-MM-YYYY"
        hoja.cell(row=fila, column=4).fill = RELLENO_ESTADO[datos.estado]
        _bordear(hoja[fila])
        fila += 1

    hoja.freeze_panes = "A2"
    _fijar_anchos(hoja, [12, 10, 14, 12])


def _construir_libro(resumen, detalle, extras, cobertura, mes):
    """Arma el libro Excel con las cuatro hojas en el orden pedido."""
    libro = Workbook()
    _hoja_resumen(libro, resumen, mes)
    _hoja_tardanzas(libro, detalle)
    _hoja_extras(libro, extras)
    _hoja_cobertura(libro, cobertura)
    return libro


def generar_resumen(ruta_reloj, ruta_turnos, mes, salida=".", minimo=25):
    """Genera el XLSX del mes y devuelve su ruta."""
    if not RE_MES.fullmatch(str(mes)):
        raise ValueError(
            f"Mes inválido '{mes}'. Use el formato AAAA-MM (por ejemplo, 2026-11)."
        )
    turnos = leer_turnos(ruta_turnos)
    reloj = leer_reloj(ruta_reloj, mes)
    registros = calcular_registros(reloj, turnos)

    resumen = resumen_por_persona(registros)
    detalle = detalle_tardanzas(registros)
    extras = resumen_extras(registros)
    cobertura = calcular_cobertura(registros, minimo)

    directorio_salida = Path(salida)
    directorio_salida.mkdir(parents=True, exist_ok=True)
    ruta_xlsx = directorio_salida / f"resumen_asistencia_{mes}.xlsx"

    libro = _construir_libro(resumen, detalle, extras, cobertura, mes)
    libro.save(ruta_xlsx)
    print(f"XLSX generado: {ruta_xlsx}")
    return ruta_xlsx


def main(argv=None):
    """Punto de entrada con validación de argumentos y códigos 0/2."""
    analizador = argparse.ArgumentParser(
        description="Analiza la asistencia del mes y genera un resumen en Excel."
    )
    analizador.add_argument(
        "--reloj",
        default=str(DIRECTORIO_BASE / "reloj_asistencia.csv"),
        help="Ruta del CSV con el reloj de asistencia.",
    )
    analizador.add_argument(
        "--turnos",
        default=str(DIRECTORIO_BASE / "turnos_esperados.csv"),
        help="Ruta del CSV con los turnos esperados.",
    )
    analizador.add_argument(
        "--mes",
        default="2026-11",
        help="Mes a analizar en formato AAAA-MM (por defecto, 2026-11).",
    )
    analizador.add_argument(
        "--salida",
        default=".",
        help="Directorio donde guardar el XLSX.",
    )
    analizador.add_argument(
        "--minimo",
        type=int,
        default=25,
        help="Dotación mínima diaria para marcar déficit (por defecto, 25).",
    )
    argumentos = analizador.parse_args(argv)

    if not RE_MES.fullmatch(argumentos.mes):
        print(
            f"Error: mes inválido '{argumentos.mes}'. "
            "Use el formato AAAA-MM (por ejemplo, 2026-11).",
            file=sys.stderr,
        )
        raise SystemExit(2)
    if argumentos.minimo < 1:
        print(
            "Error: el mínimo exigido debe ser un entero mayor o igual a 1.",
            file=sys.stderr,
        )
        raise SystemExit(2)

    try:
        generar_resumen(
            argumentos.reloj,
            argumentos.turnos,
            argumentos.mes,
            argumentos.salida,
            argumentos.minimo,
        )
    except (FileNotFoundError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        raise SystemExit(2)

    return 0


if __name__ == "__main__":
    main()
