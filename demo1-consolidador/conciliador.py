"""Conciliador de ventas multicanal (Local, Web y POS).

Normaliza las exportaciones a un esquema común, separa errores, elimina
duplicados y genera un dashboard en Excel.
"""

from __future__ import annotations

import argparse
import re
import sys
from datetime import date, datetime, time
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

COLUMNAS = ["canal", "folio", "fecha", "sku", "vendedor", "unidades", "monto"]

_CANALES: dict[str, dict[str, Any]] = {
    "local": {
        "nombre": "Local",
        "tipo": "excel",
        "columnas": {
            "folio": "Folio",
            "fecha": "Fecha",
            "sku": "SKU",
            "vendedor": "Vendedor",
            "unidades": "Unidades",
            "monto": "Monto",
        },
    },
    "web": {
        "nombre": "Web",
        "tipo": "excel",
        "columnas": {
            "folio": "Order ID",
            "fecha": "Order Date",
            "sku": "Product Code",
            "vendedor": "Sales Rep",
            "unidades": "Qty",
            "monto": "Revenue",
        },
    },
    "pos": {
        "nombre": "POS",
        "tipo": "csv",
        "columnas": {
            "folio": "ticket",
            "fecha": "timestamp",
            "sku": "sku",
            "vendedor": "cajero",
            "unidades": "cantidad",
            "monto": "total",
        },
    },
}

_ETIQUETAS_ERROR = {
    "fecha": "fecha inválida",
    "monto": "monto no numérico",
    "sku": "sku vacío",
    "unidades": "unidades no positivas",
}


def parsear_monto_chileno(valor: Any) -> float | None:
    """Convierte un monto chileno (``$1.234,56``) a float; None si no es numérico."""
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float, np.integer, np.floating)):
        return None if pd.isna(valor) else float(valor)
    if not isinstance(valor, str):
        return None

    texto = valor.strip()
    if not texto or texto.upper() in {"N/D", "N/A", "NA", "-"}:
        return None

    negativo = texto.startswith("-")
    texto = texto.lstrip("+-").replace("$", "").replace(" ", "").replace("\u00a0", "")
    if not texto:
        return None

    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")
    elif "." in texto and re.fullmatch(r"\d{1,3}(?:\.\d{3})+", texto):
        texto = texto.replace(".", "")

    if not re.fullmatch(r"\d+(?:\.\d+)?", texto):
        return None
    numero = float(texto)
    return -numero if negativo else numero


def normalizar_fecha(valor: Any) -> datetime | None:
    """Convierte fechas dd/mm/aaaa, ISO o datetime a datetime; None si es inválida."""
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(valor, pd.Timestamp):
        return valor.to_pydatetime()
    if isinstance(valor, datetime):
        return valor
    if isinstance(valor, date):
        return datetime.combine(valor, time.min)
    if not isinstance(valor, str):
        return None

    texto = valor.strip()
    if not texto:
        return None
    for formato in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(texto, formato)
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(texto)
    except ValueError:
        return None


def leer_canal(ruta: str | Path, canal: str) -> pd.DataFrame:
    """Lee una exportación de canal y la devuelve con el esquema común."""
    config = _CANALES.get(canal.strip().lower())
    if config is None:
        raise ValueError(f"Canal desconocido: {canal}. Use Local, Web o POS.")

    ruta = Path(ruta)
    nombre = config["nombre"]
    if not ruta.is_file():
        raise FileNotFoundError(f"No se encontró el archivo del canal {nombre}: {ruta}")

    try:
        if config["tipo"] == "excel":
            original = pd.read_excel(ruta)
        else:
            original = pd.read_csv(ruta, sep=";", encoding="utf-8-sig")
    except Exception as exc:
        raise ValueError(f"No se pudo leer el archivo {ruta}: {exc}") from exc

    columnas = config["columnas"]
    faltantes = [col for col in columnas.values() if col not in original.columns]
    if faltantes:
        raise ValueError(
            f"Al archivo {ruta} del canal {nombre} le faltan columnas: "
            + ", ".join(faltantes)
        )

    df = pd.DataFrame(index=original.index)
    df["canal"] = nombre
    df["folio"] = original[columnas["folio"]]
    df["fecha"] = pd.to_datetime(
        original[columnas["fecha"]].map(normalizar_fecha), errors="coerce"
    )
    df["sku"] = original[columnas["sku"]].astype("string").fillna("").str.strip()
    df["vendedor"] = (
        original[columnas["vendedor"]].astype("string").fillna("").str.strip()
    )
    df["unidades"] = (
        pd.to_numeric(original[columnas["unidades"]], errors="coerce")
        .fillna(0)
        .astype(int)
    )
    if nombre == "POS":
        montos = original[columnas["monto"]].map(parsear_monto_chileno)
    else:
        montos = pd.to_numeric(original[columnas["monto"]], errors="coerce")
    df["monto"] = pd.to_numeric(pd.Series(montos, index=original.index), errors="coerce").astype(
        float
    )
    return df[COLUMNAS]


def conciliar(
    ruta_local: str | Path, ruta_web: str | Path, ruta_pos: str | Path
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, Any]]:
    """Lee los tres canales, separa errores y elimina duplicados."""
    lecturas = {
        "Local": leer_canal(ruta_local, "Local"),
        "Web": leer_canal(ruta_web, "Web"),
        "POS": leer_canal(ruta_pos, "POS"),
    }
    filas_por_canal = {nombre: len(df) for nombre, df in lecturas.items()}
    df = pd.concat(lecturas.values(), ignore_index=True)

    condiciones = [
        df["fecha"].isna(),
        df["monto"].isna(),
        df["sku"].eq(""),
        df["unidades"] <= 0,
    ]
    etiquetas = [
        _ETIQUETAS_ERROR["fecha"],
        _ETIQUETAS_ERROR["monto"],
        _ETIQUETAS_ERROR["sku"],
        _ETIQUETAS_ERROR["unidades"],
    ]
    df["problema"] = np.select(condiciones, etiquetas, default="")
    mask_error = df["problema"] != ""
    df_errores = df.loc[mask_error, COLUMNAS + ["problema"]].reset_index(drop=True)
    df_limpio = df.loc[~mask_error, COLUMNAS].reset_index(drop=True)

    mask_exacto = df_limpio.duplicated(keep="first")
    duplicados_exactos = int(mask_exacto.sum())
    df_limpio = df_limpio.loc[~mask_exacto].reset_index(drop=True)

    df_limpio["_dia"] = df_limpio["fecha"].dt.normalize()
    df_limpio["_monto"] = df_limpio["monto"].round().astype(int)
    clave = ["sku", "_dia", "unidades", "_monto"]
    if df_limpio.empty:
        mask_cruzado = pd.Series(dtype=bool)
    else:
        canales = df_limpio.groupby(clave, dropna=False)["canal"].transform("nunique")
        mask_cruzado = (canales > 1) & df_limpio.duplicated(subset=clave, keep="first")
    duplicados_cruzados = int(mask_cruzado.sum())
    df_limpio = (
        df_limpio.loc[~mask_cruzado]
        .drop(columns=["_dia", "_monto"])
        .reset_index(drop=True)
    )

    ventas_totales = float(df_limpio["monto"].sum())
    transacciones = int(len(df_limpio))
    metricas: dict[str, Any] = {
        "filas_por_canal": filas_por_canal,
        "total_leido": int(sum(filas_por_canal.values())),
        "errores": int(len(df_errores)),
        "duplicados_exactos": duplicados_exactos,
        "duplicados_cruzados": duplicados_cruzados,
        "ventas_totales": ventas_totales,
        "transacciones": transacciones,
        "ticket_promedio": ventas_totales / transacciones if transacciones else 0.0,
        "unidades_totales": int(df_limpio["unidades"].sum()),
    }
    return df_limpio, df_errores, metricas


def _valor_celda(valor: Any) -> Any:
    """Adapta valores de pandas/numpy a tipos nativos de openpyxl."""
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(valor, pd.Timestamp):
        return valor.to_pydatetime()
    if isinstance(valor, np.generic):
        return valor.item()
    return valor


def _estilo_encabezado(celda) -> None:
    celda.font = Font(bold=True, color="FFFFFF")
    celda.fill = PatternFill(fill_type="solid", fgColor="305496")
    celda.alignment = Alignment(horizontal="center", vertical="center")


def _escribir_dataframe(ws, df: pd.DataFrame) -> None:
    columnas = list(df.columns)
    for columna, nombre in enumerate(columnas, start=1):
        _estilo_encabezado(ws.cell(row=1, column=columna, value=nombre))

    for fila, valores in enumerate(df.itertuples(index=False, name=None), start=2):
        for columna, valor in enumerate(valores, start=1):
            ws.cell(row=fila, column=columna, value=_valor_celda(valor))

    if "fecha" in columnas:
        columna_fecha = columnas.index("fecha") + 1
        for fila in range(2, len(df) + 2):
            ws.cell(row=fila, column=columna_fecha).number_format = "DD/MM/YYYY"
    if "monto" in columnas:
        columna_monto = columnas.index("monto") + 1
        for fila in range(2, len(df) + 2):
            ws.cell(row=fila, column=columna_monto).number_format = "$#,##0"

    ws.freeze_panes = "A2"
    ultima_columna = get_column_letter(len(columnas)) if columnas else "A"
    ws.auto_filter.ref = f"A1:{ultima_columna}{max(len(df) + 1, 1)}"
    for columna, nombre in enumerate(columnas, start=1):
        valores = [str(nombre)] + [
            str(_valor_celda(valor)) for valor in df[nombre].head(200)
        ]
        ws.column_dimensions[get_column_letter(columna)].width = min(
            max(len(valor) for valor in valores) + 2, 40
        )


def _periodo(df: pd.DataFrame) -> str:
    if df.empty or df["fecha"].isna().all():
        return "sin datos"
    return f"{df['fecha'].min():%Y-%m} a {df['fecha'].max():%Y-%m}"


_MESES_ES = (
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
)


def _nombre_mes(periodo: str) -> str:
    anio, mes = periodo.split("-")
    return f"{_MESES_ES[int(mes) - 1]} de {anio}"


def _formato_pct(valor: float) -> str:
    return f"{valor:.1f}%".replace(".", ",")


def _lectura_mes(df: pd.DataFrame) -> str:
    """Genera tres insights gerenciales calculados desde las ventas conciliadas."""
    if df.empty:
        return "No hay ventas conciliadas para generar la lectura del mes."

    mensual = (
        df.assign(mes=df["fecha"].dt.strftime("%Y-%m"))
        .groupby("mes")["monto"]
        .sum()
        .sort_index()
    )
    mes_actual = str(mensual.index[-1])
    monto_actual = float(mensual.iloc[-1])

    if len(mensual) >= 2:
        mes_previo = str(mensual.index[-2])
        monto_previo = float(mensual.iloc[-2])
        if monto_previo:
            variacion = (monto_actual - monto_previo) / monto_previo * 100
            if variacion > 0:
                cambio = (
                    f"un {_formato_pct(variacion)} más que "
                    f"{_nombre_mes(mes_previo)} ({_formato_clp(monto_previo)})"
                )
            elif variacion < 0:
                cambio = (
                    f"un {_formato_pct(abs(variacion))} menos que "
                    f"{_nombre_mes(mes_previo)} ({_formato_clp(monto_previo)})"
                )
            else:
                cambio = (
                    f"sin variación respecto de {_nombre_mes(mes_previo)} "
                    f"({_formato_clp(monto_previo)})"
                )
            frase_ventas = (
                f"Las ventas de {_nombre_mes(mes_actual)} totalizaron "
                f"{_formato_clp(monto_actual)}, {cambio}."
            )
        else:
            frase_ventas = (
                f"Las ventas de {_nombre_mes(mes_actual)} totalizaron "
                f"{_formato_clp(monto_actual)}, sin una base previa comparable."
            )

        por_canal = df.assign(mes=df["fecha"].dt.strftime("%Y-%m")).pivot_table(
            index="canal", columns="mes", values="monto", aggfunc="sum", fill_value=0.0
        )
        if mes_actual in por_canal.columns and mes_previo in por_canal.columns:
            diferencias = por_canal[mes_actual] - por_canal[mes_previo]
            if (diferencias < 0).any():
                canal = str(diferencias.idxmin())
                etiqueta = "mayor caída"
            else:
                canal = str(diferencias.idxmax())
                etiqueta = "mayor alza"
            base_canal = float(por_canal.loc[canal, mes_previo])
            actual_canal = float(por_canal.loc[canal, mes_actual])
            variacion_canal = (
                (actual_canal - base_canal) / base_canal * 100 if base_canal else 0.0
            )
            if variacion_canal < 0:
                movimiento = f"un {_formato_pct(abs(variacion_canal))} menos"
            elif variacion_canal > 0:
                movimiento = f"un {_formato_pct(variacion_canal)} más"
            else:
                movimiento = "sin variación"
            frase_canal = (
                f"El canal con {etiqueta} fue {canal}, que pasó de "
                f"{_formato_clp(base_canal)} a {_formato_clp(actual_canal)} "
                f"({movimiento})."
            )
        else:
            frase_canal = (
                "No hay un período anterior comparable para medir la "
                "evolución por canal."
            )
    else:
        frase_ventas = (
            f"Las ventas de {_nombre_mes(mes_actual)} totalizaron "
            f"{_formato_clp(monto_actual)}, sin un período anterior comparable."
        )
        frase_canal = (
            "No hay un período anterior comparable para medir la evolución "
            "por canal."
        )

    por_sku = df.groupby("sku")["monto"].sum()
    sku = str(por_sku.idxmax())
    monto_sku = float(por_sku.loc[sku])
    total = float(df["monto"].sum())
    participacion = monto_sku / total * 100 if total else 0.0
    frase_producto = (
        f"El producto estrella del período fue {sku}, con "
        f"{_formato_clp(monto_sku)} y una participación del "
        f"{_formato_pct(participacion)} del total."
    )
    return " ".join((frase_ventas, frase_canal, frase_producto))


def _crear_grafico(
    titulo: str,
    eje_categoria: str,
    categorias: Reference,
    datos: Reference,
    titulos_datos: bool,
) -> BarChart:
    grafico = BarChart()
    grafico.type = "col"
    grafico.title = titulo
    grafico.x_axis.title = eje_categoria
    grafico.y_axis.title = "Monto (CLP)"
    grafico.y_axis.numFmt = "$#,##0"
    grafico.y_axis.scaling.min = 0
    grafico.y_axis.majorGridlines = None
    grafico.legend = None
    grafico.varyColors = False
    grafico.visible_cells_only = False
    grafico.graphical_properties = GraphicalProperties(
        ln=LineProperties(noFill=True)
    )
    grafico.add_data(datos, titles_from_data=titulos_datos)
    grafico.set_categories(categorias)
    serie = grafico.series[0]
    serie.graphicalProperties.solidFill = "1F4E79"
    serie.graphicalProperties.line.solidFill = "1F4E79"
    serie.dLbls = DataLabelList(
        showVal=True,
        showCatName=False,
        showSerName=False,
        showLegendKey=False,
        showPercent=False,
        showBubbleSize=False,
        numFmt="$#,##0",
    )
    grafico.width = 11
    grafico.height = 7.5
    return grafico


def _escribir_dashboard(
    ws, df_limpio: pd.DataFrame, metricas: dict[str, Any]
) -> None:
    for columna, ancho in {
        "A": 24,
        "B": 16,
        "C": 3,
        "D": 16,
        "E": 14,
        "F": 16,
        "G": 18,
        "H": 14,
    }.items():
        ws.column_dimensions[columna].width = ancho

    ws.merge_cells("A1:H1")
    ws["A1"] = "Conciliación de ventas multicanal"
    ws["A1"].font = Font(size=20, bold=True, color="1F3864")
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:H2")
    ws["A2"] = f"Período: {_periodo(df_limpio)}"
    ws["A2"].font = Font(size=11, italic=True, color="595959")

    kpis = [
        ("Ventas totales", metricas["ventas_totales"], "$#,##0"),
        ("Transacciones", metricas["transacciones"], "#,##0"),
        ("Ticket promedio", metricas["ticket_promedio"], "$#,##0"),
        ("Unidades", metricas["unidades_totales"], "#,##0"),
    ]
    for (inicio, fin), (titulo, valor, formato) in zip(
        (("A", "B"), ("C", "D"), ("E", "F"), ("G", "H")), kpis
    ):
        ws.merge_cells(f"{inicio}4:{fin}4")
        ws.merge_cells(f"{inicio}5:{fin}5")
        etiqueta = ws[f"{inicio}4"]
        etiqueta.value = titulo
        etiqueta.font = Font(bold=True, color="FFFFFF")
        etiqueta.fill = PatternFill(fill_type="solid", fgColor="305496")
        etiqueta.alignment = Alignment(horizontal="center", vertical="center")
        celda = ws[f"{inicio}5"]
        celda.value = valor
        celda.font = Font(size=16, bold=True, color="1F3864")
        celda.fill = PatternFill(fill_type="solid", fgColor="D9E1F2")
        celda.alignment = Alignment(horizontal="center", vertical="center")
        celda.number_format = formato
    ws.row_dimensions[4].height = 18
    ws.row_dimensions[5].height = 28

    top = (
        df_limpio.groupby("vendedor", as_index=False)["monto"]
        .sum()
        .sort_values("monto", ascending=False)
        .head(5)
    )
    ws.merge_cells("A7:B7")
    ws["A7"] = "Top 5 vendedores"
    ws["A7"].font = Font(bold=True, size=12, color="1F3864")
    for columna, encabezado in enumerate(["Vendedor", "Monto"], start=1):
        _estilo_encabezado(ws.cell(row=8, column=columna, value=encabezado))
    for fila, (vendedor, monto) in enumerate(
        top.itertuples(index=False, name=None), start=9
    ):
        ws.cell(row=fila, column=1, value=vendedor)
        ws.cell(row=fila, column=2, value=float(monto)).number_format = "$#,##0"
    if not top.empty:
        ws.conditional_formatting.add(
            f"B9:B{8 + len(top)}",
            DataBarRule(
                start_type="num", start_value=0, end_type="max", color="5B9BD5"
            ),
        )

    comparativa = (
        df_limpio.groupby("canal", as_index=False)
        .agg(transacciones=("monto", "size"), monto=("monto", "sum"))
        .sort_values("monto", ascending=False)
    )
    comparativa["ticket_promedio"] = comparativa["monto"] / comparativa["transacciones"]
    ws.merge_cells("D7:G7")
    ws["D7"] = "Comparativa por canal"
    ws["D7"].font = Font(bold=True, size=12, color="1F3864")
    for columna, encabezado in enumerate(
        ["Canal", "Transacciones", "Monto", "Ticket promedio"], start=4
    ):
        _estilo_encabezado(ws.cell(row=8, column=columna, value=encabezado))
    for fila, valores in enumerate(
        comparativa.itertuples(index=False, name=None), start=9
    ):
        canal, transacciones, monto, ticket = valores
        ws.cell(row=fila, column=4, value=canal)
        ws.cell(row=fila, column=5, value=int(transacciones)).number_format = "#,##0"
        ws.cell(row=fila, column=6, value=float(monto)).number_format = "$#,##0"
        ws.cell(row=fila, column=7, value=float(ticket)).number_format = "$#,##0"
    if not comparativa.empty:
        ws.conditional_formatting.add(
            f"F9:F{8 + len(comparativa)}",
            ColorScaleRule(
                start_type="min",
                start_color="FFFFFF",
                end_type="max",
                end_color="9DC3E6",
            ),
        )

    ws.row_dimensions[7].height = 20
    ws.row_dimensions[8].height = 20
    for fila in range(9, 14):
        ws.row_dimensions[fila].height = 18

    ws.merge_cells("A15:H15")
    ws["A15"] = "Lectura del mes"
    ws["A15"].font = Font(bold=True, size=12, color="1F3864")
    ws.merge_cells("A16:H18")
    lectura = ws["A16"]
    lectura.value = _lectura_mes(df_limpio)
    lectura.font = Font(size=10.5, color="404040")
    lectura.alignment = Alignment(
        horizontal="left", vertical="top", wrap_text=True
    )
    for fila in range(16, 19):
        ws.row_dimensions[fila].height = 36

    meses = (
        df_limpio.assign(mes=df_limpio["fecha"].dt.strftime("%Y-%m"))
        .groupby("mes", as_index=False)["monto"]
        .sum()
        .sort_values("mes")
    )
    for fila, (mes, monto) in enumerate(
        meses.itertuples(index=False, name=None), start=1
    ):
        ws.cell(row=fila, column=26, value=mes)
        ws.cell(row=fila, column=27, value=float(monto)).number_format = "$#,##0"
    ws.column_dimensions["Z"].hidden = True
    ws.column_dimensions["AA"].hidden = True

    if not meses.empty:
        grafico_mes = _crear_grafico(
            "Ventas por mes",
            "Mes",
            Reference(ws, min_col=26, min_row=1, max_row=len(meses)),
            Reference(ws, min_col=27, min_row=1, max_row=len(meses)),
            False,
        )
        ws.add_chart(grafico_mes, "A20")

    if not comparativa.empty:
        grafico_canal = _crear_grafico(
            "Ventas por canal",
            "Canal",
            Reference(
                ws, min_col=4, min_row=9, max_row=8 + len(comparativa)
            ),
            Reference(
                ws, min_col=6, min_row=8, max_row=8 + len(comparativa)
            ),
            True,
        )
        ws.add_chart(grafico_canal, "E20")

    ws.print_area = "A1:H36"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True


def exportar_dashboard(
    df_limpio: pd.DataFrame,
    df_errores: pd.DataFrame,
    metricas: dict[str, Any],
    ruta: str | Path,
) -> Path:
    """Crea el Excel con las hojas Dashboard, Consolidado y Errores."""
    ruta = Path(ruta)
    ruta.parent.mkdir(parents=True, exist_ok=True)

    libro = Workbook()
    hoja_dashboard = libro.active
    hoja_dashboard.title = "Dashboard"
    hoja_consolidado = libro.create_sheet("Consolidado")
    hoja_errores = libro.create_sheet("Errores")

    _escribir_dataframe(hoja_consolidado, df_limpio)
    if df_errores.empty:
        hoja_errores["A1"] = "Sin errores"
        hoja_errores["A1"].font = Font(bold=True)
    else:
        _escribir_dataframe(hoja_errores, df_errores)
    _escribir_dashboard(hoja_dashboard, df_limpio, metricas)

    libro.save(ruta)
    return ruta.resolve()


def _formato_clp(valor: float) -> str:
    return f"${valor:,.0f}".replace(",", ".")


def _resumen(metricas: dict[str, Any], ruta: Path) -> str:
    lineas = ["", "Resumen de conciliación"]
    for canal, filas in metricas["filas_por_canal"].items():
        lineas.append(f"  {canal}: {filas} filas leídas")
    lineas.extend(
        [
            f"  Total leído: {metricas['total_leido']} filas",
            f"  Errores: {metricas['errores']} (quedan en la hoja Errores)",
            f"  Duplicados exactos eliminados: {metricas['duplicados_exactos']}",
            f"  Duplicados entre canales eliminados: {metricas['duplicados_cruzados']}",
            f"  Ventas totales: {_formato_clp(metricas['ventas_totales'])}",
            f"  Transacciones: {metricas['transacciones']}",
            f"  Ticket promedio: {_formato_clp(metricas['ticket_promedio'])}",
            f"  Unidades totales: {metricas['unidades_totales']}",
            f"  Archivo generado: {ruta.resolve()}",
        ]
    )
    return "\n".join(lineas)


def main(argv: list[str] | None = None) -> int:
    """Punto de entrada de la CLI en español."""
    parser = argparse.ArgumentParser(
        description="Concilia ventas de Local, Web y POS en un dashboard Excel."
    )
    parser.add_argument(
        "--local",
        default="datos/ventas_local.xlsx",
        help="Excel del canal Local (default: datos/ventas_local.xlsx)",
    )
    parser.add_argument(
        "--web",
        default="datos/ventas_web.xlsx",
        help="Excel del canal Web (default: datos/ventas_web.xlsx)",
    )
    parser.add_argument(
        "--pos",
        default="datos/ventas_pos.csv",
        help="CSV del canal POS (default: datos/ventas_pos.csv)",
    )
    parser.add_argument(
        "--salida",
        default="dashboard_ventas.xlsx",
        help="Archivo Excel de salida (default: dashboard_ventas.xlsx)",
    )
    args = parser.parse_args(argv)

    entradas = [("Local", args.local), ("Web", args.web), ("POS", args.pos)]
    faltantes = [(nombre, ruta) for nombre, ruta in entradas if not Path(ruta).is_file()]
    if faltantes:
        for nombre, ruta in faltantes:
            print(
                f"Error: no se encontró el archivo de {nombre}: {ruta}",
                file=sys.stderr,
            )
        return 1

    try:
        df_limpio, df_errores, metricas = conciliar(
            args.local, args.web, args.pos
        )
        ruta = exportar_dashboard(df_limpio, df_errores, metricas, args.salida)
    except (FileNotFoundError, ValueError, OSError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(_resumen(metricas, ruta))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
