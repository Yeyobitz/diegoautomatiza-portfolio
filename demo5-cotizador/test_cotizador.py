"""Tests del generador de cotizaciones (demo 5)."""

import re
import shutil
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

import cotizador

ITEMS_OK = """descripcion,unidad,cantidad,precio_unitario
Servicio de montaje de estructura metálica,hh,10,1000
Confección e instalación de pasamanos,ml,4,2500
"""


def _items_csv(tmp_path, texto):
    ruta = tmp_path / "items.csv"
    ruta.write_text(texto, encoding="utf-8")
    return ruta


def _generar(tmp_path, texto=ITEMS_OK, **opciones):
    """Genera una cotización de prueba con los parámetros indicados."""
    ruta = _items_csv(tmp_path / "entrada", texto)
    opciones.setdefault("fecha", "2026-11-25")
    opciones.setdefault("salida", tmp_path)
    opciones.setdefault("pdf", "no")
    argv = ["--items", str(ruta)]
    for clave, valor in opciones.items():
        argv += [f"--{clave.replace('_', '-')}", str(valor)]
    return cotizador.main(argv)


# ---------------------------------------------------------------- folios

def test_folio_consecutivo_entre_corridas(tmp_path):
    """Dos corridas en el mismo directorio siguen la numeración."""
    ruta = _items_csv(tmp_path, ITEMS_OK)
    r1 = cotizador.main(["--items", str(ruta), "--salida", str(tmp_path),
                         "--pdf", "no", "--fecha", "2026-11-25"])
    r2 = cotizador.main(["--items", str(ruta), "--salida", str(tmp_path),
                         "--pdf", "no", "--fecha", "2026-11-25"])
    assert r1["folio"] == "COT-2026-0001"
    assert r2["folio"] == "COT-2026-0002"


def test_folio_continua_desde_archivo_previo(tmp_path):
    """Si ya hay una cotización COT-2026-0007, la nueva es COT-2026-0008."""
    (tmp_path / "cotizacion_COT-2026-0007.pdf").write_bytes(b"previa")
    ruta = _items_csv(tmp_path, ITEMS_OK)
    resultado = cotizador.main(["--items", str(ruta), "--salida", str(tmp_path),
                                "--pdf", "no", "--fecha", "2026-11-25"])
    assert resultado["folio"] == "COT-2026-0008"


def test_formato_folio(tmp_path):
    ruta = _items_csv(tmp_path, ITEMS_OK)
    resultado = cotizador.main(["--items", str(ruta), "--salida", str(tmp_path),
                                "--pdf", "no", "--fecha", "2026-11-25"])
    assert re.fullmatch(r"COT-\d{4}-\d{4}", resultado["folio"])
    assert (tmp_path / f"cotizacion_{resultado['folio']}.xlsx").is_file()


# ---------------------------------------------------------------- IVA y redondeo

def test_iva_19_por_ciento():
    items = [{"descripcion": "Servicio", "unidad": "hh",
              "cantidad": 100.0, "precio_unitario": 1000.0}]
    totales = cotizador.calcular_totales(items)
    assert totales["subtotal"] == 100_000
    assert totales["iva"] == 19_000
    assert totales["total"] == 119_000


def test_redondeo_half_up():
    """56,5 → 57 (no el redondeo bancario de Python); y el IVA se redondea igual."""
    assert cotizador.redondear("56.5") == 57
    assert cotizador.redondear(12.5) == 13
    assert cotizador.redondear(2.4) == 2
    items = [{"descripcion": "Servicio", "unidad": "un",
              "cantidad": 1.0, "precio_unitario": 56.5}]
    totales = cotizador.calcular_totales(items)
    assert totales["lineas"][0]["importe"] == 57      # no 56 (banker's)
    assert totales["subtotal"] == 57
    assert totales["iva"] == 11                       # 10,83 → 11
    assert totales["total"] == 68


def test_total_cuadra_con_suma(tmp_path):
    """SUBTOTAL = suma de importes; TOTAL = SUBTOTAL + IVA; y cuadra en el xlsx."""
    ruta = _items_csv(tmp_path, ITEMS_OK)
    resultado = cotizador.main(["--items", str(ruta), "--salida", str(tmp_path),
                                "--pdf", "no", "--fecha", "2026-11-25"])
    totales = resultado["totales"]
    assert totales["subtotal"] == sum(l["importe"] for l in totales["lineas"])
    assert totales["total"] == totales["subtotal"] + totales["iva"]

    hoja = load_workbook(tmp_path / f"cotizacion_{resultado['folio']}.xlsx").active
    celdas = {}
    for fila in hoja.iter_rows(min_col=1, max_col=6):
        for celda in fila:
            if isinstance(celda.value, str) and celda.value.startswith(("SUBTOTAL", "IVA", "TOTAL")):
                celdas[celda.value.split(" (")[0]] = hoja.cell(celda.row, 6).value
    assert celdas["SUBTOTAL"] == totales["subtotal"]
    assert celdas["IVA"] == totales["iva"]
    assert celdas["TOTAL"] == totales["total"]


# ---------------------------------------------------------------- CSV inválido

def test_csv_malformado_rechazado(tmp_path):
    """Un CSV malformado (columnas de más en una fila) falla con mensaje claro."""
    ruta = _items_csv(tmp_path, ITEMS_OK + "Servicio extra con,coma,hh,1,1000,de más\n")
    with pytest.raises(SystemExit) as excinfo:
        cotizador.leer_items(ruta)
    assert "malformado" in str(excinfo.value)


def test_csv_columnas_faltantes_rechazado(tmp_path):
    ruta = _items_csv(tmp_path, "descripcion,cantidad\nServicio,10\n")
    df = cotizador.leer_items(ruta)
    with pytest.raises(SystemExit) as excinfo:
        cotizador.validar_items(df, ruta)
    assert "precio_unitario" in str(excinfo.value)


def test_cantidad_negativa_rechazada(tmp_path):
    ruta = _items_csv(tmp_path, ITEMS_OK + "Servicio de escombros,viaje,-5,1000\n")
    df = cotizador.leer_items(ruta)
    with pytest.raises(SystemExit) as excinfo:
        cotizador.validar_items(df, ruta)
    assert "cantidad" in str(excinfo.value)


def test_cantidad_no_numerica_rechazada(tmp_path):
    ruta = _items_csv(tmp_path, "descripcion,unidad,cantidad,precio_unitario\n"
                                "Servicio,hh,muchos,1000\n")
    df = cotizador.leer_items(ruta)
    with pytest.raises(SystemExit):
        cotizador.validar_items(df, ruta)


def test_precio_negativo_rechazado(tmp_path):
    ruta = _items_csv(tmp_path, "descripcion,unidad,cantidad,precio_unitario\n"
                                "Servicio,hh,10,-1000\n")
    df = cotizador.leer_items(ruta)
    with pytest.raises(SystemExit) as excinfo:
        cotizador.validar_items(df, ruta)
    assert "precio" in str(excinfo.value)


# ---------------------------------------------------------------- documento

SOFFICE = shutil.which("soffice")


@pytest.mark.skipif(SOFFICE is None, reason="LibreOffice (soffice) no está disponible")
def test_pdf_existe_y_no_vacio(tmp_path):
    ruta = _items_csv(tmp_path, ITEMS_OK)
    resultado = cotizador.main(["--items", str(ruta), "--salida", str(tmp_path),
                                "--pdf", "si", "--fecha", "2026-11-25"])
    pdf = Path(resultado["pdf"])
    assert pdf.is_file()
    assert pdf.stat().st_size > 0
    assert pdf.read_bytes()[:5] == b"%PDF-"
    assert pdf.name == f"cotizacion_{resultado['folio']}.pdf"


def test_plantilla_de_marca_presente(tmp_path):
    """La marca ficticia, su RUT y el formato del encabezado están en el xlsx."""
    ruta = _items_csv(tmp_path, ITEMS_OK)
    resultado = cotizador.main(["--items", str(ruta), "--salida", str(tmp_path),
                                "--pdf", "no", "--fecha", "2026-11-25"])
    hoja = load_workbook(tmp_path / f"cotizacion_{resultado['folio']}.xlsx").active
    textos = [str(celda.value) for fila in hoja.iter_rows() for celda in fila
              if celda.value is not None]
    todo = " | ".join(textos)
    assert cotizador.EMPRESA.upper() in todo
    assert cotizador.RUT_EMPRESA in todo
    assert "COTIZACIÓN" in todo
    assert resultado["folio"] in todo
    assert cotizador.CLIENTE_DEFAULT in todo
    # Encabezado de la tabla con el relleno de marca.
    celda_descripcion = next(
        celda for fila in hoja.iter_rows() for celda in fila if celda.value == "Descripción"
    )
    assert celda_descripcion.fill.fgColor.rgb.endswith(cotizador.AZUL_OSCURO)
    assert celda_descripcion.font.color.rgb.endswith("FFFFFF")


def test_validez_15_dias_por_defecto(tmp_path):
    ruta = _items_csv(tmp_path, ITEMS_OK)
    resultado = cotizador.main(["--items", str(ruta), "--salida", str(tmp_path),
                                "--pdf", "no", "--fecha", "2026-11-25"])
    assert resultado["validez"] == 15
    hoja = load_workbook(tmp_path / f"cotizacion_{resultado['folio']}.xlsx").active
    textos = [str(celda.value) for fila in hoja.iter_rows() for celda in fila
              if celda.value is not None]
    validez = next(texto for texto in textos if "Validez de la oferta" in texto)
    assert "15 días" in validez
    assert "10 de diciembre de 2026" in validez  # 2026-11-25 + 15 días


def test_linea_de_firma_presente(tmp_path):
    ruta = _items_csv(tmp_path, ITEMS_OK)
    resultado = cotizador.main(["--items", str(ruta), "--salida", str(tmp_path),
                                "--pdf", "no", "--fecha", "2026-11-25"])
    hoja = load_workbook(tmp_path / f"cotizacion_{resultado['folio']}.xlsx").active
    textos = [str(celda.value) for fila in hoja.iter_rows() for celda in fila
              if celda.value is not None]
    assert "Aceptación del cliente" in textos
    assert any("Por Ingeniería Austral SpA" in texto for texto in textos)
    assert any(texto.startswith("___") for texto in textos)
