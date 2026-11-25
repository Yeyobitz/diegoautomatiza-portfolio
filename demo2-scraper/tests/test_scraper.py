"""Tests del monitor de precios de catálogo.

Todo se resuelve con fixtures HTML locales (tests/fixtures): cero red.
Un producto agotado ("Out of stock"), precios con £ y ratings como clase CSS
("Three"/"Four") son parte del contrato del parser.
"""
from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import load_workbook

import scraper

FIXTURES = Path(__file__).parent / "fixtures"


def fixture_url(nombre):
    """URL file:// de una fixture: parsea sin red por el mismo camino que http."""
    return (FIXTURES / nombre).as_uri()


def leer_pagina(nombre):
    return scraper.fetch_soup(fixture_url(nombre))


# ---------------------------------------------------------------- parseo ---

def test_parse_price_extrae_libras():
    assert scraper.parse_price("£51.77") == 51.77
    assert scraper.parse_price("£23.90") == pytest.approx(23.90)
    assert scraper.parse_price("£1,299.00") == pytest.approx(1299.00)


def test_parse_rating_mapea_clases_css():
    assert scraper.parse_rating(["star-rating", "Three"]) == 3
    assert scraper.parse_rating(["star-rating", "Four"]) == 4
    for clase, valor in [("One", 1), ("Two", 2), ("Five", 5)]:
        assert scraper.parse_rating(["star-rating", clase]) == valor


def test_parse_availability_distingue_stock():
    assert scraper.parse_availability("In stock") is True
    assert scraper.parse_availability("Out of stock") is False


def test_parse_catalog_page_extrae_8_productos():
    productos = scraper.parse_catalog_page(leer_pagina("catalogo_viejo.html"), fixture_url("catalogo_viejo.html"))
    assert len(productos) == 8
    primero = productos[0]
    assert primero["nombre"] == "A Light in the Attic"
    assert primero["precio"] == pytest.approx(51.77)
    assert primero["rating"] == 3
    assert primero["disponible"] is True
    assert primero["url"].endswith("a-light-in-the-attic_1000/index.html")


def test_scrape_catalog_file_url_paginacion_y_agotado():
    """file:// recorre la paginación completa sin red y detecta el agotado."""
    productos = scraper.scrape_catalog(fixture_url("catalogo_viejo.html"))
    assert len(productos) == 12  # 8 + 4 de la segunda página
    rooster = next(p for p in productos if p["nombre"] == "The Rooster Bar")
    assert rooster["disponible"] is False


# ------------------------------------------------------------ cambios -----

def _mini_catalogo():
    return (
        [
            {"nombre": "Alpha", "precio": 50.00, "rating": 3, "disponible": True, "url": "u1"},
            {"nombre": "Beta", "precio": 40.00, "rating": 4, "disponible": True, "url": "u2"},
            {"nombre": "Gamma", "precio": 30.00, "rating": 5, "disponible": True, "url": "u3"},
        ],
        [
            {"nombre": "Alpha", "precio": 60.00, "rating": 3, "disponible": True, "url": "u1"},
            {"nombre": "Beta", "precio": 30.00, "rating": 4, "disponible": True, "url": "u2"},
            {"nombre": "Delta", "precio": 20.00, "rating": 2, "disponible": True, "url": "u4"},
        ],
    )


def test_detect_changes_clasifica_y_calcula_variacion():
    anteriores, actuales = _mini_catalogo()
    cambios = scraper.detect_changes(anteriores, actuales)
    subieron = {c["nombre"]: c for c in cambios["subieron"]}
    bajaron = {c["nombre"]: c for c in cambios["bajaron"]}
    assert cambios["desaparecidos"][0]["nombre"] == "Gamma"
    assert cambios["nuevos"][0]["nombre"] == "Delta"
    assert subieron["Alpha"]["variacion_pct"] == pytest.approx(0.20)
    assert bajaron["Beta"]["variacion_pct"] == pytest.approx(-0.25)


def test_detect_changes_sobre_fixtures_cuenta_movimientos():
    anteriores = scraper.scrape_catalog(fixture_url("catalogo_viejo.html"))
    actuales = scraper.scrape_catalog(fixture_url("catalogo_nuevo.html"))
    cambios = scraper.detect_changes(anteriores, actuales)
    assert len(cambios["subieron"]) == 3
    assert len(cambios["bajaron"]) == 2
    assert len(cambios["desaparecidos"]) == 2
    assert len(cambios["nuevos"]) == 2
    # mayor movimiento primero
    variaciones = [abs(c["variacion_pct"]) for c in cambios["subieron"]]
    assert variaciones == sorted(variaciones, reverse=True)


# ------------------------------------------------------------- excel ------

def test_write_snapshot_formato_profesional(tmp_path):
    productos = scraper.scrape_catalog(fixture_url("catalogo_viejo.html"))
    ruta = tmp_path / "catalogo_snapshot.xlsx"
    scraper.write_snapshot(productos, str(ruta), datetime(2026, 11, 24, 7, 0))
    wb = load_workbook(ruta)
    assert wb.sheetnames == ["Catálogo", "Resumen"]
    hoja = wb["Catálogo"]
    encabezado = hoja["A1"]
    assert encabezado.value == "Producto"
    assert encabezado.font.bold is True
    assert encabezado.fill.start_color.rgb.endswith("1F4E78")
    assert encabezado.font.color.rgb.endswith("FFFFFF")
    assert hoja.auto_filter.ref is not None
    assert hoja.freeze_panes == "A2"
    assert hoja.column_dimensions["A"].width >= 40
    assert hoja["B2"].number_format == '"£"#,##0.00'
    assert hoja["B2"].value == pytest.approx(51.77)
    # el gráfico de distribución no sobrevive a load_workbook: se valida su tabla
    resumen = wb["Resumen"]
    textos = {c.value for fila in resumen.iter_rows() for c in fila if isinstance(c.value, str)}
    assert "£20–£30" in textos or any("£" in t and "–" in t for t in textos)
    assert "Precio promedio" in textos


def test_historial_acumulado_y_cambios_xlsx(tmp_path):
    productos_viejos = scraper.scrape_catalog(fixture_url("catalogo_viejo.html"))
    productos_nuevos = scraper.scrape_catalog(fixture_url("catalogo_nuevo.html"))
    historial_ruta = tmp_path / "historial_precios.xlsx"

    historial = scraper.append_history(productos_viejos, str(historial_ruta), datetime(2026, 11, 23, 9, 0))
    assert len(historial) == 12
    historial = scraper.append_history(productos_nuevos, str(historial_ruta), datetime(2026, 11, 24, 9, 0))
    assert len(historial) == 24  # acumulado: dos corridas
    leido = scraper.read_history(str(historial_ruta))
    assert leido[0]["producto"] == "A Light in the Attic"
    assert leido[0]["precio"] == pytest.approx(51.77)
    assert leido[0]["fecha"] == datetime(2026, 11, 23, 9, 0)

    cambios = scraper.detect_changes(productos_viejos, productos_nuevos)
    ruta = tmp_path / "cambios_precios.xlsx"
    scraper.write_changes(cambios, historial, str(ruta))
    wb = load_workbook(ruta)
    assert wb.sheetnames == ["Subieron", "Bajaron", "Desaparecidos", "Nuevos", "Historial"]
    subieron = wb["Subieron"]
    assert subieron["A1"].font.bold is True
    assert subieron["D2"].number_format.startswith("+0.0%")
    assert subieron["D2"].font.color.rgb.endswith("C00000")
    bajaron = wb["Bajaron"]
    assert bajaron["D2"].font.color.rgb.endswith("2E7D32")
    historial_hoja = wb["Historial"]
    assert historial_hoja["A2"].number_format == "dd-mm-yyyy hh:mm"
    assert historial_hoja.max_row == 25  # encabezado + 24 observaciones
