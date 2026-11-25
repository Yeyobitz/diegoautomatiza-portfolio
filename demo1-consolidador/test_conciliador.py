"""Tests del conciliador de ventas multicanal."""

from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

import conciliador as conciliador


COLUMNAS_LOCAL = ["Folio", "Fecha", "SKU", "Vendedor", "Unidades", "Monto"]
COLUMNAS_WEB = ["Order ID", "Order Date", "Product Code", "Sales Rep", "Qty", "Revenue"]
COLUMNAS_POS = ["ticket", "timestamp", "sku", "cajero", "cantidad", "total"]


def _crear_local(ruta: Path, filas: list[dict]) -> Path:
    pd.DataFrame(filas, columns=COLUMNAS_LOCAL).to_excel(ruta, index=False)
    return ruta


def _crear_web(ruta: Path, filas: list[dict]) -> Path:
    pd.DataFrame(filas, columns=COLUMNAS_WEB).to_excel(ruta, index=False)
    return ruta


def _crear_pos(ruta: Path, filas: list[dict]) -> Path:
    pd.DataFrame(filas, columns=COLUMNAS_POS).to_csv(
        ruta, sep=";", index=False, encoding="utf-8-sig"
    )
    return ruta


def _tres_canales(tmp_path: Path, local: list[dict], web: list[dict], pos: list[dict]):
    return (
        _crear_local(tmp_path / "local.xlsx", local),
        _crear_web(tmp_path / "web.xlsx", web),
        _crear_pos(tmp_path / "pos.csv", pos),
    )


@pytest.mark.parametrize(
    ("entrada", "esperado"),
    [
        ("$1.234,56", 1234.56),
        ("$ 41.990,00", 41990.0),
        ("41.990", 41990.0),
        ("1234.56", 1234.56),
        (1234, 1234.0),
        (1234.56, 1234.56),
        ("N/D", None),
        ("", None),
        (None, None),
        ("texto", None),
    ],
)
def test_parsear_monto_chileno(entrada, esperado):
    assert conciliador.parsear_monto_chileno(entrada) == esperado


def test_normalizar_fecha_formatos_validos():
    assert conciliador.normalizar_fecha("05/11/2025") == datetime(2025, 11, 5)
    assert conciliador.normalizar_fecha("2025-11-05") == datetime(2025, 11, 5)
    assert conciliador.normalizar_fecha("2025-11-05T15:32:41") == datetime(
        2025, 11, 5, 15, 32, 41
    )


@pytest.mark.parametrize("entrada", ["32/13/2025", "2025-02-30", "", None])
def test_normalizar_fecha_invalida(entrada):
    assert conciliador.normalizar_fecha(entrada) is None


def test_leer_canal_esquema_comun(tmp_path):
    local, web, pos = _tres_canales(
        tmp_path,
        [
            {
                "Folio": 1,
                "Fecha": "05/11/2025",
                "SKU": "A-1",
                "Vendedor": "Ana",
                "Unidades": 2,
                "Monto": 1000,
            }
        ],
        [
            {
                "Order ID": 101,
                "Order Date": "2025-11-06",
                "Product Code": "B-2",
                "Sales Rep": "Bea",
                "Qty": 1,
                "Revenue": 5000,
            }
        ],
        [
            {
                "ticket": "TK-9",
                "timestamp": "2025-11-07T10:15:00",
                "sku": "C-3",
                "cajero": "Carla",
                "cantidad": 3,
                "total": "$12.500,00",
            }
        ],
    )

    casos = [
        (conciliador.leer_canal(local, "Local"), "Local", "1", datetime(2025, 11, 5), "A-1", "Ana", 2, 1000.0),
        (conciliador.leer_canal(web, "Web"), "Web", "101", datetime(2025, 11, 6), "B-2", "Bea", 1, 5000.0),
        (conciliador.leer_canal(pos, "POS"), "POS", "TK-9", datetime(2025, 11, 7, 10, 15), "C-3", "Carla", 3, 12500.0),
    ]

    for df, canal, folio, fecha, sku, vendedor, unidades, monto in casos:
        assert list(df.columns) == conciliador.COLUMNAS
        assert len(df) == 1
        fila = df.iloc[0]
        assert fila["canal"] == canal
        assert str(fila["folio"]) == folio
        assert fila["fecha"] == pd.Timestamp(fecha)
        assert fila["sku"] == sku
        assert fila["vendedor"] == vendedor
        assert fila["unidades"] == unidades
        assert fila["monto"] == pytest.approx(monto)
        assert pd.api.types.is_datetime64_any_dtype(df["fecha"])
        assert pd.api.types.is_integer_dtype(df["unidades"])
        assert pd.api.types.is_float_dtype(df["monto"])


def test_errores_quedan_fuera_del_limpio(tmp_path):
    local, web, pos = _tres_canales(
        tmp_path,
        [
            {"Folio": 1, "Fecha": "05/11/2025", "SKU": "A-1", "Vendedor": "Ana", "Unidades": 1, "Monto": 1000},
            {"Folio": 2, "Fecha": "32/13/2025", "SKU": "A-2", "Vendedor": "Ana", "Unidades": 1, "Monto": 2000},
            {"Folio": 3, "Fecha": "06/11/2025", "SKU": "A-3", "Vendedor": "Ana", "Unidades": 1, "Monto": "N/D"},
        ],
        [],
        [],
    )

    limpio, errores, metricas = conciliador.conciliar(local, web, pos)

    assert len(limpio) == 1
    assert len(errores) == 2
    assert set(errores["problema"]) == {"fecha inválida", "monto no numérico"}
    assert metricas["errores"] == 2
    assert set(limpio["folio"]) == {1}


def test_duplicados_exactos_dentro_del_canal(tmp_path):
    fila = {"Folio": 1, "Fecha": "05/11/2025", "SKU": "A-1", "Vendedor": "Ana", "Unidades": 2, "Monto": 1000}
    local, web, pos = _tres_canales(tmp_path, [fila, fila], [], [])

    limpio, _, metricas = conciliador.conciliar(local, web, pos)

    assert len(limpio) == 1
    assert metricas["duplicados_exactos"] == 1
    assert metricas["duplicados_cruzados"] == 0


def test_duplicados_entre_canales(tmp_path):
    local, web, pos = _tres_canales(
        tmp_path,
        [{"Folio": 1, "Fecha": "05/11/2025", "SKU": "A-1", "Vendedor": "Ana", "Unidades": 2, "Monto": 1000}],
        [{"Order ID": 999, "Order Date": "2025-11-05", "Product Code": "A-1", "Sales Rep": "Bea", "Qty": 2, "Revenue": 1000}],
        [],
    )

    limpio, _, metricas = conciliador.conciliar(local, web, pos)

    assert len(limpio) == 1
    assert metricas["duplicados_cruzados"] == 1
    assert metricas["duplicados_exactos"] == 0
    assert limpio.iloc[0]["canal"] == "Local"


def test_kpis_coherentes(tmp_path):
    filas = [
        {"Folio": 1, "Fecha": "05/11/2025", "SKU": "A-1", "Vendedor": "Ana", "Unidades": 1, "Monto": 1000},
        {"Folio": 2, "Fecha": "06/11/2025", "SKU": "A-2", "Vendedor": "Ana", "Unidades": 2, "Monto": 2000},
        {"Folio": 3, "Fecha": "07/11/2025", "SKU": "A-3", "Vendedor": "Ana", "Unidades": 3, "Monto": 3000},
    ]
    local, web, pos = _tres_canales(tmp_path, filas, [], [])

    _, _, metricas = conciliador.conciliar(local, web, pos)

    assert metricas["transacciones"] == 3
    assert metricas["ventas_totales"] == pytest.approx(6000.0)
    assert metricas["unidades_totales"] == 6
    assert metricas["ticket_promedio"] == pytest.approx(
        metricas["ventas_totales"] / metricas["transacciones"]
    )
    assert metricas["filas_por_canal"]["Local"] == 3
    assert metricas["total_leido"] == 3


def test_integracion_genera_dashboard(tmp_path):
    local, web, pos = _tres_canales(
        tmp_path,
        [{"Folio": 1, "Fecha": "05/11/2025", "SKU": "A-1", "Vendedor": "Ana", "Unidades": 2, "Monto": 10000}],
        [{"Order ID": 101, "Order Date": "2025-11-06", "Product Code": "B-2", "Sales Rep": "Bea", "Qty": 1, "Revenue": 5000}],
        [{"ticket": "TK-1", "timestamp": "2025-11-07T10:00:00", "sku": "C-3", "cajero": "Carla", "cantidad": 3, "total": "$7.500,00"}],
    )
    salida = tmp_path / "dashboard_ventas.xlsx"
    _, _, metricas = conciliador.conciliar(local, web, pos)

    codigo = conciliador.main(
        ["--local", str(local), "--web", str(web), "--pos", str(pos), "--salida", str(salida)]
    )

    assert codigo == 0
    assert salida.exists()
    libro = load_workbook(salida)
    assert libro.sheetnames == ["Dashboard", "Consolidado", "Errores"]

    hoja = libro["Dashboard"]
    assert len(hoja._charts) == 2
    reglas = [regla for grupo in hoja.conditional_formatting for regla in grupo.rules]
    assert len(reglas) >= 1
    assert hoja.page_setup.orientation == "landscape"
    assert hoja.page_setup.fitToWidth == 1
    assert hoja.page_setup.fitToHeight == 1

    consolidado = libro["Consolidado"]
    encabezados = [celda.value for celda in consolidado[1]]
    columna_monto = encabezados.index("monto") + 1
    total = sum(
        consolidado.cell(row=fila, column=columna_monto).value or 0
        for fila in range(2, consolidado.max_row + 1)
    )
    assert total == pytest.approx(metricas["ventas_totales"])


def test_leer_canal_falta_columna(tmp_path):
    ruta = tmp_path / "incompleto.xlsx"
    pd.DataFrame({"Folio": [1]}).to_excel(ruta, index=False)

    with pytest.raises(ValueError, match="faltan columnas"):
        conciliador.leer_canal(ruta, "Local")


def test_main_avisa_archivo_faltante(tmp_path, capsys):
    codigo = conciliador.main(
        [
            "--local", str(tmp_path / "no_existe.xlsx"),
            "--web", str(tmp_path / "no_existe_web.xlsx"),
            "--pos", str(tmp_path / "no_existe_pos.csv"),
            "--salida", str(tmp_path / "salida.xlsx"),
        ]
    )

    assert codigo == 1
    assert "no se encontró" in capsys.readouterr().err.lower()
