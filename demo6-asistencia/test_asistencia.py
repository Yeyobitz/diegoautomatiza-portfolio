"""Tests del analizador de asistencia y turnos."""

from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

import asistencia

DIRECTORIO = Path(__file__).resolve().parent
RELOJ = DIRECTORIO / "reloj_asistencia.csv"
TURNOS = DIRECTORIO / "turnos_esperados.csv"
MES = "2026-11"


def _registros():
    """Carga el reloj de ejemplo y lo cruza con los turnos."""
    turnos = asistencia.leer_turnos(TURNOS)
    reloj = asistencia.leer_reloj(RELOJ, MES)
    return asistencia.calcular_registros(reloj, turnos)


def _escribir_reloj(ruta, filas, columnas=None):
    """Escribe un CSV mínimo para probar la validación de entrada."""
    columnas = columnas or asistencia.COLUMNAS_RELOJ
    lineas = [",".join(columnas)] + [",".join(fila) for fila in filas]
    ruta.write_text("\n".join(lineas), encoding="utf-8")


@pytest.mark.parametrize(
    ("hora_real", "hora_ideal", "tolerancia", "esperado"),
    [
        ("06:07", "06:00", 5, 2),
        ("06:05", "06:00", 5, 0),
        ("06:04", "06:00", 5, 0),
        ("08:00", "08:00", 10, 0),
        ("08:20", "08:00", 10, 10),
        ("06:40", "06:00", 5, 35),
    ],
)
def test_minutos_tardanza(hora_real, hora_ideal, tolerancia, esperado):
    assert (
        asistencia.minutos_tardanza(hora_real, hora_ideal, tolerancia) == esperado
    )


@pytest.mark.parametrize(
    ("salida_real", "salida_ideal", "gracia", "esperado"),
    [
        ("14:30", "14:00", 15, 30),
        ("14:15", "14:00", 15, 0),
        ("14:10", "14:00", 15, 0),
        ("13:55", "14:00", 15, 0),
        ("19:30", "17:00", 15, 150),
        ("14:40", "14:00", 30, 40),
    ],
)
def test_minutos_horas_extra(salida_real, salida_ideal, gracia, esperado):
    assert (
        asistencia.minutos_horas_extra(salida_real, salida_ideal, gracia)
        == esperado
    )


def test_minutos_trabajados_cruza_medianoche():
    assert asistencia.minutos_trabajados("06:00", "14:00") == 480
    assert asistencia.minutos_trabajados("22:00", "06:00") == 480
    assert asistencia.minutos_trabajados("23:30", "00:30") == 60


def test_resumen_detecta_ausencia():
    resumen = asistencia.resumen_por_persona(_registros())

    fila = resumen[resumen["empleado"] == "EMP-002"].iloc[0]
    assert int(fila["ausencias"]) == 1
    # La ausencia del 13 de noviembre no genera tardanza ni extra.
    assert int(fila["minutos_tardanza"]) == 0
    assert int(fila["minutos_extra"]) == 0


def test_olvido_marcaje_se_cuenta_aparte():
    registros = _registros()
    resumen = asistencia.resumen_por_persona(registros)

    fila = resumen[resumen["empleado"] == "EMP-003"].iloc[0]
    assert int(fila["olvidos"]) == 1
    assert int(fila["ausencias"]) == 0

    detalle = asistencia.detalle_tardanzas(registros)
    assert not (
        (detalle["empleado"] == "EMP-003")
        & (detalle["fecha"] == date(2026, 11, 11))
    ).any()

    # Para cobertura, el olvido cuenta como presente.
    cobertura = asistencia.calcular_cobertura(registros, 25)
    dia = cobertura[cobertura["fecha"] == date(2026, 11, 11)].iloc[0]
    assert int(dia["presentes"]) == 30


def test_csv_sin_columna_falla(tmp_path):
    ruta = tmp_path / "reloj_sin_estado.csv"
    columnas = [columna for columna in asistencia.COLUMNAS_RELOJ if columna != "estado"]
    _escribir_reloj(
        ruta,
        [["EMP-001", "2026-11-02", "Mañana", "06:00", "14:00"]],
        columnas=columnas,
    )

    with pytest.raises(ValueError, match="estado"):
        asistencia.leer_reloj(ruta, MES)


def test_csv_con_hora_invalida_falla(tmp_path):
    ruta = tmp_path / "reloj_hora_mala.csv"
    _escribir_reloj(
        ruta,
        [["EMP-001", "2026-11-02", "Mañana", "Asistió", "25:99", "14:00"]],
    )

    with pytest.raises(ValueError) as error:
        asistencia.leer_reloj(ruta, MES)
    assert "Fila 2" in str(error.value)
    assert "25:99" in str(error.value)


def test_totales_cuadran(tmp_path):
    ruta = asistencia.generar_resumen(RELOJ, TURNOS, MES, salida=tmp_path)
    libro = load_workbook(ruta)

    resumen = libro["Resumen por Persona"]
    fila_total = next(
        fila
        for fila in resumen.iter_rows(min_row=2)
        if fila[0].value == "TOTALES"
    )
    total_tardanza = fila_total[4].value
    total_extra = fila_total[6].value

    detalle = libro["Tardanzas del Mes"]
    suma_detalle = sum(
        fila[5].value for fila in detalle.iter_rows(min_row=2) if fila[5].value
    )

    extras = libro["Extras"]
    fila_total_extras = next(
        fila for fila in extras.iter_rows(min_row=2) if fila[0].value == "TOTALES"
    )

    assert total_tardanza == suma_detalle
    assert total_extra == pytest.approx(fila_total_extras[2].value)


def test_xlsx_tiene_formato(tmp_path):
    ruta = asistencia.generar_resumen(RELOJ, TURNOS, MES, salida=tmp_path)
    libro = load_workbook(ruta)

    assert libro.sheetnames == [
        "Resumen por Persona",
        "Tardanzas del Mes",
        "Extras",
        "Cobertura",
    ]

    for hoja in libro.worksheets:
        encabezado = hoja["A1"]
        assert encabezado.fill.fgColor.rgb.endswith("1F4E78")
        assert encabezado.font.bold
        assert encabezado.font.color.rgb.endswith("FFFFFF")
        congelado_esperado = (
            "A3" if hoja.title == "Resumen por Persona" else "A2"
        )
        assert hoja.freeze_panes == congelado_esperado
        assert hoja.sheet_view.showGridLines is False

    resumen = libro["Resumen por Persona"]
    reglas = [
        regla
        for rango in resumen.conditional_formatting
        for regla in rango.rules
    ]
    assert any(regla.type == "colorScale" for regla in reglas)
    assert resumen.auto_filter.ref is not None

    # Título con periodo y encabezado de la columna extra con unidad.
    assert resumen["A1"].value == "Resumen de Asistencia — Noviembre 2026"
    encabezados = [resumen.cell(row=2, column=c).value for c in range(1, 10)]
    assert "HH extra (horas)" in encabezados

    # Nota metodológica al pie: línea 1 con la regla de tolerancia y
    # línea 2 con la gracia de 15 min para las extras.
    textos = [
        str(celda.value)
        for fila in resumen.iter_rows()
        for celda in fila
        if celda.value is not None
    ]
    notas = [t for t in textos if t.startswith("Reglas de cálculo")]
    assert len(notas) == 1
    assert "tolerancia" in notas[0]
    assert any(
        t.startswith("Horas extra") and "15 minutos de gracia" in t for t in textos
    )


def test_cobertura_marca_deficit_y_ok(tmp_path):
    ruta = asistencia.generar_resumen(RELOJ, TURNOS, MES, salida=tmp_path)
    libro = load_workbook(ruta)
    hoja = libro["Cobertura"]

    cobertura = {
        fila[0].value.date(): (fila[1].value, fila[3].value)
        for fila in hoja.iter_rows(min_row=2)
    }

    assert cobertura[date(2026, 11, 13)] == (24, "DEFICIT")
    assert cobertura[date(2026, 11, 24)] == (24, "DEFICIT")
    assert cobertura[date(2026, 11, 2)] == (30, "OK")

    # Los días con déficit quedan al inicio y pintados de rojo.
    primera = hoja["A2"].value.date()
    assert primera in {date(2026, 11, 13), date(2026, 11, 24)}
    fila_deficit = next(
        fila
        for fila in hoja.iter_rows(min_row=2)
        if fila[3].value == "DEFICIT"
    )
    assert fila_deficit[3].fill.fgColor.rgb.endswith("FFC7CE")


def test_mes_sin_registros_sale_con_codigo_2(capsys):
    with pytest.raises(SystemExit) as salida:
        asistencia.main(["--mes", "2027-01"])
    assert salida.value.code == 2

    error = capsys.readouterr().err
    assert "No hay registros para el mes 2027-01" in error
