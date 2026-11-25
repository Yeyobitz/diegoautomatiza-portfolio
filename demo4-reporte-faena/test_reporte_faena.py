"""Tests del generador de reporte diario de faena."""

import shutil
from pathlib import Path

import pytest
from openpyxl import load_workbook

import reporte_faena

DIRECTORIO = Path(__file__).resolve().parent
REPORTE_REAL = DIRECTORIO / "reporte_jefe.txt"
HORAS_REALES = DIRECTORIO / "horas_cuadrillas.xlsx"
FECHA = "2026-11-23"


def _valores_hoja(ruta_xlsx, nombre_hoja):
    """Junta todos los valores de una hoja para inspeccionarlos."""
    libro = load_workbook(ruta_xlsx)
    hoja = libro[nombre_hoja]
    return [celda.value for fila in hoja.iter_rows() for celda in fila]


def test_parseo_frases_clave():
    reporte = reporte_faena.parsear_reporte(REPORTE_REAL, FECHA)

    permisos = {permiso["pt"]: permiso for permiso in reporte["permisos"]}
    assert permisos["PT-4471"]["estado"] == "Aprobado"
    assert permisos["PT-4482"]["estado"] == "Rechazado"

    assert reporte["avance"]["PLANTA-NORTE"] == 82
    assert "colacion" in reporte["cuadrillas_ausentes"][3]
    assert any(evento["tipo"] == "Casi Incidente" for evento in reporte["eventos"])


def test_clasificacion_severidad():
    assert (
        reporte_faena.clasificar_severidad(
            "casi incidente: herramienta cayo 3 m, sin lesionados"
        )
        == "Baja"
    )
    assert reporte_faena.clasificar_severidad("incidente con lesionados leve") == "Alta"
    assert (
        reporte_faena.clasificar_severidad("incidente: daño menor, sin lesionados")
        == "Media"
    )


def test_suma_hh():
    horas = reporte_faena.cargar_horas(HORAS_REALES, FECHA)
    assert horas["HH_normales"].sum() == 450
    assert horas["HH_extras"].sum() == 38
    assert horas["n_trabajadores"].sum() == 50


@pytest.mark.parametrize("fecha", ["23/11/2026", "2026-02-31"])
def test_fecha_invalida(fecha):
    with pytest.raises(SystemExit):
        reporte_faena.main(["--fecha", fecha])


def test_sin_incidentes_no_rompe(tmp_path):
    texto = "\n".join(
        [
            "23/11/26, 00:47 - turno noche corriendo normal, dotacion 42 en faena",
            "23/11/26, 05:02 - avance PLANTA-NORTE 82%",
            "23/11/26, 05:30 - permiso PT-4485 aprobado para trabajo en caliente en planta PLANTA-NORTE",
        ]
    )
    ruta_txt = tmp_path / "reporte.txt"
    ruta_txt.write_text(texto, encoding="utf-8")

    resultado = reporte_faena.generar_reporte(
        ruta_txt, HORAS_REALES, FECHA, salida=tmp_path
    )

    assert resultado["xlsx"].is_file()
    valores = _valores_hoja(resultado["xlsx"], "Incidentes y Casi Incidentes")
    assert any(
        valor and "Sin incidentes" in str(valor) for valor in valores
    )


@pytest.mark.skipif(
    shutil.which("soffice") is None,
    reason="soffice (LibreOffice) no está disponible",
)
def test_pdf_existe(tmp_path):
    resultado = reporte_faena.generar_reporte(
        REPORTE_REAL, HORAS_REALES, FECHA, salida=tmp_path
    )
    assert resultado["pdf"] is not None
    assert resultado["pdf"].is_file()


def test_turno_detectado():
    reporte = reporte_faena.parsear_reporte(REPORTE_REAL, FECHA)
    assert reporte["turno"] == "Noche"


def test_resumen_documento_controlado(tmp_path):
    resultado = reporte_faena.generar_reporte(
        REPORTE_REAL,
        HORAS_REALES,
        FECHA,
        salida=tmp_path,
        clima="Nublado",
        estado_faena="Restringida",
    )
    libro = load_workbook(resultado["xlsx"])
    hoja = libro["Resumen del Día"]
    textos = [
        str(celda.value)
        for fila in hoja.iter_rows()
        for celda in fila
        if celda.value is not None
    ]

    assert any("RDF-2026-084" in texto for texto in textos)
    assert any("RDF-01 Rev. 2" in texto for texto in textos)
    assert "Clima del día:" in textos and "Nublado" in textos
    assert "Estado de faena:" in textos and "Restringida" in textos
    assert "A1:F1" in {str(rango) for rango in hoja.merged_cells.ranges}
    assert hoja.row_dimensions[1].height == 36

    # Los frentes sin avance y la fila TOTALES muestran "—", no celdas vacías.
    avances = [
        hoja.cell(row=fila, column=6).value
        for fila in range(1, hoja.max_row + 1)
        if str(hoja.cell(row=fila, column=1).value).upper() in {"PATIO-1", "TOTALES"}
    ]
    assert avances == ["—", "—"]
