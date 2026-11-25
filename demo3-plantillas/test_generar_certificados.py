"""Pruebas del emisor de certificados con folio."""

import csv
from datetime import date
import shutil
from pathlib import Path

from openpyxl import load_workbook
import pandas as pd
import pytest

import generar_certificados as gc


COLUMNAS = ["nombre", "curso", "fecha", "nota", "tipo"]
SOFFICE = shutil.which("soffice")


@pytest.fixture
def plantillas(tmp_path):
    directorio = tmp_path / "plantillas"
    gc.construir_todas(directorio)
    return directorio


def escribir_csv(ruta, filas, columnas=COLUMNAS):
    with open(ruta, "w", encoding="utf-8", newline="") as archivo:
        escritor = csv.writer(archivo)
        escritor.writerow(columnas)
        escritor.writerows(filas)


def emitir(tmp_path, plantillas, filas, nombre_csv="participantes.csv"):
    entrada = tmp_path / nombre_csv
    escribir_csv(entrada, filas)
    salida = tmp_path / "salida"
    resumen = gc.main(
        [
            "--input",
            str(entrada),
            "--output",
            str(salida),
            "--plantilla",
            str(plantillas),
        ]
    )
    return resumen, salida


def test_folios_unicos_y_consecutivos(tmp_path, plantillas):
    filas = [
        (f"Persona {numero}", "Curso", "2026-03-15", 6.0, "participacion")
        for numero in range(1, 6)
    ]
    filas += [
        (f"Distinguida {numero}", "Curso", "2026-03-15", 7.0, "distincion")
        for numero in range(1, 3)
    ]

    resumen, _ = emitir(tmp_path, plantillas, filas)
    año = date.today().year
    esperados = [f"CERT-{año}-{n:04d}" for n in range(1, 6)]
    esperados += [f"DIST-{año}-{n:04d}" for n in range(1, 3)]

    assert resumen["folios"] == esperados
    assert len(set(resumen["folios"])) == 7


def test_marcadores_reemplazados(tmp_path, plantillas):
    filas = [
        ("Camila Aravena", "Análisis de Datos con Excel", "2026-03-15", 6.5, "participacion")
    ]

    resumen, _ = emitir(tmp_path, plantillas, filas)
    libro = load_workbook(resumen["archivos"][0])
    textos = [
        celda.value
        for hoja in libro.worksheets
        for fila in hoja.iter_rows()
        for celda in fila
        if isinstance(celda.value, str)
    ]

    assert textos
    assert not any("{{" in texto for texto in textos)
    contenido = "\n".join(textos)
    for esperado in (
        "Camila Aravena",
        "Análisis de Datos con Excel",
        "6.5",
        resumen["folios"][0],
        "Santiago, Chile",
        "15 de marzo de 2026",
    ):
        assert esperado in contenido


def test_csv_malformado(tmp_path, plantillas):
    sin_nota = tmp_path / "sin_nota.csv"
    escribir_csv(
        sin_nota,
        [("Ana Pérez", "Curso", "2026-03-15", "participacion")],
        columnas=["nombre", "curso", "fecha", "tipo"],
    )
    with pytest.raises(SystemExit) as error:
        gc.main(
            [
                "--input",
                str(sin_nota),
                "--output",
                str(tmp_path / "salida"),
                "--plantilla",
                str(plantillas),
            ]
        )
    mensaje = str(error.value).lower()
    assert "nota" in mensaje
    assert "fila" in mensaje

    nota_invalida = tmp_path / "nota_invalida.csv"
    escribir_csv(
        nota_invalida,
        [
            ("Ana Pérez", "Curso", "2026-03-15", 6.0, "participacion"),
            ("Beto Soto", "Curso", "2026-03-15", "abc", "participacion"),
        ],
    )
    with pytest.raises(SystemExit) as error:
        gc.main(
            [
                "--input",
                str(nota_invalida),
                "--output",
                str(tmp_path / "salida"),
                "--plantilla",
                str(plantillas),
            ]
        )
    mensaje = str(error.value).lower()
    assert "fila 3" in mensaje
    assert "nota" in mensaje

    fecha_invalida = tmp_path / "fecha_invalida.csv"
    escribir_csv(
        fecha_invalida,
        [
            ("Ana Pérez", "Curso", "2026-03-15", 6.0, "participacion"),
            ("Beto Soto", "Curso", "2026-13-01", 6.0, "participacion"),
        ],
    )
    with pytest.raises(SystemExit) as error:
        gc.main(
            [
                "--input",
                str(fecha_invalida),
                "--output",
                str(tmp_path / "salida"),
                "--plantilla",
                str(plantillas),
            ]
        )
    mensaje = str(error.value).lower()
    assert "fila 3" in mensaje
    assert "fecha" in mensaje


def test_plantilla_inexistente(tmp_path):
    entrada = tmp_path / "participantes.csv"
    escribir_csv(
        entrada,
        [("Ana Pérez", "Curso", "2026-03-15", 6.0, "participacion")],
    )
    salida = tmp_path / "salida"

    with pytest.raises(SystemExit) as error:
        gc.main(
            [
                "--input",
                str(entrada),
                "--output",
                str(salida),
                "--plantilla",
                str(tmp_path / "no_existe"),
            ]
        )

    assert "plantilla" in str(error.value).lower()
    assert not salida.exists()


@pytest.mark.skipif(SOFFICE is None, reason="LibreOffice (soffice) no está disponible")
def test_pdf_generado_existe_y_no_vacio(tmp_path, plantillas):
    filas = [
        ("Ana Pérez", "Curso", "2026-03-15", 6.0, "participacion"),
        ("Luis Soto", "Curso", "2026-03-15", 6.5, "distincion"),
    ]

    resumen, _ = emitir(tmp_path, plantillas, filas)

    assert len(resumen["archivos"]) == 2
    for xlsx in resumen["archivos"]:
        pdf = Path(xlsx).with_suffix(".pdf")
        assert pdf.is_file()
        assert pdf.stat().st_size > 1000
        assert pdf.read_bytes()[:4] == b"%PDF"


def test_registro_completo(tmp_path, plantillas):
    filas = [
        ("Ana Pérez", "Curso A", "2026-03-15", 6.0, "participacion"),
        ("Luis Soto", "Curso B", "2026-04-15", 6.5, "distincion"),
    ]

    _, salida = emitir(tmp_path, plantillas, filas)
    libro = load_workbook(salida / "registro_emision.xlsx")
    hoja = libro.active

    assert [celda.value for celda in hoja[1]] == [
        "folio",
        "nombre",
        "curso",
        "fecha",
        "fecha_emision",
        "archivo",
    ]
    assert hoja.max_row - 1 == len(filas)
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        assert (salida / fila[5]).is_file()


def test_folio_continua_con_registro_previo(tmp_path, plantillas):
    primera = [
        ("Ana Pérez", "Curso A", "2026-03-15", 6.0, "participacion"),
        ("Luis Soto", "Curso A", "2026-03-15", 6.5, "participacion"),
    ]
    segunda = [
        ("Eva Rojas", "Curso B", "2026-04-15", 6.8, "participacion"),
        ("Tomás Díaz", "Curso B", "2026-04-15", 6.9, "participacion"),
    ]

    resumen_uno, salida = emitir(tmp_path, plantillas, primera, "uno.csv")
    resumen_dos, _ = emitir(tmp_path, plantillas, segunda, "dos.csv")

    año = date.today().year
    assert resumen_uno["folios"] == [f"CERT-{año}-0001", f"CERT-{año}-0002"]
    assert resumen_dos["folios"] == [f"CERT-{año}-0003", f"CERT-{año}-0004"]

    hoja = load_workbook(salida / "registro_emision.xlsx").active
    assert hoja.max_row - 1 == 4


def test_fecha_iso_ambigua_no_se_invierte(tmp_path):
    df = pd.DataFrame(
        [("Ana Pérez", "Curso", "2026-04-12", 6.0, "participacion")],
        columns=COLUMNAS,
    )

    resultado = gc.validar(df, tmp_path / "participantes.csv")

    assert resultado.iloc[0]["fecha"] == date(2026, 4, 12)


def test_fecha_larga():
    assert gc.fecha_larga(date(2026, 3, 15)) == "15 de marzo de 2026"
