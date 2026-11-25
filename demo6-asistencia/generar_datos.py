"""Genera los CSV de ejemplo del analizador de asistencia (datos ficticios).

El resultado es determinista: misma semilla, mismos archivos. Los casos
(tardanzas, salidas antes, extras, ausencias y olvidos) están declarados
explícitamente para que el demo muestre situaciones reales repartidas por
el mes.
"""

import csv
import random
from datetime import date
from pathlib import Path

DIRECTORIO = Path(__file__).resolve().parent
SEMILLA = 20261102

DIAS_HABILES = [
    date(2026, 11, dia)
    for dia in (2, 3, 4, 5, 6, 9, 10, 11, 12, 13, 16, 17, 18, 19, 20, 23, 24, 25, 26, 27)
]

TURNOS = {
    "Mañana": {"entrada": "06:00", "salida": "14:00", "tolerancia": 5},
    "Día": {"entrada": "08:00", "salida": "17:00", "tolerancia": 10},
    "Tarde": {"entrada": "14:00", "salida": "22:00", "tolerancia": 10},
}

EMPLEADOS = {
    **{f"EMP-{numero:03d}": "Mañana" for numero in range(1, 11)},
    **{f"EMP-{numero:03d}": "Día" for numero in range(11, 23)},
    **{f"EMP-{numero:03d}": "Tarde" for numero in range(23, 31)},
}

# Días con ausencias: 13 y 24 con seis cada uno para que la cobertura caiga
# bajo el mínimo; el resto entre 0 y 2.
AUSENCIAS = {
    "2026-11-05": {"EMP-018"},
    "2026-11-10": {"EMP-008", "EMP-025"},
    "2026-11-13": {"EMP-002", "EMP-007", "EMP-012", "EMP-016", "EMP-024", "EMP-029"},
    "2026-11-17": {"EMP-022"},
    "2026-11-20": {"EMP-001", "EMP-013"},
    "2026-11-24": {"EMP-004", "EMP-009", "EMP-014", "EMP-019", "EMP-026", "EMP-030"},
    "2026-11-26": {"EMP-028"},
}

OLVIDOS = {
    ("EMP-003", "2026-11-11"),
    ("EMP-015", "2026-11-18"),
    ("EMP-021", "2026-11-05"),
    ("EMP-027", "2026-11-25"),
}

# Minutos de atraso respecto de la entrada ideal (siempre sobre la tolerancia).
TARDANZAS = {
    ("EMP-001", "2026-11-03"): 18,
    ("EMP-003", "2026-11-05"): 25,
    ("EMP-005", "2026-11-09"): 40,
    ("EMP-006", "2026-11-17"): 14,
    ("EMP-008", "2026-11-24"): 22,
    ("EMP-010", "2026-11-13"): 31,
    ("EMP-012", "2026-11-06"): 12,
    ("EMP-014", "2026-11-11"): 37,
    ("EMP-016", "2026-11-20"): 16,
    ("EMP-018", "2026-11-04"): 28,
    ("EMP-019", "2026-11-26"): 33,
    ("EMP-022", "2026-11-10"): 19,
    ("EMP-024", "2026-11-18"): 24,
    ("EMP-026", "2026-11-02"): 15,
    ("EMP-028", "2026-11-23"): 40,
    ("EMP-001", "2026-11-19"): 20,
    ("EMP-005", "2026-11-25"): 13,
    ("EMP-012", "2026-11-27"): 29,
}

# Atrasos dentro de la tolerancia: no deben aparecer en Tardanzas del Mes.
TARDANZAS_LEVES = {
    ("EMP-011", "2026-11-03"): 7,
    ("EMP-023", "2026-11-05"): 8,
    ("EMP-007", "2026-11-19"): 4,
}

# Minutos antes de la salida ideal (más que la tolerancia del turno).
SALIDAS_ANTES = {
    ("EMP-002", "2026-11-09"): 35,
    ("EMP-007", "2026-11-16"): 42,
    ("EMP-013", "2026-11-11"): 28,
    ("EMP-020", "2026-11-24"): 50,
    ("EMP-025", "2026-11-06"): 33,
    ("EMP-029", "2026-11-19"): 60,
}

# Minutos después de la salida ideal (siempre más que la gracia de 15 min).
EXTRAS = {
    ("EMP-003", "2026-11-02"): 75,
    ("EMP-004", "2026-11-06"): 120,
    ("EMP-008", "2026-11-12"): 55,
    ("EMP-011", "2026-11-16"): 90,
    ("EMP-015", "2026-11-03"): 45,
    ("EMP-017", "2026-11-18"): 150,
    ("EMP-021", "2026-11-09"): 65,
    ("EMP-023", "2026-11-20"): 100,
    ("EMP-027", "2026-11-11"): 80,
    ("EMP-030", "2026-11-04"): 110,
}


def _a_minutos(hora):
    """Convierte 'HH:MM' a minutos desde la medianoche."""
    horas, minutos = (int(parte) for parte in hora.split(":"))
    return horas * 60 + minutos


def _a_hora(minutos):
    """Convierte minutos desde la medianoche a 'HH:MM' con dos dígitos."""
    return f"{minutos // 60:02d}:{minutos % 60:02d}"


def _estado(empleado, dia):
    """Devuelve estado y horas de la fila según los casos declarados."""
    clave = (empleado, dia.isoformat())
    if empleado in AUSENCIAS.get(dia.isoformat(), set()):
        return "Ausencia", "", ""
    if clave in OLVIDOS:
        return "Olvido marcaje", "", ""

    turno = TURNOS[EMPLEADOS[empleado]]
    entrada = _a_minutos(turno["entrada"])
    salida = _a_minutos(turno["salida"])

    # Jitter leve para que los registros no se vean perfectos.
    entrada += random.randint(-4, 2)
    salida += random.randint(-3, 3)

    if clave in TARDANZAS:
        entrada = _a_minutos(turno["entrada"]) + TARDANZAS[clave]
    elif clave in TARDANZAS_LEVES:
        entrada = _a_minutos(turno["entrada"]) + TARDANZAS_LEVES[clave]

    if clave in SALIDAS_ANTES:
        salida = _a_minutos(turno["salida"]) - SALIDAS_ANTES[clave]
    elif clave in EXTRAS:
        salida = _a_minutos(turno["salida"]) + EXTRAS[clave]

    return "Asistió", _a_hora(entrada), _a_hora(salida)


def main():
    """Escribe reloj_asistencia.csv y turnos_esperados.csv."""
    random.seed(SEMILLA)

    with (DIRECTORIO / "reloj_asistencia.csv").open(
        "w", encoding="utf-8", newline=""
    ) as archivo:
        escritor = csv.writer(archivo)
        escritor.writerow(
            ["empleado", "fecha", "turno", "estado", "hora_entrada", "hora_salida"]
        )
        for empleado, turno in EMPLEADOS.items():
            for dia in DIAS_HABILES:
                estado, entrada, salida = _estado(empleado, dia)
                escritor.writerow(
                    [empleado, dia.isoformat(), turno, estado, entrada, salida]
                )

    with (DIRECTORIO / "turnos_esperados.csv").open(
        "w", encoding="utf-8", newline=""
    ) as archivo:
        escritor = csv.writer(archivo)
        escritor.writerow(
            ["turno", "hora_entrada_ideal", "hora_salida_ideal", "tolerancia_min"]
        )
        for turno, datos in TURNOS.items():
            escritor.writerow(
                [turno, datos["entrada"], datos["salida"], datos["tolerancia"]]
            )

    print("Datos generados: reloj_asistencia.csv y turnos_esperados.csv")


if __name__ == "__main__":
    main()
