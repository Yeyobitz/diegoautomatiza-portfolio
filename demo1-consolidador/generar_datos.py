#!/usr/bin/env python3
"""Genera las tres exportaciones de ejemplo del Conciliador de ventas multicanal.

Simula lo que llega a la empresa todos los lunes por correo:
  - Local  -> ventas_local.xlsx  [Folio, Fecha (dd/mm/aaaa), SKU, Vendedor, Unidades, Monto]
  - Web    -> ventas_web.xlsx    [Order ID, Order Date (aaaa-mm-dd), Product Code, Sales Rep, Qty, Revenue]
  - POS    -> ventas_pos.csv     [ticket, timestamp ISO, sku, cajero, cantidad, total "$1.234,56"]

Todo es reproducible (semilla fija) y los errores incluidos son los típicos de
una exportación mal hecha: fechas imposibles (32/13/2025), montos con texto
("N/D") y filas repetidas entre canales porque alguien cargó la misma venta dos
veces en sistemas distintos.
"""

import csv
import random
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

SEED = 20251124
CARPETA = Path(__file__).parent / "datos"

# Catálogo de material de construcción (códigos verosímiles de faena).
CATALOGO = {
    "TOR-4X25": 12990,   # tornillo autoperforante 4x25, caja x500
    "TOR-3X16": 9490,    # tornillo drywall 3x16, caja x1000
    "CAT-4X25": 4590,    # cantoneta metálica 4x25
    "MIX-118": 39990,    # mezclador de mortero 118
    "DRY-STD15": 8790,   # plancha drywall estándar 15 mm
    "PER-7030": 6290,    # perfil estructurero 70x30 x 6 m
    "CEM-B25": 4390,     # cemento bolsa 25 kg
    "AIS-LV100": 16990,  # lana de vidrio 100 mm
    "PVC-110": 10490,    # tubo PVC sanitario 110 mm
    "PNT-INT20": 34990,  # pintura interior 20 L
    "IMP-ACR20": 41990,  # impermeabilizante acrílico 20 L
    "SIL-AC280": 2190,   # sellador acrílico 280 ml
    "PLA-OND6": 24990,   # plancha ondulada 6 m
    "GUA-NIT9": 1590,    # guante nitrilo talla 9 (par)
    "FIJ-TEX25": 5190,   # fijador textil 25 kg
}

VENDEDORES_LOCAL = [
    "Camila Rojas", "Matías Fuentes", "Carolina Soto",
    "Rodrigo Salinas", "Valentina Muñoz", "Ignacio Pérez",
]
VENDEDORES_WEB = [
    "Camila Rojas", "Matías Fuentes", "Carolina Soto",
    "Fernanda Urbina", "Sebastián Herrera", "Valentina Muñoz",
]
CAJEROS_POS = ["J. Cifuentes", "M. Espinoza", "B. Quilodrán", "H. Vargas"]

INICIO = datetime(2025, 8, 1)   # meses cruzados: agosto a noviembre 2025
DIAS = 118                      # hasta fines de noviembre


def fecha_aleatoria(rng):
    return INICIO + timedelta(days=rng.randrange(DIAS))


def monto_venta(rng, sku, unidades, exacto=False):
    base = CATALOGO[sku] * unidades
    if exacto:
        return int(base)
    return int(round(base * rng.uniform(0.97, 1.03) / 10) * 10)


def monto_chileno(valor):
    """Formatea 123456.5 -> '$123.456,50' (punto miles, coma decimal)."""
    texto = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"${texto}"


def generar_local(rng, duplicados_cruzados):
    filas = []
    folio = 24801
    # 1.895 ventas válidas + 3 duplicados exactos dentro del mismo archivo
    for _ in range(1895):
        sku = rng.choice(list(CATALOGO))
        unidades = rng.randint(1, 6) if CATALOGO[sku] > 15000 else rng.randint(1, 40)
        filas.append({
            "Folio": folio,
            "Fecha": fecha_aleatoria(rng).strftime("%d/%m/%Y"),
            "SKU": sku,
            "Vendedor": rng.choice(VENDEDORES_LOCAL),
            "Unidades": unidades,
            "Monto": monto_venta(rng, sku, unidades),
        })
        folio += 1
    for _ in range(3):
        filas.append(dict(filas[rng.randrange(1500)]))
    # 1 fecha imposible
    filas[-1] = {"Folio": folio, "Fecha": "32/13/2025", "SKU": "CEM-B25",
                 "Vendedor": "Carolina Soto", "Unidades": 10, "Monto": 43900}
    # 1 monto con texto
    filas.append({"Folio": folio + 1, "Fecha": "14/10/2025", "SKU": "TOR-4X25",
                  "Vendedor": "Camila Rojas", "Unidades": 3, "Monto": "N/D"})
    # duplicados inter-canal: la misma venta también quedó registrada en otro canal
    for fila in duplicados_cruzados["Local"]:
        filas.append(fila)
    return filas


def generar_web(rng, duplicados_cruzados):
    filas = []
    order = 88201
    for _ in range(1748):
        sku = rng.choice(list(CATALOGO))
        unidades = rng.randint(1, 4) if CATALOGO[sku] > 15000 else rng.randint(1, 30)
        filas.append({
            "Order ID": order,
            "Order Date": fecha_aleatoria(rng).strftime("%Y-%m-%d"),
            "Product Code": sku,
            "Sales Rep": rng.choice(VENDEDORES_WEB),
            "Qty": unidades,
            "Revenue": monto_venta(rng, sku, unidades),
        })
        order += 1
    filas.append(dict(filas[200]))
    # 1 fecha imposible (30 de febrero)
    filas[-1] = {"Order ID": order, "Order Date": "2025-02-30", "Product Code": "MIX-118",
                 "Sales Rep": "Sebastián Herrera", "Qty": 1, "Revenue": 39990}
    for fila in duplicados_cruzados["Web"]:
        filas.append(fila)
    return filas


def generar_pos(rng, duplicados_cruzados):
    filas = []
    for _ in range(1396):
        sku = rng.choice(list(CATALOGO))
        unidades = rng.randint(1, 3) if CATALOGO[sku] > 15000 else rng.randint(1, 25)
        fecha = fecha_aleatoria(rng)
        hora = fecha.replace(hour=rng.randint(9, 19), minute=rng.randint(0, 59),
                             second=rng.randint(0, 59))
        filas.append({
            "ticket": f"TK-{rng.randrange(100000, 999999)}",
            "timestamp": hora.strftime("%Y-%m-%dT%H:%M:%S"),
            "sku": sku,
            "cajero": rng.choice(CAJEROS_POS),
            "cantidad": unidades,
            "total": monto_chileno(monto_venta(rng, sku, unidades)),
        })
    filas.append(dict(filas[150]))
    filas.append(dict(filas[700]))
    # 1 total con texto
    filas[-2] = {"ticket": "TK-100233", "timestamp": "2025-11-05T16:40:12",
                 "sku": "GUA-NIT9", "cajero": "H. Vargas", "cantidad": 4, "total": "N/D"}
    # 1 timestamp vacío
    filas[-1] = {"ticket": "TK-100234", "timestamp": "", "sku": "CEM-B25",
                 "cajero": "M. Espinoza", "cantidad": 2, "total": "$8.780,00"}
    for fila in duplicados_cruzados["POS"]:
        filas.append(fila)
    return filas


def crear_duplicados_cruzados(rng):
    """22 ventas registradas en dos canales a la vez (misma fecha, sku,
    unidades, monto y responsable), el caso clásico de 'no cuadra el cuadre'."""
    dups = {"Local": [], "Web": [], "POS": []}
    pares = [("Local", "Web", 10), ("Local", "POS", 6), ("Web", "POS", 6)]
    for canal_a, canal_b, cantidad in pares:
        for _ in range(cantidad):
            sku = rng.choice(list(CATALOGO))
            unidades = rng.randint(1, 5)
            fecha = fecha_aleatoria(rng)
            monto = monto_venta(rng, sku, unidades, exacto=True)
            vendedor = rng.choice(VENDEDORES_LOCAL + VENDEDORES_WEB)
            cajero = rng.choice(CAJEROS_POS)
            fecha_ddmm = fecha.strftime("%d/%m/%Y")
            fecha_iso = fecha.strftime("%Y-%m-%d")
            timestamp = fecha_iso + f"T{rng.randint(9, 19):02d}:17:33"

            def fila_para(canal):
                if canal == "Local":
                    return {
                        "Folio": rng.randrange(90000, 99999),
                        "Fecha": fecha_ddmm,
                        "SKU": sku, "Vendedor": vendedor,
                        "Unidades": unidades, "Monto": monto,
                    }
                if canal == "Web":
                    return {
                        "Order ID": rng.randrange(700000, 799999),
                        "Order Date": fecha_iso,
                        "Product Code": sku, "Sales Rep": vendedor,
                        "Qty": unidades, "Revenue": monto,
                    }
                return {
                    "ticket": f"TK-{rng.randrange(100000, 999999)}",
                    "timestamp": timestamp,
                    "sku": sku, "cajero": cajero,
                    "cantidad": unidades, "total": monto_chileno(monto),
                }

            dups[canal_a].append(fila_para(canal_a))
            dups[canal_b].append(fila_para(canal_b))
    return dups


def main():
    rng = random.Random(SEED)
    CARPETA.mkdir(exist_ok=True)
    dups = crear_duplicados_cruzados(rng)

    pd.DataFrame(generar_local(rng, dups)).to_excel(CARPETA / "ventas_local.xlsx", index=False)
    pd.DataFrame(generar_web(rng, dups)).to_excel(CARPETA / "ventas_web.xlsx", index=False)
    with open(CARPETA / "ventas_pos.csv", "w", newline="", encoding="utf-8-sig") as archivo:
        escritor = csv.DictWriter(archivo, fieldnames=["ticket", "timestamp", "sku",
                                                       "cajero", "cantidad", "total"],
                                  delimiter=";")
        escritor.writeheader()
        escritor.writerows(generar_pos(rng, dups))

    for nombre in ("ventas_local.xlsx", "ventas_web.xlsx", "ventas_pos.csv"):
        ruta = CARPETA / nombre
        print(f"{nombre}: {ruta.stat().st_size:,} bytes")


if __name__ == "__main__":
    main()
