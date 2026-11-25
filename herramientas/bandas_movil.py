#!/usr/bin/env python3
"""Regenera las bandas -movil anclando cada corte a una franja blanca cercana.
Modo 'fino': vale cualquier separacion >=3 px (entre filas de tabla).
Modo 'ancho': solo separaciones de bloque >=22 px (entre parrafos)."""
from PIL import Image
import os

DOCS = '/home/yeyo/ingresos-ia/portfolio/docs/img'
GAP = 18
AIRE_FIN = 24  # aire blanco tras la ultima tira
BLANK = 250

# nombre: (fracciones, modo)
BANDAS = {
    'consolidador':    ([(0.00, 0.48)], 'ancho'),
    'precios-cambios': ([(0.00, 1.00)], 'ancho'),
    'certificados':    ([(0.00, 0.78)], 'ancho'),
    'faena':           ([(0.00, 1.00)], 'ancho'),
    'cotizacion':      ([(0.00, 0.145), (0.28, 0.665)], 'fino'),
    'asistencia':      ([(0.00, 0.25), (0.906, 1.00)], 'fino'),
}

def gap_y(g, min_h):
    w, h = g.size
    px = g.load()
    franjas, inicio = [], None
    for y in range(h):
        blancos = sum(1 for x in range(0, w, 3) if px[x, y] >= BLANK)
        llena = blancos / (w / 3) > 0.92
        if llena and inicio is None:
            inicio = y
        elif not llena and inicio is not None:
            if y - inicio >= min_h:
                franjas.append((inicio + y) / 2)
            inicio = None
    return franjas

for nombre, (bandas, modo) in BANDAS.items():
    path = os.path.join(DOCS, nombre + '.png')
    im = Image.open(path).convert('RGB')
    w, h = im.size
    cortes = gap_y(im.convert('L'), 3 if modo == 'fino' else 22)
    ventana = 0.08 * h
    tiras = []
    for a, b in bandas:
        y0, y1 = round(a * h), round(b * h)
        if a > 0:
            cerca = [c for c in cortes if abs(c - y0) <= ventana]
            if cerca:
                y0 = round(min(cerca, key=lambda c: abs(c - y0)))
        if b < 1:
            cerca = [c for c in cortes if abs(c - y1) <= ventana]
            if cerca:
                y1 = round(min(cerca, key=lambda c: abs(c - y1)))
        tiras.append(im.crop((0, y0, w, y1)))
    alto = sum(t.height for t in tiras) + GAP * (len(tiras) - 1)
    if tiras[-1].height + (tiras[0].height if len(tiras) == 1 else 0) < h:
        alto += AIRE_FIN  # la ultima tira va seguida de aire, no de un corte a filo
    banda = Image.new('RGB', (w, alto), (255, 255, 255))
    y = 0
    for t in tiras:
        banda.paste(t, (0, y))
        y += t.height + GAP
    out = os.path.join(DOCS, nombre + '-movil.png')
    banda.save(out, optimize=True)
    print(nombre, '->', banda.size, os.path.getsize(out) // 1024, 'KB')
