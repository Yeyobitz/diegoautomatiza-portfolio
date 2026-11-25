#!/usr/bin/env python3
"""Recorta margenes muertos de las capturas de demos (pasada 3 del jurado).

- Cada PNG se recorta a su caja de contenido (borde blanco >= 242) + 8 px de aire.
- Se generan variantes -movil.png: bandas horizontales con las 3-4 filas mas
  contundentes de cada documento, para la vista previa scrolleable en telefono.
"""
from PIL import Image
import os

DOCS = '/home/yeyo/ingresos-ia/portfolio/docs/img'
PAD = 8
BLANK = 242

def bbox_contenido(im):
    g = im.convert('L')
    w, h = g.size
    px = g.load()
    def fila_vacia(y):
        return all(px[x, y] >= BLANK for x in range(w))
    def col_vacia(x):
        return all(px[x, y] >= BLANK for y in range(h))
    top = 0
    while top < h - 1 and fila_vacia(top):
        top += 1
    bot = h - 1
    while bot > top and fila_vacia(bot):
        bot -= 1
    izq = 0
    while izq < w - 1 and col_vacia(izq):
        izq += 1
    der = w - 1
    while der > izq and col_vacia(der):
        der -= 1
    return (izq, top, der + 1, bot + 1)

# bandas moviles: lista de (y0, y1) como fraccion del alto recortado
BANDAS = {
    'consolidador':    [(0.00, 0.52)],
    'precios-cambios': [(0.00, 1.00)],
    'certificados':    [(0.00, 0.55)],
    'faena':           [(0.00, 1.00)],
    'cotizacion':      [(0.00, 0.145), (0.28, 0.70)],
    'asistencia':      [(0.00, 0.30), (0.885, 0.985)],
}
GAP = 18

for nombre in BANDAS:
    path = os.path.join(DOCS, nombre + '.png')
    im = Image.open(path).convert('RGB')
    caja = bbox_contenido(im)
    caja = (max(0, caja[0] - PAD), max(0, caja[1] - PAD),
            min(im.width, caja[2] + PAD), min(im.height, caja[3] + PAD))
    recorte = im.crop(caja)
    recorte.save(path, optimize=True)
    print(f'{nombre}: {im.size} -> {recorte.size}')

    # banda movil
    w, h = recorte.size
    tiras = [recorte.crop((0, round(a * h), w, round(b * h))) for a, b in BANDAS[nombre]]
    alto = sum(t.height for t in tiras) + GAP * (len(tiras) - 1)
    banda = Image.new('RGB', (w, alto), (255, 255, 255))
    y = 0
    for t in tiras:
        banda.paste(t, (0, y))
        y += t.height + GAP
    out = os.path.join(DOCS, nombre + '-movil.png')
    banda.save(out, optimize=True)
    print(f'  {nombre}-movil: {banda.size} {os.path.getsize(out)//1024} KB')
