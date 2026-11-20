#!/usr/bin/env python3
"""Genera favicon.ico (16/32/48), favicon-32.png y apple-touch-icon.png (180) en app/static a partir del escudo.
Para cambiar el icono: reemplace scripts/escudo_fuente.png (idealmente cuadrado, >= 512 px; si no tiene transparencia
el blanco de fondo se vuelve transparente) y ejecute:  python scripts/generar_iconos.py [ruta_fuente.png]"""
import os, sys
import numpy as np
from PIL import Image

AQUI = os.path.dirname(os.path.abspath(__file__))
DESTINO = os.path.join(AQUI, '..', 'app', 'static')


def preparar(img):
    """Devuelve RGBA cuadrado con fondo transparente (si la fuente es opaca, el blanco pasa a transparente)."""
    img = img.convert('RGBA'); a = np.array(img).astype(float)
    if a[..., 3].min() == 255:                      # fuente opaca: blanco -> alfa segun luminosidad
        lum = a[..., :3].min(axis=2)
        alfa = np.clip((255 - lum) / (255 - 200) * 255, 0, 255)   # 255 = blanco puro -> 0; <=200 -> opaco
        a[..., 3] = alfa
        # des-premultiplicar el halo blanco para bordes limpios
        k = np.maximum(alfa / 255.0, 1e-3)[..., None]
        a[..., :3] = np.clip(255 - (255 - a[..., :3]) / k, 0, 255)
    img = Image.fromarray(a.astype('uint8'), 'RGBA')
    bbox = img.getchannel('A').point(lambda v: 255 if v > 8 else 0).getbbox()
    if bbox: img = img.crop(bbox)
    lado = max(img.size); lienzo = Image.new('RGBA', (lado, lado), (0, 0, 0, 0))
    lienzo.paste(img, ((lado - img.width) // 2, (lado - img.height) // 2), img)
    return lienzo


def main():
    fuente = sys.argv[1] if len(sys.argv) > 1 else os.path.join(AQUI, 'escudo_fuente.png')
    base = preparar(Image.open(fuente))
    def tam(n, margen=0.0):
        inner = int(round(n * (1 - 2 * margen))); im = base.resize((inner, inner), Image.LANCZOS)
        c = Image.new('RGBA', (n, n), (0, 0, 0, 0)); c.paste(im, ((n - inner) // 2, (n - inner) // 2), im); return c
    os.makedirs(DESTINO, exist_ok=True)
    grande = base.resize((256, 256), Image.LANCZOS)
    grande.save(os.path.join(DESTINO, 'favicon.ico'), sizes=[(16, 16), (32, 32), (48, 48)])
    tam(32).save(os.path.join(DESTINO, 'favicon-32.png'), optimize=True)
    ap = Image.new('RGBA', (180, 180), (255, 255, 255, 255)); e = tam(180, 0.08); ap.paste(e, (0, 0), e)   # iOS: fondo opaco
    ap.convert('RGB').save(os.path.join(DESTINO, 'apple-touch-icon.png'), optimize=True)
    print('Iconos generados en', os.path.normpath(DESTINO))


if __name__ == '__main__':
    main()
