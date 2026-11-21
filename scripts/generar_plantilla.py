#!/usr/bin/env python3
"""Regenera app/static/plantilla_comunicaciones_internas.xlsx (plantilla vacia, sin datos). Uso: python scripts/generar_plantilla.py [ruta_destino]"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app import plantilla

destino = sys.argv[1] if len(sys.argv) > 1 else plantilla.RUTA_ESTATICA
print(plantilla.generar_plantilla(destino))
