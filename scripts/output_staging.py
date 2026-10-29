#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
output_staging.py
Módulo centralizado para redirección de salidas a la carpeta temporal de staging
durante la ejecución de actualizar_dashboard.bat.
Garantiza que la generación y la activación estén completamente separadas.
"""

import os

STAGING_DIR = os.environ.get("CODIAGRO_STAGING_DIR", "").strip()

def get_dest(path):
    """
    Retorna la ruta destino dentro de STAGING_DIR si está definido,
    o la ruta original si se ejecuta en modo estándar sin staging.
    Crea automáticamente los directorios padre necesarios.
    """
    if STAGING_DIR:
        # Si la ruta ya empieza con STAGING_DIR, no duplicar
        norm_path = os.path.normpath(path)
        norm_staging = os.path.normpath(STAGING_DIR)
        if norm_path.startswith(norm_staging):
            target = path
        else:
            target = os.path.join(STAGING_DIR, path)
        parent = os.path.dirname(target)
        if parent:
            os.makedirs(parent, exist_ok=True)
        return target
    return path

