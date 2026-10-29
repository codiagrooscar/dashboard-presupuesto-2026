#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
output_staging.py
Módulo centralizado para redirección de salidas y entradas durante la generación en staging.
Garantiza que la generación y la activación estén completamente separadas y que los scripts
encadenados consuman las versiones generadas en la misma corrida.
"""

import os
import sys

STAGING_DIR = os.environ.get("CODIAGRO_STAGING_DIR", "").strip()

def is_staging_active():
    return bool(STAGING_DIR)

def get_dest(path):
    """
    Retorna la ruta destino dentro de STAGING_DIR si está definido,
    o la ruta original si se ejecuta en modo estándar sin staging.
    Crea automáticamente los directorios padre necesarios.
    """
    if STAGING_DIR:
        norm_path = os.path.normpath(path)
        norm_staging = os.path.normpath(STAGING_DIR)
        if norm_path.startswith(norm_staging):
            target = path
        else:
            if os.path.isabs(path):
                try:
                    rel = os.path.relpath(path, os.getcwd())
                    if not rel.startswith(".."):
                        target = os.path.join(STAGING_DIR, rel)
                    else:
                        target = path
                except Exception:
                    target = path
            else:
                target = os.path.join(STAGING_DIR, path)
        parent = os.path.dirname(target)
        if parent:
            os.makedirs(parent, exist_ok=True)
        return target
    return path

def get_source(path):
    """
    Si STAGING_DIR está activo y el archivo existe dentro de STAGING_DIR,
    retorna la ruta dentro de STAGING_DIR para encadenar scripts durante la generación.
    Si no existe en STAGING_DIR, retorna path original.
    """
    if STAGING_DIR:
        staged = get_dest(path)
        if os.path.exists(staged):
            return staged
    return path
