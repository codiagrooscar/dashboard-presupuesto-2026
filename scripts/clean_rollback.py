#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
clean_rollback.py
Gestor transaccional de copias de seguridad y reversión limpia para actualizar_dashboard.bat.
Garantiza que, ante cualquier fallo en la generación o validación:
  1. Se restauren exactamente los archivos previos válidos.
  2. Se purgue cualquier archivo nuevo generado durante la ejecución fallida
     (incluyendo archivos Excel generados en web_dashboard o web_dashboard_direccion).
"""

import sys
import os
import shutil
import json

BACKUP_DIR = os.path.join("scratch", "last_valid_backup")
MANIFEST_FILE = os.path.join(BACKUP_DIR, "manifest.json")

TARGET_DIRS = [
    "web_dashboard",
    "web_dashboard_direccion"
]
ROOT_FILES = [
    "Seguimiento_Presupuesto_Sep_2026.xlsx",
    "Orden_Fabricacion_MRP_Codiagro.xlsx"
]

def take_snapshot():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    manifest = {
        "files": []
    }
    
    # 1. Snapshot de directorios web
    for d in TARGET_DIRS:
        dest_d = os.path.join(BACKUP_DIR, d)
        os.makedirs(dest_d, exist_ok=True)
        for root, dirs, files in os.walk(dest_d, topdown=False):
            for f in files:
                try:
                    os.remove(os.path.join(root, f))
                except Exception:
                    pass
            for dir_name in dirs:
                try:
                    os.rmdir(os.path.join(root, dir_name))
                except Exception:
                    pass
        
        if os.path.exists(d):
            for root, _, files in os.walk(d):
                for f in files:
                    src_f = os.path.join(root, f)
                    rel_f = os.path.relpath(src_f, ".")
                    manifest["files"].append(rel_f)
                    
                    dest_f = os.path.join(BACKUP_DIR, rel_f)
                    os.makedirs(os.path.dirname(dest_f), exist_ok=True)
                    shutil.copy2(src_f, dest_f)

    # 2. Snapshot de archivos raíz
    for rf in ROOT_FILES:
        if os.path.exists(rf):
            manifest["files"].append(rf)
            dest_rf = os.path.join(BACKUP_DIR, rf)
            shutil.copy2(rf, dest_rf)

    with open(MANIFEST_FILE, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print(f"[SNAPSHOT] Copia de seguridad guardada ({len(manifest['files'])} archivos en {BACKUP_DIR}).")

def restore_snapshot():
    if not os.path.exists(MANIFEST_FILE):
        print(f"[AVISO] No se encontró manifiesto en {MANIFEST_FILE}. Nada que restaurar.")
        return 0

    with open(MANIFEST_FILE, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    valid_files_set = set(os.path.normpath(p) for p in manifest.get("files", []))

    # 1. Purgar cualquier archivo nuevo que no estuviese en el snapshot
    purged_count = 0
    for d in TARGET_DIRS:
        if os.path.exists(d):
            for root, _, files in os.walk(d):
                for f in files:
                    current_path = os.path.normpath(os.path.join(root, f))
                    if current_path not in valid_files_set:
                        try:
                            os.remove(current_path)
                            purged_count += 1
                        except Exception as e:
                            print(f"[AVISO] No se pudo eliminar archivo huérfano {current_path}: {e}")

    for rf in ROOT_FILES:
        current_rf = os.path.normpath(rf)
        if os.path.exists(rf) and current_rf not in valid_files_set:
            try:
                os.remove(rf)
                purged_count += 1
            except Exception as e:
                print(f"[AVISO] No se pudo eliminar archivo raíz nuevo {rf}: {e}")

    # 2. Restaurar todos los archivos válidos desde el backup
    restored_count = 0
    for rel_f in manifest.get("files", []):
        backup_src = os.path.join(BACKUP_DIR, rel_f)
        if os.path.exists(backup_src):
            parent_dir = os.path.dirname(rel_f)
            if parent_dir:
                os.makedirs(parent_dir, exist_ok=True)
            shutil.copy2(backup_src, rel_f)
            restored_count += 1

    print(f"[RESTAURACIÓN LIMPIA] {restored_count} archivos restaurados; {purged_count} archivos huérfanos purgados.")
    return 0

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "restore":
        sys.exit(restore_snapshot())
    else:
        take_snapshot()
        sys.exit(0)

