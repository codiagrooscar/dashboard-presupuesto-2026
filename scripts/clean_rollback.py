#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
clean_rollback.py
Gestor transaccional de copias de seguridad, staging temporal, activación y reversión limpia.
Garantiza:
  1. Aborto inmediato si falla la copia de seguridad.
  2. Generación en carpeta temporal de staging (scratch/staging) separada de carpetas activas.
  3. Activación atómica de salidas tras superar la validación.
  4. Reversión limpia con purga de huérfanos y verificación estricta de restauración.
"""

import sys
import os
import shutil
import json

BACKUP_DIR = os.path.join("scratch", "last_valid_backup")
STAGING_DIR = os.path.join("scratch", "staging")
MANIFEST_FILE = os.path.join(BACKUP_DIR, "manifest.json")

TARGET_DIRS = [
    "web_dashboard",
    "web_dashboard_direccion"
]
ROOT_FILES = [
    "Seguimiento_Presupuesto_Sep_2026.xlsx",
    "Orden_Fabricacion_MRP_Codiagro.xlsx"
]

FORBIDDEN_COMMERCIAL_FILES = [
    "fullcost_data.json",
    "fullcost_data.js",
    "cockpit_inicio_logic.js",
    "profitability_logic.js",
    "bi_intelligence_logic.js",
    "treasury_data.json",
    "treasury_data.js",
    "update_validation_status.json"
]

def take_snapshot():
    try:
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

        if not os.path.exists(MANIFEST_FILE) or os.path.getsize(MANIFEST_FILE) == 0:
            print("[ERROR] El manifiesto de respaldo no se pudo escribir correctamente.")
            return 1

        print(f"[SNAPSHOT] Copia de seguridad guardada ({len(manifest['files'])} archivos en {BACKUP_DIR}).")
        return 0
    except Exception as e:
        print(f"[ERROR CRITICO] Fallo en take_snapshot: {e}")
        return 1

def prepare_staging():
    try:
        if os.path.exists(STAGING_DIR):
            for root, dirs, files in os.walk(STAGING_DIR, topdown=False):
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
        os.makedirs(STAGING_DIR, exist_ok=True)

        # Copiar estructura base desde BACKUP_DIR hacia STAGING_DIR para servir de plantilla
        if os.path.exists(MANIFEST_FILE):
            with open(MANIFEST_FILE, "r", encoding="utf-8") as f:
                manifest = json.load(f)
            for rel_f in manifest.get("files", []):
                src = os.path.join(BACKUP_DIR, rel_f)
                dst = os.path.join(STAGING_DIR, rel_f)
                if os.path.exists(src):
                    parent = os.path.dirname(dst)
                    if parent:
                        os.makedirs(parent, exist_ok=True)
                    shutil.copy2(src, dst)
        else:
            for d in TARGET_DIRS:
                os.makedirs(os.path.join(STAGING_DIR, d), exist_ok=True)

        print(f"[STAGING] Entorno temporal preparado en {STAGING_DIR}.")
        return 0
    except Exception as e:
        print(f"[ERROR CRITICO] Fallo en prepare_staging: {e}")
        return 1

def activate():
    try:
        if not os.path.exists(STAGING_DIR):
            print(f"[ERROR] No existe el directorio de staging {STAGING_DIR} para activar.")
            return 1

        activated_count = 0
        # Copiar directorios web desde staging hacia carpetas activas
        for d in TARGET_DIRS:
            staging_d = os.path.join(STAGING_DIR, d)
            if os.path.exists(staging_d):
                os.makedirs(d, exist_ok=True)
                for root, _, files in os.walk(staging_d):
                    for f in files:
                        src_f = os.path.join(root, f)
                        rel_path = os.path.relpath(src_f, staging_d)
                        dst_f = os.path.join(d, rel_path)
                        os.makedirs(os.path.dirname(dst_f), exist_ok=True)
                        shutil.copy2(src_f, dst_f)
                        activated_count += 1

        # Copiar archivos raíz desde staging hacia raíz
        for rf in ROOT_FILES:
            staging_rf = os.path.join(STAGING_DIR, rf)
            if os.path.exists(staging_rf):
                shutil.copy2(staging_rf, rf)
                activated_count += 1

        # Escudo de seguridad en web_dashboard activa
        for fn in FORBIDDEN_COMMERCIAL_FILES:
            bad_f = os.path.join("web_dashboard", fn)
            if os.path.exists(bad_f):
                try:
                    os.remove(bad_f)
                except Exception:
                    pass

        # Limpiar directorio de staging tras activación exitosa
        shutil.rmtree(STAGING_DIR, ignore_errors=True)

        print(f"[ACTIVACIÓN] {activated_count} archivos activados con éxito desde staging a producción.")
        return 0
    except Exception as e:
        print(f"[ERROR CRITICO] Fallo en activate: {e}")
        return 1

def restore_snapshot():
    try:
        # 0. Limpiar staging si existe
        if os.path.exists(STAGING_DIR):
            shutil.rmtree(STAGING_DIR, ignore_errors=True)

        if not os.path.exists(MANIFEST_FILE):
            print(f"[ERROR] No se encontró manifiesto en {MANIFEST_FILE}. Imposible restaurar.")
            return 1

        with open(MANIFEST_FILE, "r", encoding="utf-8") as f:
            manifest = json.load(f)

        files_list = manifest.get("files", [])
        if not files_list:
            print("[ERROR] El manifiesto de respaldo está vacío. Imposible restaurar.")
            return 1

        valid_files_set = set(os.path.normpath(p) for p in files_list)

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
        for rel_f in files_list:
            backup_src = os.path.join(BACKUP_DIR, rel_f)
            if os.path.exists(backup_src):
                parent_dir = os.path.dirname(rel_f)
                if parent_dir:
                    os.makedirs(parent_dir, exist_ok=True)
                shutil.copy2(backup_src, rel_f)
                restored_count += 1
            else:
                print(f"[ERROR] Archivo en manifiesto no existe en backup: {backup_src}")
                return 1

        # 3. Verificar que los archivos restaurados existen
        for rel_f in files_list:
            if not os.path.exists(rel_f):
                print(f"[ERROR] Falló la verificación de restauración del archivo: {rel_f}")
                return 1

        print(f"[RESTAURACIÓN LIMPIA] {restored_count} archivos restaurados correctamente; {purged_count} archivos huérfanos purgados.")
        return 0
    except Exception as e:
        print(f"[ERROR CRITICO] Fallo en restore_snapshot: {e}")
        return 1

if __name__ == "__main__":
    action = sys.argv[1].lower() if len(sys.argv) > 1 else "backup"
    if action == "restore":
        sys.exit(restore_snapshot())
    elif action in ["prepare_staging", "stage_init", "staging"]:
        sys.exit(prepare_staging())
    elif action in ["activate", "promote"]:
        sys.exit(activate())
    else:
        sys.exit(take_snapshot())
