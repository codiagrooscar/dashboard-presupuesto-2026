# -*- coding: utf-8 -*-
"""
Script: apply_option_4_budget_reduction.py
Aplica la Opción 4 (Mix Equilibrado y Quirúrgico) para ajustar el presupuesto
de ventas 2027 de 17.738.160 € a exactamente 17.180.000 € (-558.160 €).

Ajustes:
1. Egipto (García): Dotra Chemicals y Diplomats (-270.000 € / Factor: 0.93454999)
2. Turquía (Mehmet): Çitar Çiçek y Doga Tarim (-140.000 € / Factor: 0.79457102)
3. Chile (Ricardo): Bayer, S.A. (-120.000 € / Factor: 0.93831393)
4. Grecia (García): Emphyton / Poultsidis (-28.160,29 € / Factor: 0.91897368)
"""

import openpyxl
import shutil
import os
import subprocess

print("=== 1. CREANDO COPIAS DE SEGURIDAD ===")
backup_def = 'Presupuesto_Ventas_2027_Definitivo_BACKUP_PRE_17180K.xlsx'
shutil.copy2('Presupuesto_Ventas_2027_Definitivo.xlsx', backup_def)
print(f"Copia de seguridad creada: {backup_def}")

if os.path.exists('Presupuesto_Ventas_2027.xlsx'):
    backup_std = 'Presupuesto_Ventas_2027_BACKUP_PRE_17180K.xlsx'
    shutil.copy2('Presupuesto_Ventas_2027.xlsx', backup_std)
    print(f"Copia de seguridad creada: {backup_std}")

f_egipto = 0.93454999
f_turquia = 0.79457102
f_chile = 0.93831393
f_grecia = 0.91897368

def process_budget_file(filepath):
    print(f"\n=== Procesando archivo: {filepath} ===")
    wb = openpyxl.load_workbook(filepath, data_only=False)
    
    # ----------------------------------------------------
    # A. Actualizar Previsión Matriz Horizontal
    # ----------------------------------------------------
    ws_h = wb['Previsión Matriz Horizontal']
    col_u_indices = [36, 38, 40, 42, 44, 46, 48, 50, 52, 54, 56, 58]
    col_e_indices = [37, 39, 41, 43, 45, 47, 49, 51, 53, 55, 57, 59]
    
    h_modified_rows = 0
    h_modified_cells = 0
    
    for r in range(2, ws_h.max_row + 1):
        com = str(ws_h.cell(r, 1).value or '').strip()
        if not com or 'TOTAL' in com.upper():
            continue
        pais = str(ws_h.cell(r, 2).value or '').strip().upper()
        cli = str(ws_h.cell(r, 3).value or '').strip().upper()
        pr27 = float(ws_h.cell(r, 9).value or 0.0)
        
        factor = 1.0
        if ('DOTRA CHEMICALS' in cli or 'DIPLOMATS' in cli) and pais == 'EGIPTO':
            factor = f_egipto
        elif ('ÇITAR' in cli or 'CITAR' in cli or 'DOGA TARIM' in cli) and pais == 'TURQUÍA':
            factor = f_turquia
        elif 'BAYER' in cli and pais == 'CHILE':
            factor = f_chile
        elif ('EMPHYTON' in cli or 'POULTSIDIS' in cli) and pais == 'GRECIA':
            factor = f_grecia

        if factor != 1.0:
            h_modified_rows += 1
            for col_u, col_e in zip(col_u_indices, col_e_indices):
                val_u = ws_h.cell(r, col_u).value
                try:
                    val_u = float(val_u or 0.0)
                except:
                    val_u = 0.0
                    
                if val_u > 0:
                    new_u = int(round(val_u * factor))
                    new_e = round(new_u * pr27, 2)
                    ws_h.cell(r, col_u, new_u)
                    ws_h.cell(r, col_e, new_e)
                    h_modified_cells += 1

    print(f"  Filas modificadas en Previsión Matriz Horizontal: {h_modified_rows}")
    print(f"  Celdas mensuales modificadas: {h_modified_cells}")

    # ----------------------------------------------------
    # B. Actualizar Detalle Líneas Clasificadas
    # ----------------------------------------------------
    if 'Detalle Líneas Clasificadas' in wb.sheetnames:
        ws_d = wb['Detalle Líneas Clasificadas']
        d_modified_rows = 0
        for r in range(2, ws_d.max_row + 1):
            yr = ws_d.cell(r, 1).value
            if yr != 2027:
                continue
            pais = str(ws_d.cell(r, 4).value or '').strip().upper()
            cli = str(ws_d.cell(r, 5).value or '').strip().upper()
            u = float(ws_d.cell(r, 15).value or 0.0)
            
            factor = 1.0
            if ('DOTRA CHEMICALS' in cli or 'DIPLOMATS' in cli) and pais == 'EGIPTO':
                factor = f_egipto
            elif ('ÇITAR' in cli or 'CITAR' in cli or 'DOGA TARIM' in cli) and pais == 'TURQUÍA':
                factor = f_turquia
            elif 'BAYER' in cli and pais == 'CHILE':
                factor = f_chile
            elif ('EMPHYTON' in cli or 'POULTSIDIS' in cli) and pais == 'GRECIA':
                factor = f_grecia
                
            if factor != 1.0 and u > 0:
                new_u = int(round(u * factor))
                ws_d.cell(r, 15, new_u)
                d_modified_rows += 1

        print(f"  Filas modificadas en Detalle Líneas Clasificadas: {d_modified_rows}")

    # ----------------------------------------------------
    # C. Actualizar Resumen Ejecutivo Previsión
    # ----------------------------------------------------
    if 'Resumen Ejecutivo Previsión' in wb.sheetnames:
        ws_re = wb['Resumen Ejecutivo Previsión']
        # Actualizar celda E5 (Presupuesto 2027 Total) y E6 (Unidades y Precio medio)
        ws_re.cell(5, 5, '17,180,000.00 €')
        ws_re.cell(6, 5, '6,317,709 uds (2,72 €/ud)')
        print("  Encabezados de Resumen Ejecutivo Previsión actualizados.")

    wb.save(filepath)
    print(f"  Guardado exitoso de {filepath}!")

process_budget_file('Presupuesto_Ventas_2027_Definitivo.xlsx')
if os.path.exists('Presupuesto_Ventas_2027.xlsx'):
    process_budget_file('Presupuesto_Ventas_2027.xlsx')

print("\n=== AJUSTE PRESUPUESTARIO COMPLETADO EXITOSAMENTE ===")
