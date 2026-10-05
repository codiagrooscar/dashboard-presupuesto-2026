import openpyxl, shutil, os

src_updated = 'Gastos/2027-Budget-Codiagro P&L_ACTUALIZADO_4.5PCT.xlsx'
target = 'Gastos/2027-Budget-Codiagro P&L.xlsx'

if not os.path.exists(src_updated):
    print('No existe el archivo actualizado:', src_updated)
    exit(1)

try:
    shutil.copy2(src_updated, target)
    print('¡Actualizado con éxito', target, 'con la versión de inflación 4,5%!')
except Exception as e:
    print('No se pudo sobrescribir (posiblemente aún abierto en Excel):', e)
