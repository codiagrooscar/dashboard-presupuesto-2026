# -*- coding: utf-8 -*-
import openpyxl
import json
import os
import datetime
import calendar
import unicodedata
from collections import defaultdict

def normalize(text):
    if not text:
        return ""
    text = unicodedata.normalize('NFKD', text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    for char in [",", ".", "-", "'", "´", "`", "\"", "/"]:
        text = text.replace(char, " ")
    return " ".join(text.upper().split())

# Catálogo Oficial de Condiciones y Formas de Pago aportado por la dirección financiera
TERMS_TABLE = [
    ('430000001', 'Nacional', 'JAVIER', 'COMERCIAL TÉCNICA FITOCARTHAGO, S.L.', 'GIRO A 120 DÍAS F.F.', 120, 124.50),
    ('430000003', 'Nacional', '', 'AGROQUIMICAS DRAGO S.L.', 'TRANSFERENCIA BANCARIA F.F.', 0, 3.50),
    ('430000004', 'Nacional', 'GARCIA / IREN', 'COOP.AGRIC. SAN ISIDRO S.C.A.', 'CONFIRMING A 60 DÍAS F.F.', 60, 70.03),
    ('430000005', 'C. Valenciana', 'JAVIER', 'ASESORAMIENTO TECNICO DE CULTIVOS, S.L.', 'GIRO A 120 DÍAS F.F.', 120, 124.00),
    ('430000006', 'Nacional', 'JAVIER', 'FITOMURCIA S.L.', 'GIRO A 120 DÍAS F.F.', 120, 124.79),
    ('430000008', 'Export', 'GARCIA', 'DOTRA CHEMICALS', 'TRANSF. BANCARIA 30 D. BL (MARGEN SEGURIDAD 60 D)', 30, 60.00),
    ('430000009', 'Export', 'GARCÍA', 'DIPLOMATS ORCHARDS FOR TRADE AGENCIES', 'TRANSFERENCIA BANCARIA 30 D. BL', 30, 28.52),
    ('430000011', 'Export', 'RICARDO', 'ECUAQUIMICA', 'TRANSF. BANCARIA A 120 D. F. ARRIBO', 150, 164.80),
    ('430000015', 'Nacional', 'JAVIER', 'CABRERA MORALES FRANCISCO', 'TRANSF. BANCARIA A 90 DÍAS F.F.', 90, 104.03),
    ('430000018', 'C. Valenciana', '', 'TRANSPORTES CAMPILLO, S.A.', 'TRANSFERENCIA BANCARIA F.F.', 0, None),
    ('430000019', 'Nacional', 'GARCIA / IREN', 'CAMPOEJIDO S.C.A.', 'TRANSF. BANC. ENTRE EL 15 Y 20 FF', 20, 14.94),
    ('430000021', 'C. Valenciana', '', 'FITOXENT S.C.L.V.', 'TRANSFERENCIA BANCARIA F.F.', 0, 8.98),
    ('430000023', 'Nacional', 'ALFONSO', 'MAJ AGROQUIMICOS S.L.', 'GIRO A 90 DÍAS F.F.', 90, 94.80),
    ('430000024', 'Nacional', 'JAVIER', 'FITOSANITARIOS DRAGO S.L.U', 'TRANSFERENCIA BANCARIA F.F.', 0, None),
    ('430000032', 'Nacional', '', 'CAMP-ALT S.L.', 'GIRO A 90 DÍAS F.F.', 90, 95.50),
    ('430000033', 'Nacional', 'JAVIER', 'CODIMED EXPORT, S.L.', 'TRANSF. BANCARIA A 120 DÍAS F.F.', 120, 160.69),
    ('430000037', 'Export', 'RICARDO', 'EUROFERTIL, S.A.', 'TRANSF. BANCARIA A 120 D. F. ARRIBO', 120, 80.44),
    ('430000058', 'Nacional', 'ALFONSO', 'AGROTÉCNICA DEL NORTE, S.L', 'GIRO A 60 DÍAS F.F.', 60, 63.82),
    ('430000064', 'Europa', '', 'AIFAR S.P.A', 'TRANSF. BANCARIA A 60 DÍAS F.F.', 60, None),
    ('430000065', 'Nacional', 'ALFONSO', 'AGRICOLA DEL JUCAR DE CUENCA, S.L.', 'GIRO A 90 DÍAS F.F.', 90, 94.33),
    ('430000067', 'Nacional', 'ALFONSO', 'PLANTALDIA SLU', 'GIRO A 90 DÍAS F.F.', 90, None),
    ('430000071', 'Nacional', 'ALFONSO', 'CITAL, S.L.', 'TRANSF. BANCARIA A 60 DÍAS F.F.', 60, 62.56),
    ('430000075', 'Nacional', 'ALFONSO', 'AGROVEYCA,S.L', 'GIRO A LA VISTA', 60, 50.32),
    ('430000080', 'C. Valenciana', '', 'SAFIAGRO, S.A.', 'PAGARÉ A 60 DÍAS F.F.', 60, 87.80),
    ('430000089', 'Nacional', 'ALFONSO', 'T.J AGRICULTURA S.L.', 'TRANSF. BANCARIA A 60 DÍAS F.F.', 60, 73.89),
    ('430000095', 'Nacional', 'ALFONSO', 'HORTICOLAS Y PATATAS HNOS. MORAN, S.A.T.', 'GIRO A 85 DÍAS', 85, 89.00),
    ('430000102', 'Nacional', 'ALFONSO', 'BELLO ALIAGA ERNESTO', 'GIRO A 75 DÍAS F.F.', 75, 81.00),
    ('430000105', 'Nacional', 'GARCIA', 'TRATAMIENTOS AGRICOLAS BRENES, S.L.', 'GIRO A 85 DÍAS', 85, 88.57),
    ('430000134', 'Nacional', '', 'HORTOFRUTICOLA COSTA DE ALMERÍA, S.L.', 'GIRO A 60 DÍAS F.F.', 60, None),
    ('430000143', 'Nacional', '', 'INIESTA DEL PINO DISTRIBUCIONES S.L.U', 'TRANSFERENCIA BANCARIA F.F.', 0, None),
    ('430000182', 'Nacional', 'ALFONSO', 'BARRIOS LÓPEZ, JUSTO', 'GIRO A 60 DÍAS F.F.', 60, None),
    ('430000190', 'Nacional', 'GARCIA / IREN', 'SUCA, S.C.A.', 'PAGARÉ A 60 DÍAS F.F.', 60, 61.45),
    ('430000205', 'C. Valenciana', '', 'AGROCAMP CANET S.L.', 'CONFIRMING A 60 DÍAS F.F.', 60, 52.00),
    ('430000219', 'Nacional', '', 'S.A.T. N. 9989 PEREGRIN', 'TRANSFERENCIA/CONFIRMING 60 DÍAS', 60, None),
    ('430000237', 'Nacional', 'GARCIA', 'AGRUPACION DE LABRADORES DEL POZUELO, SA', 'GIRO A 60 DÍAS F.F.', 60, None),
    ('430000243', 'Europa', 'MEHMED', 'DOGA TARIM TIC.SAN.VE TUR. LTD. STI', 'TRANSF. BANCARIA A 120 DÍAS F.F.', 120, 117.60),
    ('430000261', 'Nacional', '', 'UNICOM SERVICIOS Y GESTION S.L.', 'GIRO A 60 DÍAS F.F.', 60, 62.30),
    ('430000265', 'Nacional', 'GARCIA', 'AGROTRAPICHE, S.L.U', 'GIRO A 60 DÍAS F.F.', 60, None),
    ('430000267', 'C. Valenciana', '', 'RECICLADOS Y DERRIBOS LLORENS, S.L.', 'TRANSFERENCIA BANCARIA A 30 DÍAS F.F.', 30, None),
    ('430000290', 'Export', 'GARCÍA', 'EURL BOUAZZA FILAHA', 'REMESA DOCUMENTARIA A 59 D. F. B/L', 60, 84.00),
    ('430000292', 'Nacional', 'GARCIA', 'AGROSER BERNARDINO, S.L.', 'GIRO A 60 DÍAS F.F.', 60, 79.73),
    ('430000295', 'Nacional', 'GARCIA / IREN', 'GUIVARTO AGRICOLA, S.L.', 'PAGARE A 90 DÍAS F.F.', 90, 89.97),
    ('430000307', 'Nacional', '', 'FRUITURE ADVISORS, S.L.', 'GIRO A 120 DÍAS F.F.', 120, 124.00),
    ('430000311', 'C. Valenciana', '', 'INDUSTRIAS CEVIMA VILAFAMES, S.L.', 'TRANSF. BANCARIA A 30 DÍAS F.F.', 30, None),
    ('430000316', 'Nacional', 'ALFONSO', 'CARDONA Y CELMA, S.L.', 'GIRO A 60 DÍAS F.F.', 60, None),
    ('430000326', 'Nacional', '', 'FINCA DOÑA ANA C.B.', 'TRANSF. BANCARIA A 60 DÍAS F.F.', 60, 82.00),
    ('430000332', 'Nacional', 'GARCIA', 'SERVIAGRI´ 97, S.L.', 'GIRO A 90 DÍAS F.F.', 90, 114.03),
    ('430000334', 'Nacional', 'GARCIA / IREN', 'HORTOCCAMPO, S.A.', 'GIRO A 60 DÍAS F.F.', 60, None),
    ('430000335', 'Nacional', '', 'AGROGIMEDEL, S.L.U.', 'GIRO A 60 DÍAS F.F.', 60, None),
    ('430000344', 'Export', 'RICARDO', 'CAMPO ABIERTO, S.A.S.', 'TRANSFERENCIA BANCARIA F.F.', 120, 135.50),
    ('430000345', 'Export', 'GARCÍA', 'GEORGIOS POULTSIDIS', 'TRANSF. BANC. CONF. 90, 120 y 150 FF', 120, 143.50),
    ('430000380', 'Export', 'RICARDO', 'LEKKERBIO B.V.', 'TRANSFERENCIA BANCARIA F.F.', 0, None),
    ('430000382', 'Export', 'GARCÍA', 'STE BBMAGRI, SARL', '30% ANTI 20% LLEGA PEDIDO 50% 90 D', 50, 81.36),
    ('430000383', 'Export', 'RICARDO', 'BAYER, S.A.', 'TRANSF. BANCARIA A 60 DÍAS F.F.', 65, 91.22),
    ('430000384', 'Nacional', 'GARCIA', 'AGRONEW FITO, S.L.', 'TRANSF. BANCARIA A 90 DÍAS F.F.', 90, 107.77),
    ('430000385', 'C. Valenciana', '', 'MUNDOHUERTO-CARMEN SOLAZ BLANCO', 'GIRO A 60 DÍAS F.F.', 60, 63.56),
    ('430000386', 'Nacional', 'ALFONSO', 'VILLALTA SATIVUM, S.L.', 'GIRO A 120 DÍAS F.F.', 120, 103.71),
    ('430000389', 'Nacional', 'ALFONSO', 'ASOC. HORTICULTORES DE VILLA DEL PRADO', 'GIRO A 60 DÍAS F.F.', 60, 64.61),
    ('430000398', 'Export', 'MEHMED', 'PLANTA SANA, D.O.O.', 'TRANSFERENCIA BANCARIA F.F.', 0, None),
    ('430000399', 'Export', 'GARCÍA', 'EMPHYTON SERVICE P.C.', 'TRANSF. BANC. CONF. 90, 120 y 150 FF', 120, 171.29),
    ('430000405', 'Nacional', 'ALFONSO', 'AGRITEC MEDINA, S.L.U.', 'GIRO A 60 DÍAS F.F.', 60, 63.78),
    ('430000407', 'Export', 'RICARDO', 'BCOMERCE E.I.R.L.', 'TRANSFERENCIA BANCARIA A 150 DÍAS', 150, None),
    ('430000414', 'Nacional', 'GARCIA', 'BOSCO ABASCAL, S.L.', 'GIRO A 60 DÍAS F.F.', 60, None),
    ('430000417', 'Export', 'GARCÍA', 'KUWAIT FARM COMPANY', 'TRANSFERENCIA BANCARIA F.F.', 0, None),
    ('430000419', 'Nacional', '', 'AGROREBOLLO, S.L.', 'GIRO A 60 DÍAS F.F.', 60, 64.33),
    ('430000421', 'Nacional', 'ALFONSO', 'ACTIV. PRODUC. Y COMERC. LA VEGUILLA, SL', 'PAGARÉ A 60 DÍAS F.F.', 60, None),
    ('430000424', 'Europa', 'MEHMED', 'ÇITAR ÇIÇEK TARIM TIC. LTD. STI', 'CRÉDITO DOCUMENTARIO A 75 DIAS', 75, 91.67),
    ('430000425', 'Export', 'GARCÍA', 'DOTRA GULF TRADING COMPANY', 'TRANSFERENCIA BANCARIA 150 D. BL', 150, 148.00),
    ('430000427', 'Nacional', '', 'SUSTAINABLE AGRO SOLUTIONS SAU', 'TRANSFERENCIA BANCARIA F.F.', 60, 87.00),
    ('430000028', 'Europa', 'GARCÍA', 'SC MARCOSER S.R.L.', 'CONTADO', 0, None),
    ('430000097', 'Nacional', '', 'AGRÍCOLA DE GOZÓN S.L', 'TRANSF. BANCARIA A 30 DÍAS F.F.', 30, None),
    ('430000269', 'Nacional', 'GARCIA / IREN', 'SERVICIOS AGRÍCOLAS CAMPO RIO, S.L.', 'GIRO A 90 DÍAS F.F.', 90, 96.00),
    ('430000350', 'Nacional', 'ALFONSO', 'BARAHONA HERNAEZ, DAVID', 'TRANSF. BANCARIA A 60 DÍAS F.F.', 60, 61.60),
    ('430000381', 'Nacional', 'ALFONSO', 'BARDENAS COMERCIAL, S.L.', 'GIRO A 60 DÍAS F.F.', 60, 63.58),
    ('430000387', 'Nacional', 'GARCIA', 'EL PINAR BERRIES, S.L.', 'TRANSF. BANCARIA A 60 DÍAS F.F.', 60, 60.25),
    ('430000388', 'Nacional', 'ALFONSO', 'MARTINEZ S.C.', 'GIRO A 60 DÍAS F.F.', 60, 63.00),
    ('430000393', 'Nacional', '', 'COMERCIAL NANCLARES, S.L.', 'GIRO A 60 DÍAS F.F.', 60, 63.20),
    ('430000394', 'Nacional', 'ALFONSO', 'COMERCIAL NORTE AGROCAMPO, S.L.', 'GIRO A 60 DÍAS F.F.', 60, 65.50),
    ('430000396', 'Export', 'MEHMED', 'ZTB AQRO MMC', 'TRANSF. BANC. 50% PEDIDO 50% CARG', 0, None),
    ('430000430', 'Nacional', '', 'ALMENDRALIA IBÉRICA, S.L.U.', 'CONTADO', 0, None),
    ('430000431', 'Nacional', '', 'COMERCIAL DEL ENVASE RECICLADO, SL', 'TRANSF. BANCARIA A 90 DÍAS F.F.', 90, None),
]

# Ventas Reales de Septiembre 2026 proporcionadas oficialmente
SEPT_SALES_REAL = [
    ('430000001', 'COMERCIAL TÉCNICA FITOCARTHAGO, S.L.', 5240, 14014),
    ('430000003', 'AGROQUIMICAS DRAGO S.L.', 8400, 21544),
    ('430000004', 'COOP.AGRIC. SAN ISIDRO S.C.A.', 0, 0),
    ('430000005', 'ASESORAMIENTO TECNICO DE CULTIVOS, S.L.', 1920, 6125),
    ('430000006', 'FITOMURCIA S.L.', 1400, 6607),
    ('430000008', 'DOTRA CHEMICALS', 21000, 81270),
    ('430000009', 'DIPLOMATS ORCHARDS FOR TRADE AGENCIES', -6, -37),
    ('430000011', 'ECUAQUIMICA', 8800, 14916),
    ('430000015', 'CABRERA MORALES FRANCISCO', 6855, 14348),
    ('430000018', 'TRANSPORTES CAMPILLO, S.A.', 80, 216),
    ('430000019', 'CAMPOEJIDO S.C.A.', 12464, 52104),
    ('430000021', 'FITOXENT S.C.L.V.', 0, 0),
    ('430000023', 'MAJ AGROQUIMICOS S.L.', 60, 108),
    ('430000024', 'FITOSANITARIOS DRAGO S.L.U', 14784, 40585),
    ('430000028', 'SC MARCOSER S.R.L.', 7932, 43438),
    ('430000033', 'CODIMED EXPORT, S.L.', 1460, 7774),
    ('430000037', 'EUROFERTIL, S.A.', 20050, 58212),
    ('430000058', 'AGROTÉCNICA DEL NORTE, S.L', 1200, 4852),
    ('430000064', 'AIFAR S.P.A', 3000, 12000),
    ('430000065', 'AGRICOLA DEL JUCAR DE CUENCA, S.L.', 300, 1255),
    ('430000075', 'AGROVEYCA,S.L', 720, 2702),
    ('430000080', 'SAFIAGRO, S.A.', 640, 1837),
    ('430000089', 'T.J AGRICULTURA S.L.', 1100, 3131),
    ('430000105', 'TRATAMIENTOS AGRICOLAS BRENES, S.L.', 1100, 2035),
    ('430000143', 'INIESTA DEL PINO DISTRIBUCIONES S.L.U', 0, 0),
    ('430000190', 'SUCA, S.C.A.', 640, 3323),
    ('430000192', 'SERVICIOS AGROPECUARIOS EL LLANO, S. A.', 12872, 48446),
    ('430000205', 'AGROCAMP CANET S.L.', 0, 0),
    ('430000237', 'AGRUPACION DE LABRADORES DEL POZUELO, SA', 600, 3112),
    ('430000243', 'DOGA TARIM TIC.SAN.VE TUR. LTD. STI', 0, 0),
    ('430000269', 'SERVICIOS AGRÍCOLAS CAMPO RIO, S.L.', 0, 0),
    ('430000292', 'AGROSER BERNARDINO, S.L.', 1100, 1903),
    ('430000295', 'GUIVARTO AGRICOLA, S.L.', 10370, 32009),
    ('430000326', 'FINCA DOÑA ANA C.B.', 1120, 4213),
    ('430000332', 'SERVIAGRI´ 97, S.L.', 3716, 13311),
    ('430000335', 'AGROGIMEDEL, S.L.U.', 0, 0),
    ('430000344', 'CAMPO ABIERTO, S.A.S.', 0, 0),
    ('430000350', 'BARAHONA HERNAEZ, DAVID', 0, 0),
    ('430000372', 'SUMINISTROS AGRÍCOLAS SALAMPACK, S.L.', 0, 0),
    ('430000379', 'ALCADA EXPORT, S.L.', 0, 0),
    ('430000380', 'LEKKERBIO B.V.', 0, 0),
    ('430000381', 'BARDENAS COMERCIAL, S.L.', 600, 2070),
    ('430000382', 'STE BBMAGRI, SARL', -2600, -1212),
    ('430000384', 'AGRONEW FITO, S.L.', 3480, 8176),
    ('430000385', 'MUNDOHUERTO-CARMEN SOLAZ BLANCO', 164, 575),
    ('430000386', 'VILLALTA SATIVUM, S.L.', 3460, 6193),
    ('430000387', 'EL PINAR BERRIES, S.L.', 1100, 2409),
    ('430000389', 'ASOC. HORTICULTORES DE VILLA DEL PRADO', 210, 1013),
    ('430000394', 'COMERCIAL NORTE AGROCAMPO, S.L.', 220, 682),
    ('430000396', 'ZTB AQRO MMC', 15, 0),
    ('430000405', 'AGRITEC MEDINA, S.L.U.', 0, 0),
    ('430000414', 'BOSCO ABASCAL, S.L.', 0, 0),
    ('430000417', 'KUWAIT FARM COMPANY', 0, 0),
    ('430000418', 'AUDITORIA Y CERTIFICACION EN GESTION DE', 0, 0),
    ('430000419', 'AGROREBOLLO, S.L.', 0, 0),
    ('430000420', 'SALES AMPOSTA, JAUME', 0, 0),
    ('430000421', 'ACTIV. PRODUC. Y COMERC. LA VEGUILLA, SL', 2560, 6771),
    ('430000425', 'DOTRA GULF TRADING COMPANY', 20284, 88497),
    ('430000429', 'FITOSANITARIOS CARCAIXENT, S.L.', 8700, 26197),
    ('430000430', 'ALMENDRALIA IBÉRICA, S.L.U.', 1100, 1958)
]

month_map = {
    'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'ago': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dic': 12,
    '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '6': 6,
    '7': 7, '8': 8, '9': 9, '10': 10, '11': 11, '12': 12
}

def get_fac_date(mes_str, is_nac):
    m = month_map.get(str(mes_str).lower(), 9)
    y = 2026
    if is_nac:
        # Clientes Nacionales: se factura a fin de mes y el plazo empieza a contar desde ahí
        last_day = calendar.monthrange(y, m)[1]
        return datetime.date(y, m, last_day)
    else:
        # Clientes Internacionales / Europa: se factura a fecha de operación (15 de mes)
        return datetime.date(y, m, 15)

def build_treasury_dataset():
    src_file = 'temp_cf_diplomats.xlsx'
    if not os.path.exists(src_file):
        src_file = 'temp_cashflow_septiembre.xlsx'
    if not os.path.exists(src_file):
        src_file = r'c:\Users\oscar.ocampo\OneDrive - Agroquímica Codiagro S.L\Escritorio\REPORTING SIMPLIFICADO\SEPTIEMBRE\Cashflow\Cashflow previsión cierre 2026.xlsx'

    print(f"Loading cash flow from: {src_file}")
    wb = openpyxl.load_workbook(src_file, data_only=True)
    ws_cf = wb['CF']
    ws_c = wb['Clients']

    # Clientes con problemas de pago / riesgo de impago excluidos de la previsión de tesorería
    EXCLUDED_CLIENTS_RISK = ['SUSTAINABLE', 'CODIMED', 'EUROFERTIL']

    def is_client_excluded(name_str):
        n = normalize(str(name_str))
        return any(exc in n for exc in EXCLUDED_CLIENTS_RISK)

    # 1. Construir Diccionario Maestro de Condiciones y Retrasos de Clientes
    terms_by_code = {}
    terms_by_name = {}

    for code, tip, com, cli, fp, dias_pact, dias_real in TERMS_TABLE:
        # Si tiene días reales observados en el histórico: se respetan
        # Si no tiene días observados (-): se aplica retraso mínimo de 5 días (salvo pago al contado estricto)
        if dias_real is not None:
            dias_efectivos = int(round(dias_real))
        else:
            dias_efectivos = int(dias_pact) + (5 if int(dias_pact) > 0 else 0)

        if tip in ['Nacional', 'C. Valenciana']:
            canal = 'Cobros por canal Clientes nacionales'
            is_nacional = True
            pais_def = 'ES'
        elif tip == 'Europa':
            canal = 'Cobros por canal U.E.'
            is_nacional = False
            pais_def = 'UE'
        else:
            canal = 'Cobros por canal Clientes internac.'
            is_nacional = False
            pais_def = 'EXT'

        retraso = dias_efectivos - int(dias_pact)
        excluido = is_client_excluded(cli)

        info = {
            'codigo': code,
            'tipo': tip,
            'comercial': com,
            'cliente': cli,
            'forma_pago': fp,
            'dias_pactados': int(dias_pact),
            'dias_reales_obs': dias_real,
            'dias_efectivos': dias_efectivos,
            'retraso_dias': retraso,
            'canal': canal,
            'is_nacional': is_nacional,
            'pais': pais_def,
            'total_facturacion': 0.0,
            'total_cobros': 0.0,
            'excluido_prevision': excluido,
            'motivo_exclusion': 'Excluido por problemas de pago / riesgo de impago' if excluido else ''
        }
        terms_by_code[str(code)] = info
        terms_by_name[normalize(cli)] = info

    # 2. Extraer líneas de vencimientos aplicando las nuevas reglas:
    #    A. Cartera Real Pendiente Oficial por Cliente (Ageing Codiagro Sep.xlsx - Desglose)
    #    B. Incorporar SÓLO las ventas previstas desde el 1 de octubre cuyo vencimiento caiga en oct, nov o dic
    #    C. Excluir Sustainable Agro Solutions, Codimed y Eurofertil
    client_vencimientos = []

    # A. Cartera Real Pendiente (Ageing Desglose)
    ageing_file = r'c:\Users\oscar.ocampo\OneDrive - Agroquímica Codiagro S.L\Escritorio\REPORTING SIMPLIFICADO\SEPTIEMBRE\Cobros\Ageing Codiagro Sep.xlsx'
    if not os.path.exists(ageing_file):
        ageing_file = 'Ageing Codiagro Sep.xlsx'

    if os.path.exists(ageing_file):
        print(f"Loading real pending balances from: {ageing_file}")
        wb_ageing = openpyxl.load_workbook(ageing_file, data_only=True)
        ws_desg = wb_ageing['Desglose']
        for r in range(5, ws_desg.max_row + 1):
            cod = ws_desg.cell(r, 2).value
            cli = str(ws_desg.cell(r, 3).value or '').strip()
            fra = ws_desg.cell(r, 4).value
            ffra = ws_desg.cell(r, 5).value
            imp = ws_desg.cell(r, 8).value

            if not cli or imp is None or is_client_excluded(cli):
                continue

            val_imp = float(imp)
            term = terms_by_code.get(str(cod)) or terms_by_name.get(normalize(cli))
            if term:
                dias = term['dias_efectivos']
                dias_pact = term['dias_pactados']
                dias_obs = term['dias_reales_obs']
                canal_str = term['canal']
                com_str = term['comercial']
                pais_str = term['pais']
                cod_str = term['codigo']
            else:
                dias = 65
                dias_pact = 60
                dias_obs = None
                canal_str = 'Cobros por canal Clientes nacionales'
                com_str = ''
                pais_str = 'ES'
                cod_str = str(cod or '')

            if isinstance(ffra, datetime.datetime):
                dt_fra = ffra.date()
            elif isinstance(ffra, datetime.date):
                dt_fra = ffra
            else:
                dt_fra = datetime.date(2026, 9, 30)

            vto_calc = dt_fra + datetime.timedelta(days=dias)
            # Facturas con vencimiento teórico anterior a cierre sep: saldo pendiente exigible en Octubre
            if vto_calc <= datetime.date(2026, 9, 30):
                mes_vto = '10-2026'
                vto_efectiva = datetime.date(2026, 10, 15)
            else:
                mes_vto = f"{vto_calc.month}-{vto_calc.year}"
                vto_efectiva = vto_calc

            client_vencimientos.append({
                'codigo_cliente': cod_str,
                'cliente': cli,
                'comercial': com_str,
                'canal': canal_str,
                'pais': pais_str,
                'mes_venta': dt_fra.strftime('%b').lower(),
                'fecha_factura': dt_fra.strftime('%d/%m/%Y'),
                'fecha_operacion': dt_fra.strftime('%d/%m/%Y'),
                'num_factura': str(fra or ''),
                'facturacion': val_imp,
                'iva': 0.0,
                'importe': val_imp,
                'dias_pactados': dias_pact,
                'dias_reales_obs': dias_obs,
                'dias_efectivos': dias,
                'dias': dias,
                'vto_fecha': vto_efectiva.strftime('%d/%m/%Y'),
                'mes_vto': mes_vto,
                'tipo_origen': 'Cartera Real Pendiente'
            })

            if cod_str in terms_by_code:
                terms_by_code[cod_str]['total_facturacion'] += val_imp
                terms_by_code[cod_str]['total_cobros'] += val_imp

    # B. Incorporar SÓLO las ventas previstas desde el 1 de octubre cuyo vencimiento caiga en oct, nov o dic
    m_map_q4 = {'oct': 10, 'nov': 11, 'dic': 12, '10': 10, '11': 11, '12': 12}
    for r in range(3, ws_c.max_row + 1):
        tipo = ws_c.cell(r, 1).value
        cli = str(ws_c.cell(r, 2).value or '').strip()
        canal = ws_c.cell(r, 3).value
        pais = ws_c.cell(r, 4).value
        mes_v = str(ws_c.cell(r, 5).value or '').strip().lower()
        fact = ws_c.cell(r, 7).value
        iva = ws_c.cell(r, 8).value
        imp = ws_c.cell(r, 9).value

        # Condición estricta: sólo previsiones (BG) a partir del 1 de octubre (oct, nov, dic)
        if tipo != 'BG' or mes_v not in m_map_q4 or not cli or imp is None or is_client_excluded(cli):
            continue

        m_num = m_map_q4[mes_v]
        term = terms_by_name.get(normalize(cli))
        if term:
            is_nac = term['is_nacional']
            dias = term['dias_efectivos']
            dias_pact = term['dias_pactados']
            canal_str = term['canal']
            comercial_str = term['comercial']
            pais_str = term['pais']
            dias_obs = term['dias_reales_obs']
            cod_str = term['codigo']
        else:
            is_nac = 'nacional' in str(canal or '').lower()
            dias_orig = ws_c.cell(r, 10).value
            dias_pact = int(dias_orig) if dias_orig is not None else 60
            dias = dias_pact + (5 if dias_pact > 0 else 0)
            canal_str = str(canal) if canal else 'Cobros por canal Clientes nacionales'
            comercial_str = ''
            pais_str = str(pais or ('ES' if is_nac else 'EXT'))
            dias_obs = None
            cod_str = ''

        if is_nac:
            last_day = calendar.monthrange(2026, m_num)[1]
            fac_date = datetime.date(2026, m_num, last_day)
        else:
            fac_date = datetime.date(2026, m_num, 15)

        vto = fac_date + datetime.timedelta(days=dias)
        mes_vto = f"{vto.month}-{vto.year}"

        # REGLA: Sólo incorporar si el vencimiento cae en oct, nov o dic
        if mes_vto not in ['10-2026', '11-2026', '12-2026']:
            continue

        tot_val = float(imp)
        fact_val = float(fact) if fact is not None else tot_val
        iva_val = float(iva) if iva is not None else 0.0

        client_vencimientos.append({
            'codigo_cliente': cod_str,
            'cliente': cli,
            'comercial': comercial_str,
            'canal': canal_str,
            'pais': pais_str,
            'mes_venta': mes_v,
            'fecha_factura': fac_date.strftime('%d/%m/%Y'),
            'fecha_operacion': fac_date.strftime('%d/%m/%Y'),
            'num_factura': f"Prev. {mes_v.upper()}",
            'facturacion': fact_val,
            'iva': iva_val,
            'importe': tot_val,
            'dias_pactados': dias_pact,
            'dias_reales_obs': dias_obs,
            'dias_efectivos': dias,
            'dias': dias,
            'vto_fecha': vto.strftime('%d/%m/%Y'),
            'mes_vto': mes_vto,
            'tipo_origen': 'Venta Prevista Q4'
        })

        if cod_str in terms_by_code:
            terms_by_code[cod_str]['total_facturacion'] += fact_val
            terms_by_code[cod_str]['total_cobros'] += tot_val

    # 3. Sumar Cobros Proyectados de Q4 por Canal (meses 10-2026, 11-2026, 12-2026)
    q4_cobros_proyectados = {
        '2026-10': defaultdict(float),
        '2026-11': defaultdict(float),
        '2026-12': defaultdict(float)
    }

    for cv in client_vencimientos:
        mv = cv['mes_vto']
        if mv in ['10-2026', '11-2026', '12-2026']:
            q4_cobros_proyectados[f"2026-{mv.split('-')[0].zfill(2)}"][cv['canal']] += cv['importe']

    q4_cobros_summary = {
        m: dict(q4_cobros_proyectados[m]) for m in q4_cobros_proyectados
    }

    print("\nNUEVOS COBROS PROYECTADOS Q4 (CON CONDICIONES REALES Y RETRASOS):")
    for m, cdict in sorted(q4_cobros_summary.items()):
        print(f"  {m}: Total = {sum(cdict.values()):,.2f} € | {cdict}")

    # 4. Construir Filas de Cash Flow con Proyecciones de Cobros Actualizadas
    month_keys = [f"2026-{m:02d}" for m in range(1, 13)]
    month_labels = ["Ene 26", "Feb 26", "Mar 26", "Abr 26", "May 26", "Jun 26", "Jul 26", "Ago 26", "Sep 26", "Oct 26", "Nov 26", "Dic 26"]

    cf_rows = []
    for r in range(6, 55):
        category = ws_cf.cell(r, 3).value
        label = ws_cf.cell(r, 6).value
        if not label:
            continue
        
        lbl_str = str(label).strip()
        cat_str = str(category).strip() if category else ''

        budget_anual = ws_cf.cell(r, 19).value
        try:
            budget_anual_val = float(budget_anual) if budget_anual is not None else 0.0
        except Exception:
            budget_anual_val = 0.0

        # Ene - Sep Reales (cols 21 a 29 en hoja CF)
        monthly_values = {}
        for m_idx in range(1, 10):
            m_key = f"2026-{m_idx:02d}"
            v = ws_cf.cell(r, 20 + m_idx).value
            try:
                monthly_values[m_key] = float(v) if v is not None else 0.0
            except Exception:
                monthly_values[m_key] = 0.0

        # Oct - Dic (Proyectado con las proyecciones calculadas si es cobro)
        budget_q4 = {}
        for idx_q, m_idx in enumerate([10, 11, 12]):
            m_key = f"2026-{m_idx:02d}"
            v_b = ws_cf.cell(r, 30 + idx_q).value
            if v_b is None:
                v_b = ws_cf.cell(r, 16 + idx_q).value
            try:
                val_q4 = float(v_b) if v_b is not None else 0.0
            except Exception:
                val_q4 = 0.0

            # Sobrescribir con cobros proyectados si coincide el canal
            if lbl_str == 'Cobros por canal Clientes nacionales':
                val_q4 = q4_cobros_summary[m_key].get('Cobros por canal Clientes nacionales', 0.0)
            elif lbl_str == 'Cobros por canal Clientes internac.':
                val_q4 = q4_cobros_summary[m_key].get('Cobros por canal Clientes internac.', 0.0)
            elif lbl_str == 'Cobros por canal U.E.':
                val_q4 = q4_cobros_summary[m_key].get('Cobros por canal U.E.', 0.0)
            
            budget_q4[m_key] = val_q4

        cf_rows.append({
            'row_idx': r,
            'category': cat_str,
            'label': lbl_str,
            'budget_annual': budget_anual_val,
            'reales': monthly_values,
            'budget_q4': budget_q4
        })

    # Recalcular totales de Cobros, Cash Flow y Saldo en cf_rows
    row_map = {row['label']: row for row in cf_rows}

    if 'Total Cobros' in row_map:
        for m_key in ['2026-10', '2026-11', '2026-12']:
            tot_c = (
                row_map.get('Cobros por canal Clientes nacionales', {}).get('budget_q4', {}).get(m_key, 0.0) +
                row_map.get('Cobros por canal Clientes internac.', {}).get('budget_q4', {}).get(m_key, 0.0) +
                row_map.get('Cobros por canal U.E.', {}).get('budget_q4', {}).get(m_key, 0.0)
            )
            row_map['Total Cobros']['budget_q4'][m_key] = tot_c
            if 'Total Cobros con ajuste' in row_map:
                row_map['Total Cobros con ajuste']['budget_q4'][m_key] = tot_c

    salidas_q4 = {
        '2026-10': -1165000.0,
        '2026-11': -340000.0,
        '2026-12': -975000.0
    }

    saldo_inicial_octubre = 3203215.24
    current_saldo = saldo_inicial_octubre

    if 'Cash Flow del periodo' in row_map:
        for m_key in ['2026-10', '2026-11', '2026-12']:
            tot_cobros = row_map['Total Cobros']['budget_q4'][m_key]
            cf_periodo = tot_cobros + salidas_q4[m_key]
            row_map['Cash Flow del periodo']['budget_q4'][m_key] = cf_periodo

    if 'Tesorería inicial' in row_map:
        row_map['Tesorería inicial']['budget_q4']['2026-10'] = saldo_inicial_octubre

    for m_idx, m_key in enumerate(['2026-10', '2026-11', '2026-12']):
        cf_per = row_map['Cash Flow del periodo']['budget_q4'][m_key] if 'Cash Flow del periodo' in row_map else 0.0
        final_saldo = current_saldo + cf_per
        if 'Tesorería final' in row_map:
            row_map['Tesorería final']['budget_q4'][m_key] = final_saldo
        next_m_key = f"2026-{11 + m_idx:02d}"
        if next_m_key in ['2026-11', '2026-12'] and 'Tesorería inicial' in row_map:
            row_map['Tesorería inicial']['budget_q4'][next_m_key] = final_saldo
        current_saldo = final_saldo

    client_master_list = list(terms_by_code.values())

    output_data = {
        'metadata': {
            'generated_at': datetime.datetime.now().isoformat(),
            'source_file': 'Cashflow previsión cierre 2026.xlsx',
            'sociedad': 'CODIAGRO S.L.',
            'moneda': 'EUR',
            'corte_real': '2026-09',
            'saldo_cierre_sep_real': saldo_inicial_octubre,
            'months': [
                {'key': k, 'label': lbl, 'is_real': idx < 9, 'idx': idx + 1}
                for idx, (k, lbl) in enumerate(zip(month_keys, month_labels))
            ]
        },
        'cf_rows': cf_rows,
        'q4_cobros_proyectados': q4_cobros_summary,
        'client_master': client_master_list,
        'client_vencimientos': client_vencimientos
    }

    out_json = 'web_dashboard/treasury_data.json'
    out_js = 'web_dashboard/treasury_data.js'

    with open(out_json, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)

    with open(out_js, 'w', encoding='utf-8') as f:
        f.write("window.TREASURY_DATA = " + json.dumps(output_data, ensure_ascii=False) + ";\n")

    print(f"\nGenerated {out_json} and {out_js} successfully!")
    return output_data

if __name__ == '__main__':
    build_treasury_dataset()
