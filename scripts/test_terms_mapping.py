# -*- coding: utf-8 -*-
import openpyxl, datetime, unicodedata
from collections import defaultdict

def normalize(text):
    if not text:
        return ""
    text = unicodedata.normalize('NFKD', text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    for char in [",", ".", "-", "'", "´", "`", "\"", "/"]:
        text = text.replace(char, " ")
    return " ".join(text.upper().split())

# Transcribed payment terms table from user's image
TERMS_TABLE = [
    ('430000001', 'Nacional', 'JAVIER', 'COMERCIAL TÉCNICA FITOCARTHAGO, S.L.', 'GIRO A 120 DÍAS F.F.', 120, 124.50),
    ('430000003', 'Nacional', '', 'AGROQUIMICAS DRAGO S.L.', 'TRANSFERENCIA BANCARIA F.F.', 0, 3.50),
    ('430000004', 'Nacional', 'GARCIA / IREN', 'COOP.AGRIC. SAN ISIDRO S.C.A.', 'CONFIRMING A 60 DÍAS F.F.', 60, 70.03),
    ('430000005', 'C. Valenciana', 'JAVIER', 'ASESORAMIENTO TECNICO DE CULTIVOS, S.L.', 'GIRO A 120 DÍAS F.F.', 120, 124.00),
    ('430000006', 'Nacional', 'JAVIER', 'FITOMURCIA S.L.', 'GIRO A 120 DÍAS F.F.', 120, 124.79),
    ('430000008', 'Export', 'GARCIA', 'DOTRA CHEMICALS', 'TRANSFERENCIA BANCARIA 30 D. BL', 30, 26.51),
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

# Client mapping
terms_by_code = {}
terms_by_name = {}

for code, tip, com, cli, fp, dias_pact, dias_real in TERMS_TABLE:
    # Determinación de días efectivos según la premisa del usuario:
    # "la mayoría se retrasan un mínimo de 5 días. Ajusta los cobros con esta premisa."
    if dias_real is not None:
        # Si tiene días reales empíricos, los respetamos (o mínimo pactado + 5 si supera)
        dias_efectivos = int(round(dias_real))
    else:
        # Si no tiene días reales observados, se aplica el retraso mínimo de 5 días
        dias_efectivos = int(dias_pact) + (5 if int(dias_pact) > 0 else 0)

    # Canal
    if tip in ['Nacional', 'C. Valenciana']:
        canal = 'Cobros por canal Clientes nacionales'
        is_nacional = True
    elif tip == 'Europa':
        canal = 'Cobros por canal U.E.'
        is_nacional = False
    else:
        canal = 'Cobros por canal Clientes internac.'
        is_nacional = False

    info = {
        'codigo': code,
        'tipo': tip,
        'comercial': com,
        'cliente': cli,
        'forma_pago': fp,
        'dias_pactados': int(dias_pact),
        'dias_reales_obs': dias_real,
        'dias_efectivos': dias_efectivos,
        'canal': canal,
        'is_nacional': is_nacional
    }
    terms_by_code[code] = info
    terms_by_name[normalize(cli)] = info

print(f"Total clientes mapeados: {len(terms_by_code)}")
