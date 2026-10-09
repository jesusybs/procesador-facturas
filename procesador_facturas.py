#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
Motor Determinista para Procesamiento de Facturas y Listas de Empaque
Con Auto-Aprendizaje y Diccionario Inteligente de Repuestos
Ubicación Permanente: C:\Facturas_Repuestos\procesador_facturas.py
"""

import sys
import os
import re
import json
from decimal import Decimal, ROUND_HALF_UP
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Configuración de consola Windows para soporte UTF-8 completo
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='backslashreplace')

# Ruta del diccionario persistente
DICT_PATH = os.path.join(os.path.dirname(__file__), "diccionario_repuestos.json")

# Variables de caché en memoria para el diccionario
_DICTIONARY_CACHE = None

def load_dictionary():
    """Carga el diccionario desde el archivo JSON persistente."""
    global _DICTIONARY_CACHE
    if _DICTIONARY_CACHE is not None:
        return _DICTIONARY_CACHE

    if os.path.exists(DICT_PATH):
        try:
            with open(DICT_PATH, "r", encoding="utf-8") as f:
                _DICTIONARY_CACHE = json.load(f)
                return _DICTIONARY_CACHE
        except Exception as e:
            print(f"[AVISO] No se pudo leer {DICT_PATH}: {e}", file=sys.stderr)

    # Fallback por defecto si no existe
    _DICTIONARY_CACHE = {"exact_mappings": {}, "learned_mappings": {}}
    return _DICTIONARY_CACHE

_DICTIONARY_DIRTY = False

def flush_dictionary():
    """Escribe los cambios del diccionario a disco si hubo modificaciones."""
    global _DICTIONARY_CACHE, _DICTIONARY_DIRTY
    if not _DICTIONARY_DIRTY or _DICTIONARY_CACHE is None:
        return
    try:
        with open(DICT_PATH, "w", encoding="utf-8") as f:
            json.dump(_DICTIONARY_CACHE, f, ensure_ascii=False, indent=2)
        _DICTIONARY_DIRTY = False
    except Exception as e:
        print(f"[ERROR] Al guardar diccionario en disco: {e}", file=sys.stderr)

def save_learned_mapping(raw_key, standardized_value, flush=True):
    """Guarda un nuevo repuesto aprendido en el archivo JSON (con opción de buffer diferido)."""
    global _DICTIONARY_CACHE, _DICTIONARY_DIRTY
    d = load_dictionary()
    if "learned_mappings" not in d:
        d["learned_mappings"] = {}

    if d["learned_mappings"].get(raw_key) == standardized_value:
        return

    d["learned_mappings"][raw_key] = standardized_value
    _DICTIONARY_DIRTY = True

    if flush:
        flush_dictionary()
    else:
        # Registro en memoria sin I/O bloqueante
        pass

# Patrones y reglas automotrices generales para estandarización instantánea
GENERAL_AUTO_PATTERNS = [
    (r'\bTOPE\s+Y\s+CUBRE\s*POLVO\b', 'TOPE Y CUBREPOLVO AMORTIGUADOR'),
    (r'\bAMORT(?:IGUADOR)?[\.\s]+DEL\b', 'AMORTIGUADOR DELANTERO'),
    (r'\bAMORT(?:IGUADOR)?[\.\s]+(?:TRAS|TRS)\b', 'AMORTIGUADOR TRASERO'),
    (r'\b(?:KIT|JG)[\.\s]+PIN\s+(?:C[\/\s]|Y\s+)?BUJES\b', 'KIT PIN BUJES'),
    (r'\bHUB\s+DEL\b', 'HUB DELANTERO'),
    (r'\bHUB\s+TRAS\b', 'HUB TRASERO'),
    (r'\bHUB\s+BAL\b', 'HUB CON BALINERA'),
    (r'\b(YOYOS\s+)?CINTA\s+(?:DE\s+)?TIMON\b', 'CINTA DE TIMON'),
    (r'\b(?:TERM(?:INAL)?[\.\s]+(?:DE\s+)?CREM(?:AYERA|ALLERA)?|T[\/\.\s]+CREM(?:AYERA|ALLERA)?)\b', 'TERMINAL DE CREMALLERA'),
    (r'\bTERM(?:INAL)?[\.\s]+(?:DE\s+)?ESTAB(?:ILIZADOR)?(?:[\.\s]+DEL|[\.\s]+TRAS)?\b', 'TERMINAL ESTABILIZADOR'),
    (r'\bTERM(?:INAL)?\.?\s*(?:EXT|EX)\b', 'TERMINAL EXTERIOR'),
    (r'\bTERM(?:INAL)?\.?\s*CREM\b', 'TERMINAL DE CREMALLERA'),
    (r'\bTERM(?:INAL)?\.?\s*INT\b', 'TERMINAL INTERIOR'),
    (r'\bTIJERA\s+TRAS\b', 'MESETA/TIJERA TRASERA'),
    (r'\bTIJERA\b', 'MESETA/TIJERA'),
    (r'\b(?:BRZO|BRAZO)[\.\s]+INF(?:ERIOR)?\b', 'MESETA/TIJERA'),
    (r'\b(?:BOMBA?[\.\s]+(?:DE\s+)?GAS(?:OLINA)?|B[\.\/\s]+GAS(?:OLINA)?)\b', 'BOMBA DE GASOLINA'),
    (r'\b(?:BOMBA?[\.\s]+ACEITE|B[\.\/\s]+ACEITE)\b', 'BOMBA DE ACEITE'),
    (r'\bBOMBA[\.\s]+HIDRAUL(?:ICA)?\b', 'BOMBA HIDRAULICA'),
    (r'\bBOMBA[\.\s]+AGUA\b', 'BOMBA DE AGUA'),
    (r'\bB\/AGUA\b', 'BOMBA DE AGUA'),
    (r'\bBOMBA[\.\s]+FRENO\b', 'BOMBA FRENO'),
    (r'\b(?:TAPA|CARCASA\s+TAPA)[\.\s]+RAD(?:IADOR)?\b', 'TAPA DE RADIADOR'),
    (r'\bTOMA[\.\s]+(?:DE\s+)?RAD(?:IADOR)?\b', 'TOMA DE RADIADOR'),
    (r'\b(?:CARCAZA|CARCASA)[\.\s]+(?:DE\s+)?TERMOSTATO\b', 'CARCAZA TERMOSTATO'),
    (r'\bTOMA[\.\s]+(?:DE\s+)?TERMOSTATO\b', 'TOMA TERMOSTATO'),
    (r'\bCUELLO[\.\s]+(?:DE\s+)?RELLENO\s+COOLANT\b', 'CUELLO RELLENO COOLANT'),
    (r'\bVALVULA\s+DE\s+AGUA\b', 'VALVULA DE AGUA'),
    (r'\bRAD(?:IADOR)?[\.\s]', 'RADIADOR DE AGUA'),
    (r'\bCREM(?:ALLERA)?[\.,\s]', 'CREMALLERA DE DIRECCION'),
    (r'\bCREMALLERA\b', 'CREMALLERA DE DIRECCION'),
    (r'\bCUERPO[\.\s]+(?:DE\s+)?ACEL(?:ERACION)?\b', 'CUERPO DE ACELERACION'),
    (r'\bIN[YJ]ECTOR(?:ES)?\b', 'INYECTOR'),
    (r'\bPUNTA[\.\s]+FLECHA\b', 'PUNTA FLECHA'),
    (r'\bDISCO[\.\s]+FRENO\b', 'DISCO DE FRENO'),
    (r'\bDISC(?:O)?[\.\s]+(?:DE\s+)?(?:EMBRAG(?:UE)?|CLUTCH)\b', 'DISCO DE EMBRAGUE'),
    (r'\bDISCO\b', 'DISCO DE FRENO'),
    (r'\bTAMBOR\s+(?:DE\s+)?FRENO\b', 'TAMBOR DE FRENO'),
    (r'\bCRUCETA\b', 'CRUCETA'),
    (r'\bESPIRAL\b', 'ESPIRAL / RESORTE'),
    (r'\bBASE\s+(?:AMORT|MOTOR)?\b', 'BASE/SOPORTE MOTOR'),
    (r'\bBOLA\b', 'ROTULA DE SUSPENSION'),
    (r'\bCALIPER\b', 'CALIPER DE FRENO'),
    (r'\bMU[Ñ\uFFFD]EQUILLA\b', 'TERMINAL ESTABILIZADOR'),
    (r'\bBAL[\.\s]+AMORT\b', 'BALINERA/RODAMIENTO AMORTIGUADOR DELANTERO'),
    (r'\bBAL(?:INERA)?[\/\.\s]+(?:RDA[\/\.\s]+)?DEL\b', 'BALINERA DELANTERA'),
    (r'\bBAL(?:INERA)?[\/\.\s]+(?:RDA[\/\.\s]+)?TRAS\b', 'BALINERA TRASERA'),
    (r'\bBAL[\/\.\s]+RDA\b', 'BALINERA'),
    (r'\bBAL(?:INERA|INERAS)?\b', 'BALINERA'),
    (r'\bBEARING\b', 'BALINERA'),
    (r'\bBOTA\s+FLECHA\b', 'GUARDAPOLVO FLECHA'),
    (r'\bBUJE\s+BARRA\b', 'BUJE DE BARRA ESTABILIZADORA'),
    (r'\bBUJE\b', 'BUJE'),
    (r'\bCIL(?:INDRO)?[\.\s]+FRENO\b', 'CILINDRO DE FRENO'),
    (r'\bFREE\s+WHEEL\b', 'CUBO MANUAL / FREE WHEEL'),
    (r'\bLINK\b', 'TERMINAL ESTABILIZADOR'),
    (r'\b(?:TUBO|TUBERIA|MANGUERA)\s+(?:MET[\.\s]+)?RET(?:ORNO)?[\.\s]+AGUA\b', 'TUBO DE RETORNO DE AGUA'),
    (r'\b(?:TUBO|TUBERIA|MANGUERA)\s+(?:MET(?:AL)?[\.\s]+)?AGUA\b', 'TUBO DE AGUA'),
    (r'\bTUBO\s+(?:DE\s+)?CALEFACCION\b', 'TUBO DE CALEFACCION'),
    (r'\b(?:RETEN|RETENES|RETENEDOR|RETENEDORA|RETENEDORES|RETENEDORAS|OIL\s+SEAL|SEAL\s+OIL)\b', 'RETENEDOR'),
    (r'(?<!MET\.)\bRET[\.\/](?!AGUA|ACEIT)', 'RETENEDOR'),
    (r'\b(?:EMP(?:ACADURA)?|EMPAQUE|JUNTA)[\.\s]+(?:DE\s+)?TAPA[\.\s]+(?:DE\s+)?VALV(?:ULA)?\b', 'EMPAQUE TAPA VALVULA'),
    (r'\bTAPA[\.\s]+(?:DE\s+)?VALV(?:ULA)?\b', 'TAPA DE VALVULA'),
    (r'\bEMP(?:ACADURA)?\.?\s*COMPLETO\b', 'EMPACADURA COMPLETA (JUEGO)'),
    (r'\bEMP(?:AQUE)?[\.\s]+OVER(?:HAUL)?\b', 'EMPACADURA COMPLETA (JUEGO)'),
    (r'\bEMP(?:ACADURA)?[\.\s]+(?:CULATA|CABEZOTE)\b', 'EMPACADURA DE CULATA'),
    (r'\b(?:JGO[\.\s]+)?CADENA[\.\/\s]+(?:DE\s+)?TIEMPO\b', 'CADENA DE TIEMPO'),
    (r'\bCADENA\s+(?:TPO|TIEMPO)\b', 'CADENA DE TIEMPO'),
    (r'\bCULATA(?:\s+MOTOR)?\b', 'CULATA DE MOTOR'),
    (r'\bTAPON\s+(?:DE\s+)?(?:CARTER|RAD|RADIADOR)\b', 'TAPON'),
    (r'\bSELLO(?:S)?\s+(?:DE\s+)?VALV(?:ULA)?\b', 'SELLOS DE VALVULA'),
    (r'\bSELLO(?:S)?\s+(?:DE\s+)?BUJIA(?:S)?\b', 'SELLOS DE BUJIA'),
    (r'\bTACO\s+(?:DE\s+)?FRENO\b', 'PASTILLA/TACO DE FRENO'),
    (r'\bASPAS?\b', 'ASPA DE VENTILADOR'),
    (r'\bASPA\s+VENTILADORA?\b', 'ASPA DE VENTILADOR'),
    (r'\b(?:BOMBILLO|FOCO|HALOGENO|F\/HALOGENO)\b', 'BOMBILLO / FOCO'),
    (r'\bSOCKET\b', 'SOCKET / CONECTOR'),
    (r'\bPOLEA[\.\s]+(?:AJST|AJUSTE|TENS|TENSOR|TENSORA)[\.\s]*CORREA\b', 'POLEA TENSORA'),
    (r'\bPOLEA\s+(?:TENS|TENSOR|TENSORA)\b', 'POLEA TENSORA'),
    (r'\bPOLEA[\.\s]+(?:DE\s+)?CIG(?:UE[Ñ\uFFFD]AL)?\b', 'POLEA CIGUEÑAL'),
    (r'\bPOLEA\b', 'POLEA'),
    (r'\b(?:TENSOR|TENSIONER)\b', 'TENSOR'),
    (r'\bPLATO\s+(?:DE\s+)?(?:CLUTCH|EMBRAG(?:UE)?)\b', 'PLATO DE EMBRAGUE'),
    (r'\b(?:CIL(?:INDRO)?[\.\s]+ESCLAVO|REP[\.\s]+ESCLAVO|CIL[\.\s]+AUXL[\.\s]+EMBRAG)\b', 'BOMBA AUXILIAR DE EMBRAGUE'),
    (r'\bCIL(?:INDRO)?[\.\s]+(?:DE\s+)?EMBRA[GQ](?:UE)?\b', 'CILINDRO DE EMBRAGUE'),
    (r'\bCORREA\s+(?:TPO|TIEMPO|T[\.\s]|T\b)', 'CORREA DE TIEMPO'),
    (r'\bCORREA\s+(?:ALT|ALTERNADOR)\b', 'CORREA DE ALTERNADOR'),
    (r'\bCORREA\s+(?:A\/C|AC|CLIMA)\b', 'CORREA DE AIRE ACONDICIONADO'),
    (r'\bCORREA\s+MULTICANAL\b', 'CORREA MULTICANAL'),
    (r'\bCORREA\b', 'CORREA'),
    (r'\bVALV(?:ULA)?[\.\s]+(?:DE\s+)?PURGA\b', 'VALVULA DE PURGA'),
    (r'\bVALV(?:ULA)?[\.\s]+ADM(?:ISION)?\b', 'VALVULA DE ADMISION'),
    (r'\bVALV(?:ULA)?[\.\s]+ESC(?:APE)?\b', 'VALVULA DE ESCAPE'),
    (r'\bVALV(?:ULA)?[\.\s]+PCV\b', 'VALVULA PCV'),
    (r'\bCASQ(?:UILLO)?[\.\s]+BIELA\b', 'CASQUILLO DE BIELA'),
    (r'\bCASQ(?:UILLO)?[\.\s]+(?:BANC(?:ADA)?|CIG(?:UE[Ñ\uFFFD]AL)?)\b', 'CASQUILLO DE BANCADA'),
    (r'\b(?:JG[\.\s]*)?A?NIO?LLOS\b', 'ANILLOS'),
    (r'\bPISTON\b', 'PISTON'),
    (r'\bFILT(?:RO)?[\.\s]+AIRE\b', 'FILTRO DE AIRE'),
    (r'\bFILT(?:RO)?[\.\s]+(?:DE\s+)?GAS(?:OLINA)?\b', 'FILTRO DE GASOLINA'),
    (r'\bBOTA[\.\s]+CREMALL(?:ERA)?\b', 'GUARDAPOLVO DE CREMALLERA'),
    (r'\bGUIA[\.\s]+.*?\bVALV(?:ULA)?\b', 'GUIA DE VALVULA'),
    (r'\b[VB]ARILLA[\.\s]+(?:MEDIDOR(?:A)?[\.\s]+)?ACE[IT]', 'VARILLA DE ACEITE'),
    (r'\bP[\/\.\s]+EJE[\.\s]+INT\b', 'PUNTA DE EJE INTERIOR'),
    (r'\bCUBO[\.\s]+(?:RDA[\.\s]+)?DEL\b', 'HUB DELANTERO'),
    (r'\b(?:MOTOR[\.\s]+ARRANQ(?:UE)?|ARRANQUE)\b', 'MOTOR DE ARRANQUE'),
    (r'\b(?:AUT[\.\s]+ARR|AUTOMATICO[\.\s]+(?:DE\s+)?ARRANQUE)\b', 'AUTOMATICO DE ARRANQUE'),
    (r'\bBENDIX(?:\s+ARRANQUE)?\b', 'BENDIX DE ARRANQUE'),
    (r'\bCARBONERA(?:\s+ARR)?\b', 'CARBONERA'),
    (r'\bALTERNADOR\b', 'ALTERNADOR'),
    (r'\b(?:PUENTE[\.\s]+)?SOPORTE[\.\s]+CARDAN\b', 'SOPORTE DE CARDAN'),
    (r'\bCARBURADOR\b', 'CARBURADOR'),
    (r'\bCAUCHO(?:S)?[\.\s]+(?:DE\s+)?FRENO\b', 'CAUCHO DE FRENO'),
    (r'\bREFORZADOR[\.\s]+(?:DE\s+)?FRENO\b', 'REFORZADOR DE FRENO'),
    (r'\bREFORZADOR[\.\s]+(?:DE\s+)?EMBRA[GQ](?:UE)?\b', 'REFORZADOR DE EMBRAGUE'),
    (r'\bFAN[\.\s]+CLUTCH\b', 'FAN CLUTCH'),
    (r'\bHORQUILLA[\.\s]+(?:DE\s+)?EMBRA(?:GUE)?\b', 'HORQUILLA DE EMBRAGUE'),
    (r'\bCABLE(?:S)?[\.\s]+(?:DE\s+)?BUJIA(?:S)?\b', 'CABLES DE BUJIA'),
    (r'\bBUJIA(?:S)?\b', 'BUJIA'),
    (r'\bCAMISA[\.\s]+(?:DE\s+)?(?:MOTOR|CILINDRO)?\b', 'CAMISA DE MOTOR'),
    (r'\bCABALLITOS?\b', 'CABALLITOS / BALANCINES'),
    (r'\bARO[\.\s]+SINCRONIZ(?:ADOR)?\b', 'ARO SINCRONIZADOR'),
    (r'\bBULBO[\.\s]+(?:DE\s+)?TEMP(?:ERATURA)?\b', 'SENSOR DE TEMPERATURA'),
    (r'\bTAPA[\.\s]+(?:DE\s+)?DIST(?:RIB(?:UIDOR)?)?\b', 'TAPA DE DISTRIBUIDOR'),
    (r'\bROTOR[\.\s]+(?:DE\s+)?DIST(?:RIB(?:UIDOR)?)?\b', 'ROTOR DE DISTRIBUIDOR'),
    (r'\bCONDENSADOR\b', 'CONDENSADOR'),
    (r'\b(?:JG[\.\s]+)?AXIALES\b', 'AXIALES'),
    (r'\b(?:JG[\.\s]+)?ZAPATAS?\b', 'ZAPATAS DE FRENO'),
    (r'\b(?:SWITCH|SENSOR)[\.\s]+(?:DE\s+)?(?:REVERSA|RETROCESO)\b', 'SENSOR DE RETROCESO'),
    (r'\bSENSOR[\.\s]+(?:DE\s+)?VELOCIDAD\b', 'SENSOR DE VELOCIDAD'),
    (r'\bSENSOR[\.\s]+(?:DE\s+)?(?:POSIC(?:ION)?[\.\s]+)?EJE[\.\s]+(?:DE\s+)?LEVA[S]?\b', 'SENSOR DE EJE DE LEVAS'),
    (r'\bSENSOR[\.\s]+(?:DE\s+)?(?:POSIC(?:ION)?[\.\s]+)?CIG(?:UE[Ñ\uFFFD]AL)?\b', 'SENSOR DE CIGUEÑAL'),
    (r'\bCIG(?:UE[Ñ\uFFFD]AL)?\b', 'CIGUEÑAL'),
    (r'\bDEPOSITO[\.\s]+LI[QD][\.\s]+HIDRAUL(?:ICO)?\b', 'DEPOSITO LIQUIDO HIDRAULICO'),
    (r'\bTUBERIA[\.\s]+.*EMBRAG(?:UE)?\b', 'TUBERIA DE EMBRAGUE'),
    (r'\bBIELA[\.\s]+(?:DE\s+)?MOTOR\b', 'BIELA DE MOTOR'),
    (r'\bMOTOR[\.\s]+TANQ(?:UE)?[\.\s]+WIPER\b', 'BOMBA DE LIMPIAPARABRISAS'),
    (r'\bTERMOSWITCH\b', 'TERMOSWITCH')
]

def auto_standardize_unknown_part(raw_desc):
    """
    Motor Heurístico de Estandarización de Repuestos Nuevos:
    Elimina códigos de motor, medidas, modelos, años y referencias OEM,
    extrayendo exclusivamente la categoría o nombre genérico base.
    """
    d = raw_desc.upper().strip()

    # 1. Eliminar códigos entre paréntesis o corchetes
    d = re.sub(r'\(.*?\)', '', d)
    d = re.sub(r'\[.*?\]', '', d)

    # 2. Separar términos con puntos comunes (ej. BOMBA.GAS -> BOMBA GAS)
    d = re.sub(r'\b([A-Z]+)\.([A-Z]+)\b', r'\1 \2', d)

    # 3. Eliminar dimensiones y especificaciones técnicas
    d = re.sub(r'\b\d+(\.\d+)?\s*(MM|CM|LBS|KG|KPA|T|V|CYL|L)\b', '', d)
    d = re.sub(r'\b\d+\s*[xX]\s*\d+(\s*[xX]\s*\d+)?\b', '', d)
    d = re.sub(r'\bØ\d+(\.\d+)?\s*(MM)?\b', '', d)
    d = re.sub(r'\b(4X4|4WD|AWD|2WD|16V|8V|24V|V6|V8|TURBO|DIESEL)\b', '', d)

    # 4. Eliminar patrones de años
    d = re.sub(r'\b(19|20)?\d{2}\s*[-~/]\s*(19|20)?\d{2}\b', '', d)
    d = re.sub(r'\b(19|20)\d{2}\s*[-~]?\b', '', d)
    d = re.sub(r'\b\d{2}\s*[-~]\b', '', d)
    d = re.sub(r'\b(19|20)\d{2}\b', '', d)
    d = re.sub(r'\b\d{2}/\d{2}(/\d{2})?\b', '', d)
    d = re.sub(r'\b\d{2}/~', '', d)
    d = re.sub(r'\b\d{2}-\d{2}\b', '', d)
    d = re.sub(r'\b\d{4}/\d{4}\b', '', d)

    # 5. Eliminar marcas de vehículos
    d = re.sub(r'\b(TOYOTA|TOY|NISSAN|NIS|MITSUBISHI|MIT|HONDA|HON|SUZUKI|SUZ|HYUNDAI|HYU|CHEVROLET|CHEV|ISUZU|ISU|MAZDA|FORD|KIA|DAIHATSU|DAI|UNIVERSAL|JAC|RENAULT|REN|HINO|JEEP|VOLKSWAGEN|VW)\b\.?', '', d)

    # 6. Eliminar modelos de vehículos comunes
    d = re.sub(r'\b(COROLLA|HILUX|YARIS|TERCEL|CAMRY|PRADO|4RUNNER|TACOMA|FORTUNER|HIACE|COASTER|RAV-4|RAV4|VITARA|GRAND VITARA|CELERIO|SWIFT|ALTO|ERTIGA|BALENO|CIAZ|SX4|CIVIC|CR-V|CRV|ACCENT|ELANTRA|TUCSON|PICANTO|RIO|SPORTAGE|L200|MONTERO|NATIVA|LANCER|FRONTIER|PATHFINDER|TIIDA|ALTIMA|SENTRA|ESCAPE|RANGER|BT-50|AVEO|NEW SAIL|APV|KAVAK|MERU|STOUT|DUTRO|KIZASHI|RIDGELINE|SONET|QASHQAI|XTRAIL|IPSUM|AVENSIS|PROBOX|TRAX|URBAN)\b', '', d)

    # 7. Eliminar códigos de motor (ej. 1KD, 2KD, 4AFE, 1ZZ, 2TR, 5VZ)
    d = re.sub(r'\b[0-9][A-Z]{1,3}(-?[A-Z0-9]+)?\b', '', d)

    # 8. Eliminar orientaciones redundantes
    d = re.sub(r'\b(RH/LH|LH/RH|AMBOS LADOS)\b', '', d)
    d = re.sub(r'\b(RH|LH)\b', '', d)

    # 9. Eliminar términos de relleno
    d = re.sub(r'\b(AT|MT|CERAMICA|SEMI-CERAMICA|ORIGINAL|GENUINO|OEM|REFORZADO)\b', '', d)

    # 10. Limpiar puntuación sobrante
    d = re.sub(r'[/\\#.,\-_]+', ' ', d)
    cleaned = " ".join(d.split()).strip()

    # Normalizaciones comunes de prefijos abreviados
    if cleaned.startswith("FILTRO ACEITE"):
        return "FILTRO DE ACEITE"
    if cleaned.startswith("FILTRO AIRE"):
        return "FILTRO DE AIRE"
    if cleaned.startswith("FILTRO COMB"):
        return "FILTRO DE COMBUSTIBLE"
    if cleaned.startswith("PASTILLA") or cleaned.startswith("PASTILLAS"):
        return "PASTILLAS DE FRENO"
    if cleaned.startswith("DISCO FRENO DEL") or cleaned == "DISCO DEL":
        return "DISCO FRENO DELANTERO"
    if cleaned.startswith("DISCO FRENO TRAS") or cleaned == "DISCO TRAS":
        return "DISCO FRENO TRASERO"
    if cleaned.startswith("AMORTIGUADOR DEL") or cleaned.startswith("AMORT DEL"):
        return "AMORTIGUADOR DELANTERO"
    if cleaned.startswith("AMORTIGUADOR TRAS") or cleaned.startswith("AMORT TRAS"):
        return "AMORTIGUADOR TRASERO"
    if cleaned.startswith("HUB TRAS"):
        return "HUB TRASERO"
    if cleaned.startswith("HUB DEL"):
        return "HUB DELANTERO"
    if cleaned.startswith("TAMBOR FRENO"):
        return "TAMBOR DE FRENO"
    if cleaned.startswith("RETEN") or cleaned.startswith("RET ") or cleaned.startswith("OIL SEAL"):
        return "RETENEDOR"
    if cleaned.startswith("BOMBA GAS"):
        return "BOMBA DE GASOLINA"
    if cleaned.startswith("GUARDAPOLVO"):
        return "GUARDAPOLVO"
    if cleaned.startswith("MANIJA"):
        return "MANIJA DE PUERTA"
    if cleaned.startswith("FARO"):
        return "FARO DELANTERO"
    if cleaned.startswith("STOP") or cleaned.startswith("FAROL TRAS"):
        return "FAROL TRASERO"
    if cleaned.startswith("RADIADOR"):
        return "RADIADOR DE AGUA"
    if cleaned.startswith("CRUCETA"):
        return "CRUCETA"

    return cleaned if len(cleaned) > 2 else raw_desc.strip().upper()

def format_number_es(val, decimals=None):
    """Formatea números usando la coma (,) como separador decimal según requerimiento."""
    if val is None:
        return ""
    if decimals is not None:
        fmt = f"{{:.{decimals}f}}".format(val)
        return fmt.replace(".", ",")
    else:
        if val == int(val):
            return str(int(val))
        return str(val).replace(".", ",")

def clean_detail(desc):
    """
    Estandariza la descripción del repuesto:
    1. Revisa mappings exactos y reglas prioritarias en el diccionario JSON.
    2. Revisa si ya fue aprendido anteriormente.
    3. Revisa patrones automotrices generales comunes (STAR_MAPPINGS).
    4. Si es un repuesto nuevo y desconocido, aplica auto-aprendizaje y lo guarda en el diccionario.
    """
    if not desc:
        return ""
    raw = str(desc).strip()
    d_upper = raw.upper()

    dictionary = load_dictionary()
    exact_map = dictionary.get("exact_mappings", {})
    learned_map = dictionary.get("learned_mappings", {})

    # 1. Reglas fijas y prioritarias del diccionario
    # Se ordenan las claves por longitud descendente para que 'BALINERA DE CLUTCH' gane a 'BALINERA'
    sorted_keys = sorted(exact_map.keys(), key=len, reverse=True)
    for key in sorted_keys:
        if key in d_upper:
            return exact_map[key]

    # 2. Reglas de patrones automotrices generales frecuentes
    for pat, rep in GENERAL_AUTO_PATTERNS:
        if re.search(pat, d_upper):
            return rep

    # 3. Revisar si la descripción completa ya fue aprendida manualmente
    if d_upper in learned_map:
        return learned_map[d_upper]

    # 4. REPUESTO NUEVO Y DESCONOCIDO: Motor de Auto-Estandarización Inteligente
    standardized = auto_standardize_unknown_part(raw)
    
    # Guardar en memoria (se escribe a disco al final del proceso)
    save_learned_mapping(d_upper, standardized, flush=False)
    
    return standardized

def parse_decimal_safe(val):
    if val is None or val == "":
        return Decimal(0)
    val_str = str(val).replace(",", ".").strip()
    try:
        return Decimal(val_str)
    except:
        return Decimal(0)

def detect_provider(pdf_path):
    """
    Detecta automáticamente el proveedor del documento PDF:
    - STAR AUTO PARTS, S.A.
    - JAPAN INTERNATIONAL (Pedidos WO- / WO_LY-)
    - PERFECT TRADING
    """
    try:
        import pypdf
        reader = pypdf.PdfReader(pdf_path)
        if reader.pages:
            p0 = (reader.pages[0].extract_text() or "").upper()
            if "STAR AUTO PARTS" in p0 or "1158018-1-573227" in p0 or ("ZONA LIBRE DE COL" in p0 and "STAR" in p0):
                return "STAR_AUTO_PARTS"
            if "JAPAN INTERNATIONAL" in p0 or "WO-" in p0 or "WO_LY-" in p0 or ("PEDIDO:" in p0 and "COTIZAC" in p0) or "ENX00" in p0:
                return "JAPAN_INTERNATIONAL"
            if "PERFECT TRADING" in p0 or "COLON/OUT/" in p0:
                return "PERFECT_TRADING"
    except Exception:
        pass

    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            if pdf.pages:
                p0 = (pdf.pages[0].extract_text() or "").upper()
                if "STAR AUTO PARTS" in p0 or "1158018-1-573227" in p0:
                    return "STAR_AUTO_PARTS"
                if "AJ INTERNATIONAL" in p0:
                    return "AJ_INTERNATIONAL"
                if "JAPAN INTERNATIONAL" in p0 or "WO-" in p0 or "WO_LY-" in p0 or ("PEDIDO:" in p0 and "COTIZAC" in p0) or "ENX00" in p0:
                    return "JAPAN_INTERNATIONAL"
    except Exception:
        pass

    return "PERFECT_TRADING"

def parse_star_auto_parts(pdf_path):
    """
    Extrae facturas y repuestos para STAR AUTO PARTS, S.A.
    Utiliza visitor_text de pypdf para coordenadas exactas,
    extrayendo rangos de bultos, bultos mixtos, código, descripción y marcas.
    """
    import pypdf
    reader = pypdf.PdfReader(pdf_path)

    # 1. Extraer cabecera (Factura y Cliente) de la primera página
    p0_chunks = []
    def visitor0(text, cm, tm, fontDict, fontSize):
        t = text.strip()
        if t:
            p0_chunks.append({'text': t, 'x': round(tm[4], 1), 'y': round(tm[5], 1)})
    reader.pages[0].extract_text(visitor_text=visitor0)

    n_fact = "S/N"
    cliente = "DESCONOCIDO"

    fact_c = next((c for c in p0_chunks if c['text'] == 'Factura'), None)
    if fact_c:
        vals = [c['text'] for c in p0_chunks if abs(c['y'] - fact_c['y']) < 3 and c['x'] > fact_c['x']]
        if vals:
            n_fact = vals[0].strip()

    nom_c = next((c for c in p0_chunks if c['text'] == 'Nombre'), None)
    if nom_c:
        vals = [c['text'] for c in p0_chunks if abs(c['y'] - nom_c['y']) < 3 and nom_c['x'] < c['x'] < 500]
        if vals:
            cliente = ' '.join(vals).strip()

    # Fallback de cabecera si fallara por coordenadas
    if n_fact == "S/N" or cliente == "DESCONOCIDO":
        full_text = "\n".join([p.extract_text() or "" for p in reader.pages])
        if n_fact == "S/N":
            m_f = re.search(r'Factura\s*[:\s]*(\d+)', full_text)
            if m_f:
                n_fact = m_f.group(1).strip()
        if cliente == "DESCONOCIDO":
            m_c = re.search(r'Vendio A:\s*([^\n]+)', full_text)
            if m_c:
                cliente = m_c.group(1).strip()

    # 2. Extraer todos los fragmentos con coordenadas por página
    all_chunks = []
    for p_idx, page in enumerate(reader.pages):
        page_chunks = []
        def visitor(text, cm, tm, fontDict, fontSize):
            t = text.strip()
            if t:
                page_chunks.append({'text': t, 'x': round(tm[4], 1), 'y': round(tm[5], 1), 'p': p_idx + 1})
        page.extract_text(visitor_text=visitor)
        all_chunks.extend(page_chunks)

    pages_dict = {}
    for ch in all_chunks:
        pages_dict.setdefault(ch['p'], []).append(ch)

    current_bultos = []
    raw_items = []

    for p in sorted(pages_dict.keys()):
        chs = pages_dict[p]
        rows = []
        for ch in sorted(chs, key=lambda x: (x['y'], x['x'])):
            if not rows or abs(rows[-1][0]['y'] - ch['y']) > 3:
                rows.append([ch])
            else:
                rows[-1].append(ch)

        for r in rows:
            r_text = ' '.join(c['text'] for c in r)

            # Detección de línea de bulto
            m_b = re.search(r'Cantidad\s*B\.?\s*(\d+)\s*Bulto\s*(.*?)\s+([\d.,]+)\s+([\d.,]+)$', r_text, re.I)
            if not m_b:
                m_b = re.search(r'Bulto\s*(.*?)\s+([\d.,]+)\s+([\d.,]+)$', r_text, re.I)

            if m_b and 'Total Bultos' not in r_text:
                if len(m_b.groups()) == 4:
                    b_str = m_b.group(2)
                else:
                    b_str = m_b.group(1)

                b_nums = [int(x) for x in re.findall(r'\b\d+\b', b_str)]
                if len(b_nums) == 2:
                    current_bultos = list(range(b_nums[0], b_nums[1] + 1))
                elif len(b_nums) == 1:
                    current_bultos = [b_nums[0]]
                continue

            # Detección de fila de producto
            pza_list = [i for i, c in enumerate(r) if c['text'] in ['PZA', 'JGO', 'SET']]
            if pza_list and current_bultos:
                pza_idx = pza_list[0]
                try:
                    qty = Decimal(r[pza_idx-1]['text'].replace(',', ''))
                except:
                    continue
                uni = r[pza_idx]['text']

                # Código (elimina número de renglón inicial si existe)
                code_chunks = [c['text'] for c in r if 30 <= c['x'] < 140 and c['text'] not in ['Reng', 'Referencia']]
                raw_code = ' '.join(code_chunks)
                code = re.sub(r'^\d+\s+', '', raw_code).strip()

                # Descripción limpia
                desc_chunks = [c['text'] for c in r if 140 <= c['x'] < 420 and c['text'] not in ['CHINA', 'KOREA', 'MALASIA', 'THAILANDIA', 'TAIWAN', 'JAPAN', 'Descripcin', 'Descripcion', 'Descripción']]
                raw_desc = ' '.join(desc_chunks).strip()
                detalle = clean_detail(raw_desc)

                # Marca
                marca_chunks = [c['text'] for c in r if 420 <= c['x'] < 480 and c['text'] not in ['PZA', 'JGO', 'SET', 'Marca']]
                marca = ' '.join(marca_chunks).strip()
                if not marca and pza_idx > 2:
                    marca = r[pza_idx-2]['text'].strip()

                # Peso y CBM de la fila
                row_peso = Decimal(0)
                row_cbm = Decimal(0)
                nums = []
                for c in r[pza_idx+1:]:
                    try:
                        nums.append(Decimal(c['text'].replace(',', '')))
                    except:
                        pass
                if len(nums) >= 2:
                    row_peso = nums[-2]
                    row_cbm = nums[-1]
                elif len(nums) == 1:
                    row_peso = nums[0]

                # Distribución en bultos (rango o bulto individual)
                n_bultos = len(current_bultos)
                if n_bultos == 1:
                    raw_items.append({
                        'nb': current_bultos[0],
                        'code': code,
                        'detalle': detalle,
                        'uni': uni,
                        'marca': marca,
                        'qty': qty,
                        'peso': row_peso,
                        'cbm': row_cbm
                    })
                else:
                    qty_per = qty / n_bultos
                    kg_per = (row_peso / n_bultos).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                    cbm_per = (row_cbm / n_bultos).quantize(Decimal('0.0001'), rounding=ROUND_HALF_UP)

                    kg_accum = Decimal(0)
                    cbm_accum = Decimal(0)

                    for idx, b in enumerate(current_bultos):
                        is_last = (idx == n_bultos - 1)
                        if is_last:
                            curr_kg = row_peso - kg_accum
                            curr_cbm = row_cbm - cbm_accum
                        else:
                            curr_kg = kg_per
                            curr_cbm = cbm_per
                            kg_accum += curr_kg
                            cbm_accum += curr_cbm

                        raw_items.append({
                            'nb': b,
                            'code': code,
                            'detalle': detalle,
                            'uni': uni,
                            'marca': marca,
                            'qty': qty_per,
                            'peso': curr_kg,
                            'cbm': curr_cbm
                        })

    if not raw_items:
        print(f"\n[AVISO] No se encontraron ítems de empaque en {os.path.basename(pdf_path)}. (Es una factura comercial de precios sin lista de empaque).", file=sys.stderr)
        return []

    return [{
        'proveedor': 'STAR AUTO PARTS, S.A.',
        'n_fact': n_fact,
        'cliente': cliente,
        'items': raw_items
    }]

def parse_perfect_trading(pdf_path):
    """Extrae las facturas y repuestos desde un archivo PDF de PERFECT TRADING usando pdfplumber."""
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        full_text = "\n".join([page.extract_text() or "" for page in pdf.pages])
        lines = [l.strip() for l in full_text.split("\n") if l.strip()]

        proveedor = "PERFECT TRADING"
        n_fact = ""
        cliente = ""

        for i, l in enumerate(lines):
            if l.startswith("COLON/OUT/"):
                n_fact = l.replace("COLON/OUT/", "").strip()
            elif "Dirección de entrega:" in l or "Direccin de entrega:" in l:
                if i + 1 < len(lines):
                    cliente = lines[i+1].strip()

        raw_items = []
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or not row[0] or row[0] == 'Producto':
                        continue
                    if row[0] is None or row[0] == '':
                        continue

                    prod_text = str(row[0]).strip()
                    m = re.search(r'\[(.*?)\]\s*(.*)', prod_text, re.DOTALL)
                    if m:
                        code = m.group(1).strip()
                        raw_desc = m.group(2).strip().replace("\n", " ")
                    else:
                        code = prod_text
                        raw_desc = prod_text

                    qty = parse_decimal_safe(row[1])
                    marca = str(row[2]).replace("\n", " ").strip() if len(row) > 2 and row[2] else ""
                    kg = parse_decimal_safe(row[7]) if len(row) > 7 else Decimal(0)
                    m3 = parse_decimal_safe(row[8]) if len(row) > 8 else Decimal(0)
                    nb_str = str(row[10]).strip() if len(row) > 10 and row[10] else "1"

                    detalle = clean_detail(raw_desc)
                    uni = "PZA"

                    # Regla 4: Distribución matemática de bultos múltiples (ej. "1-20", "1-3")
                    range_match = re.match(r'^(\d+)\s*[-aA]\s*(\d+)$', nb_str)
                    if range_match:
                        start_nb = int(range_match.group(1))
                        end_nb = int(range_match.group(2))
                        count = end_nb - start_nb + 1

                        qty_per = qty / count
                        kg_per = (kg / count).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                        m3_per = (m3 / count).quantize(Decimal('0.0001'), rounding=ROUND_HALF_UP)

                        # Absorción de residuo matemático en el último bulto (Regla de Oro)
                        kg_diff = kg - (kg_per * count)
                        m3_diff = m3 - (m3_per * count)

                        for idx, b_num in enumerate(range(start_nb, end_nb + 1)):
                            cur_kg = kg_per + (kg_diff if idx == count - 1 else Decimal(0))
                            cur_m3 = m3_per + (m3_diff if idx == count - 1 else Decimal(0))
                            raw_items.append({
                                'nb': b_num,
                                'code': code,
                                'detalle': detalle,
                                'uni': uni,
                                'marca': marca,
                                'qty': qty_per,
                                'peso': cur_kg,
                                'cbm': cur_m3
                            })
                    else:
                        nb = int(re.sub(r'\D', '', nb_str)) if re.sub(r'\D', '', nb_str) else 1
                        raw_items.append({
                            'nb': nb,
                            'code': code,
                            'detalle': detalle,
                            'uni': uni,
                            'marca': marca,
                            'qty': qty,
                            'peso': kg,
                            'cbm': m3
                        })

        return [{
            'proveedor': proveedor,
            'n_fact': n_fact,
            'cliente': cliente,
            'items': raw_items
        }]

COUNTRIES_JAPAN_INT = [
    "JAPON", "JAPÓN", "TAIWAN", "TAIWÁN", "CHINA", "INDIA", 
    "COREA DEL SUR", "COREA", "TAILANDIA", "POLONIA", "ALEMANIA", 
    "BRASIL", "MALASIA", "INDONESIA", "MEXICO", "MÉXICO", "USA", "EEUU"
]

def parse_japan_international(pdf_path):
    """
    Extrae facturas y repuestos para JAPAN INTERNATIONAL (Pedidos tipo WO- y WO_LY-).
    Soporta formato por bultos TAG y formato directo/rango. Estandariza descripciones
    y garantiza cuadre matemático exacto de CANTIDAD, PESO y CBM al centavo.
    """
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        all_text = ""
        for p in pdf.pages:
            all_text += (p.extract_text() or "") + "\n"

        cliente_m = re.search(r'Cliente:\s*(.*?)\s+Fecha:', all_text)
        pedido_m = re.search(r'Pedido:\s*([^\n\r]+)', all_text)
        bultos_m = re.search(r'Bultos:\s*(\d+)', all_text)
        cant_m = re.search(r'Cantidad:\s*(\d+)', all_text)
        peso_b_m = re.search(r'Peso Bruto:\s*([\d\.]+)', all_text)
        cbm_m = re.search(r'Cubicaje m3:\s*([\d\.]+)', all_text)
        marca_hdr_m = re.search(r'Marca:\s*([^\n\r]+?)(?:\s+Peso Neto:|\s+Cubicaje|$)', all_text)

        cliente = cliente_m.group(1).strip() if cliente_m else ""
        pedido = pedido_m.group(1).strip() if pedido_m else ""
        hdr_bultos = int(bultos_m.group(1)) if bultos_m else 0
        hdr_cant = Decimal(cant_m.group(1)) if cant_m else Decimal(0)
        hdr_peso = Decimal(peso_b_m.group(1)) if peso_b_m else Decimal(0)
        hdr_cbm = Decimal(cbm_m.group(1)) if cbm_m else Decimal(0)
        hdr_marca = marca_hdr_m.group(1).strip() if marca_hdr_m else ""
        if "AUTOREPUESTOS SAN" in hdr_marca:
            hdr_marca = "AUTOREPUESTOS SAN JOSE"
        if "PESO NETO" in hdr_marca.upper():
            hdr_marca = ""

        lines = []
        for p in pdf.pages:
            for l in (p.extract_text() or "").split("\n"):
                l = l.strip()
                if not l:
                    continue
                if any(l.startswith(k) for k in ["FECHA:", "USUARIO:", "Lista de Empaque", "Cliente:", "Bultos:", "Cantidad:", "Peso Bruto:", "Peso Neto:", "Cubicaje"]):
                    continue
                if re.search(r'P[aá\uFFFD]gina\s+\d+\s+de\s+\d+', l, re.IGNORECASE):
                    continue
                if l.startswith("No.") or l.startswith("Bulto Referencia") or l == ".":
                    continue
                lines.append(l)

        raw_items = []
        cur_bulto = 1
        cur_bulto_peso = Decimal(0)
        cur_bulto_cbm = Decimal(0)
        cur_is_tag_box = False
        bulto_items_temp = []

        def flush_bulto():
            nonlocal bulto_items_temp, cur_bulto_peso, cur_bulto_cbm
            if not bulto_items_temp:
                return
            tot_q = sum(it['qty'] for it in bulto_items_temp)
            if tot_q == 0:
                tot_q = Decimal(1)
            accum_p = Decimal(0)
            accum_c = Decimal(0)
            n_items = len(bulto_items_temp)
            for idx, it in enumerate(bulto_items_temp):
                if idx == n_items - 1:
                    p = cur_bulto_peso - accum_p
                    c = cur_bulto_cbm - accum_c
                else:
                    p = (cur_bulto_peso * (it['qty'] / tot_q)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                    c = (cur_bulto_cbm * (it['qty'] / tot_q)).quantize(Decimal('0.0001'), rounding=ROUND_HALF_UP)
                    accum_p += p
                    accum_c += c
                it['peso'] = p
                it['cbm'] = c
                raw_items.append(it)
            bulto_items_temp = []

        i = 0
        while i < len(lines):
            l = lines[i]

            box_m = re.match(r'^(\d+)\s+R\d+\s+TAG\w+\s+(\d+)(?:\s+([\d\.]+)\s+([\d\.]+))?', l)
            if box_m:
                flush_bulto()
                cur_bulto = int(box_m.group(1))
                cur_is_tag_box = True
                cur_bulto_peso = Decimal(box_m.group(3)) if box_m.group(3) else Decimal(0)
                cur_bulto_cbm = Decimal(box_m.group(4)) if box_m.group(4) else Decimal(0)
                i += 1
                continue

            style_b_m = re.search(r'(\d+)\s+(\d+)\s+UND$', l)
            style_a_m = re.search(r'(?:(\d+)\s+)?(\d+)\s+([\d\.]+)\s+([\d\.]+)$', l)

            if style_b_m:
                cant = Decimal(style_b_m.group(2))
                raw_prefix = l[:style_b_m.start()].strip()
                tokens = raw_prefix.split()
                pais_found = ""
                marca_found = hdr_marca
                code = tokens[0] if tokens else ""

                if len(tokens) >= 3 and tokens[-2].upper() == "COREA" and tokens[-1].upper() == "DEL":
                    pais_found = "Corea del Sur"
                    marca_found = tokens[-3]
                    raw_desc = " ".join(tokens[1:-3])
                elif len(tokens) >= 4 and tokens[-3].upper() == "COREA" and tokens[-2].upper() == "DEL" and tokens[-1].upper() == "SUR":
                    pais_found = "Corea del Sur"
                    marca_found = tokens[-4]
                    raw_desc = " ".join(tokens[1:-4])
                elif len(tokens) >= 2 and any(c in tokens[-1].upper() for c in COUNTRIES_JAPAN_INT):
                    pais_found = tokens[-1]
                    marca_found = tokens[-2]
                    raw_desc = " ".join(tokens[1:-2])
                else:
                    marca_found = tokens[-1] if tokens else hdr_marca
                    raw_desc = " ".join(tokens[1:-1])

                while i + 1 < len(lines):
                    next_l = lines[i+1].strip()
                    if re.match(r'^(\d+)\s+R\d+\s+TAG', next_l) or re.search(r'(\d+)\s+(\d+)\s+UND$', next_l) or re.search(r'(?:(\d+)\s+)?(\d+)\s+([\d\.]+)\s+([\d\.]+)$', next_l):
                        break
                    next_tokens = next_l.split()
                    if next_tokens:
                        first_t = next_tokens[0]
                        if marca_found and (code + first_t).endswith(marca_found):
                            code = code + first_t
                            next_tokens = next_tokens[1:]
                        elif marca_found and (code + "-" + first_t).endswith(marca_found):
                            code = code + "-" + first_t
                            next_tokens = next_tokens[1:]
                        elif len(first_t) <= 3 and not code.endswith("-") and (code.endswith("-KI") and first_t == "C" or code.endswith("-JO") and first_t == "MO"):
                            code = code + first_t
                            next_tokens = next_tokens[1:]
                    if next_tokens and next_tokens[-1].upper() == 'SUR' and 'COREA' in pais_found.upper():
                        next_tokens = next_tokens[:-1]
                    if next_tokens:
                        raw_desc += " " + " ".join(next_tokens)
                    i += 1

                detalle = clean_detail(raw_desc)

                bulto_items_temp.append({
                    'nb': cur_bulto,
                    'code': code,
                    'detalle': detalle,
                    'uni': 'PZA',
                    'marca': marca_found or hdr_marca,
                    'qty': cant,
                    'peso': Decimal(0),
                    'cbm': Decimal(0)
                })

            elif style_a_m and not re.match(r'^\d+\s+R\d+', l):
                qty = Decimal(style_a_m.group(2))
                peso = Decimal(style_a_m.group(3))
                cbm = Decimal(style_a_m.group(4))
                raw_prefix = l[:style_a_m.start()].strip()
                while i + 1 < len(lines):
                    next_l = lines[i+1]
                    if re.match(r'^(\d+)\s+R\d+\s+TAG', next_l) or re.search(r'(\d+)\s+(\d+)\s+UND$', next_l) or re.search(r'(\d+)\s+([\d\.]+)\s+([\d\.]+)$', next_l):
                        break
                    raw_prefix += " " + next_l
                    i += 1

                if cur_is_tag_box:
                    tokens = raw_prefix.split()
                    if tokens and tokens[-1].isdigit():
                        tokens = tokens[:-1]
                    code = tokens[0] if tokens else ""
                    raw_desc = " ".join(tokens[1:])
                    detalle = clean_detail(raw_desc)

                    raw_items.append({
                        'nb': cur_bulto,
                        'code': code,
                        'detalle': detalle,
                        'uni': 'PZA',
                        'marca': hdr_marca,
                        'qty': qty,
                        'peso': peso,
                        'cbm': cbm
                    })
                else:
                    range_m = re.match(r'^(\d+)\s*[-aA]\s*(\d+)\s+(\S+)\s+(.*)', raw_prefix)
                    single_nb_m = re.match(r'^(\d+)\s+([A-Za-z0-9][A-Za-z0-9\-_]*)\s+(.*)', raw_prefix)

                    if range_m:
                        start_b = int(range_m.group(1))
                        end_b = int(range_m.group(2))
                        count_b = end_b - start_b + 1
                        code = range_m.group(3)
                        desc_raw = range_m.group(4)
                        desc = re.sub(r'\s+\d+$', '', desc_raw).strip()
                        detalle = clean_detail(desc)

                        qty_per = qty / count_b
                        kg_per = (peso / count_b).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                        cbm_per = (cbm / count_b).quantize(Decimal('0.0001'), rounding=ROUND_HALF_UP)
                        kg_diff = peso - (kg_per * count_b)
                        cbm_diff = cbm - (cbm_per * count_b)

                        for b_idx, b_num in enumerate(range(start_b, end_b + 1)):
                            raw_items.append({
                                'nb': b_num,
                                'code': code,
                                'detalle': detalle,
                                'uni': 'PZA',
                                'marca': hdr_marca,
                                'qty': qty_per,
                                'peso': kg_per + (kg_diff if b_idx == count_b - 1 else Decimal(0)),
                                'cbm': cbm_per + (cbm_diff if b_idx == count_b - 1 else Decimal(0))
                            })
                    elif single_nb_m:
                        b_num = int(single_nb_m.group(1))
                        code = single_nb_m.group(2)
                        desc = re.sub(r'\s+\d+$', '', single_nb_m.group(3)).strip()
                        detalle = clean_detail(desc)
                        raw_items.append({
                            'nb': b_num,
                            'code': code,
                            'detalle': detalle,
                            'uni': 'PZA',
                            'marca': hdr_marca,
                            'qty': qty,
                            'peso': peso,
                            'cbm': cbm
                        })
                    else:
                        tokens = raw_prefix.split()
                        code = tokens[0] if tokens else ""
                        raw_desc = " ".join(tokens[1:])
                        detalle = clean_detail(raw_desc)
                        raw_items.append({
                            'nb': cur_bulto,
                            'code': code,
                            'detalle': detalle,
                            'uni': 'PZA',
                            'marca': hdr_marca,
                            'qty': qty,
                            'peso': peso,
                            'cbm': cbm
                        })

            i += 1

        flush_bulto()

        sum_p = sum(it['peso'] for it in raw_items)
        sum_c = sum(it['cbm'] for it in raw_items)
        diff_p = hdr_peso - sum_p
        diff_c = hdr_cbm - sum_c

        if raw_items:
            raw_items[-1]['peso'] += diff_p
            raw_items[-1]['cbm'] += diff_c

        return [{
            'proveedor': 'JAPAN INTERNATIONAL',
            'n_fact': pedido,
            'cliente': cliente,
            'items': raw_items
        }]

def parse_aj_international(pdf_path):
    """
    Extrae facturas y repuestos desde archivos PDF de AJ INTERNATIONAL GROUP, S.A.
    Maneja múltiples bultos secuenciales y saltos de página con encabezados repetidos.
    """
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        all_pages = []
        for p in pdf.pages:
            all_pages.append((p.extract_text() or '').split('\n'))

    full_text = '\n'.join('\n'.join(lines) for lines in all_pages)

    proveedor = "AJ INTERNATIONAL GROUP, S.A."
    n_fact_m = re.search(r'FACTURA\s+No\.:\s*(\d+)', full_text)
    cliente_m = re.search(r'Cliente:\s*(?:[\w\-]+\s*-\s*)?([^\n\r]+?)(?:\s+R\.U\.C|\s+Fecha|\n|$)', full_text)

    n_fact = n_fact_m.group(1).strip() if n_fact_m else ""
    cliente = cliente_m.group(1).strip() if cliente_m else ""

    tot_p_m = re.search(r'Peso Final \(Kg\):\s*([\d\.]+)', full_text)
    tot_c_m = re.search(r'Total M3:\s*([\d\.]+)', full_text)

    hdr_peso = parse_decimal_safe(tot_p_m.group(1)) if tot_p_m else Decimal(0)
    hdr_cbm = parse_decimal_safe(tot_c_m.group(1)) if tot_c_m else Decimal(0)

    KNOWN_BRANDS = ['DDT USA', 'PSH - POLLAND', 'VULKO', 'NTS', 'ZM', 'GAUSS', 'ENZO', 'ZEN']

    bulto_dict = {}
    current_bulto_key = None
    bulto_seq = 0

    for p_lines in all_pages:
        i = 0
        while i < len(p_lines):
            l = p_lines[i].strip()
            if not l:
                i += 1
                continue

            b_m = re.search(r'(#\w+)\s+Bulto:(\d+)\s+Peso\s*\(Kg\):([\d\.]+)\s+M3:([\d\.]+)', l)
            if b_m:
                tag = b_m.group(1)
                b_loc = int(b_m.group(2))
                key = (tag, b_loc)
                if key not in bulto_dict:
                    bulto_seq += 1
                    bulto_dict[key] = {
                        'nb': bulto_seq,
                        'peso': parse_decimal_safe(b_m.group(3)),
                        'cbm': parse_decimal_safe(b_m.group(4)),
                        'items': []
                    }
                current_bulto_key = key
                i += 1
                continue

            if any(l.startswith(k) for k in ['Página', 'AJ INTERNATIONAL', 'R.U.C.', 'Dirección:', 'Télefonos:', 'LISTA DE EMPAQUE', 'Cliente:', 'Teléfono:', 'Observación:', 'Código Referencia', 'Peso Bultos', 'Bultos Totales']):
                i += 1
                continue

            it_m = re.search(r'^([A-Z0-9\-]+)\s+(\S+)\s+(.*?)\s+(\d+\.\d{2})$', l)
            if it_m and current_bulto_key:
                code = it_m.group(1)
                ref = it_m.group(2)
                rest = it_m.group(3)
                qty = parse_decimal_safe(it_m.group(4))

                marca = ""
                desc = rest
                for b_cand in sorted(KNOWN_BRANDS, key=len, reverse=True):
                    if rest.startswith(b_cand):
                        marca = b_cand
                        desc = rest[len(b_cand):].strip()
                        break
                if not marca:
                    tokens = rest.split()
                    marca = tokens[0] if tokens else "AJ"
                    desc = " ".join(tokens[1:])

                bulto_dict[current_bulto_key]['items'].append({
                    'code': code,
                    'ref': ref,
                    'desc': desc,
                    'marca': marca,
                    'qty': qty
                })
                i += 1
                continue

            i += 1

    raw_items = []
    for (tag, loc), bdata in sorted(bulto_dict.items(), key=lambda x: x[1]['nb']):
        nb = bdata['nb']
        b_p = bdata['peso']
        b_c = bdata['cbm']
        items = bdata['items']
        tot_q = sum(it['qty'] for it in items)
        if tot_q == 0:
            tot_q = Decimal(1)

        accum_p = Decimal(0)
        accum_c = Decimal(0)
        n_it = len(items)

        for idx, it in enumerate(items):
            if idx == n_it - 1:
                p = b_p - accum_p
                c = b_c - accum_c
            else:
                p = (b_p * (it['qty'] / tot_q)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                c = (b_c * (it['qty'] / tot_q)).quantize(Decimal('0.0001'), rounding=ROUND_HALF_UP)
                accum_p += p
                accum_c += c

            detalle = clean_detail(it['desc'])
            raw_items.append({
                'nb': nb,
                'code': it['code'],
                'detalle': detalle,
                'uni': 'PZA',
                'marca': it['marca'],
                'qty': it['qty'],
                'peso': p,
                'cbm': c
            })

    if hdr_peso > 0 and raw_items:
        sum_p = sum(it['peso'] for it in raw_items)
        sum_c = sum(it['cbm'] for it in raw_items)
        raw_items[-1]['peso'] += (hdr_peso - sum_p)
        raw_items[-1]['cbm'] += (hdr_cbm - sum_c)

    return [{
        'proveedor': proveedor,
        'n_fact': n_fact,
        'cliente': cliente,
        'items': raw_items
    }]

def parse_pdf(pdf_path):
    """Detecta automáticamente el proveedor del PDF y ejecuta el parser correspondiente."""
    prov = detect_provider(pdf_path)
    if prov == "STAR_AUTO_PARTS":
        return parse_star_auto_parts(pdf_path)
    elif prov == "AJ_INTERNATIONAL":
        return parse_aj_international(pdf_path)
    elif prov == "JAPAN_INTERNATIONAL":
        return parse_japan_international(pdf_path)
    else:
        return parse_perfect_trading(pdf_path)

def parse_adk_excel(excel_path, wb=None):
    """Extrae facturas y repuestos desde archivos Excel de ADK CORPORATION."""
    if wb is None:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active

    proveedor = 'ADK CORPORATION'
    factura = ''
    cliente = ''

    for r in range(1, min(16, ws.max_row + 1)):
        for c in range(1, min(65, ws.max_column + 1)):
            v = ws.cell(r, c).value
            if v == 'Factura':
                for offset in [8, 1, 2, 7]:
                    val = ws.cell(r, c + offset).value
                    if val and str(val).strip():
                        factura = str(val).strip()
                        break
            elif v == 'Consignatario:':
                for c2 in range(c + 1, min(c + 20, ws.max_column + 1)):
                    v2 = ws.cell(r, c2).value
                    if v2 and str(v2).strip():
                        cliente = str(v2).strip()
                        break

    if not factura:
        factura = str(ws.cell(12, 50).value or '').strip()
    if not cliente:
        cliente = str(ws.cell(10, 16).value or '').strip()
    if not factura:
        m = re.search(r'([A-Za-z0-9\-]+)', os.path.basename(excel_path))
        if m:
            factura = m.group(1)

    tot_peso_hdr = Decimal(0)
    tot_cbm_hdr = Decimal(0)

    for r in range(ws.max_row, max(1, ws.max_row - 15), -1):
        if str(ws.cell(r, 3).value or '').strip() == 'Total Bultos:':
            tot_peso_hdr = parse_decimal_safe(ws.cell(r, 54).value)
            tot_cbm_hdr = parse_decimal_safe(ws.cell(r, 66).value)
            break

    bultos = {}
    cur_b_num = 1
    cur_b_peso = Decimal(0)
    cur_b_cbm = Decimal(0)

    r = 16
    while r <= ws.max_row:
        if str(ws.cell(r, 3).value or '').strip() == 'Total Bultos:':
            break

        if ws.cell(r, 19).value == 'Bulto':
            cur_b_num = int(ws.cell(r, 22).value or cur_b_num)
            cur_b_peso = parse_decimal_safe(ws.cell(r, 51).value)
            cur_b_cbm = parse_decimal_safe(ws.cell(r, 61).value)
            bultos[cur_b_num] = {
                'peso': cur_b_peso,
                'cbm': cur_b_cbm,
                'items': []
            }
            r += 1
            continue

        c2 = ws.cell(r, 2).value
        c5 = ws.cell(r, 5).value
        if c2 is not None and isinstance(c2, (int, float)) and c5:
            code = str(c5).strip()
            desc = str(ws.cell(r, 16).value or ws.cell(r, 21).value or '').strip()
            marca = str(ws.cell(r, 33).value or ws.cell(r, 34).value or '').strip()

            qty = Decimal(0)
            if r + 1 <= ws.max_row and ws.cell(r + 1, 33).value is not None:
                qty = parse_decimal_safe(ws.cell(r + 1, 33).value)

            uni = 'PZA'
            if cur_b_num not in bultos:
                bultos[cur_b_num] = {'peso': Decimal(0), 'cbm': Decimal(0), 'items': []}

            bultos[cur_b_num]['items'].append({
                'code': code,
                'desc': desc,
                'marca': marca,
                'qty': qty,
                'uni': uni
            })
            r += 2
            continue

        r += 1

    raw_items = []
    for b_num, b_data in bultos.items():
        b_items = b_data['items']
        b_peso = b_data['peso']
        b_cbm = b_data['cbm']
        tot_q = sum(it['qty'] for it in b_items)
        if tot_q == 0:
            tot_q = Decimal(1)

        accum_p = Decimal(0)
        accum_c = Decimal(0)
        n_it = len(b_items)
        for idx, it in enumerate(b_items):
            if idx == n_it - 1:
                p = b_peso - accum_p
                c = b_cbm - accum_c
            else:
                p = (b_peso * (it['qty'] / tot_q)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                c = (b_cbm * (it['qty'] / tot_q)).quantize(Decimal('0.0001'), rounding=ROUND_HALF_UP)
                accum_p += p
                accum_c += c

            detalle = clean_detail(it['desc'])
            raw_items.append({
                'nb': b_num,
                'code': it['code'],
                'detalle': detalle,
                'uni': it['uni'],
                'marca': it['marca'],
                'qty': it['qty'],
                'peso': p,
                'cbm': c
            })

    if tot_peso_hdr > 0 and raw_items:
        sum_p = sum(it['peso'] for it in raw_items)
        sum_c = sum(it['cbm'] for it in raw_items)
        diff_p = tot_peso_hdr - sum_p
        diff_c = tot_cbm_hdr - sum_c
        raw_items[-1]['peso'] += diff_p
        raw_items[-1]['cbm'] += diff_c

    return [{
        'proveedor': proveedor,
        'n_fact': factura,
        'cliente': cliente,
        'items': raw_items
    }]

def parse_nikomoto(excel_path):
    """
    Extrae facturas y repuestos desde archivos Excel (.xls y .xlsx) de NIKOMOTO, S.A.
    Soporta rangos de bultos (ej. '14 Al 63'), pesos por bulto y pesos por fila,
    y extrae con precisión número de factura y cliente.
    """
    ext = os.path.splitext(excel_path)[1].lower()
    if ext == '.xls':
        import xlrd
        wb = xlrd.open_workbook(excel_path)
        sh = wb.sheet_by_index(0)
        nrows, ncols = sh.nrows, sh.ncols
        get_val = lambda r, c: sh.cell_value(r, c)
    else:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        sh = wb.active
        nrows, ncols = sh.max_row, sh.max_column
        get_val = lambda r, c: sh.cell(r + 1, c + 1).value or ''

    proveedor = 'NIKOMOTO, S.A.'
    n_fact = ''
    cliente = ''

    # Extracción de cabecera (Factura y Cliente)
    for r in range(min(15, nrows)):
        for c in range(ncols):
            v = str(get_val(r, c)).strip()
            if v.lower() in ['nombe', 'nombre']:
                for c2 in range(c + 1, min(c + 15, ncols)):
                    v2 = get_val(r, c2)
                    if v2:
                        cliente = str(v2).strip()
                        break
            elif v == 'Lista de Empaque':
                for c2 in range(0, c):
                    v2 = get_val(r, c2)
                    if v2:
                        val_str = str(v2).strip()
                        if val_str.endswith('.0'):
                            val_str = val_str[:-2]
                        n_fact = val_str
                        break

    if not n_fact or n_fact == 'REPUESTOS':
        m = re.search(r'00\d{5}-E|\d{5}', os.path.basename(excel_path))
        if m:
            n_fact = m.group(0)
        elif not n_fact:
            n_fact = "S/F"

    if not cliente:
        base = os.path.basename(excel_path)
        m = re.search(r'^(.*?)\s*-\s*NIKOMOTO', base, re.IGNORECASE)
        if m:
            cliente = m.group(1).strip()

    # Totales de la cabecera / pie de página
    tot_p_hdr = Decimal(0)
    tot_c_hdr = Decimal(0)
    tot_q_hdr = Decimal(0)
    tot_b_hdr = Decimal(0)

    for r in range(nrows - 1, max(-1, nrows - 25), -1):
        for c in range(ncols):
            if 'total de bultos' in str(get_val(r, c)).lower():
                vals = [parse_decimal_safe(get_val(r, c2)) for c2 in range(ncols) if parse_decimal_safe(get_val(r, c2)) > 0]
                if len(vals) >= 4:
                    tot_b_hdr = vals[0]
                    tot_q_hdr = vals[1]
                    tot_p_hdr = vals[2]
                    tot_c_hdr = vals[3]
                break
        if tot_q_hdr > 0:
            break

    # Detección dinámica de bultos y lectura de repuestos
    raw_bultos = []
    cur_b_start = 1
    cur_b_end = 1
    cur_hdr_p = Decimal(0)
    cur_hdr_c = Decimal(0)
    cur_items = []
    in_items = False

    for r in range(nrows):
        if 'total de bultos' in str(get_val(r, 1)).lower():
            break

        is_bulto_row = False
        for c in range(ncols):
            if str(get_val(r, c)).strip().upper() == 'BULTO':
                is_bulto_row = True
                if cur_items:
                    raw_bultos.append((cur_b_start, cur_b_end, cur_hdr_p, cur_hdr_c, cur_items))
                    cur_items = []

                nums = []
                for c2 in range(c + 1, min(c + 10, ncols)):
                    val = str(get_val(r, c2)).strip()
                    if val and val.upper() != 'AL':
                        try:
                            nums.append(int(float(val)))
                        except:
                            pass
                if len(nums) == 1:
                    cur_b_start = nums[0]
                    cur_b_end = nums[0]
                elif len(nums) >= 2:
                    cur_b_start = nums[0]
                    cur_b_end = nums[1]
                else:
                    cur_b_start += 1
                    cur_b_end = cur_b_start

                cur_hdr_p = parse_decimal_safe(get_val(r, 33))
                cur_hdr_c = parse_decimal_safe(get_val(r, 37))
                in_items = True
                break

        if is_bulto_row:
            continue

        if not in_items:
            continue

        code = str(get_val(r, 3)).strip()
        desc = str(get_val(r, 18)).strip()
        qty = parse_decimal_safe(get_val(r, 29))
        p_row = parse_decimal_safe(get_val(r, 33))
        c_row = parse_decimal_safe(get_val(r, 37))
        marca = str(get_val(r, 14)).strip()

        if code and desc and qty > 0 and 'PRODUCTO' not in code.upper() and 'TOTAL' not in code.upper():
            cur_items.append({
                'code': code,
                'desc': desc,
                'uni': 'PZA',
                'marca': marca or 'NIKOMOTO',
                'qty': qty,
                'peso_row': p_row,
                'cbm_row': c_row
            })

    if cur_items:
        raw_bultos.append((cur_b_start, cur_b_end, cur_hdr_p, cur_hdr_c, cur_items))

    # Expansión de bultos y distribución exacta de pesos y volúmenes (Regla 3)
    raw_items = []
    for b_start, b_end, hdr_p, hdr_c, items in raw_bultos:
        sum_row_p = sum(it['peso_row'] for it in items)
        sum_row_c = sum(it['cbm_row'] for it in items)
        b_p = hdr_p if hdr_p > 0 else sum_row_p
        b_c = hdr_c if hdr_c > 0 else sum_row_c

        b_list = list(range(b_start, b_end + 1))
        n_b = len(b_list)
        tot_q = sum(it['qty'] for it in items)
        if tot_q == 0:
            tot_q = Decimal(1)

        for it in items:
            detalle = clean_detail(it['desc'])
            q_per = (it['qty'] / Decimal(n_b)).quantize(Decimal('1'), rounding=ROUND_HALF_UP) if it['qty'] % Decimal(n_b) == 0 else (it['qty'] / Decimal(n_b))
            q_diff = it['qty'] - (q_per * Decimal(n_b))

            it_p_tot = b_p * (it['qty'] / tot_q)
            it_c_tot = b_c * (it['qty'] / tot_q)
            p_per = (it_p_tot / Decimal(n_b)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            c_per = (it_c_tot / Decimal(n_b)).quantize(Decimal('0.0001'), rounding=ROUND_HALF_UP)
            p_diff = it_p_tot - (p_per * Decimal(n_b))
            c_diff = it_c_tot - (c_per * Decimal(n_b))

            for b_idx, b_num in enumerate(b_list):
                raw_items.append({
                    'nb': b_num,
                    'code': it['code'],
                    'detalle': detalle,
                    'uni': it['uni'],
                    'marca': it['marca'],
                    'qty': q_per + (q_diff if b_idx == n_b - 1 else Decimal(0)),
                    'peso': p_per + (p_diff if b_idx == n_b - 1 else Decimal(0)),
                    'cbm': c_per + (c_diff if b_idx == n_b - 1 else Decimal(0))
                })

    # Cuadre final con totales de cabecera si están presentes
    if tot_p_hdr > 0 and raw_items:
        sum_p = sum(it['peso'] for it in raw_items)
        sum_c = sum(it['cbm'] for it in raw_items)
        diff_p = tot_p_hdr - sum_p
        diff_c = tot_c_hdr - sum_c
        raw_items[-1]['peso'] += diff_p
        raw_items[-1]['cbm'] += diff_c

    return [{
        'proveedor': proveedor,
        'n_fact': n_fact,
        'cliente': cliente,
        'items': raw_items
    }]

def parse_nikomoto_xlsx(excel_path, wb=None):
    """Compatibilidad: Llama al parser unificado de NIKOMOTO."""
    return parse_nikomoto(excel_path)

def parse_excel_xls(excel_path):
    """Compatibilidad: Llama al parser unificado de NIKOMOTO."""
    return parse_nikomoto(excel_path)

def parse_standard_excel(excel_path, wb=None):
    """Extrae las facturas y repuestos desde el formato estándar de Excel."""
    if wb is None:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
    all_invoices = []

    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        invoices_in_sheet = {}
        current_fact = None

        for r in range(1, sheet.max_row + 1):
            c1 = str(sheet.cell(row=r, column=1).value or '').strip()
            c2 = str(sheet.cell(row=r, column=2).value or '').strip()

            if c1.startswith("FACTURA"):
                current_fact = c1.replace("FACTURA", "").strip()
                continue

            if c1 in ["PROVEEDOR", ""] and c2 in ["N° FACT", "N FACT", ""]:
                continue

            prov = c1 or "PERFECT TRADING"
            n_fact = c2 or current_fact
            cliente = str(sheet.cell(row=r, column=3).value or '').strip()
            nb_val = sheet.cell(row=r, column=4).value
            code = str(sheet.cell(row=r, column=5).value or '').strip()
            raw_desc = str(sheet.cell(row=r, column=6).value or '').strip()
            uni = str(sheet.cell(row=r, column=7).value or 'PZA').strip()
            marca = str(sheet.cell(row=r, column=8).value or '').strip()
            qty_val = sheet.cell(row=r, column=9).value
            peso_val = sheet.cell(row=r, column=10).value
            cbm_val = sheet.cell(row=r, column=11).value

            if not code and not raw_desc:
                continue

            if not n_fact:
                continue

            if n_fact not in invoices_in_sheet:
                invoices_in_sheet[n_fact] = {
                    'proveedor': prov,
                    'n_fact': n_fact,
                    'cliente': cliente,
                    'items': []
                }

            qty = parse_decimal_safe(qty_val)
            peso = parse_decimal_safe(peso_val)
            cbm = parse_decimal_safe(cbm_val)
            nb = int(nb_val) if nb_val is not None and str(nb_val).isdigit() else 1

            detalle = clean_detail(raw_desc)

            invoices_in_sheet[n_fact]['items'].append({
                'nb': nb,
                'code': code,
                'detalle': detalle,
                'uni': uni,
                'marca': marca,
                'qty': qty,
                'peso': peso,
                'cbm': cbm
            })

        all_invoices.extend(invoices_in_sheet.values())

    return all_invoices

def parse_excel(excel_path):
    """
    Extrae facturas y repuestos desde archivos Excel (.xlsx o .xls).
    Detecta automáticamente si es ADK CORPORATION, NIKOMOTO, S.A. o formato estándar.
    """
    ext = os.path.splitext(excel_path)[1].lower()
    if ext == '.xls':
        return parse_excel_xls(excel_path)

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active

    header_text = ""
    for r in range(1, min(15, ws.max_row + 1)):
        for c in range(1, min(25, ws.max_column + 1)):
            v = str(ws.cell(r, c).value or '')
            if v:
                header_text += " " + v.upper()

    if "ADK CORPORATION" in header_text or "COCO SOLITO" in header_text or "CLAVE 8752" in header_text:
        return parse_adk_excel(excel_path, wb=wb)
    elif "NIKOMOTO" in header_text or "NIKOMOTOZL.COM" in header_text or "TAIHO" in header_text:
        return parse_nikomoto(excel_path)
    else:
        return parse_standard_excel(excel_path, wb=wb)

def parse_aisin_docx(file_path):
    """
    Extrae facturas y repuestos desde archivos Word (.docx) de AISIN SALES LATIN AMERICA, S.A.
    Maneja bloques de bultos '( BULTOS#: 1/ 1 ---> 16.00 Kilos )', ítems y líneas MARCA->.
    Aplica la Regla 3 para división equitativa de bultos múltiples.
    """
    import docx

    doc = docx.Document(file_path)
    paras = [p.text for p in doc.paragraphs if p.text.strip()]
    full_text = "\n".join(paras)

    proveedor = "AISIN SALES LATIN AMERICA, S.A."
    cliente = ""
    n_fact = ""

    senores_m = re.search(r'SENORES\s*:\s*([^\n\r]+)', full_text)
    if senores_m:
        raw_senores = senores_m.group(1).strip()
        m_num = re.search(r'\s+(\d{3,6})$', raw_senores)
        if m_num:
            n_fact = m_num.group(1)
            cliente = raw_senores[:m_num.start()].strip()
        else:
            cliente = raw_senores

    if not n_fact:
        m = re.search(r'(\d{4,})', os.path.basename(file_path))
        if m:
            n_fact = m.group(1)
        else:
            n_fact = "S/F"

    tot_p_m = re.search(r'TOTAL PIEZAS:\s*([\d,\.]+)', full_text)
    tot_w_m = re.search(r'TOTAL PESO\.\.:\s*([\d,\.]+)', full_text)
    tot_b_m = re.search(r'TOTAL BULTOS:\s*([\d,\.]+)', full_text)

    hdr_cant = parse_decimal_safe(tot_p_m.group(1)) if tot_p_m else Decimal(0)
    hdr_peso = parse_decimal_safe(tot_w_m.group(1)) if tot_w_m else Decimal(0)
    hdr_bultos = int(parse_decimal_safe(tot_b_m.group(1))) if tot_b_m else 0

    raw_bultos = []
    cur_b_start = 1
    cur_b_end = 1
    cur_b_peso = Decimal(0)
    cur_items = []

    i = 0
    while i < len(paras):
        p = paras[i]
        b_box_m = re.search(r'\(\s*BULTOS#:\s*(\d+)/\s*(\d+)\s*--->\s*([\d,\.]+)\s*Kilos\s*\)', p)
        if b_box_m:
            if cur_items:
                raw_bultos.append((cur_b_start, cur_b_end, cur_b_peso, cur_items))
                cur_items = []
            cur_b_start = int(b_box_m.group(1))
            cur_b_end = int(b_box_m.group(2))
            cur_b_peso = parse_decimal_safe(b_box_m.group(3))
            i += 1
            continue

        it_m = re.search(r'^\s*(\d+)/\s*(\d+)\s+(\d+)\s+(\S+)\s+(.*?)\s+([\d,\.]+)\s+([A-Z]{2})\s*$', p)
        if it_m:
            code = it_m.group(4)
            middle = it_m.group(5).strip()
            qty = parse_decimal_safe(it_m.group(6))
            uni = it_m.group(7)

            marca = "AISIN"
            if i + 1 < len(paras):
                m_marca = re.search(r'MARCA->\s*([^/\n\r]+)', paras[i+1])
                if m_marca:
                    marca = m_marca.group(1).strip()

            parts = re.split(r'\s{2,}', middle, maxsplit=1)
            raw_desc = parts[1] if len(parts) == 2 else parts[0]

            detalle = clean_detail(raw_desc)
            cur_items.append({
                'code': code,
                'detalle': detalle,
                'uni': 'PZA' if uni in ['PZ', 'UND', 'UN'] else uni,
                'marca': marca,
                'qty': qty
            })

        i += 1

    if cur_items:
        raw_bultos.append((cur_b_start, cur_b_end, cur_b_peso, cur_items))

    raw_items = []
    for b_start, b_end, b_peso, items in raw_bultos:
        b_list = list(range(b_start, b_end + 1))
        n_b = len(b_list)
        tot_q = sum(it['qty'] for it in items)
        if tot_q == 0:
            tot_q = Decimal(1)

        for it in items:
            q_per = (it['qty'] / Decimal(n_b)).quantize(Decimal('1'), rounding=ROUND_HALF_UP) if it['qty'] % Decimal(n_b) == 0 else (it['qty'] / Decimal(n_b))
            q_diff = it['qty'] - (q_per * Decimal(n_b))

            it_p_tot = b_peso * (it['qty'] / tot_q)
            p_per = (it_p_tot / Decimal(n_b)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            p_diff = it_p_tot - (p_per * Decimal(n_b))

            for b_idx, b_num in enumerate(b_list):
                raw_items.append({
                    'nb': b_num,
                    'code': it['code'],
                    'detalle': it['detalle'],
                    'uni': it['uni'],
                    'marca': it['marca'],
                    'qty': q_per + (q_diff if b_idx == n_b - 1 else Decimal(0)),
                    'peso': p_per + (p_diff if b_idx == n_b - 1 else Decimal(0)),
                    'cbm': Decimal(0)
                })

    if hdr_peso > 0 and raw_items:
        sum_p = sum(it['peso'] for it in raw_items)
        diff_p = hdr_peso - sum_p
        raw_items[-1]['peso'] += diff_p

    return [{
        'proveedor': proveedor,
        'n_fact': n_fact,
        'cliente': cliente,
        'items': raw_items
    }]

def parse_docx(docx_path):
    """Detecta el proveedor del documento Word .docx y ejecuta el parser correspondiente."""
    return parse_aisin_docx(docx_path)

def group_invoice_items(items):
    """
    Agrupa repuestos dentro del mismo bulto por su tipo de repuesto (NB, DETALLE),
    sin importar que pertenezcan a distintos fabricantes/marcas.

    Reglas:
    1. Si el repuesto es el mismo en el mismo bulto, se unen en una sola fila.
    2. Se toma el código del repuesto que tenga la mayor cantidad.
    3. Se conserva la descripción (y marca) del repuesto con mayor cantidad.
    4. Las cantidades, pesos y CBM se suman con cuadre exacto.
    """
    groups = {}
    for it in items:
        key = (it['nb'], it['detalle'])
        if key not in groups:
            groups[key] = []
        groups[key].append(it)

    result = []
    for (nb, detalle), group_items in groups.items():
        # Ítem dominante con mayor cantidad (si hay empate en cantidad, prefiere con código o más específico)
        dominant = max(group_items, key=lambda x: (x['qty'], 1 if x.get('code') else 0, len(str(x.get('detalle', '')))))

        tot_qty = sum((it['qty'] for it in group_items), Decimal(0))
        tot_peso = sum((it['peso'] for it in group_items), Decimal(0))
        tot_cbm = sum((it['cbm'] for it in group_items), Decimal(0))

        result.append({
            'nb': nb,
            'code': dominant['code'],
            'detalle': dominant['detalle'],
            'uni': dominant['uni'],
            'marca': dominant['marca'],
            'qty': tot_qty,
            'peso': tot_peso,
            'cbm': tot_cbm
        })

    return sorted(result, key=lambda x: (x['nb'], x['detalle']))

def process_and_generate_markdown(invoice_data):
    """
    Agrupa por (NB, DETALLE, MARCA), calcula totales y genera la tabla Markdown pura.
    """
    proveedor = invoice_data['proveedor']
    n_fact = invoice_data['n_fact']
    cliente = invoice_data['cliente']
    sorted_rows = group_invoice_items(invoice_data['items'])
    if not sorted_rows:
        return f"### FACTURA {n_fact}\n*Sin ítems de empaque.*"

    # Regla 5: Totales y cuadre perfecto
    total_qty = sum((r['qty'] for r in sorted_rows), Decimal(0))
    total_peso = sum((r['peso'] for r in sorted_rows), Decimal(0))
    total_cbm = sum((r['cbm'] for r in sorted_rows), Decimal(0))
    last_nb = max(r['nb'] for r in sorted_rows) if sorted_rows else 0

    # Construcción de la tabla Markdown pura (11 columnas)
    lines = []
    lines.append(f"### FACTURA {n_fact}")
    lines.append("| PROVEEDOR | N° FACT | CLIENTE | NB | CODIGO | DETALLE | UNI | MARCA | CANTIDA | PESO | CBM |")
    lines.append("|---|---|---|---|---|---|---|---|---|---|---|")

    for r in sorted_rows:
        q_str = format_number_es(r['qty'])
        p_str = format_number_es(r['peso'].quantize(Decimal('0.01'), rounding=ROUND_HALF_UP), 2)
        c_str = format_number_es(r['cbm'].quantize(Decimal('0.0001'), rounding=ROUND_HALF_UP), 4)
        lines.append(f"| {proveedor} | {n_fact} | {cliente} | {r['nb']} | {r['code']} | {r['detalle']} | {r['uni']} | {r['marca']} | {q_str} | {p_str} | {c_str} |")

    # Fila de totales en negrita
    tot_q_str = format_number_es(total_qty)
    tot_p_str = format_number_es(total_peso.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP), 2)
    tot_c_str = format_number_es(total_cbm.quantize(Decimal('0.0001'), rounding=ROUND_HALF_UP), 4)
    lines.append(f"| **{proveedor}** | **{n_fact}** | **{cliente}** | **{last_nb}** | | | | | **{tot_q_str}** | **{tot_p_str}** | **{tot_c_str}** |")

    return "\n".join(lines)

def deduplicate_invoices(invoices):
    """
    Si existen facturas duplicadas (mismo proveedor y n_fact),
    conserva la que tenga mayor cantidad de bultos o pesos
    (por ejemplo, si se cargó la Lista de Empaque y la Factura comercial del mismo número).
    """
    merged = {}
    for inv in invoices:
        key = (inv.get('proveedor', ''), str(inv.get('n_fact', '')).strip())
        if key not in merged:
            merged[key] = inv
        else:
            curr_weight = sum(it.get('peso', Decimal(0)) for it in inv.get('items', []))
            prev_weight = sum(it.get('peso', Decimal(0)) for it in merged[key].get('items', []))
            if curr_weight > prev_weight or len(inv.get('items', [])) > len(merged[key].get('items', [])):
                merged[key] = inv
    return list(merged.values())

def export_invoices_to_excel(invoices_data, output_excel_path):
    """
    Genera un archivo Excel consolidado (.xlsx) con todas las facturas procesadas,
    siguiendo exactamente la estructura de 'ejmplo de resultado esperado.xlsx'.
    """
    invoices_data = deduplicate_invoices(invoices_data)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CONSOLIDADO"

    font_title = Font(name="Calibri", size=11, bold=True)
    font_header = Font(name="Calibri", size=11, bold=True)
    font_data = Font(name="Calibri", size=11)
    font_total = Font(name="Calibri", size=11, bold=True)

    fill_header = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_total = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    border_thin = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )

    current_row = 1

    for inv in invoices_data:
        proveedor = inv['proveedor']
        n_fact = inv['n_fact']
        cliente = inv['cliente']
        sorted_rows = group_invoice_items(inv['items'])

        # Fila 1 de la factura: Titulo (FACTURA [N°])
        ws.cell(row=current_row, column=1, value=f"FACTURA {n_fact}").font = font_title
        current_row += 1

        # Fila 2 de la factura: Cabeceras de columnas
        headers = ['PROVEEDOR', 'N° FACT', 'CLIENTE', 'NB', 'CODIGO', 'DETALLE', 'UNI', 'MARCA', 'CANTIDAD', 'PESO', 'CBM']
        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=h)
            cell.font = font_header
            cell.alignment = align_center
            cell.fill = fill_header
            cell.border = border_thin
        current_row += 1

        start_data_row = current_row
        # Filas de datos
        for r in sorted_rows:
            vals = [
                (proveedor, align_left, "@"),
                (n_fact, align_center, "@"),
                (cliente, align_left, "@"),
                (r['nb'], align_center, "0"),
                (r['code'], align_left, "@"),
                (r['detalle'], align_left, "@"),
                (r['uni'], align_center, "@"),
                (r['marca'], align_left, "@"),
                (float(r['qty']), align_right, "#,##0"),
                (float(r['peso']), align_right, "#,##0.00"),
                (float(r['cbm']), align_right, "#,##0.0000")
            ]
            for col_idx, (v, al, num_fmt) in enumerate(vals, start=1):
                cell = ws.cell(row=current_row, column=col_idx, value=v)
                cell.font = font_data
                cell.alignment = al
                cell.number_format = num_fmt
                cell.border = border_thin
            current_row += 1
        end_data_row = current_row - 1

        # Fila de totales en negrita con fórmulas dinámicas de Excel
        if end_data_row >= start_data_row:
            formula_nb = f"=MAX(D{start_data_row}:D{end_data_row})"
            formula_qty = f"=SUM(I{start_data_row}:I{end_data_row})"
            formula_peso = f"=SUM(J{start_data_row}:J{end_data_row})"
            formula_cbm = f"=SUM(K{start_data_row}:K{end_data_row})"
        else:
            formula_nb = 0
            formula_qty = 0
            formula_peso = 0
            formula_cbm = 0

        tot_vals = [
            (proveedor, align_left, "@"),
            (n_fact, align_center, "@"),
            (cliente, align_left, "@"),
            (formula_nb, align_center, "0"),
            ("", align_left, "@"),
            ("", align_left, "@"),
            ("", align_center, "@"),
            ("", align_left, "@"),
            (formula_qty, align_right, "#,##0"),
            (formula_peso, align_right, "#,##0.00"),
            (formula_cbm, align_right, "#,##0.0000")
        ]
        for col_idx, (v, al, num_fmt) in enumerate(tot_vals, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=v if v != "" else None)
            cell.font = font_total
            cell.alignment = al
            cell.number_format = num_fmt
            cell.fill = fill_total
            cell.border = border_thin
        current_row += 1

        # Fila en blanco entre facturas
        current_row += 1

    # Ajuste dinámico de columnas
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

    os.makedirs(os.path.dirname(output_excel_path), exist_ok=True)
    wb.save(output_excel_path)
    return output_excel_path

def process_file(file_path):
    """Detecta la extensión y procesa el archivo."""
    if not os.path.exists(file_path):
        print(f"Error: El archivo no existe: {file_path}", file=sys.stderr)
        return None

    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.pdf':
        invoices = parse_pdf(file_path)
    elif ext in ['.xlsx', '.xls']:
        invoices = parse_excel(file_path)
    elif ext == '.docx':
        invoices = parse_docx(file_path)
    else:
        print(f"Error: Formato no soportado ({ext}). Debe ser PDF, Word (.docx) o Excel (.xlsx/.xls).", file=sys.stderr)
        return None

    results = []
    for inv in invoices:
        md = process_and_generate_markdown(inv)
        results.append(md)

    flush_dictionary()
    return "\n\n".join(results)

def process_folder(input_dir, output_dir):
    """
    Procesa todos los archivos PDF, Word y Excel en la carpeta de Entrada,
    genera los archivos Markdown individuales y crea el CONSOLIDADO_FACTURAS.xlsx con todos.
    """
    os.makedirs(output_dir, exist_ok=True)
    supported_exts = ('.pdf', '.xlsx', '.xls', '.docx')
    files = [f for f in os.listdir(input_dir) if f.lower().endswith(supported_exts) and not f.startswith("~$")]

    if not files:
        print(f"No se encontraron archivos PDF, Word o Excel en: {input_dir}")
        return

    print(f"Encontrados {len(files)} archivos para procesar en: {input_dir}\n")
    all_invoices = []

    for f in sorted(files):
        in_path = os.path.join(input_dir, f)
        base_name = os.path.splitext(f)[0]
        out_path = os.path.join(output_dir, f"{base_name}_resultado.md")

        print(f"-> Procesando: {f} ...", end=" ", flush=True)
        try:
            ext = os.path.splitext(f)[1].lower()
            if ext == '.pdf':
                invs = parse_pdf(in_path)
            elif ext == '.docx':
                invs = parse_docx(in_path)
            else:
                invs = parse_excel(in_path)

            all_invoices.extend(invs)

            # Generar Markdown individual
            md_list = [process_and_generate_markdown(inv) for inv in invs]
            with open(out_path, "w", encoding="utf-8") as out_f:
                out_f.write("\n\n".join(md_list))
            print(f"[OK] Markdown guardado")
        except Exception as e:
            print(f"[ERROR]: {e}")

    flush_dictionary()

    # Generar el Excel consolidado con TODOS los archivos
    if all_invoices:
        consolidated_excel = os.path.join(output_dir, "CONSOLIDADO_FACTURAS.xlsx")
        export_invoices_to_excel(all_invoices, consolidated_excel)
        print(f"\n========================================================")
        print(f"  EXCEL CONSOLIDADO GENERADO CON ÉXITO:")
        print(f"  {consolidated_excel}")
        print(f"  Total de facturas consolidadas: {len(all_invoices)}")
        print(f"========================================================")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso:")
        print("  1. Archivo individual: python procesador_facturas.py <ruta_archivo>")
        print("  2. Carpeta completa:   python procesador_facturas.py --carpeta <carpeta_entrada> <carpeta_salida>")
        sys.exit(1)

    if sys.argv[1] == "--carpeta":
        in_d = sys.argv[2] if len(sys.argv) > 2 else r"C:\Facturas_Repuestos\Entrada"
        out_d = sys.argv[3] if len(sys.argv) > 3 else r"C:\Facturas_Repuestos\Salida"
        process_folder(in_d, out_d)
    else:
        target_file = sys.argv[1]
        ext = os.path.splitext(target_file)[1].lower()
        if ext == '.pdf':
            invs = parse_pdf(target_file)
        elif ext in ['.xlsx', '.xls']:
            invs = parse_excel(target_file)
        elif ext == '.docx':
            invs = parse_docx(target_file)
        else:
            print(f"Error: Formato no soportado ({ext}). Debe ser PDF, Word (.docx) o Excel.", file=sys.stderr)
            sys.exit(1)

        if invs:
            base_dir = os.path.dirname(target_file)
            parent_dir = os.path.dirname(base_dir)
            salida_candidate = os.path.join(parent_dir, "Salida")
            out_dir = salida_candidate if os.path.exists(salida_candidate) else base_dir
            excel_filename = os.path.splitext(os.path.basename(target_file))[0] + "_resultado.xlsx"
            out_excel_path = os.path.join(out_dir, excel_filename)
            export_invoices_to_excel(invs, out_excel_path)

            result = "\n\n".join([process_and_generate_markdown(inv) for inv in invs])
            print(result)
