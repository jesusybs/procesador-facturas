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

def save_learned_mapping(raw_key, standardized_value):
    """Guarda un nuevo repuesto aprendido en el archivo JSON."""
    global _DICTIONARY_CACHE
    d = load_dictionary()
    if "learned_mappings" not in d:
        d["learned_mappings"] = {}

    d["learned_mappings"][raw_key] = standardized_value

    try:
        with open(DICT_PATH, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
        print(f"[AUTO-APRENDIZAJE] Nuevo repuesto aprendido y guardado: '{raw_key}' -> '{standardized_value}'")
    except Exception as e:
        print(f"[ERROR] Al guardar repuesto aprendido: {e}", file=sys.stderr)

# Patrones y reglas automotrices generales para estandarización instantánea
GENERAL_AUTO_PATTERNS = [
    (r'\bAMORT(?:IGUADOR)?\s+DEL\b', 'AMORTIGUADOR DELANTERO'),
    (r'\bAMORT(?:IGUADOR)?\s+TRAS\b', 'AMORTIGUADOR TRASERO'),
    (r'\bKIT[\.\s]+PIN\s+BUJES\b', 'KIT PIN BUJES'),
    (r'\bHUB\s+DEL\b', 'HUB DELANTERO'),
    (r'\bHUB\s+TRAS\b', 'HUB TRASERO'),
    (r'\bHUB\s+BAL\b', 'HUB CON BALINERA'),
    (r'\b(YOYOS\s+)?CINTA\s+(?:DE\s+)?TIMON\b', 'CINTA DE TIMON'),
    (r'\bTERM(?:INAL)?\.?\s*(?:EXT|EX)\b', 'TERMINAL EXTERIOR'),
    (r'\bTERM(?:INAL)?\.?\s*CREM\b', 'TERMINAL DE CREMALLERA'),
    (r'\bTERM(?:INAL)?\.?\s*INT\b', 'TERMINAL INTERIOR'),
    (r'\bTIJERA\s+TRAS\b', 'MESETA/TIJERA TRASERA'),
    (r'\bTIJERA\b', 'MESETA/TIJERA'),
    (r'\bBOMBA[\.\s]+GAS\b', 'BOMBA DE GASOLINA'),
    (r'\bBOMBA[\.\s]+FRENO\b', 'BOMBA FRENO'),
    (r'\bCREMALLERA\b', 'CREMALLERA DE DIRECCION'),
    (r'\bPUNTA[\.\s]+FLECHA\b', 'PUNTA FLECHA'),
    (r'\bDISCO[\.\s]+FRENO\b', 'DISCO DE FRENO'),
    (r'\bDISCO\b', 'DISCO DE FRENO'),
    (r'\bTAMBOR\s+(?:DE\s+)?FRENO\b', 'TAMBOR DE FRENO'),
    (r'\bCRUCETA\b', 'CRUCETA'),
    (r'\bESPIRAL\b', 'ESPIRAL / RESORTE'),
    (r'\bBASE\s+(?:AMORT|MOTOR)?\b', 'BASE/SOPORTE MOTOR'),
    (r'\bBOLA\b', 'ROTULA DE SUSPENSION'),
    (r'\bCALIPER\b', 'CALIPER DE FRENO'),
    (r'\bMU[Ñ\uFFFD]EQUILLA\b', 'TERMINAL ESTABILIZADOR'),
    (r'\bBAL[\.\s]+AMORT\b', 'BALINERA/RODAMIENTO AMORTIGUADOR DELANTERO'),
    (r'\bBAL(?:INERA)?\s+DEL\b', 'BALINERA DELANTERA'),
    (r'\bBOTA\s+FLECHA\b', 'GUARDAPOLVO FLECHA'),
    (r'\bBUJE\s+BARRA\b', 'BUJE DE BARRA ESTABILIZADORA'),
    (r'\bCIL(?:INDRO)?[\.\s]+FRENO\b', 'CILINDRO DE FRENO'),
    (r'\bFREE\s+WHEEL\b', 'CUBO MANUAL / FREE WHEEL'),
    (r'\bLINK\b', 'TERMINAL ESTABILIZADOR'),
    (r'\bRETEN\s+RDA\b', 'RETENEDOR')
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
    
    # Guardar en memoria y persistir en JSON para las próximas facturas
    save_learned_mapping(d_upper, standardized)
    
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
    - PERFECT TRADING
    """
    try:
        import pypdf
        reader = pypdf.PdfReader(pdf_path)
        if reader.pages:
            p0 = (reader.pages[0].extract_text() or "").upper()
            if "STAR AUTO PARTS" in p0 or "1158018-1-573227" in p0 or ("ZONA LIBRE DE COL" in p0 and "STAR" in p0):
                return "STAR_AUTO_PARTS"
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

def parse_pdf(pdf_path):
    """Detecta automáticamente el proveedor del PDF y ejecuta el parser correspondiente."""
    prov = detect_provider(pdf_path)
    if prov == "STAR_AUTO_PARTS":
        return parse_star_auto_parts(pdf_path)
    else:
        return parse_perfect_trading(pdf_path)

def parse_excel(excel_path):
    """Extrae las facturas y repuestos desde un archivo Excel."""
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

def group_invoice_items(items):
    """Agrupa por (NB, DETALLE, MARCA) y conserva el primer código."""
    groups = {}
    for it in items:
        key = (it['nb'], it['detalle'], it['marca'])
        if key not in groups:
            groups[key] = {
                'nb': it['nb'],
                'code': it['code'],
                'detalle': it['detalle'],
                'uni': it['uni'],
                'marca': it['marca'],
                'qty': Decimal(0),
                'peso': Decimal(0),
                'cbm': Decimal(0)
            }
        groups[key]['qty'] += it['qty']
        groups[key]['peso'] += it['peso']
        groups[key]['cbm'] += it['cbm']

    return sorted(groups.values(), key=lambda x: (x['nb'], x['detalle'], x['marca']))

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
    else:
        print(f"Error: Formato no soportado ({ext}). Debe ser PDF o Excel (.xlsx/.xls).", file=sys.stderr)
        return None

    results = []
    for inv in invoices:
        md = process_and_generate_markdown(inv)
        results.append(md)

    return "\n\n".join(results)

def process_folder(input_dir, output_dir):
    """
    Procesa todos los archivos PDF y Excel en la carpeta de Entrada,
    genera los archivos Markdown individuales y crea el CONSOLIDADO_FACTURAS.xlsx con todos.
    """
    os.makedirs(output_dir, exist_ok=True)
    supported_exts = ('.pdf', '.xlsx', '.xls')
    files = [f for f in os.listdir(input_dir) if f.lower().endswith(supported_exts) and not f.startswith("~$")]

    if not files:
        print(f"No se encontraron archivos PDF o Excel en: {input_dir}")
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
        else:
            print(f"Error: Formato no soportado ({ext}). Debe ser PDF o Excel.", file=sys.stderr)
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
