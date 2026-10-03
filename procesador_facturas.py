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

    # 2. Eliminar dimensiones y especificaciones técnicas
    d = re.sub(r'\b\d+(\.\d+)?\s*(MM|CM|LBS|KG|KPA|T|V|CYL|L)\b', '', d)
    d = re.sub(r'\b\d+\s*[xX]\s*\d+(\s*[xX]\s*\d+)?\b', '', d)
    d = re.sub(r'\bØ\d+(\.\d+)?\s*(MM)?\b', '', d)
    d = re.sub(r'\b(4X4|4WD|AWD|2WD|16V|8V|24V|V6|V8|TURBO|DIESEL)\b', '', d)

    # 3. Eliminar patrones de años
    d = re.sub(r'\b\d{2}/\d{2}(/\d{2})?\b', '', d)
    d = re.sub(r'\b\d{2}/~', '', d)
    d = re.sub(r'\b\d{2}-\d{2}\b', '', d)
    d = re.sub(r'\b\d{4}/\d{4}\b', '', d)

    # 4. Eliminar marcas de vehículos
    d = re.sub(r'\b(TOYOTA|TOY|NISSAN|NIS|MITSUBISHI|MIT|HONDA|HON|SUZUKI|SUZ|HYUNDAI|HYU|CHEVROLET|CHEV|ISUZU|ISU|MAZDA|FORD|KIA|DAIHATSU|DAI|UNIVERSAL|JAC|RENAULT|REN)\b\.?', '', d)

    # 5. Eliminar modelos de vehículos comunes
    d = re.sub(r'\b(COROLLA|HILUX|YARIS|TERCEL|CAMRY|PRADO|4RUNNER|TACOMA|FORTUNER|HIACE|COASTER|RAV-4|RAV4|VITARA|CELERIO|SWIFT|ALTO|CIVIC|CR-V|CRV|ACCENT|ELANTRA|TUCSON|PICANTO|RIO|SPORTAGE|L200|MONTERO|NATIVA|LANCER|FRONTIER|PATHFINDER|TIIDA|ALTIMA|SENTRA|ESCAPE|RANGER|BT-50|AVEO|NEW SAIL|APV|KAVAK|MERU)\b', '', d)

    # 6. Eliminar códigos de motor (ej. 1KD, 2KD, 4AFE, 1ZZ, 2TR, 5VZ)
    d = re.sub(r'\b[0-9][A-Z]{1,3}(-?[A-Z0-9]+)?\b', '', d)

    # 7. Eliminar orientaciones redundantes
    d = re.sub(r'\b(RH/LH|LH/RH|AMBOS LADOS)\b', '', d)
    d = re.sub(r'\b(RH|LH)\b', '', d)

    # 8. Eliminar términos de relleno
    d = re.sub(r'\b(AT|MT|CERAMICA|SEMI-CERAMICA|ORIGINAL|GENUINO|OEM|REFORZADO)\b', '', d)

    # 9. Limpiar puntuación sobrante
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
    if cleaned.startswith("AMORTIGUADOR DEL"):
        return "AMORTIGUADOR DELANTERO"
    if cleaned.startswith("AMORTIGUADOR TRAS"):
        return "AMORTIGUADOR TRASERO"
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
    3. Si es un repuesto nuevo y desconocido, aplica auto-aprendizaje y lo guarda en el diccionario.
    """
    if not desc:
        return ""
    raw = str(desc).strip()
    d_upper = raw.upper()

    dictionary = load_dictionary()
    exact_map = dictionary.get("exact_mappings", {})
    learned_map = dictionary.get("learned_mappings", {})

    # 1. Revisar si la descripción completa ya fue aprendida
    if d_upper in learned_map:
        return learned_map[d_upper]

    # 2. Reglas fijas y prioritarias del diccionario
    # Se ordenan las claves por longitud descendente para que 'BALINERA DE CLUTCH' gane a 'BALINERA'
    sorted_keys = sorted(exact_map.keys(), key=len, reverse=True)
    for key in sorted_keys:
        if key in d_upper:
            return exact_map[key]

    # 3. REPUESTO NUEVO Y DESCONOCIDO: Motor de Auto-Estandarización Inteligente
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

def parse_pdf(pdf_path):
    """Extrae las facturas y repuestos desde un archivo PDF usando pdfplumber."""
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

    # Regla 5: Totales y cuadre perfecto
    total_qty = sum(r['qty'] for r in sorted_rows)
    total_peso = sum(r['peso'] for r in sorted_rows)
    total_cbm = sum(r['cbm'] for r in sorted_rows)
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

def export_invoices_to_excel(invoices_data, output_excel_path):
    """
    Genera un archivo Excel consolidado (.xlsx) con todas las facturas procesadas,
    siguiendo exactamente la estructura de 'ejmplo de resultado esperado.xlsx'.
    """
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

        # Fila de totales en negrita
        total_qty = float(sum(r['qty'] for r in sorted_rows))
        total_peso = float(sum(r['peso'] for r in sorted_rows))
        total_cbm = float(sum(r['cbm'] for r in sorted_rows))
        last_nb = max(r['nb'] for r in sorted_rows) if sorted_rows else 0

        tot_vals = [
            (proveedor, align_left, "@"),
            (n_fact, align_center, "@"),
            (cliente, align_left, "@"),
            (last_nb, align_center, "0"),
            ("", align_left, "@"),
            ("", align_left, "@"),
            ("", align_center, "@"),
            ("", align_left, "@"),
            (total_qty, align_right, "#,##0"),
            (total_peso, align_right, "#,##0.00"),
            (total_cbm, align_right, "#,##0.0000")
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
