#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
Servidor Web Local para la Interfaz Gráfica de Procesamiento de Facturas
Ubicación: C:\Facturas_Repuestos\app\app.py
"""

import os
import sys
from decimal import Decimal
import tempfile
import webbrowser
from threading import Timer

from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from werkzeug.exceptions import HTTPException

# Agregar directorio padre para importar procesador_facturas
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from procesador_facturas import (
    parse_pdf,
    parse_excel,
    group_invoice_items,
    export_invoices_to_excel,
    format_number_es,
    load_dictionary,
    save_learned_mapping,
    flush_dictionary,
    deduplicate_invoices
)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 64 * 1024 * 1024  # 64 MB max

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
OUTPUT_FOLDER = os.path.join(os.path.dirname(__file__), "..", "Salida")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Guardar la última ruta de excel generado en memoria de la sesión
LAST_EXCEL_PATH = os.path.join(OUTPUT_FOLDER, "CONSOLIDADO_FACTURAS.xlsx")
LAST_INVOICES_CACHE = []
LAST_FILES_INFO_CACHE = []

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/version')
def get_version():
    return jsonify({
        'version': '2.3.0',
        'build': 'batch-safe',
        'providers': ['PERFECT TRADING', 'STAR AUTO PARTS, S.A.', 'JAPAN INTERNATIONAL', 'ADK CORPORATION', 'NIKOMOTO, S.A.'],
        'status': 'online'
    })

@app.route('/api/clear', methods=['POST'])
def clear_session():
    global LAST_INVOICES_CACHE, LAST_FILES_INFO_CACHE
    LAST_INVOICES_CACHE = []
    LAST_FILES_INFO_CACHE = []
    return jsonify({'success': True})

@app.route('/api/update_item', methods=['POST'])
def update_item():
    """
    Actualiza la descripción de un repuesto en memoria y en el diccionario,
    y regenera el Excel consolidado con los cambios aplicados en tiempo real.
    """
    global LAST_INVOICES_CACHE, LAST_EXCEL_PATH
    data = request.get_json() or {}
    factura = str(data.get('factura', '')).strip()
    code = str(data.get('code', '')).strip()
    old_detalle = str(data.get('old_detalle', '')).strip()
    new_detalle = str(data.get('new_detalle', '')).strip().upper()

    if not new_detalle:
        return jsonify({'error': 'La descripción no puede estar vacía'}), 400

    # 1. Guardar en diccionario persistente para auto-aprendizaje
    if old_detalle and old_detalle != new_detalle:
        save_learned_mapping(old_detalle, new_detalle)

    # 2. Actualizar en la lista de facturas en memoria
    updated_count = 0
    for inv in LAST_INVOICES_CACHE:
        if not factura or str(inv.get('n_fact', '')).strip() == factura:
            for it in inv.get('items', []):
                if (code and it.get('code') == code) or (it.get('detalle') == old_detalle):
                    it['detalle'] = new_detalle
                    updated_count += 1

    # 3. Regenerar el Excel consolidado con los cambios
    if LAST_INVOICES_CACHE:
        output_excel = os.path.join(OUTPUT_FOLDER, "CONSOLIDADO_FACTURAS.xlsx")
        export_invoices_to_excel(LAST_INVOICES_CACHE, output_excel)
        LAST_EXCEL_PATH = output_excel

    return jsonify({
        'success': True,
        'updated_count': updated_count,
        'new_detalle': new_detalle,
        'learned': True
    })

@app.route('/api/process', methods=['POST'])
def process_files():
    global LAST_EXCEL_PATH, LAST_INVOICES_CACHE, LAST_FILES_INFO_CACHE
    
    append_mode = request.form.get('append', 'false').lower() == 'true'

    if 'files[]' not in request.files:
        return jsonify({'error': 'No se enviaron archivos'}), 400

    uploaded_files = request.files.getlist('files[]')
    if not uploaded_files or uploaded_files[0].filename == '':
        return jsonify({'error': 'No se seleccionó ningún archivo'}), 400

    new_invoices = []
    processed_files_info = []

    for file in uploaded_files:
        filename = secure_filename(file.filename)
        if not filename:
            continue

        ext = os.path.splitext(filename)[1].lower()
        if ext not in ['.pdf', '.xlsx', '.xls']:
            continue

        saved_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(saved_path)

        try:
            if ext == '.pdf':
                invs = parse_pdf(saved_path)
            else:
                invs = parse_excel(saved_path)
            
            new_invoices.extend(invs)
            processed_files_info.append({
                'name': filename,
                'invoices_count': len(invs),
                'status': 'success'
            })
        except Exception as e:
            processed_files_info.append({
                'name': filename,
                'status': 'error',
                'error': str(e)
            })

    if append_mode:
        combined = deduplicate_invoices(LAST_INVOICES_CACHE + new_invoices)
        LAST_FILES_INFO_CACHE.extend(processed_files_info)
    else:
        combined = deduplicate_invoices(new_invoices)
        LAST_FILES_INFO_CACHE = list(processed_files_info)

    all_invoices = [inv for inv in combined if inv.get('items')]
    if not all_invoices:
        return jsonify({'error': 'No se encontraron facturas o listas de empaque válidas para procesar.'}), 400

    LAST_INVOICES_CACHE = all_invoices
    flush_dictionary()

    # Generar el Excel consolidado con todas las facturas
    output_excel = os.path.join(OUTPUT_FOLDER, "CONSOLIDADO_FACTURAS.xlsx")
    export_invoices_to_excel(all_invoices, output_excel)
    LAST_EXCEL_PATH = output_excel

    # Preparar datos estructurados para la vista previa en la tabla web
    preview_data = []
    grand_total_qty = Decimal(0)
    grand_total_peso = Decimal(0)
    grand_total_cbm = Decimal(0)
    grand_total_bultos = 0

    for inv in all_invoices:
        prov = inv['proveedor']
        n_fact = inv['n_fact']
        cliente = inv['cliente']
        sorted_rows = group_invoice_items(inv['items'])

        inv_qty = sum(r['qty'] for r in sorted_rows)
        inv_peso = sum(r['peso'] for r in sorted_rows)
        inv_cbm = sum(r['cbm'] for r in sorted_rows)
        last_nb = max(r['nb'] for r in sorted_rows) if sorted_rows else 0

        grand_total_qty += inv_qty
        grand_total_peso += inv_peso
        grand_total_cbm += inv_cbm
        grand_total_bultos += last_nb

        inv_rows = []
        for r in sorted_rows:
            inv_rows.append({
                'proveedor': prov,
                'n_fact': n_fact,
                'cliente': cliente,
                'nb': r['nb'],
                'codigo': r['code'],
                'detalle': r['detalle'],
                'uni': r['uni'],
                'marca': r['marca'],
                'cantidad': format_number_es(r['qty']),
                'peso': format_number_es(r['peso'], 2),
                'cbm': format_number_es(r['cbm'], 4)
            })

        preview_data.append({
            'proveedor': prov,
            'n_fact': n_fact,
            'cliente': cliente,
            'rows': inv_rows,
            'total_nb': last_nb,
            'total_qty': format_number_es(inv_qty),
            'total_peso': format_number_es(inv_peso, 2),
            'total_cbm': format_number_es(inv_cbm, 4)
        })

    return jsonify({
        'success': True,
        'files_info': LAST_FILES_INFO_CACHE,
        'invoices': preview_data,
        'summary': {
            'total_facturas': len(all_invoices),
            'total_bultos': grand_total_bultos,
            'total_piezas': format_number_es(grand_total_qty),
            'total_peso': format_number_es(grand_total_peso, 2),
            'total_cbm': format_number_es(grand_total_cbm, 4)
        },
        'download_url': '/download'
    })

@app.errorhandler(Exception)
def handle_exception(e):
    if isinstance(e, HTTPException):
        return e
    import traceback
    traceback.print_exc()
    return jsonify({'error': f'Error en el servidor: {str(e)}'}), 500

@app.route('/api/dictionary', methods=['GET', 'POST'])
def manage_dictionary():
    if request.method == 'GET':
        d = load_dictionary()
        return jsonify(d)

    data = request.get_json()
    if not data or 'key' not in data or 'standard' not in data:
        return jsonify({'error': 'Parámetros inválidos'}), 400

    k = data['key'].strip().upper()
    val = data['standard'].strip().upper()
    save_learned_mapping(k, val)
    return jsonify({'success': True, 'key': k, 'standard': val})

@app.route('/download')
def download_excel():
    global LAST_EXCEL_PATH
    if os.path.exists(LAST_EXCEL_PATH):
        return send_file(
            LAST_EXCEL_PATH,
            as_attachment=True,
            download_name="CONSOLIDADO_FACTURAS.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    return "El archivo no se encuentra disponible", 404

def open_browser():
    webbrowser.open_new("http://127.0.0.1:5000")

if __name__ == '__main__':
    print("=" * 60)
    print("  SERVIDOR ACTIVO:")
    print("  -> Acceso Local:   http://localhost:5000")
    print("  -> Acceso en Red:  http://0.0.0.0:5000 (o tu IP local:5000)")
    print("=" * 60)
    Timer(1.2, open_browser).start()
    app.run(host='0.0.0.0', port=5000, debug=False)
