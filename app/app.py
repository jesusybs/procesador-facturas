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

# Agregar directorio padre para importar procesador_facturas
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from procesador_facturas import (
    parse_pdf,
    parse_excel,
    group_invoice_items,
    export_invoices_to_excel,
    format_number_es,
    load_dictionary,
    save_learned_mapping
)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 64 * 1024 * 1024  # 64 MB max

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
OUTPUT_FOLDER = os.path.join(os.path.dirname(__file__), "..", "Salida")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Guardar la última ruta de excel generado en memoria de la sesión
LAST_EXCEL_PATH = os.path.join(OUTPUT_FOLDER, "CONSOLIDADO_FACTURAS.xlsx")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/process', methods=['POST'])
def process_files():
    global LAST_EXCEL_PATH
    
    if 'files[]' not in request.files:
        return jsonify({'error': 'No se enviaron archivos'}), 400

    uploaded_files = request.files.getlist('files[]')
    if not uploaded_files or uploaded_files[0].filename == '':
        return jsonify({'error': 'No se seleccionó ningún archivo'}), 400

    all_invoices = []
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
            
            all_invoices.extend(invs)
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

    if not all_invoices:
        return jsonify({'error': 'No se pudo procesar ninguna factura de los archivos cargados.'}), 400

    # Generar el Excel consolidado
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
        'files_info': processed_files_info,
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
