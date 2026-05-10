from flask import Flask, render_template, request, send_file, jsonify
import xlrd
import openpyxl
import io
import os
from datetime import datetime

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32MB max


def parse_cassette_balances(file_bytes):
    wb = xlrd.open_workbook(file_contents=file_bytes)
    ws = wb.sheets()[0]
    datemode = wb.datemode

    terminals = {}
    # Row 3 (index 3) = headers, data starts at row 4 (index 4)
    for row_idx in range(4, ws.nrows):
        row = ws.row_values(row_idx)
        terminal_id = str(row[0]).strip()

        if not terminal_id:
            continue
        # Skip footer rows like "Total Outstanding Cash Balance:"
        if 'Total' in terminal_id or 'Outstanding' in terminal_id:
            continue

        est_cash_out = None
        if isinstance(row[7], float) and row[7] > 40000:
            try:
                est_cash_out = xlrd.xldate_as_datetime(row[7], datemode)
            except Exception:
                est_cash_out = None

        terminals[terminal_id] = {
            'terminal_id': terminal_id,
            'site_name': str(row[2]).strip(),
            'cash_balance': row[4] if isinstance(row[4], (int, float)) else 0,
            'est_cash_out': est_cash_out,
            'last_error': str(row[8]).strip() if row[8] else '',
            'last_communication': str(row[9]).strip() if row[9] else '',
            'last_cash_wd': str(row[10]).strip() if row[10] else '',
            '_row_order': row_idx,
        }

    return terminals


def parse_cash_projection(file_bytes):
    wb = xlrd.open_workbook(file_contents=file_bytes)
    ws = wb.sheets()[0]

    projections = {}
    # Row 2 (index 2) = headers, data starts at row 3 (index 3)
    for row_idx in range(3, ws.nrows):
        row = ws.row_values(row_idx)
        terminal_id = str(row[0]).strip()

        if not terminal_id:
            continue

        projections[terminal_id] = {
            'days_until_load': int(row[6]) if isinstance(row[6], (int, float)) else None,
            'suggested_cash_add': int(row[8]) if isinstance(row[8], (int, float)) else None,
        }

    return projections


def get_city_lookup(wb):
    ws = wb['City']
    city_lookup = {}
    # Headers on row 3 (index 2 with min_row=3), data from row 4
    for row in ws.iter_rows(min_row=4, values_only=True):
        # Col A = blank, Col B (index 1) = Terminal ID, Col D (index 3) = City
        if row[1]:
            terminal_id = str(row[1]).strip()
            city = str(row[3]).strip() if row[3] else ''
            city_lookup[terminal_id] = city
    return city_lookup


def is_tulsa(city):
    return bool(city) and 'tulsa' in city.lower()


def clear_data_rows(ws, start_row, num_cols):
    if ws.max_row >= start_row:
        for row_idx in range(start_row, ws.max_row + 1):
            for col_idx in range(1, num_cols + 1):
                ws.cell(row=row_idx, column=col_idx).value = None


def write_sheet_rows(ws, rows, col_keys, start_row=3):
    clear_data_rows(ws, start_row, len(col_keys))
    for i, row_data in enumerate(rows):
        row_num = start_row + i
        for col_offset, key in enumerate(col_keys, start=1):
            ws.cell(row=row_num, column=col_offset).value = row_data.get(key, '')


def update_template(template_bytes, cassette_data, projection_data):
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes), keep_vba=True)

    city_lookup = get_city_lookup(wb)

    # Preserve original sort order from cassette report (cash balance ascending)
    sorted_terminals = sorted(cassette_data.values(), key=lambda x: x['_row_order'])

    all_rows = []
    for t in sorted_terminals:
        tid = t['terminal_id']
        proj = projection_data.get(tid, {})
        city = city_lookup.get(tid, '')

        all_rows.append({
            'terminal_id': tid,
            'site_name': t['site_name'],
            'city': city,
            'days_until_load': proj.get('days_until_load', ''),
            'suggested_cash_add': proj.get('suggested_cash_add', ''),
            'cash_remaining': t['cash_balance'],
            'last_transaction': t['last_cash_wd'],
            'est_cash_out': t['est_cash_out'],
            'last_error': t['last_error'],
            'last_communication': t['last_communication'],
        })

    cash_pos_cols = [
        'terminal_id', 'site_name', 'city',
        'days_until_load', 'suggested_cash_add',
        'cash_remaining', 'last_transaction',
    ]
    write_sheet_rows(wb['CashPosition'], all_rows, cash_pos_cols)

    tulsa_rows = [r for r in all_rows if is_tulsa(r['city'])]
    write_sheet_rows(wb['Tulsa'], tulsa_rows, cash_pos_cols)

    error_rep_cols = [
        'terminal_id', 'site_name', 'cash_remaining',
        'est_cash_out', 'last_error',
        'last_communication', 'last_transaction',
    ]
    write_sheet_rows(wb['Error Rep.'], all_rows, error_rep_cols)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, len(all_rows), len(tulsa_rows)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    try:
        cassette_file = request.files.get('cassette')
        projection_file = request.files.get('projection')
        template_file = request.files.get('template')

        missing = []
        if not cassette_file or not cassette_file.filename:
            missing.append('Cassette Balances report')
        if not projection_file or not projection_file.filename:
            missing.append('Cash Projection report')
        if not template_file or not template_file.filename:
            missing.append('Cash POS template')

        if missing:
            return jsonify({'error': f"Missing files: {', '.join(missing)}"}), 400

        cassette_data = parse_cassette_balances(cassette_file.read())
        projection_data = parse_cash_projection(projection_file.read())
        template_bytes = template_file.read()

        output, total_count, tulsa_count = update_template(
            template_bytes, cassette_data, projection_data
        )

        today = datetime.now().strftime('%m-%d-%Y')
        filename = f'Cash_Pos_{today}.xlsm'

        response = send_file(
            output,
            mimetype='application/vnd.ms-excel.sheet.macroEnabled.12',
            as_attachment=True,
            download_name=filename,
        )
        response.headers['X-Terminal-Count'] = str(total_count)
        response.headers['X-Tulsa-Count'] = str(tulsa_count)
        return response

    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f'\n  ATM Cash POS Generator')
    print(f'  Open http://localhost:{port} in your browser\n')
    app.run(host='0.0.0.0', port=port, debug=False)
