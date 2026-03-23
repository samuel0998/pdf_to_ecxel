import os
import re
import uuid
from flask import Flask, request, jsonify, send_file, render_template
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = '/tmp/uploads'
app.config['DOWNLOAD_FOLDER'] = '/tmp/downloads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)


def parse_pdf(pdf_path):
    """Parse the PDF and extract order items."""
    items = []
    current_product = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            lines = text.split('\n')
            for line in lines:
                line = line.strip()
                if not line:
                    continue

                # Detect product header line
                prod_match = re.match(
                    r'Produto:\s*(\S+)\s*-\s*(.+)', line)
                if prod_match:
                    current_product = {
                        'code': prod_match.group(1).strip(),
                        'description': prod_match.group(2).strip()
                    }
                    continue

                # Skip header/footer lines
                if any(x in line for x in [
                    'Nro. Pedido', 'Emissão', 'Produto:', 'Total do Ítem',
                    'Total Geral', 'Impresso em', 'Pag.:', 'Período',
                    'Relação de Pedidos'
                ]):
                    # Capture "Total Geral"
                    tg = re.search(r'Total Geral\s*[:\s]*([\d.,]+)', line)
                    if tg:
                        items.append({
                            'order_number': '',
                            'date': '',
                            'client': 'TOTAL GERAL',
                            'product_code': '',
                            'product_description': '',
                            'qty': '',
                            'unit_price': '',
                            'total': tg.group(1).replace('.', '').replace(',', '.')
                        })
                    continue

                if current_product is None:
                    continue

                # Try to match order line:
                # order_num  date  client  qty  unit  total
                order_match = re.match(
                    r'^(\d{5,})\s+'                         # order number
                    r'(\d{2}/\d{2}/\d{4})\s+'              # date
                    r'(.+?)\s+'                             # client name
                    r'([\d]+,\d{2})\s+'                    # qty
                    r'([\d]+,\d{2,})\s+'                   # unit price
                    r'([\d.,]+)$',                         # total
                    line
                )
                if order_match:
                    def br_to_float(v):
                        return v.replace('.', '').replace(',', '.')

                    items.append({
                        'order_number': order_match.group(1),
                        'date': order_match.group(2),
                        'client': order_match.group(3).strip(),
                        'product_code': current_product['code'],
                        'product_description': current_product['description'],
                        'qty': float(br_to_float(order_match.group(4))),
                        'unit_price': float(br_to_float(order_match.group(5))),
                        'total': float(br_to_float(order_match.group(6)))
                    })

    return items


def build_excel(items, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Pedidos"

    # ── Styles ──────────────────────────────────────────────
    header_font   = Font(name='Arial', bold=True, color='FFFFFF', size=10)
    header_fill   = PatternFill('solid', start_color='1F4E79')
    product_font  = Font(name='Arial', bold=True, color='FFFFFF', size=9)
    product_fill  = PatternFill('solid', start_color='2E75B6')
    subtotal_font = Font(name='Arial', bold=True, size=9)
    subtotal_fill = PatternFill('solid', start_color='D6E4F0')
    total_font    = Font(name='Arial', bold=True, color='FFFFFF', size=10)
    total_fill    = PatternFill('solid', start_color='1F4E79')
    data_font     = Font(name='Arial', size=9)
    center_align  = Alignment(horizontal='center', vertical='center')
    left_align    = Alignment(horizontal='left',   vertical='center')
    right_align   = Alignment(horizontal='right',  vertical='center')

    thin = Side(style='thin', color='BFBFBF')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Column setup ────────────────────────────────────────
    headers = ['Nº Pedido', 'Emissão', 'Cliente',
               'Cód. Produto', 'Descrição do Produto',
               'Quantidade', 'Preço Unit.', 'Total']
    col_widths = [12, 13, 40, 14, 50, 12, 13, 15]

    for i, (h, w) in enumerate(zip(headers, col_widths), 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── Title row ───────────────────────────────────────────
    ws.merge_cells('A1:H1')
    title_cell = ws['A1']
    title_cell.value = 'RELAÇÃO DE PEDIDOS - ÍTENS'
    title_cell.font  = Font(name='Arial', bold=True, color='FFFFFF', size=13)
    title_cell.fill  = PatternFill('solid', start_color='1F4E79')
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    # Subtitle
    ws.merge_cells('A2:H2')
    ws['A2'].value = 'Período: 01/02/2026 a 03/03/2026'
    ws['A2'].font  = Font(name='Arial', bold=True, size=10, color='1F4E79')
    ws['A2'].alignment = Alignment(horizontal='center')
    ws.row_dimensions[2].height = 18

    # ── Column headers ──────────────────────────────────────
    ws.row_dimensions[3].height = 20
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = center_align
        cell.border    = border

    # ── Data rows ───────────────────────────────────────────
    row = 4
    alt = False

    # Group items by product
    product_groups = {}
    for item in items:
        if item['client'] == 'TOTAL GERAL':
            continue
        key = item['product_code'] + ' - ' + item['product_description']
        product_groups.setdefault(key, []).append(item)

    data_font_alt_fill = PatternFill('solid', start_color='EBF3FB')

    for prod_key, prod_items in product_groups.items():
        # Product header row
        ws.merge_cells(f'A{row}:H{row}')
        ph = ws[f'A{row}']
        ph.value     = prod_key
        ph.font      = product_font
        ph.fill      = product_fill
        ph.alignment = left_align
        ph.border    = border
        ws.row_dimensions[row].height = 16
        row += 1

        # Order rows
        qty_start = row
        for item in prod_items:
            alt = not alt
            row_fill = data_font_alt_fill if alt else PatternFill('solid', start_color='FFFFFF')
            row_data = [
                item['order_number'], item['date'], item['client'],
                item['product_code'], item['product_description'],
                item['qty'], item['unit_price'], item['total']
            ]
            for col, val in enumerate(row_data, 1):
                cell = ws.cell(row=row, column=col, value=val)
                cell.font   = data_font
                cell.fill   = row_fill
                cell.border = border
                if col in (6, 7, 8):
                    cell.alignment = right_align
                    if col in (7, 8):
                        cell.number_format = '#,##0.00'
                    else:
                        cell.number_format = '#,##0.00'
                elif col in (1, 4):
                    cell.alignment = center_align
                else:
                    cell.alignment = left_align
            ws.row_dimensions[row].height = 15
            row += 1

        # Subtotal row
        qty_end = row - 1
        st_cells = [None, None, None, None, 'Total do Item:',
                    f'=SUM(F{qty_start}:F{qty_end})',
                    None,
                    f'=SUM(H{qty_start}:H{qty_end})']
        for col, val in enumerate(st_cells, 1):
            if val is None:
                continue
            cell = ws.cell(row=row, column=col, value=val)
            cell.font      = subtotal_font
            cell.fill      = subtotal_fill
            cell.alignment = right_align if col in (6, 7, 8) else left_align
            cell.border    = border
            if col in (6, 8):
                cell.number_format = '#,##0.00'
        ws.row_dimensions[row].height = 15
        row += 1

    # ── Grand Total row ─────────────────────────────────────
    ws.merge_cells(f'A{row}:E{row}')
    ws[f'A{row}'].value = 'TOTAL GERAL'
    ws[f'A{row}'].font  = total_font
    ws[f'A{row}'].fill  = total_fill
    ws[f'A{row}'].alignment = right_align
    ws[f'A{row}'].border = border

    for col in (6, 7, 8):
        cell = ws.cell(row=row, column=col)
        if col == 6:
            cell.value = f'=SUM(F4:F{row-1})'
        elif col == 8:
            cell.value = f'=SUM(H4:H{row-1})'
        cell.font   = total_font
        cell.fill   = total_fill
        cell.alignment = right_align
        cell.border = border
        cell.number_format = '#,##0.00'
    ws.row_dimensions[row].height = 22

    # Freeze panes
    ws.freeze_panes = 'A4'

    wb.save(output_path)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado'}), 400

    f = request.files['file']
    if not f.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Apenas arquivos PDF são aceitos'}), 400

    uid = str(uuid.uuid4())[:8]
    pdf_path  = os.path.join(app.config['UPLOAD_FOLDER'], f'{uid}.pdf')
    xlsx_path = os.path.join(app.config['DOWNLOAD_FOLDER'], f'pedidos_{uid}.xlsx')

    f.save(pdf_path)

    try:
        items = parse_pdf(pdf_path)
        if not items:
            return jsonify({'error': 'Nenhum dado encontrado no PDF'}), 400

        build_excel(items, xlsx_path)

        # Count stats
        data_rows  = [i for i in items if i['client'] != 'TOTAL GERAL']
        products   = len(set(i['product_code'] for i in data_rows))
        clients    = len(set(i['client'] for i in data_rows))
        total_val  = sum(i['total'] for i in data_rows)

        return jsonify({
            'success': True,
            'filename': f'pedidos_{uid}.xlsx',
            'stats': {
                'rows': len(data_rows),
                'products': products,
                'clients': clients,
                'total': f'R$ {total_val:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
            }
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        if os.path.exists(pdf_path):
            os.remove(pdf_path)


@app.route('/download/<filename>')
def download(filename):
    path = os.path.join(app.config['DOWNLOAD_FOLDER'], filename)
    if not os.path.exists(path):
        return jsonify({'error': 'Arquivo não encontrado'}), 404
    return send_file(path, as_attachment=True,
                     download_name='Relacao_Pedidos.xlsx')


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
