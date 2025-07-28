from flask import Flask, render_template, request, jsonify, send_file
import csv
from io import StringIO, BytesIO
from spreadsheet import Spreadsheet

app = Flask(__name__, static_folder='static', template_folder='templates')

# Global spreadsheet instance (for demo; use session/db for real app)
spreadsheet = Spreadsheet()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/grid')
def get_grid():
    start_row = int(request.args.get('start_row', 0))
    end_row = int(request.args.get('end_row', start_row + 40))
    data = spreadsheet.to_dict(start_row, end_row)
    return jsonify(data)

@app.route('/api/cell', methods=['POST'])
def set_cell():
    data = request.json
    row = data['row']
    col = data['col']
    value = data.get('value')
    font_size = data.get('font_size')
    bold = data.get('bold')
    italic = data.get('italic')
    bg_color = data.get('bg_color')
    text_color = data.get('text_color')
    cell = spreadsheet.set_cell(row, col, value, font_size, bold, italic, bg_color, text_color)
    return jsonify(cell.to_dict())

@app.route('/api/rows', methods=['POST'])
def add_rows():
    data = request.json
    count = data.get('count', 20)
    spreadsheet.add_rows(count)
    return jsonify({'rows': spreadsheet.rows})

@app.route('/api/import', methods=['POST'])
def import_csv():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    if not file.filename.endswith('.csv'):
        return jsonify({'error': 'Invalid file type'}), 400
    stream = StringIO(file.stream.read().decode('utf-8'))
    reader = csv.reader(stream)
    data = list(reader)
    # Update spreadsheet grid
    num_rows = len(data)
    num_cols = max(len(row) for row in data) if data else 0
    spreadsheet._ensure_size(num_rows, num_cols)
    for r, row in enumerate(data):
        for c, value in enumerate(row):
            spreadsheet.set_cell(r, c, value=value)
    return jsonify({'status': 'success'})

@app.route('/api/export')
def export_csv():
    output = StringIO()
    writer = csv.writer(output)
    for row in spreadsheet.grid:
        writer.writerow([cell.value for cell in row])
    csv_data = output.getvalue().encode('utf-8')  # Convert to bytes
    output.close()
    return send_file(
        BytesIO(csv_data),
        mimetype='text/csv',
        as_attachment=True,
        download_name='spreadsheet.csv'
    )

if __name__ == '__main__':
    app.run(debug=True) 