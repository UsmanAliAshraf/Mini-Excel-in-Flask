from flask import Flask, render_template, request, jsonify
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
    cell = spreadsheet.set_cell(row, col, value, font_size, bold, italic)
    return jsonify(cell.to_dict())

@app.route('/api/rows', methods=['POST'])
def add_rows():
    data = request.json
    count = data.get('count', 20)
    spreadsheet.add_rows(count)
    return jsonify({'rows': spreadsheet.rows})

if __name__ == '__main__':
    app.run(debug=True) 