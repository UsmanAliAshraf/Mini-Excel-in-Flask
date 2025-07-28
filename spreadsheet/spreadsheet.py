from .cell import Cell

class Spreadsheet:
    def __init__(self, cols=20, rows=20):
        self.cols = cols
        self.rows = rows
        self.grid = [[Cell() for _ in range(cols)] for _ in range(rows)]

    def get_cell(self, row, col):
        self._ensure_size(row+1, col+1)
        return self.grid[row][col]

    def set_cell(self, row, col, value=None, font_size=None, bold=None, italic=None, bg_color=None, text_color=None):
        # Used by import to set cell values in bulk and for highlight features
        self._ensure_size(row+1, col+1)
        cell = self.grid[row][col]
        if value is not None:
            cell.value = value
        if font_size is not None:
            cell.font_size = font_size
        if bold is not None:
            cell.bold = bold
        if italic is not None:
            cell.italic = italic
        if bg_color is not None:
            cell.bg_color = bg_color
        if text_color is not None:
            cell.text_color = text_color
        return cell

    def add_rows(self, count):
        for _ in range(count):
            self.grid.append([Cell() for _ in range(self.cols)])
        self.rows += count

    def _ensure_size(self, min_rows, min_cols):
        if min_rows > self.rows:
            self.add_rows(min_rows - self.rows)
        if min_cols > self.cols:
            for row in self.grid:
                row.extend([Cell() for _ in range(min_cols - self.cols)])
            self.cols = min_cols

    def to_dict(self, start_row=0, end_row=None):
        if end_row is None or end_row > self.rows:
            end_row = self.rows
        return {
            'cols': self.cols,
            'rows': self.rows,
            'grid': [
                [cell.to_dict() for cell in row]
                for row in self.grid[start_row:end_row]
            ]
        }
