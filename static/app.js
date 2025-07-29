class SpreadsheetUI {
  
  initColumnResizing() {
    this.thead.querySelectorAll('th').forEach((th, colIdx) => {
      if (colIdx === 0) return; // skip corner cell
  
      const resizer = document.createElement('div');
      resizer.className = 'col-resizer';
      th.appendChild(resizer);
  
      let startX, startWidth;
  
      resizer.addEventListener('mousedown', (e) => {
        e.preventDefault();
        startX = e.pageX;
        startWidth = th.offsetWidth;
  
        const onMouseMove = (eMove) => {
          const diff = eMove.pageX - startX;
          const newWidth = Math.max(50, startWidth + diff);
          th.style.width = newWidth + 'px';
  
          // update all cells in this column
          this.tbody.querySelectorAll(`td[data-col='${colIdx-1}']`).forEach(td => {
            td.style.width = newWidth + 'px';
          });
        };
  
        const onMouseUp = () => {
          document.removeEventListener('mousemove', onMouseMove);
          document.removeEventListener('mouseup', onMouseUp);
        };
  
        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
      });
    });
  }
  
  constructor(tableElem) {
    this.table = tableElem;
    this.thead = tableElem.querySelector('thead tr');
    this.tbody = tableElem.querySelector('tbody');
    this.COLS = 20;
    this.visibleRows = 40;
    this.startRow = 0;
    this.selectedCells = [];
    this.lastSelected = null;
    this.isLoading = false;
    this.initHeaders();
    this.fetchAndRenderGrid();
    this.initEvents();
    this.initRibbon();
    this.initInfiniteScroll();
    this.isDragging = false;
    this.startCell = null;
    this.selectionBox = null;
    this.searchMatches = [];
    this.currentMatchIdx = 0;
  }

  initHeaders() {
    this.thead.innerHTML = '';
    const corner = document.createElement('th');
    corner.className = 'corner';
    this.thead.appendChild(corner);
    for (let c = 0; c < this.COLS; c++) {
      const th = document.createElement('th');
      th.textContent = String.fromCharCode(65 + c);
      this.thead.appendChild(th);
    }
  }

  async fetchAndRenderGrid() {
    this.isLoading = true;
    const res = await fetch(`/api/grid?start_row=${this.startRow}&end_row=${this.startRow + this.visibleRows}`);
    const data = await res.json();
    this.renderGrid(data.grid, data.rows, data.cols);
    this.isLoading = false;
  }

  renderGrid(grid, totalRows, totalCols) {
    this.tbody.innerHTML = '';
    for (let r = 0; r < grid.length; r++) {
      const tr = document.createElement('tr');
      const rowHeader = document.createElement('th');
      rowHeader.textContent = this.startRow + r + 1;
      tr.appendChild(rowHeader);
      for (let c = 0; c < totalCols; c++) {
        const cellData = grid[r][c];
        const td = document.createElement('td');
        td.tabIndex = 0;
        td.dataset.row = this.startRow + r;
        td.dataset.col = c;
        td.textContent = cellData.value;
        td.style.fontSize = cellData.font_size + 'px';
        td.style.fontWeight = cellData.bold ? 'bold' : 'normal';
        td.style.fontStyle = cellData.italic ? 'italic' : 'normal';
        td.classList.add('text-wrap');
        // Apply cell background and text color if set
        if (cellData.bg_color) {
          td.style.backgroundColor = cellData.bg_color;
        }
        if (cellData.text_color) {
          td.style.color = cellData.text_color;
        }
        tr.appendChild(td);
      }
      this.tbody.appendChild(tr);
    }
    // After rendering, re-apply search highlights if needed
    if (this.lastSearchTerm) {
      this.highlightSearchMatches(this.lastSearchTerm);
    }
  }

  clearSelection() {
    this.selectedCells.forEach(cell => {
      cell.classList.remove('selected', 'selected-multi');
      cell.style.width = ''; // ✅ reset for each cell
      if (this.selectionBox) {
        this.selectionBox.remove();
        this.selectionBox = null;
      }
      
    });
    this.selectedCells = [];
  }
  
initColumnResizing() {
  this.thead.querySelectorAll('th').forEach((th, colIdx) => {
    if (colIdx === 0) return; // skip corner cell

    const resizer = document.createElement('div');
    resizer.className = 'col-resizer';
    th.appendChild(resizer);

    let startX, startWidth;

    resizer.addEventListener('mousedown', (e) => {
      e.preventDefault();
      startX = e.pageX;
      startWidth = th.offsetWidth;

      const onMouseMove = (eMove) => {
        const diff = eMove.pageX - startX;
        const newWidth = Math.max(50, startWidth + diff);
        th.style.width = newWidth + 'px';

        // update all cells in this column
        this.tbody.querySelectorAll(`td[data-col='${colIdx-1}']`).forEach(td => {
          td.style.width = newWidth + 'px';
        });
      };

      const onMouseUp = () => {
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);
      };

      document.addEventListener('mousemove', onMouseMove);
      document.addEventListener('mouseup', onMouseUp);
    });
  });
}

  selectCell(cell, multi = false) {
    if (!multi) this.clearSelection();

if (!this.selectedCells.includes(cell)) {
  cell.classList.add(multi ? 'selected-multi' : 'selected');

  // 👇 measure text content
  const contentWidth = this.measureContentWidth(cell.textContent);
  cell.style.width = Math.max(180, contentWidth + 20) + 'px'; // 20px padding buffer

  this.selectedCells.push(cell);
}

this.lastSelected = cell;
cell.focus();

  }
  measureContentWidth(text) {
    const span = document.createElement('span');
    span.style.visibility = 'hidden';
    span.style.position = 'absolute';
    span.style.whiteSpace = 'nowrap';
    span.style.font = 'inherit';
    span.textContent = text || '';
    document.body.appendChild(span);
    const width = span.offsetWidth;
    document.body.removeChild(span);
    return width;
  }
  
  selectRange(cell1, cell2) {
    this.clearSelection();

    const r1 = Math.min(+cell1.dataset.row, +cell2.dataset.row);
    const r2 = Math.max(+cell1.dataset.row, +cell2.dataset.row);
    const col = +cell1.dataset.col;  // only take column of start cell

    for (let r = r1; r <= r2; r++) {
        const cell = this.tbody.querySelector(`td[data-row='${r}'][data-col='${col}']`);
        if (cell) {
            cell.classList.add('selected-multi');
            this.selectedCells.push(cell);
        }
    }

    // draw selection box stays as-is 👇
    const rect1 = cell1.getBoundingClientRect();
    const rect2 = cell2.getBoundingClientRect();
    const container = this.table.parentElement;
    const containerRect = container.getBoundingClientRect();

    const scrollTop = container.scrollTop;
    const scrollLeft = container.scrollLeft;

    const top = Math.min(rect1.top, rect2.top) - containerRect.top + scrollTop;
    const left = rect1.left - containerRect.left + scrollLeft;
    const bottom = Math.max(rect1.bottom, rect2.bottom) - containerRect.top + scrollTop;
    const right = rect1.right - containerRect.left + scrollLeft;

    if (!this.selectionBox) {
        this.selectionBox = document.createElement('div');
        this.selectionBox.className = 'selection-box';
        this.table.parentElement.appendChild(this.selectionBox);
    }

    this.selectionBox.style.top = `${top}px`;
    this.selectionBox.style.left = `${left}px`;
    this.selectionBox.style.width = `${right - left}px`;
    this.selectionBox.style.height = `${bottom - top}px`;
}


async updateCell(cell, updates) {
  const newValue = updates.value;

  if (typeof newValue === 'string' && newValue.startsWith('=')) {
    // check if it's direct math (like =A1+B1) or range-based
    if (/^=[A-Z]+\d+\s*[\+\-\*/]\s*[A-Z]+\d+$/.test(newValue.trim())) {
        this.handleDirectFormula(cell);
    } else {
        await this.handleFormula(cell, newValue);
    }
    return;
}


  const row = +cell.dataset.row;
  const col = +cell.dataset.col;
  const payload = { row, col, ...updates };

  const res = await fetch('/api/cell', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
  });
  const data = await res.json();

  cell.textContent = data.value;
  cell.style.fontSize = data.font_size + 'px';
  cell.style.fontWeight = data.bold ? 'bold' : 'normal';
  cell.style.fontStyle = data.italic ? 'italic' : 'normal';
}

handleDirectFormula(cell) {
  const formula = cell.textContent.trim();
  if (!formula.startsWith('=')) return;

  const match = formula.match(/^=([A-Z]+[0-9]+)\s*([\+\-\*/])\s*([A-Z]+[0-9]+)$/);
  if (!match) {
    console.warn('Invalid formula format:', formula);
    return;
  }

  const [, ref1, operator, ref2] = match;
  const cell1 = this.getCellByRef(ref1);
  const cell2 = this.getCellByRef(ref2);

  if (!cell1 || !cell2) {
    console.warn('Invalid cell references:', ref1, ref2);
    return;
  }

  const val1 = parseFloat(cell1.textContent) || 0;
  const val2 = parseFloat(cell2.textContent) || 0;

  let result = 0;
  switch (operator) {
    case '+': result = val1 + val2; break;
    case '-': result = val1 - val2; break;
    case '*': result = val1 * val2; break;
    case '/': result = val2 !== 0 ? val1 / val2 : 'DIV/0'; break;
  }

  cell.textContent = result;
}
getCellByRef(ref) {
  const colLetter = ref.match(/[A-Z]+/)[0];
  const rowNumber = parseInt(ref.match(/[0-9]+/)[0], 10) - 1;

  const colIndex = colLetter.charCodeAt(0) - 65; // A → 0
  return this.tbody.querySelector(`td[data-row='${rowNumber}'][data-col='${colIndex}']`);
}

  initEvents() {
    this.tbody.addEventListener('mousedown', (e) => {
      if (e.target.tagName === 'TD') {
        this.clearSelection();
        this.isDragging = true;
        document.body.classList.add('dragging');
        this.startCell = e.target;
        this.selectCell(e.target);
      }
    });
    this.tbody.addEventListener('mouseover', (e) => {
      if (this.isDragging && e.target.tagName === 'TD') {
        this.selectRange(this.startCell, e.target);
      }
    });
    document.addEventListener('mouseup', () => {
      this.isDragging = false;
      this.startCell = null;
      document.body.classList.remove('dragging');

    });
            
    // Double-click: make cell editable and show tooltip
    this.tbody.addEventListener('dblclick', (e) => {
      if (e.target.tagName === 'TD') {
        const cell = e.target;
        const oldValue = cell.textContent;
        cell.contentEditable = true;
        cell.focus();
        
        const finishEdit = async () => {
          cell.contentEditable = false;
          if (cell.textContent !== oldValue) {
            await this.updateCell(cell, { value: cell.textContent });
          }
        };
        cell.onblur = finishEdit;
        cell.onkeydown = (ev) => {
          if (ev.key === 'Enter') {
            ev.preventDefault();
            cell.blur();
          }
        };
      }
    });
    this.tbody.addEventListener('focusin', (e) => {
      if (e.target.tagName === 'TD') {
        this.selectCell(e.target);
      }
    });
    this.tbody.addEventListener('keydown', (e) => {
      if (e.target.tagName === 'TD') {
        let row = +e.target.dataset.row;
        let col = +e.target.dataset.col;
        let next;
        if (e.key === 'ArrowDown') next = this.tbody.querySelector(`td[data-row='${row+1}'][data-col='${col}']`);
        if (e.key === 'ArrowUp') next = this.tbody.querySelector(`td[data-row='${row-1}'][data-col='${col}']`);
        if (e.key === 'ArrowRight') next = this.tbody.querySelector(`td[data-row='${row}'][data-col='${col+1}']`);
        if (e.key === 'ArrowLeft') next = this.tbody.querySelector(`td[data-row='${row}'][data-col='${col-1}']`);
        if (next) {
          next.focus();
          this.selectCell(next);
        }
      }
    });
 
    
  }
  initClipboard() {
    document.addEventListener('keydown', async (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
            e.preventDefault();

            if (!this.selectedCells.length) return;

            // sort cells by row/col
            const cells = this.selectedCells.slice().sort((a, b) => {
                const rDiff = +a.dataset.row - +b.dataset.row;
                if (rDiff !== 0) return rDiff;
                return +a.dataset.col - +b.dataset.col;
            });

            const rowsMap = new Map();

            for (const cell of cells) {
                const r = +cell.dataset.row;
                const c = +cell.dataset.col;
                if (!rowsMap.has(r)) rowsMap.set(r, []);
                rowsMap.get(r)[c] = cell.textContent.trim();
            }

            const output = Array.from(rowsMap.values())
                .map(row => row.join('\t'))
                .join('\n');

            await navigator.clipboard.writeText(output);

            console.log('Copied to clipboard:\n', output);
        }
    });
}



  initRibbon() {
    // Ribbon button interactivity
    const ribbonBtns = document.querySelectorAll('.ribbon-btn');
    ribbonBtns.forEach(btn => {
      btn.addEventListener('click', () => {
        if (!btn.classList.contains('font-btn') && !btn.classList.contains('style-btn')) {
          ribbonBtns.forEach(b => b.classList.remove('active'));
          btn.classList.add('active');
        }
      });
    });
    // Font size
    document.getElementById('font-increase').addEventListener('click', async () => {
      for (const cell of this.selectedCells) {
        const size = parseFloat(window.getComputedStyle(cell).fontSize);
        await this.updateCell(cell, { font_size: size + 2 });
      }
    });
    document.getElementById('font-decrease').addEventListener('click', async () => {
      for (const cell of this.selectedCells) {
        const size = parseFloat(window.getComputedStyle(cell).fontSize);
        await this.updateCell(cell, { font_size: Math.max(8, size - 2) });
      }
    });
    // Style
    document.getElementById('bold-btn').addEventListener('click', async () => {
      for (const cell of this.selectedCells) {
        const isBold = window.getComputedStyle(cell).fontWeight === '700' || window.getComputedStyle(cell).fontWeight === 'bold';
        await this.updateCell(cell, { bold: !isBold });
      }
    });
    document.getElementById('italic-btn').addEventListener('click', async () => {
      for (const cell of this.selectedCells) {
        const isItalic = window.getComputedStyle(cell).fontStyle === 'italic';
        await this.updateCell(cell, { italic: !isItalic });
      }
    });
    // Fill color logic
    const fillColorBtn = document.getElementById('fill-color-btn');
    const fillColorInput = document.getElementById('fill-color-input');
    fillColorBtn.addEventListener('click', () => fillColorInput.click());
    fillColorInput.addEventListener('input', async (e) => {
      const color = e.target.value;
      for (const cell of this.selectedCells) {
        cell.style.backgroundColor = color;
        await this.updateCell(cell, { bg_color: color });
      }
    });
    // Text color logic
    const textColorBtn = document.getElementById('text-color-btn');
    const textColorInput = document.getElementById('text-color-input');
    textColorBtn.addEventListener('click', () => textColorInput.click());
    textColorInput.addEventListener('input', async (e) => {
      const color = e.target.value;
      for (const cell of this.selectedCells) {
        cell.style.color = color;
        await this.updateCell(cell, { text_color: color });
      }
    });
    // Accessibility: keyboard navigation for ribbon
    const searchBar = document.querySelector('.search-bar');
    // Accessibility: keyboard navigation for ribbon
    searchBar.addEventListener('keydown', (e) => {
      if (e.key === 'Tab' && !e.shiftKey) {
        e.preventDefault();
        // Focus first cell
        const firstCell = this.tbody.querySelector('td[data-row="0"][data-col="0"]');
        if (firstCell) firstCell.focus();
      }
      // Search logic
      if (e.key === 'Enter') {
        const term = searchBar.value.trim();
        this.lastSearchTerm = term;
        this.highlightSearchMatches(term);
      } else if (e.key === 'Escape') {
        searchBar.value = '';
        this.lastSearchTerm = '';
        this.highlightSearchMatches('');
      }
    });
    searchBar.addEventListener('input', (e) => {
      if (!searchBar.value.trim()) {
        this.lastSearchTerm = '';
        this.highlightSearchMatches('');
      }
    });
    // Import CSV logic
    const importBtn = document.getElementById('import-btn');
    const importFile = document.getElementById('import-file');
    importBtn.addEventListener('click', () => importFile.click());
    importFile.addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      const formData = new FormData();
      formData.append('file', file);
      const res = await fetch('/api/import', {
        method: 'POST',
        body: formData
      });
      if (res.ok) {
        // Reset scroll and visible rows to show imported data
        this.startRow = 0;
        this.visibleRows = 40;
        await this.fetchAndRenderGrid();
      } else {
        alert('Import failed.');
      }
      importFile.value = '';
    });
    // Export CSV logic
    const exportBtn = document.getElementById('export-btn');
    exportBtn.addEventListener('click', async () => {
      const now = new Date();
      const pad = n => n.toString().padStart(2, '0');
      const filename = `SP_${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
      const res = await fetch('/api/export');
      if (res.ok) {
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename + '.csv';
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
      } else {
        alert('Export failed.');
      }
    });
  }
  initPaste() {
    document.addEventListener('paste', async (e) => {
        if (!this.lastSelected) return;

        e.preventDefault();
        const text = (e.clipboardData || window.clipboardData).getData('text');

        const rows = text.trim().split(/\r?\n/).map(row => row.split('\t'));

        let startRow = +this.lastSelected.dataset.row;
        let startCol = +this.lastSelected.dataset.col;

        for (let r = 0; r < rows.length; r++) {
            const rowData = rows[r];
            for (let c = 0; c < rowData.length; c++) {
                const cell = this.tbody.querySelector(
                    `td[data-row='${startRow + r}'][data-col='${startCol + c}']`
                );
                if (cell) {
                    cell.textContent = rowData[c];
                    this.updateCell(cell, { value: rowData[c] });
                }
            }
        }
    });
}

  initInfiniteScroll() {
    const container = this.table.closest('.spreadsheet-container');
    container.addEventListener('scroll', async () => {
      if (this.isLoading) return;
      if (container.scrollTop + container.clientHeight >= container.scrollHeight - 50) {
        // Ask backend to add more rows
        await fetch('/api/rows', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ count: 20 })
        });
        this.visibleRows += 20;
        await this.fetchAndRenderGrid();
      }
    });
  }
  async handleFormula(cell, formula) {
    try {
        const expression = formula.slice(1).toUpperCase(); // strip '='
        let result = '';

        if (expression.startsWith('SUM(')) {
            result = this.computeRange(expression, 'SUM');
        } else if (expression.startsWith('AVERAGE(')) {
            result = this.computeRange(expression, 'AVERAGE');
        } else if (expression.startsWith('MIN(')) {
            result = this.computeRange(expression, 'MIN');
        } else if (expression.startsWith('MAX(')) {
            result = this.computeRange(expression, 'MAX');
        } else if (expression.startsWith('COUNT(')) {
            result = this.computeRange(expression, 'COUNT');
        } else {
            // new dropdown functions
            result = this.computeFunction(expression);
        }

        cell.textContent = result;

        await this.updateCell(cell, { value: result, formula });

    } catch (err) {
        console.error('Formula error:', err);
        cell.textContent = 'ERR';
    }
}

computeFunction(expression) {
  try {
      if (expression.startsWith('IF(')) {
          const args = this.parseArgs(expression.slice(3, -1));
          const [condition, valTrue, valFalse] = args;
          const condEval = this.evalCondition(condition);
          return condEval ? valTrue : valFalse;
      }

      if (expression.startsWith('CONCATENATE(')) {
          const args = this.parseArgs(expression.slice(12, -1));
          return args.join('');
      }

      if (expression.startsWith('LEFT(')) {
          const args = this.parseArgs(expression.slice(5, -1));
          return args[0].substring(0, parseInt(args[1], 10));
      }

      if (expression.startsWith('RIGHT(')) {
          const args = this.parseArgs(expression.slice(6, -1));
          const text = args[0];
          const num = parseInt(args[1], 10);
          return text.substring(text.length - num);
      }

      if (expression.startsWith('LEN(')) {
          const args = this.parseArgs(expression.slice(4, -1));
          return args[0].length;
      }

      if (expression.startsWith('ROUND(')) {
          const args = this.parseArgs(expression.slice(6, -1));
          return Number(parseFloat(args[0]).toFixed(parseInt(args[1], 10)));
      }

      return 'ERR';
  } catch {
      return 'ERR';
  }
}
parseArgs(argStr) {
  return argStr.split(',').map(s => {
      const val = s.trim();
      if (/^[A-Z]+\d+$/.test(val)) {
          const cell = this.getCellByRef(val);
          return cell ? cell.textContent.trim() : '';
      }
      if (val.startsWith('"') && val.endsWith('"')) {
          return val.slice(1, -1);
      }
      return val;
  });
}
evalCondition(condStr) {
  const match = condStr.match(/([A-Z]+\d+|\d+)\s*(<=|>=|=|>|<)\s*([A-Z]+\d+|\d+)/);
  if (!match) return false;

  let [_, left, op, right] = match;

  if (/^[A-Z]+\d+$/.test(left)) {
      const cell = this.getCellByRef(left);
      left = cell ? parseFloat(cell.textContent) : 0;
  } else {
      left = parseFloat(left);
  }

  if (/^[A-Z]+\d+$/.test(right)) {
      const cell = this.getCellByRef(right);
      right = cell ? parseFloat(cell.textContent) : 0;
  } else {
      right = parseFloat(right);
  }

  switch (op) {
      case '=': return left === right;
      case '>': return left > right;
      case '<': return left < right;
      case '>=': return left >= right;
      case '<=': return left <= right;
      default: return false;
  }
}

// helper for range-based formulas
computeRange(expression, type) {
    const rangeMatch = expression.match(/\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)/);
    if (!rangeMatch) return 'ERR';

    const [_, colStart, rowStart, colEnd, rowEnd] = rangeMatch;
    const startCol = colStart.charCodeAt(0) - 65;
    const endCol = colEnd.charCodeAt(0) - 65;
    const startRow = parseInt(rowStart) - 1;
    const endRow = parseInt(rowEnd) - 1;

    const values = [];

    for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
            const td = this.tbody.querySelector(`td[data-row="${r}"][data-col="${c}"]`);
            if (td && !isNaN(parseFloat(td.textContent))) {
                values.push(parseFloat(td.textContent));
            }
        }
    }

    switch (type) {
        case 'SUM': return values.reduce((a, b) => a + b, 0);
        case 'AVERAGE': return values.length ? (values.reduce((a, b) => a + b, 0) / values.length).toFixed(2) : 0;
        case 'MIN': return values.length ? Math.min(...values) : 0;
        case 'MAX': return values.length ? Math.max(...values) : 0;
        case 'COUNT': return values.length;
        default: return 'ERR';
    }
}

  highlightSearchMatches(term) {
    // Remove previous highlights
    this.tbody.querySelectorAll('td.search-match').forEach(td => td.classList.remove('search-match'));
    this.searchMatches = [];
    this.currentMatchIdx = 0;
    if (!term) return;
    const lowerTerm = term.toLowerCase();
    const tds = this.tbody.querySelectorAll('td');
    for (const td of tds) {
      if (td.textContent.toLowerCase().includes(lowerTerm)) {
        td.classList.add('search-match');
        this.searchMatches.push(td);
      }
    }
    // Scroll to first match
    if (this.searchMatches.length > 0) {
      this.searchMatches[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
      this.searchMatches[0].focus();
    }
  }

  // Apply formula to selected cells
  async applyFormulaToSelectedCells(formula) {
    if (this.selectedCells.length === 0) {
      alert('Please select at least one cell first.');
      return;
    }

    for (const cell of this.selectedCells) {
      // Get the cell's position to create a range reference
      const row = +cell.dataset.row;
      const col = +cell.dataset.col;
      const colLetter = String.fromCharCode(65 + col);
      const cellRef = `${colLetter}${row + 1}`;

      // Create a range based on the selected cells
      const minRow = Math.min(...this.selectedCells.map(c => +c.dataset.row));
      const maxRow = Math.max(...this.selectedCells.map(c => +c.dataset.row));
      const minCol = Math.min(...this.selectedCells.map(c => +c.dataset.col));
      const maxCol = Math.max(...this.selectedCells.map(c => +c.dataset.col));

      const startColLetter = String.fromCharCode(65 + minCol);
      const endColLetter = String.fromCharCode(65 + maxCol);
      const range = `${startColLetter}${minRow + 1}:${endColLetter}${maxRow + 1}`;

      // Apply the formula with the range
      let appliedFormula = formula;
      if (formula.includes('A1:A10')) {
        appliedFormula = formula.replace('A1:A10', range);
      }

      cell.textContent = appliedFormula;
      await this.updateCell(cell, { value: appliedFormula });
    }
  }

  // Apply function to selected cells
  async applyFunctionToSelectedCells(functionName) {
    if (this.selectedCells.length === 0) {
      alert('Please select at least one cell first.');
      return;
    }

    for (const cell of this.selectedCells) {
      const row = +cell.dataset.row;
      const col = +cell.dataset.col;
      const colLetter = String.fromCharCode(65 + col);
      const cellRef = `${colLetter}${row + 1}`;

      // Create a range based on the selected cells
      const minRow = Math.min(...this.selectedCells.map(c => +c.dataset.row));
      const maxRow = Math.max(...this.selectedCells.map(c => +c.dataset.row));
      const minCol = Math.min(...this.selectedCells.map(c => +c.dataset.col));
      const maxCol = Math.max(...this.selectedCells.map(c => +c.dataset.col));

      const startColLetter = String.fromCharCode(65 + minCol);
      const endColLetter = String.fromCharCode(65 + maxCol);
      const range = `${startColLetter}${minRow + 1}:${endColLetter}${maxRow + 1}`;

      // Apply the function with appropriate parameters
      let appliedFunction = '';
      switch (functionName) {
        case 'IF':
          appliedFunction = `=IF(${cellRef}>0,"Positive","Negative")`;
          break;
        case 'CONCATENATE':
          appliedFunction = `=CONCATENATE(${cellRef},"_text")`;
          break;
        case 'LEFT':
          appliedFunction = `=LEFT(${cellRef},3)`;
          break;
        case 'RIGHT':
          appliedFunction = `=RIGHT(${cellRef},3)`;
          break;
        case 'LEN':
          appliedFunction = `=LEN(${cellRef})`;
          break;
        case 'ROUND':
          appliedFunction = `=ROUND(${cellRef},2)`;
          break;
        default:
          appliedFunction = `=${functionName}(${range})`;
      }

      cell.textContent = appliedFunction;
      await this.updateCell(cell, { value: appliedFunction });
    }
  }

}

document.addEventListener('DOMContentLoaded', () => {
  const ui = new SpreadsheetUI(document.getElementById('spreadsheet'));
  ui.initColumnResizing();
  ui.initClipboard();
  ui.initPaste();

  // Dropdown logic
  document.getElementById('formulas-btn').addEventListener('click', (e) => {
    const dropdown = document.getElementById('formulas-dropdown');
    const btnRect = e.target.getBoundingClientRect();
  
    dropdown.style.top = `${btnRect.bottom + window.scrollY}px`;
    dropdown.style.left = `${btnRect.left + window.scrollX}px`;
    dropdown.classList.toggle('show');
  
    document.getElementById('functions-dropdown').classList.remove('show');
  });
  
  document.getElementById('functions-btn').addEventListener('click', (e) => {
    const dropdown = document.getElementById('functions-dropdown');
    const btnRect = e.target.getBoundingClientRect();
  
    dropdown.style.top = `${btnRect.bottom + window.scrollY}px`;
    dropdown.style.left = `${btnRect.left + window.scrollX}px`;
    dropdown.classList.toggle('show');
  
    document.getElementById('formulas-dropdown').classList.remove('show');
  });

  // Formula dropdown item click handlers
  document.getElementById('formulas-dropdown').addEventListener('click', (e) => {
    if (e.target.tagName === 'DIV') {
      const formula = e.target.textContent;
      ui.applyFormulaToSelectedCells(formula);
      document.getElementById('formulas-dropdown').classList.remove('show');
    }
  });

  // Function dropdown item click handlers
  document.getElementById('functions-dropdown').addEventListener('click', (e) => {
    if (e.target.tagName === 'DIV') {
      const functionName = e.target.textContent;
      ui.applyFunctionToSelectedCells(functionName);
      document.getElementById('functions-dropdown').classList.remove('show');
    }
  });

  // Close dropdowns when clicking outside
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.ribbon-btn') && !e.target.closest('.dropdown')) {
      document.getElementById('formulas-dropdown').classList.remove('show');
      document.getElementById('functions-dropdown').classList.remove('show');
    }
  });
  });

