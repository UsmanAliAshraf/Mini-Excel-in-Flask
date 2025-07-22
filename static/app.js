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

        
    
        tr.appendChild(td);
      }
      this.tbody.appendChild(tr);
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
    // Accessibility: keyboard navigation for ribbon
    const searchBar = document.querySelector('.search-bar');
    searchBar.addEventListener('keydown', (e) => {
      if (e.key === 'Tab' && !e.shiftKey) {
        e.preventDefault();
        // Focus first cell
        const firstCell = this.tbody.querySelector('td[data-row="0"][data-col="0"]');
        if (firstCell) firstCell.focus();
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
}

document.addEventListener('DOMContentLoaded', () => {
  const ui = new SpreadsheetUI(document.getElementById('spreadsheet'));
  ui.initColumnResizing();
  ui.initClipboard();
  ui.initPaste();

});
