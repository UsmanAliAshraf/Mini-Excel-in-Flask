# Mini Excel Flask App

A modern, feature-rich mini Excel clone built with Flask and vanilla JS. Supports import/export, infinite scroll, advanced formatting, formulas, search, and more—all in a beautiful green-black theme.

---

## 🚀 How to Run

1. **Install Flask:**
   ```bash
   pip install flask
   ```
2. **Run the app:**
   ```bash
   python app.py
   ```
3. **Open your browser:**
   [http://127.0.0.1:5000/](http://127.0.0.1:5000/)

---

## 📁 Project Structure

```
Mini-Excel-in-Flask/
├── app.py
├── spreadsheet/
│   ├── __init__.py
│   ├── cell.py
│   └── spreadsheet.py
├── static/
│   ├── styles.css
│   └── app.js
├── templates/
│   └── index.html
└── README.md
```

---

## ✨ Features Overview

| Feature           | Description                                                      |
| ----------------- | ---------------------------------------------------------------- |
| Infinite Scroll   | Seamlessly add and view more rows as you scroll.                 |
| Import CSV        | Import spreadsheet data from a CSV file.                         |
| Export CSV        | Export the entire sheet to a timestamped CSV file.               |
| Cell Formatting   | Bold, italic, font size, cell fill color, and text color.        |
| Multi-Selection   | Apply formatting and highlights to multiple cells at once.       |
| Formulas          | Excel-like formulas and functions (see below).                   |
| Clipboard Support | Copy/paste data between cells and from external sources.         |
| Search            | Instantly highlight all visible cells matching your search term. |
| Responsive UI     | Modern, green-black themed interface with smooth interactions.   |

---

## 🖱️ Ribbon & UI Controls

- **Formulas/Functions Dropdowns:** Quick insert for common formulas and functions.
- **Font Size:** Increase/decrease font size for selected cells.
- **Bold/Italic:** Toggle bold and italic styles.
- **Cell Fill Color:** Highlight cell background with any color.
- **Text Color:** Highlight cell text with any color.
- **Import CSV:** Upload a CSV file to populate the sheet.
- **Export CSV:** Download the current sheet as a CSV file (auto-named with timestamp).
- **Search Bar:** Type and press Enter to highlight all visible matches.

---

## 🧮 Supported Formulas & Functions

### **Basic Formulas**

- `=A1+A2` — Add values from two cells (also supports `-`, `*`, `/`).
- `=A1-B2`, `=A1*B2`, `=A1/B2`

### **Range Functions**

- `=SUM(A1:A10)` — Sum a range
- `=AVERAGE(A1:A10)` — Average
- `=MIN(A1:A10)` — Minimum
- `=MAX(A1:A10)` — Maximum
- `=COUNT(A1:A10)` — Count numeric values

### **Advanced Functions**

- `=IF(condition, value_if_true, value_if_false)`
  - Example: `=IF(A1>10, "Yes", "No")`
- `=CONCATENATE(A1, B1, "text")` — Join values
- `=LEFT(A1, n)` — First n chars
- `=RIGHT(A1, n)` — Last n chars
- `=LEN(A1)` — Length of value
- `=ROUND(A1, n)` — Round to n decimals

#### **Condition Syntax for IF**

- Supported: `=`, `>`, `<`, `>=`, `<=` (e.g., `A1>10`)

---

## 🎨 Formatting & Highlighting

- **Bold/Italic:** Toggle for any selection.
- **Font Size:** Increase/decrease for any selection.
- **Cell Fill Color:** Use the paintbrush button to pick a background color for any cell(s).
- **Text Color:** Use the "A" button to pick a text color for any cell(s).
- **Multi-Selection:** Click and drag or Ctrl+Click to select multiple cells and apply any formatting.

---

## 🔍 Search

- **How to Use:** Type in the search bar and press Enter.
- **What Happens:** All visible cells containing the search term (case-insensitive) are highlighted in yellow. The first match is scrolled into view.
- **Clear Search:** Press Escape or clear the search bar to remove highlights.

---

## 📥 Import & 📤 Export

- **Import CSV:** Click Import, select a `.csv` file. The sheet is populated with its data.
- **Export CSV:** Click Export. The current sheet is downloaded as `SP_YYYYMMDD_HHMMSS.csv`.
- **Data Persistence:** Data is stored in memory (not saved to disk). Import/export to save/load your work.

---

## 🖱️ Clipboard & Paste

- **Copy/Paste:** Use standard keyboard shortcuts or right-click to copy/paste data between cells or from external sources (e.g., Excel, Google Sheets).
- **Multi-Cell Paste:** Paste tabular data into a selected range.

---

## 🛠️ Technical Notes

- **Backend:** Python Flask, in-memory 2D grid, no database required.
- **Frontend:** Vanilla JS, modern CSS, responsive design.
- **Session:** Data is not persistent after server restart (for demo purposes).

---

## 📋 Credits & License

- Created by [Your Name].
- MIT License (or specify your own).

---

## 💡 Future Improvements

- Full-sheet backend search
- Persistent storage (file or database)
- Row/column insert/delete
- More Excel-like features (merge, freeze, etc.)
- User authentication

---

Enjoy your mini Excel experience! 🚀
