(() => {
  'use strict';

  const PASS_KEY = '5921483';

  // Local storage settings
  const STORAGE_KEY = 'cm_kpi_dashboard_html_v1';
  let AUTO_SAVE = true;
  let saveTimer = null;

  // Edit/selection state
  let isUnlocked = false;
  let editMode = false;
  let multiSelectMode = false;
  let activeTable = null;

  // Drag-to-select state
  let dragging = false;
  let pendingDrag = false; // start dragging only after moving into another cell
  let anchorCell = null;
  let anchorTable = null;
  let anchorInfo = null;
  let dragMap = null;
  let ignoreNextClick = false;

  document.addEventListener('DOMContentLoaded', init);

  function init() {
    injectToolbar();
    ensureContentRoot();

    maybeRestoreSaved(); // Suggest restore from browser storage if found

    setEditable(false); // locked by default
    attachCellSelectionHandlers();
    attachDragSelectionHandlers();
    attachAutoSaveHandlers();
    updateButtonStates();
  }

  function ensureContentRoot() {
    // Wrap page content (except toolbar) into a single root for clean save/export
    let root = document.getElementById('cm-content-root');
    if (!root) {
      root = document.createElement('div');
      root.id = 'cm-content-root';
      const children = Array.from(document.body.children);
      for (const ch of children) {
        if (ch.id !== 'cm-editor-toolbar') {
          root.appendChild(ch);
        }
      }
      document.body.appendChild(root);
    }
    return root;
  }

  function getContentRoot() {
    return document.getElementById('cm-content-root') || ensureContentRoot();
  }

  function injectToolbar() {
    const bar = document.createElement('div');
    bar.id = 'cm-editor-toolbar';
    bar.innerHTML = `
      <div class="cm-controls">
        <button type="button" id="btnUnlock">Unlock editing</button>
        <button type="button" id="btnToggleEdit">Edit: OFF</button>
        <button type="button" id="btnMultiSelect">Multi-select: OFF</button>
        <button type="button" id="btnClearSel">Clear selection</button>
        <span class="sep"></span>
        <button type="button" id="btnAddRow">Add row</button>
        <button type="button" id="btnMerge">Merge</button>
        <button type="button" id="btnDeleteRow">Delete row</button>
        <button type="button" id="btnDeleteCell">Delete cell</button>
        <button type="button" id="btnDeleteCol">Delete column</button>
        <button type="button" id="btnAddTable">Add new table</button>
        <span class="sep"></span>
        <button type="button" id="btnDownloadHTML">Download HTML</button>
        <button type="button" id="btnDownloadExcel">Download Excel</button>
        <button type="button" id="btnDownloadPPT">Download PPT</button>
        <span class="sep"></span>
        <button type="button" id="btnSaveNow">Save to browser</button>
        <button type="button" id="btnLoadSaved">Restore saved</button>
        <button type="button" id="btnClearSaved">Clear saved</button>
        <button type="button" id="btnToggleAutosave">Auto-save: ON</button>
      </div>
    `;
    document.body.prepend(bar);

    // Toolbar actions
    document.getElementById('btnUnlock').addEventListener('click', toggleLock);
    document.getElementById('btnToggleEdit').addEventListener('click', () => {
      if (!requireUnlocked()) return;
      editMode = !editMode;
      setEditable(editMode);
      document.getElementById('btnToggleEdit').textContent = `Edit: ${editMode ? 'ON' : 'OFF'}`;
    });
    document.getElementById('btnMultiSelect').addEventListener('click', () => {
      if (!requireUnlocked()) return;
      multiSelectMode = !multiSelectMode;
      document.getElementById('btnMultiSelect').textContent = `Multi-select: ${multiSelectMode ? 'ON' : 'OFF'}`;
    });

    document.getElementById('btnClearSel').addEventListener('click', clearSelection);
    document.getElementById('btnAddRow').addEventListener('click', () => { addRowToActive(); scheduleAutoSave(); });
    document.getElementById('btnMerge').addEventListener('click', () => { mergeSelected(); scheduleAutoSave(); });
    document.getElementById('btnDeleteRow').addEventListener('click', () => { deleteSelectedRows(); scheduleAutoSave(); });
    document.getElementById('btnDeleteCell').addEventListener('click', () => { deleteSelectedCells(); scheduleAutoSave(); });
    document.getElementById('btnDeleteCol').addEventListener('click', () => { deleteSelectedColumns(); scheduleAutoSave(); });
    document.getElementById('btnAddTable').addEventListener('click', () => { addNewTable(); scheduleAutoSave(); });
    document.getElementById('btnDownloadHTML').addEventListener('click', downloadHTML);
    document.getElementById('btnDownloadExcel').addEventListener('click', downloadExcel);
    document.getElementById('btnDownloadPPT').addEventListener('click', downloadPPT);

    document.getElementById('btnSaveNow').addEventListener('click', persistToLocalStorage);
    document.getElementById('btnLoadSaved').addEventListener('click', restoreFromLocalStorage);
    document.getElementById('btnClearSaved').addEventListener('click', clearSavedFromLocalStorage);
    document.getElementById('btnToggleAutosave').addEventListener('click', () => {
      AUTO_SAVE = !AUTO_SAVE;
      document.getElementById('btnToggleAutosave').textContent = `Auto-save: ${AUTO_SAVE ? 'ON' : 'OFF'}`;
      if (AUTO_SAVE) scheduleAutoSave();
    });

    // Styles
    const style = document.createElement('style');
    style.textContent = `
      #cm-editor-toolbar { position: sticky; top: 0; z-index: 9999; background: #0d6efd; color: #fff; padding: 8px; border-bottom: 2px solid #084298; font-family: system-ui, Arial, sans-serif; }
      #cm-editor-toolbar button { margin: 2px 6px 2px 0; padding: 6px 10px; border: 0; border-radius: 4px; cursor: pointer; background: #fff; color: #0d6efd; font-weight: 600; }
      #cm-editor-toolbar button:hover { background: #e9f2ff; }
      #cm-editor-toolbar button:disabled { opacity: 0.5; cursor: not-allowed; }
      #cm-editor-toolbar .sep { display: inline-block; width: 8px; }
      td[contenteditable], th[contenteditable] { outline: none; }
      .cm-selected { outline: 2px solid #ff6b00 !important; background-color: #fff2e5; }
      table { margin-bottom: 16px; }
      body.cm-dragging { user-select: none !important; cursor: crosshair; }
    `;
    document.head.appendChild(style);
  }

  function toggleLock() {
    const btn = document.getElementById('btnUnlock');
    if (!isUnlocked) {
      const val = prompt('Enter pass key to enable editing:');
      if (val === PASS_KEY) {
        isUnlocked = true;
        editMode = true;
        setEditable(true);
        btn.textContent = 'Lock editing';
      } else {
        alert('Incorrect pass key.');
        return;
      }
    } else {
      if (!confirm('Lock editing and disable all edit actions?')) return;
      isUnlocked = false;
      editMode = false;
      setEditable(false);
      clearSelection();
      btn.textContent = 'Unlock editing';
    }
    updateButtonStates();
  }

  function requireUnlocked() {
    if (!isUnlocked) {
      alert('Unlock editing first (click "Unlock editing" and enter the pass key).');
      return false;
    }
    return true;
  }

  function updateButtonStates() {
    const setDisabled = (id, disabled) => {
      const el = document.getElementById(id);
      if (el) el.disabled = disabled;
    };
    const disabled = !isUnlocked;
    setDisabled('btnToggleEdit', disabled);
    setDisabled('btnMultiSelect', disabled);
    setDisabled('btnClearSel', disabled);
    setDisabled('btnAddRow', disabled);
    setDisabled('btnMerge', disabled);
    setDisabled('btnDeleteRow', disabled);
    setDisabled('btnDeleteCell', disabled);
    setDisabled('btnDeleteCol', disabled);
    setDisabled('btnAddTable', disabled);

    document.getElementById('btnToggleEdit').textContent = `Edit: ${editMode ? 'ON' : 'OFF'}`;
    document.getElementById('btnUnlock').textContent = isUnlocked ? 'Lock editing' : 'Unlock editing';
    document.getElementById('btnToggleAutosave').textContent = `Auto-save: ${AUTO_SAVE ? 'ON' : 'OFF'}`;
  }

  function setEditable(on) {
    const shouldEdit = on && isUnlocked;
    document.querySelectorAll('#cm-content-root table td, #cm-content-root table th').forEach(cell => {
      if (shouldEdit) cell.setAttribute('contenteditable', 'true');
      else cell.removeAttribute('contenteditable');
    });
  }

  // Click-based selection (works with typing)
  function attachCellSelectionHandlers() {
    document.addEventListener('click', (e) => {
      if (ignoreNextClick) { ignoreNextClick = false; return; }

      const cell = e.target.closest('td, th');
      if (!cell) return;

      activeTable = cell.closest('table');

      if (!isUnlocked) return;

      if (e.ctrlKey || e.metaKey || multiSelectMode) {
        cell.classList.toggle('cm-selected');
      } else {
        if (!cell.classList.contains('cm-selected')) {
          clearSelection();
        }
        cell.classList.add('cm-selected');
      }
    });
  }

  // Drag-to-select handlers (start drag only after moving into a different cell)
  function attachDragSelectionHandlers() {
    document.addEventListener('mousedown', (e) => {
      if (!isUnlocked || !editMode) return;
      if (e.button !== 0) return; // left click only
      const cell = e.target.closest('td, th');
      if (!cell) return;

      anchorTable = cell.closest('table');
      activeTable = anchorTable;
      anchorCell = cell;
      dragMap = buildCellMap(anchorTable);
      anchorInfo = dragMap.info.get(cell);
      if (!anchorInfo) return;

      pendingDrag = true;   // will switch to true dragging on movement into another cell
      dragging = false;
    });

    document.addEventListener('mousemove', (e) => {
      if (!pendingDrag) return;

      const target = document.elementFromPoint(e.clientX, e.clientY);
      const overCell = target && target.closest ? target.closest('td, th') : null;
      if (!overCell) return;
      if (!overCell.closest('table') || overCell.closest('table') !== anchorTable) return;

      // Start dragging only once we move into another cell
      if (!dragging && overCell !== anchorCell) {
        dragging = true;
        ignoreNextClick = true; // suppress click toggle that fires after drag
        document.body.classList.add('cm-dragging');
        clearSelectionWithinTable(anchorTable);
      }

      if (dragging) {
        updateDragSelectionOver(overCell);
      }
    });

    document.addEventListener('mouseup', () => {
      if (dragging) {
        dragging = false;
        document.body.classList.remove('cm-dragging');
      }
      pendingDrag = false;
    });
  }

  function updateDragSelectionOver(overCell) {
    if (!dragMap || !anchorInfo) return;
    const overInfo = dragMap.info.get(overCell);
    if (!overInfo) return;

    const r1 = Math.min(anchorInfo.row, overInfo.row);
    const r2 = Math.max(anchorInfo.row, overInfo.row);
    const c1 = Math.min(anchorInfo.col, overInfo.col);
    const c2 = Math.max(anchorInfo.col, overInfo.col);

    clearSelectionWithinTable(anchorTable);

    dragMap.info.forEach((info, cell) => {
      if (info.row >= r1 && info.row <= r2 && info.col >= c1 && info.col <= c2) {
        cell.classList.add('cm-selected');
      }
    });
  }

  function clearSelection() {
    document.querySelectorAll('.cm-selected').forEach(el => el.classList.remove('cm-selected'));
  }
  function clearSelectionWithinTable(table) {
    table.querySelectorAll('.cm-selected').forEach(el => el.classList.remove('cm-selected'));
  }

  function getSelectedCells() {
    return Array.from(document.querySelectorAll('td.cm-selected, th.cm-selected'));
  }

  function getActiveTable() {
    return activeTable || document.querySelector('#cm-content-root table');
  }

  function addRowToActive() {
    if (!requireUnlocked()) return;
    const table = getActiveTable();
    if (!table) return alert('No table found.');
    const meta = buildCellMap(table);
    const colCount = Math.max(meta.cols, 1);
    const tbody = table.tBodies[0] || table.createTBody();
    const row = tbody.insertRow(-1);
    for (let i = 0; i < colCount; i++) {
      const td = row.insertCell(-1);
      td.textContent = '-';
      if (isUnlocked && editMode) td.setAttribute('contenteditable', 'true');
    }
    row.scrollIntoView({ behavior: 'smooth', block: 'center' });
    clearSelection();
    row.cells[0]?.classList.add('cm-selected');
  }

  // Build a cell map that accounts for colspan and rowspan.
  function buildCellMap(table) {
    const rows = table.rows;
    const matrix = [];
    const info = new Map();
    let maxCols = 0;

    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (!matrix[r]) matrix[r] = [];
      let c = 0;

      for (let k = 0; k < row.cells.length; k++) {
        const cell = row.cells[k];
        while (matrix[r][c]) c++; // skip slots occupied by rowspans from above
        const cs = cell.colSpan || 1;
        const rs = cell.rowSpan || 1;

        if (!info.has(cell)) info.set(cell, { row: r, col: c, colSpan: cs, rowSpan: rs });

        for (let rr = r; rr < r + rs; rr++) {
          if (!matrix[rr]) matrix[rr] = [];
          for (let cc = c; cc < c + cs; cc++) {
            matrix[rr][cc] = cell;
          }
        }
        c += cs;
      }
      if (matrix[r].length > maxCols) maxCols = matrix[r].length;
    }

    return { matrix, info, cols: maxCols, rows: rows.length };
  }

  function mergeSelected() {
    if (!requireUnlocked()) return;
    const cells = getSelectedCells();
    if (cells.length < 2) {
      alert('Select at least two cells (drag across cells, or Ctrl/⌘-click, or turn ON Multi-select).');
      return;
    }
    const table = cells[0].closest('table');
    if (!cells.every(c => c.closest('table') === table)) {
      alert('Selected cells must be in the same table.');
      return;
    }
    if (cells.some(c => (c.colSpan || 1) > 1 || (c.rowSpan || 1) > 1)) {
      alert('Please select only plain cells (not already merged).');
      return;
    }

    const meta = buildCellMap(table);
    const sameRow = cells.every(c => c.parentElement === cells[0].parentElement);
    const firstCol = meta.info.get(cells[0]).col;
    const sameCol = cells.every(c => meta.info.get(c).col === firstCol);

    if (!sameRow && !sameCol) {
      alert('Cells must be in the same row OR the same column.');
      return;
    }

    let ordered = [...cells];
    if (sameRow) {
      ordered.sort((a, b) => meta.info.get(a).col - meta.info.get(b).col);
      for (let i = 1; i < ordered.length; i++) {
        const prevCol = meta.info.get(ordered[i - 1]).col;
        const curCol = meta.info.get(ordered[i]).col;
        if (curCol !== prevCol + 1) {
          alert('For horizontal merge, selected cells must be adjacent with no gaps.');
          return;
        }
      }
    } else {
      ordered.sort((a, b) => meta.info.get(a).row - meta.info.get(b).row);
      for (let i = 1; i < ordered.length; i++) {
        const prevRow = meta.info.get(ordered[i - 1]).row;
        const curRow = meta.info.get(ordered[i]).row;
        if (curRow !== prevRow + 1) {
          alert('For vertical merge, selected cells must be directly adjacent (no gap rows).');
          return;
        }
      }
    }

    const first = ordered[0];
    const combinedText = ordered.map(c => c.textContent.trim()).filter(Boolean).join(' ');
    if (combinedText) first.textContent = combinedText;

    if (sameRow) {
      first.colSpan = ordered.length;
    } else {
      first.rowSpan = ordered.length;
    }

    for (let i = 1; i < ordered.length; i++) {
      ordered[i].remove();
    }
    clearSelection();
    first.classList.add('cm-selected');
  }

  function deleteSelectedRows() {
    if (!requireUnlocked()) return;
    const cells = getSelectedCells();
    if (cells.length === 0) return alert('Select at least one cell.');
    const table = cells[0].closest('table');
    if (!cells.every(c => c.closest('table') === table)) {
      alert('Please select cells from a single table.');
      return;
    }

    // Identify rows to remove
    const rows = Array.from(new Set(cells.map(c => c.parentElement)));
    const rowIdxSet = new Set(rows.map(r => r.rowIndex));

    // Adjust rowSpan for cells above that span into rows being removed
    const map = buildCellMap(table);
    map.info.forEach((meta, cell) => {
      const cellRowIdx = cell.parentElement.rowIndex;
      if (rowIdxSet.has(cellRowIdx)) return; // row is being removed

      if (meta.rowSpan > 1) {
        let overlap = 0;
        for (let rr = meta.row; rr < meta.row + meta.rowSpan; rr++) {
          if (rowIdxSet.has(rr)) overlap++;
        }
        if (overlap > 0) {
          const newSpan = meta.rowSpan - overlap;
          if (newSpan <= 0) {
            cell.remove();
          } else {
            cell.rowSpan = newSpan;
          }
        }
      }
    });

    // Remove rows
    rows.forEach(r => r.remove());
    clearSelection();
  }

  function deleteSelectedCells() {
    if (!requireUnlocked()) return;
    const cells = getSelectedCells();
    if (cells.length === 0) return alert('Select at least one cell.');
    cells.forEach(c => c.remove());
    clearSelection();
  }

  function deleteSelectedColumns() {
    if (!requireUnlocked()) return;
    const cells = getSelectedCells();
    if (cells.length === 0) return alert('Select at least one cell.');
    const table = cells[0].closest('table');
    if (!cells.every(c => c.closest('table') === table)) {
      alert('Please select cells from a single table.');
      return;
    }

    const map = buildCellMap(table);
    const colsToRemove = new Set(cells.map(c => map.info.get(c).col));
    if (!confirm(`Delete ${colsToRemove.size} column(s) from this table?`)) return;
    removeColumnsFromTable(table, colsToRemove);
    clearSelection();
  }

  function removeColumnsFromTable(table, colsSet) {
    const map = buildCellMap(table);
    const toUpdate = [];
       const toRemove = [];

    map.info.forEach((meta, cell) => {
      const start = meta.col;
      const end = meta.col + meta.colSpan - 1;
      let count = 0;
      colsSet.forEach(idx => { if (idx >= start && idx <= end) count++; });
      if (count === 0) return;
      if (count >= meta.colSpan) {
        toRemove.push(cell);
      } else {
        toUpdate.push([cell, meta.colSpan - count]);
      }
    });

    toUpdate.forEach(([cell, newSpan]) => { cell.colSpan = newSpan; });
    toRemove.forEach(cell => cell.remove());
  }

  function addNewTable() {
    if (!requireUnlocked()) return;
    const table = document.createElement('table');
    table.innerHTML = `
      <tbody>
        <tr><th contenteditable="true">New Table</th></tr>
        <tr><td contenteditable="true">-</td></tr>
      </tbody>
    `;
    table.style.width = '100%';
    table.style.margin = '16px 0';
    table.style.borderCollapse = 'collapse';
    table.querySelectorAll('th, td').forEach(c => { c.style.border = '1px solid #555'; c.style.padding = '6px'; });

    getContentRoot().appendChild(table);
    activeTable = table;
    table.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  // Save/Restore (no backend; localStorage only)
  function attachAutoSaveHandlers() {
    document.addEventListener('input', (e) => {
      if (!AUTO_SAVE) return;
      const cell = e.target.closest && e.target.closest('td, th');
      if (cell) scheduleAutoSave();
    }, true);
  }

  function scheduleAutoSave() {
    if (!AUTO_SAVE) return;
    clearTimeout(saveTimer);
    saveTimer = setTimeout(persistToLocalStorage, 600);
  }

  function getCleanContentHTML() {
    // Clone content root without toolbar and without edit artifacts
    const root = getContentRoot();
    const clone = root.cloneNode(true);
    clone.querySelectorAll('.cm-selected').forEach(el => el.classList.remove('cm-selected'));
    clone.querySelectorAll('[contenteditable]').forEach(el => el.removeAttribute('contenteditable'));
    clone.querySelectorAll('script').forEach(s => s.remove());
    return clone.innerHTML;
  }

  function persistToLocalStorage() {
    try {
      const html = getCleanContentHTML();
      localStorage.setItem(STORAGE_KEY, html);
    } catch (e) {
      console.warn('Could not save to browser storage:', e);
    }
  }

  function restoreFromLocalStorage() {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (!saved) {
      alert('No saved version found in this browser.');
      return;
    }
    if (!confirm('Restore the saved version from this browser? Current unsaved changes will be lost.')) return;
    getContentRoot().innerHTML = saved;
    setEditable(editMode);
    clearSelection();
  }

  function clearSavedFromLocalStorage() {
    if (!confirm('Clear saved version from this browser?')) return;
    localStorage.removeItem(STORAGE_KEY);
    alert('Saved version cleared.');
  }

  function maybeRestoreSaved() {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved && confirm('A saved version was found in this browser. Restore it now?')) {
      getContentRoot().innerHTML = saved;
    }
  }

  function downloadHTML() {
    // Build a clean, single HTML file
    const inner = getCleanContentHTML();
    const html = [
      '<!DOCTYPE html>',
      '<html lang="en">',
      '<head>',
      '<meta charset="UTF-8">',
      '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
      '<title>Updated Dashboard</title>',
      '<style>',
      '  table { border-collapse: collapse; width: 100%; }',
      '  th, td { border: 1px solid #555; padding: 6px; }',
      '</style>',
      '</head>',
      '<body>',
      inner,
      '</body>',
      '</html>'
    ].join('\n');

    downloadBlob(html, 'updated_dashboard.html', 'text/html');
  }

  function downloadExcel() {
    // Export all tables on the page as an HTML-based .xls
    const tables = Array.from(document.querySelectorAll('#cm-content-root table')).map(t => {
      const clone = t.cloneNode(true);
      clone.querySelectorAll('.cm-selected').forEach(el => el.classList.remove('cm-selected'));
      clone.querySelectorAll('[contenteditable]').forEach(el => el.removeAttribute('contenteditable'));
      clone.querySelectorAll('*').forEach(el => el.removeAttribute('onclick'));
      return clone.outerHTML;
    });

    const excelHTML = `
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="UTF-8">
<!--[if gte mso 9]>
<xml>
 <x:ExcelWorkbook>
  <x:ExcelWorksheets>
   <x:ExcelWorksheet>
    <x:Name>KPIs</x:Name>
    <x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
   </x:ExcelWorksheet>
  </x:ExcelWorksheets>
 </x:ExcelWorkbook>
</xml>
<![endif]-->
<style>
  body, table, td, th, h1, h2, h3, h4, h5, h6 { font-family: "Times New Roman", Times, serif; }
  table { border-collapse: collapse; }
  td, th { border: 1px solid #000; padding: 4px; white-space: pre-wrap; mso-number-format:"\\@"; }
</style>
</head>
<body>
${tables.join('<br><br>')}
</body>
</html>`.trim();

    downloadBlob(excelHTML, 'tables.xls', 'application/vnd.ms-excel');
  }

  // -------- PPTX EXPORT --------

  async function downloadPPT() {
    try {
      const PptxGenJS = await loadPptxLibrary();
      const pptx = new PptxGenJS();
      pptx.layout = 'LAYOUT_16x9';

      const pageW = 13.33; // inches for 16x9
      const marginX = 0.5;
      const startYTitle = 0.4;
      const startYTable = 1.0;
      const usableW = pageW - marginX * 2;

      const tables = Array.from(document.querySelectorAll('#cm-content-root table'));
      if (tables.length === 0) {
        alert('No tables found to export.');
        return;
      }

      tables.forEach((table, idx) => {
        const map = buildCellMap(table);
        const { dataRows, skipTitle } = convertTableToPptData(table, map);

        const slide = pptx.addSlide();
        const title = deriveTableTitle(table, map, idx + 1);
        slide.addText(title, {
          x: marginX,
          y: startYTitle,
          w: usableW,
          fontFace: 'Times New Roman',
          fontSize: 20,
          bold: true,
          color: '0d6efd'
        });

        // Compute column widths (inches) based on DOM widths
        const colW = computeColumnWidthsInches(table, map, usableW);

        // Table styling
        const tableOpts = {
          x: marginX,
          y: startYTable,
          w: usableW,
          colW,
          border: { pt: 1, color: '666666' },
          fontFace: 'Times New Roman',
          fontSize: 12,
          valign: 'middle',
          autoPage: true,               // auto split long tables
          autoPageRepeatHeader: true,   // repeat header rows
          autoPageLines: 1
        };

        slide.addTable(dataRows, tableOpts);
      });

      await pptx.writeFile({ fileName: 'tables.pptx' });
    } catch (err) {
      console.error('PPT export failed:', err);
      alert('Could not generate PPT. If you are offline, please connect to the internet (to load the PPT library) or vendor it locally.');
    }
  }

  function convertTableToPptData(table, map) {
    const rowsOut = [];
    const headerFill = 'e7f0ff';
    const zebraFill = 'f9fbff';
    const headerRowsIdx = getHeaderRowIndices(table); // which rows are header-ish
    const titleRowIdx = detectTitleRowIndex(table, map); // a big spanning title row, if present
    const skipTitle = titleRowIdx === 0; // if first row is a title row, skip it inside table and use slide title

    for (let r = 0; r < map.rows; r++) {
      if (skipTitle && r === 0) continue;

      const rowCells = [];
      for (let c = 0; c < map.cols; c++) {
        const cell = map.matrix[r][c];
        if (!cell) continue;
        const info = map.info.get(cell);
        // If this position is covered by a span from another anchor, skip it
        if (info.row !== r || info.col !== c) continue;

        const text = getCellText(cell);
        const isHeader = cell.tagName.toLowerCase() === 'th' || headerRowsIdx.has(r);

        const opts = {
          fontFace: 'Times New Roman',
          fontSize: 12,
          color: isHeader ? 'FFFFFF' : '000000',
          bold: isHeader ? true : false,
          align: isHeader ? 'center' : 'left',
          valign: 'middle',
          fill: isHeader ? '0d6efd' : ((r % 2 === 1) ? zebraFill : 'FFFFFF'),
          border: [{ color: '666666', pt: 1 }]
        };

        if (info.colSpan > 1) opts.colSpan = info.colSpan;
        if (info.rowSpan > 1) opts.rowSpan = info.rowSpan;

        rowCells.push({ text, options: opts });
      }
      rowsOut.push(rowCells);
    }

    return { dataRows: rowsOut, skipTitle };
  }

  function getHeaderRowIndices(table) {
    const set = new Set();
    Array.from(table.rows).forEach((tr, idx) => {
      const hasTh = Array.from(tr.cells).some(td => td.tagName.toLowerCase() === 'th');
      if (hasTh) set.add(idx);
    });
    return set;
  }

  function detectTitleRowIndex(table, map) {
    // Heuristic: a first row with a single cell spanning all columns, containing a heading element (h1-h4)
    if (map.rows === 0) return -1;
    const row0 = table.rows[0];
    if (!row0 || row0.cells.length !== 1) return -1;
    const cell = row0.cells[0];
    const spanAll = (cell.colSpan || 1) >= map.cols;
    const hasHeading = cell.querySelector('h1,h2,h3,h4') != null;
    return spanAll && hasHeading ? 0 : -1;
  }

  function deriveTableTitle(table, map, defaultIndex) {
    const idxLabel = `Table ${defaultIndex}`;
    // 1) Prefer caption
    const cap = table.querySelector('caption');
    if (cap) return cap.innerText.trim() || idxLabel;

    // 2) If first row is a spanning heading row
    if (map.rows > 0) {
      const r0 = table.rows[0];
      if (r0 && r0.cells.length === 1) {
        const cell = r0.cells[0];
        if ((cell.colSpan || 1) >= map.cols) {
          const h = cell.querySelector('h1,h2,h3,h4');
          if (h) return h.innerText.trim() || idxLabel;
          const t = cell.innerText.trim();
          if (t) return t;
        }
      }
    }

    // 3) Try previous sibling heading near the table
    let prev = table.previousElementSibling;
    let hops = 0;
    while (prev && hops < 3) {
      const h = prev.querySelector && prev.querySelector('h1,h2,h3,h4');
      if (h) return h.innerText.trim();
      if (/H[1-4]/.test(prev.tagName)) return prev.innerText.trim();
      prev = prev.previousElementSibling;
      hops++;
    }

    return idxLabel;
  }

  function getCellText(cell) {
    // Preserve line breaks
    return (cell.innerText || '').replace(/\r\n/g, '\n').replace(/\n{2,}/g, '\n').trim();
  }

  function computeColumnWidthsInches(table, map, targetTotalInches) {
    const pxPerInch = 96;
    const cols = map.cols;
    const colPx = new Array(cols).fill(40); // default min px

    // Collect widths from DOM; divide merged cell width by its colSpan
    for (let r = 0; r < map.rows; r++) {
      for (let c = 0; c < cols; c++) {
        const cell = map.matrix[r][c];
        if (!cell) continue;
        const info = map.info.get(cell);
        if (info.row !== r || info.col !== c) continue; // skip covered positions

        let w = cell.getBoundingClientRect().width;
        if (!w || w < 10) {
          // fallback: estimate by text length
          const ch = (cell.innerText || '').trim().length;
          w = Math.max(40, Math.min(400, 7 * ch + 20));
        }
        const per = w / info.colSpan;
        for (let k = 0; k < info.colSpan; k++) {
          const ci = info.col + k;
          colPx[ci] = Math.max(colPx[ci], per);
        }
      }
    }

    // Scale to target width in inches
    const totalPx = colPx.reduce((a, b) => a + b, 0) || cols * 60;
    const scale = (targetTotalInches * pxPerInch) / totalPx;
    const colIn = colPx.map(px => Math.max(0.5, px * scale / pxPerInch)); // ensure min 0.5"
    const sumIn = colIn.reduce((a, b) => a + b, 0);
    const adjust = targetTotalInches / sumIn;
    return colIn.map(v => v * adjust);
  }

  async function loadPptxLibrary() {
    if (window.PptxGenJS) return window.PptxGenJS;
    const urls = [
      'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js',
      'https://unpkg.com/pptxgenjs/dist/pptxgen.bundle.js'
    ];
    for (const src of urls) {
      try {
        await injectScript(src);
        if (window.PptxGenJS) return window.PptxGenJS;
      } catch (e) {
        // try next
      }
    }
    throw new Error('PptxGenJS could not be loaded from CDN.');
  }

  function injectScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.async = true;
      s.onload = () => resolve();
      s.onerror = () => reject(new Error('Failed to load: ' + src));
      document.head.appendChild(s);
    });
  }

  // -------- Utilities --------

  function downloadBlob(content, filename, type) {
    const blob = new Blob([content], { type });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }
})();