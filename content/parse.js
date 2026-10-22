// Turns an Excel file or text copied from Excel into [{ barcode, qty, sourceRows }].
// Column A = barcode, column B = quantity. A header row is skipped automatically.
(() => {
  const AXF = (window.AXF = window.AXF || {});

  function cellToText(value) {
    if (value === null || value === undefined) {
      return '';
    }
    if (typeof value === 'number') {
      // Barcodes typed/scanned into a General cell become numbers; keep every digit.
      return Number.isInteger(value) ? value.toFixed(0) : String(value);
    }
    return String(value).replace(/ /g, ' ').trim();
  }

  function parseQty(text) {
    if (text === '') {
      return { qty: 1, defaulted: true };
    }
    const normalised = text.replace(/,/g, '');
    const qty = Number(normalised);
    if (!Number.isFinite(qty)) {
      return { error: `Quantity "${text}" is not a number` };
    }
    if (qty <= 0) {
      return { error: `Quantity must be more than 0 (got ${text})` };
    }
    return { qty };
  }

  function looksLikeHeader(barcodeText, qtyText) {
    const qtyIsNumber = qtyText !== '' && Number.isFinite(Number(qtyText.replace(/,/g, '')));
    return !qtyIsNumber && /[a-z]/i.test(barcodeText) && /bar|code|item|sku|qty|quantity/i.test(barcodeText + ' ' + qtyText);
  }

  /**
   * @param {Array<Array<unknown>>} table rows of cells, first two columns used
   * @param {{ mergeDuplicates: boolean }} options
   * @param {number} firstRowNumber spreadsheet row number of table[0], used in messages
   * @returns {{ items: Array<{barcode: string, qty: number, sourceRows: number[]}>, problems: Array<{row: number, message: string}> }}
   */
  function rowsToItems(table, options, firstRowNumber = 1) {
    const problems = [];
    const items = [];
    const byBarcode = new Map();

    table.forEach((cells, index) => {
      const rowNumber = firstRowNumber + index;
      const barcode = cellToText(cells[0]);
      const qtyText = cellToText(cells[1]);

      if (barcode === '' && qtyText === '') {
        return;
      }
      if (index === 0 && looksLikeHeader(barcode, qtyText)) {
        return;
      }
      if (barcode === '') {
        problems.push({ row: rowNumber, message: `Row ${rowNumber}: barcode is empty` });
        return;
      }

      const parsed = parseQty(qtyText);
      if (parsed.error) {
        problems.push({ row: rowNumber, message: `Row ${rowNumber}: ${parsed.error}` });
        return;
      }

      const existing = options.mergeDuplicates ? byBarcode.get(barcode) : null;
      if (existing) {
        existing.qty = roundQty(existing.qty + parsed.qty);
        existing.sourceRows.push(rowNumber);
        return;
      }

      const item = { barcode, qty: parsed.qty, sourceRows: [rowNumber] };
      items.push(item);
      byBarcode.set(barcode, item);
    });

    return { items, problems };
  }

  function roundQty(value) {
    return Math.round(value * 10000) / 10000;
  }

  function parsePastedText(text, options) {
    const lines = text.replace(/\r\n?/g, '\n').split('\n');
    const separator = lines.some((line) => line.includes('\t')) ? '\t' : /[,;]/;
    const table = lines.map((line) => line.split(separator));
    return rowsToItems(table, options);
  }

  async function parseFile(file, options) {
    const buffer = await file.arrayBuffer();
    const workbook = window.XLSX.read(buffer, { type: 'array' });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!firstSheet || !firstSheet['!ref']) {
      return { items: [], problems: [] };
    }
    const table = window.XLSX.utils.sheet_to_json(firstSheet, { header: 1, raw: true, defval: '', blankrows: true });
    const firstRowNumber = window.XLSX.utils.decode_range(firstSheet['!ref']).s.r + 1;
    return rowsToItems(table, options, firstRowNumber);
  }

  AXF.parse = { parsePastedText, parseFile, rowsToItems };
})();
