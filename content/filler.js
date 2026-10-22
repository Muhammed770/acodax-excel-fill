// Drives the item entry row the way a person does: pick the item, type the quantity, press +.
// Works on forms built from the shared SelectItem dropdown (Purchase Invoice, Stock Transfer, ...).
(() => {
  const AXF = (window.AXF = window.AXF || {});

  const SELECT_PREFIX = 'item-select-autocomplete';
  const OPTION_TIMEOUT_MS = 20000;
  const SETTLE_TIMEOUT_MS = 15000;
  const COMMIT_TIMEOUT_MS = 45000;
  const ROW_READY_TIMEOUT_MS = 30000;

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  async function waitFor(check, timeoutMs, intervalMs = 150) {
    const started = Date.now();
    while (Date.now() - started < timeoutMs) {
      const value = check();
      if (value) {
        return value;
      }
      await sleep(intervalMs);
    }
    return null;
  }

  function normalise(text) {
    return String(text || '').replace(/\s+/g, ' ').trim().toUpperCase();
  }

  function isVisible(element) {
    return Boolean(element && element.isConnected && element.getClientRects().length > 0 && !element.closest('.d-none'));
  }

  // ---------- API (same headers the ERP's own axios setup sends) ----------

  function apiHeaders() {
    const headers = {
      Accept: 'application/json',
      'X-Acodax-App': 'ERP',
      Timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
    };
    let user = null;
    try {
      user = JSON.parse(localStorage.getItem('ACODAX_USER') || 'null');
    } catch (error) {
      user = null;
    }
    if (user && user.token) {
      headers.Authorization = `Bearer ${user.token}`;
    }
    const transBranchId = (user && (user.trans_branch_id || (user.branch && user.branch.id))) || null;
    if (transBranchId) {
      headers['X-Acodax-Trans-Branch-Id'] = transBranchId;
    }
    const tenantId = localStorage.getItem('ACODAX_X_ACODAX_TENANT_ID');
    if (tenantId) {
      headers['X-Acodax-Tenant-Id'] = tenantId;
    }
    return headers;
  }

  async function apiGet(path, params) {
    const url = new URL(`/api/${path}`, window.location.origin);
    Object.entries(params || {}).forEach(([key, value]) => url.searchParams.set(key, value));
    const response = await fetch(url.toString(), { headers: apiHeaders(), credentials: 'include' });
    if (response.status === 401) {
      throw Object.assign(new Error('Your login has expired. Log in again and retry.'), { fatal: true });
    }
    if (response.status === 403) {
      throw Object.assign(new Error('No permission to view items (show-item).'), { forbidden: true });
    }
    if (!response.ok) {
      throw new Error(`Server answered ${response.status} for ${path}`);
    }
    return response.json();
  }

  function listFrom(body) {
    const data = body && body.data;
    if (Array.isArray(data)) {
      return data;
    }
    if (data && Array.isArray(data.data)) {
      return data.data;
    }
    return [];
  }

  function sameCode(a, b) {
    const left = normalise(a).replace(/\s/g, '');
    const right = normalise(b).replace(/\s/g, '');
    if (!left || !right) {
      return false;
    }
    if (left === right) {
      return true;
    }
    // Excel drops leading zeros from barcodes typed into a number cell.
    return /^\d+$/.test(left) && /^\d+$/.test(right) && left.replace(/^0+/, '') === right.replace(/^0+/, '');
  }

  const itemCache = new Map();

  /**
   * Finds the one item whose barcode (or item code) is exactly this value.
   * @returns {Promise<{ok: true, item: {id: string, stock_id: string, item_name: string}} | {ok: false, message: string, fatal?: boolean}>}
   */
  async function resolveBarcode(code) {
    const cacheKey = normalise(code);
    if (itemCache.has(cacheKey)) {
      return itemCache.get(cacheKey);
    }

    let result;
    try {
      const candidates = listFrom(await apiGet('inventory/items-auto', { item_name: code, exclude_kits: 1 }));
      if (candidates.length === 0) {
        result = { ok: false, message: 'No item has this barcode' };
      } else {
        const byStockId = candidates.find((candidate) => sameCode(candidate.stock_id, code));
        result = byStockId ? { ok: true, item: pick(byStockId) } : await matchByBarcode(code, candidates);
      }
    } catch (error) {
      result = { ok: false, message: error.message, fatal: Boolean(error.fatal) };
    }

    if (!result.fatal) {
      itemCache.set(cacheKey, result);
    }
    return result;
  }

  async function matchByBarcode(code, candidates) {
    const matches = [];
    // The search ranks an exact barcode above partial matches, so exact hits come first.
    for (const candidate of candidates.slice(0, 6)) {
      let details;
      try {
        details = (await apiGet(`inventory/items/${candidate.id}`)).data || {};
      } catch (error) {
        if (error.forbidden) {
          return candidates.length === 1
            ? { ok: true, item: pick(candidates[0]), note: 'barcode not double-checked (no item view permission)' }
            : { ok: false, message: 'Several items match and this user cannot open item details to check the barcode' };
        }
        throw error;
      }
      if (sameCode(details.barcode, code)) {
        matches.push(pick({ ...candidate, ...details }));
      } else if (matches.length > 0) {
        break;
      }
    }

    if (matches.length === 1) {
      return { ok: true, item: matches[0] };
    }
    if (matches.length > 1) {
      return { ok: false, message: `Barcode is on ${matches.length} items: ${matches.map((m) => m.stock_id).join(', ')}` };
    }
    return { ok: false, message: `No exact barcode match (${candidates.length} similar items)` };
  }

  function pick(item) {
    return { id: item.id, stock_id: item.stock_id, item_name: item.item_name };
  }

  // ---------- DOM ----------

  function entryRows() {
    return Array.from(document.querySelectorAll(`input.${SELECT_PREFIX}__input`))
      .map((input) => input.closest('tr'))
      .filter((tr) => tr && tr.querySelector('input[name="quantity"]') && addButton(tr));
  }

  /** The new-entry row is the last item row; edit rows for saved lines come before it. */
  function findEntryRow() {
    const rows = entryRows();
    const last = rows[rows.length - 1];
    return last && isVisible(last) ? last : null;
  }

  function addButton(row) {
    const lastCell = row.lastElementChild;
    return lastCell ? lastCell.querySelector('button.text-success') : null;
  }

  function selectInput(row) {
    return row.querySelector(`input.${SELECT_PREFIX}__input`);
  }

  function selectedLabel(row) {
    const single = row.querySelector(`.${SELECT_PREFIX}__single-value`);
    return single ? single.textContent : '';
  }

  function setNativeValue(input, value) {
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
    setter.call(input, value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
  }

  function openMenu() {
    const menus = Array.from(document.querySelectorAll(`.${SELECT_PREFIX}__menu`)).filter(isVisible);
    return menus[menus.length - 1] || null;
  }

  function menuOptions(menu) {
    return Array.from(menu.querySelectorAll(`.${SELECT_PREFIX}__option[role="option"]`));
  }

  function mouse(element, type) {
    element.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, button: 0, view: window }));
  }

  function warningPopup() {
    return Array.from(document.querySelectorAll('.modal.show, .modal.modal-with-overlay')).find(isVisible) || null;
  }

  function errorListCount() {
    return document.querySelectorAll('tbody ul.text-danger').length;
  }

  async function clearRow(row) {
    const clear = row.querySelector(`.${SELECT_PREFIX}__clear-indicator`);
    if (clear) {
      mouse(clear, 'mousedown');
      mouse(clear, 'mouseup');
    }
    const input = selectInput(row);
    if (input) {
      setNativeValue(input, '');
      input.blur();
    }
    await sleep(300);
  }

  function menuIsLoading(menu) {
    return Boolean(menu && menu.querySelector(`.${SELECT_PREFIX}__menu-notice--loading`));
  }

  async function chooseItem(row, item) {
    const input = selectInput(row);
    input.focus();
    // Focusing loads a default item list; let it finish so it is not mistaken for our search.
    await waitFor(() => !menuIsLoading(openMenu()), 5000);
    await sleep(150);

    // Search by the unique item code: an exact code match is ranked first.
    setNativeValue(input, item.stock_id);
    let typedAt = Date.now();
    const wanted = normalise(item.item_name);
    let sawLoading = false;
    let lastSignature = '';
    let stableSince = Date.now();
    let retypes = 0;

    const outcome = await waitFor(() => {
      // The form can move focus elsewhere (e.g. to Item Code after a line is added), which empties the search.
      if (document.activeElement !== input && input.value === '' && retypes < 2) {
        retypes += 1;
        input.focus();
        setNativeValue(input, item.stock_id);
        typedAt = Date.now();
        sawLoading = false;
        return null;
      }
      const menu = openMenu();
      if (!menu) {
        return null;
      }
      if (menuIsLoading(menu)) {
        sawLoading = true;
        return null;
      }
      if (menu.querySelector(`.${SELECT_PREFIX}__menu-notice--no-options`)) {
        return sawLoading || Date.now() - typedAt > 3000 ? { reason: 'The item list did not find it' } : null;
      }
      const options = menuOptions(menu);
      const signature = options.map((option) => option.textContent).join('|');
      if (signature !== lastSignature) {
        lastSignature = signature;
        stableSince = Date.now();
        return null;
      }
      // Results either came from a loading round after typing, or from cache and stayed put.
      const settled = sawLoading ? Date.now() - stableSince > 300 : Date.now() - typedAt > 2500;
      if (!settled || options.length === 0) {
        return null;
      }
      return normalise(options[0].textContent) === wanted
        ? { option: options[0] }
        : { reason: 'The item list showed a different item first' };
    }, OPTION_TIMEOUT_MS, 100);

    if (!outcome || !outcome.option) {
      await clearRow(row);
      return { ok: false, message: outcome ? outcome.reason : 'The item list did not load in time' };
    }

    mouse(outcome.option, 'mousedown');
    mouse(outcome.option, 'mouseup');
    outcome.option.click();

    const picked = await waitFor(() => normalise(selectedLabel(row)).includes(wanted), SETTLE_TIMEOUT_MS);
    if (!picked) {
      await clearRow(row);
      return { ok: false, message: 'Item did not get selected' };
    }
    return { ok: true };
  }

  /** Where the row shows the item code (PI "Item Code" box, Stock Transfer code cell), make sure it is ours. */
  function rowCodeCheck(row, stockId) {
    const codeInput = row.querySelector('input[name="stock_id"]');
    if (codeInput) {
      return sameCode(codeInput.value, stockId) ? 'match' : 'mismatch';
    }
    const codeCell = row.cells && row.cells[1];
    const cellText = codeCell && !codeCell.querySelector('input, select, button') ? codeCell.textContent.trim() : '';
    if (cellText && cellText !== '#') {
      return sameCode(cellText, stockId) ? 'match' : 'mismatch';
    }
    return 'unknown';
  }

  /** After picking an item the form loads its details and resets quantity (to 1 or 0); wait for that to finish. */
  async function waitForItemDetails(row, stockId) {
    const qtyInput = row.querySelector('input[name="quantity"]');
    let lastSignature = null;
    let stableSince = Date.now();
    return waitFor(() => {
      const signature = `${qtyInput.value}|${rowCodeCheck(row, stockId)}`;
      if (signature !== lastSignature) {
        lastSignature = signature;
        stableSince = Date.now();
        return null;
      }
      const stableFor = Date.now() - stableSince;
      return qtyInput.value !== '' ? stableFor > 500 : stableFor > 1500;
    }, SETTLE_TIMEOUT_MS, 100);
  }

  async function setQuantity(row, qty) {
    const qtyInput = row.querySelector('input[name="quantity"]');
    const wanted = String(qty);
    for (let attempt = 0; attempt < 3; attempt += 1) {
      qtyInput.focus();
      setNativeValue(qtyInput, wanted);
      await sleep(400);
      if (Number(qtyInput.value) === Number(wanted)) {
        return { ok: true };
      }
    }
    return { ok: false, actual: qtyInput.value };
  }

  async function commitRow(row, onWaitingForUser) {
    const button = await waitFor(() => {
      const candidate = addButton(row);
      return candidate && !candidate.disabled ? candidate : null;
    }, 5000);
    if (!button) {
      return { ok: false, message: 'The + button stayed disabled' };
    }

    const errorsBefore = errorListCount();
    button.click();

    let askedUser = false;
    const started = Date.now();
    while (true) {
      await sleep(200);
      if (errorListCount() > errorsBefore) {
        const text = document.querySelectorAll('tbody ul.text-danger')[errorsBefore].textContent.trim();
        return {
          ok: false,
          fatal: true,
          message: `The form rejected this line: ${text}. Fix or cancel that line in the form, then press Start again.`,
        };
      }
      if (warningPopup()) {
        if (!askedUser) {
          askedUser = true;
          onWaitingForUser('A popup needs your answer in the form');
        }
        continue;
      }
      const current = findEntryRow();
      if (current && !selectedLabel(current) && !selectInput(current).value) {
        // Let the form finish its own after-add work (it refocuses the first field ~300 ms later).
        await sleep(800);
        return { ok: true };
      }
      if (!askedUser && Date.now() - started > COMMIT_TIMEOUT_MS) {
        return { ok: false, fatal: true, message: 'The line did not finish saving in time. Check the form before continuing.' };
      }
    }
  }

  /**
   * Adds one line. `item.resolved` must hold the item from resolveBarcode().
   * @returns {Promise<{ok: boolean, message?: string, fatal?: boolean}>}
   */
  async function fillItem(item, onWaitingForUser = () => {}) {
    const row = await waitFor(findEntryRow, ROW_READY_TIMEOUT_MS);
    if (!row) {
      return {
        ok: false,
        fatal: true,
        message: 'The item entry row is not available. Pick the supplier first and close any line that is being edited.',
      };
    }
    if (selectedLabel(row)) {
      await clearRow(row);
    }

    const chosen = await chooseItem(row, item.resolved);
    if (!chosen.ok) {
      return chosen;
    }

    await waitForItemDetails(row, item.resolved.stock_id);
    if (rowCodeCheck(row, item.resolved.stock_id) === 'mismatch') {
      await clearRow(row);
      return { ok: false, message: 'The form picked a different item code' };
    }
    const quantity = await setQuantity(row, item.qty);
    if (!quantity.ok) {
      await clearRow(row);
      return { ok: false, message: `Quantity changed to ${quantity.actual || 0} by the form (not enough stock?)` };
    }

    return commitRow(row, onWaitingForUser);
  }

  /** True when the form has an item entry row at all, even while it is briefly hidden. */
  function hasEntryRow() {
    return entryRows().length > 0;
  }

  AXF.filler = { findEntryRow, hasEntryRow, resolveBarcode, fillItem };
})();
