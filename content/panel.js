// Floating "Excel Fill" panel. Shown only on pages that have an item entry row.
(() => {
  const AXF = (window.AXF = window.AXF || {});
  if (AXF.panelMounted) {
    return;
  }
  AXF.panelMounted = true;

  const RESOLVE_CONCURRENCY = 4;
  const PI_LINE_LIMIT = 100;

  const STYLE = `
    :host { all: initial; }
    * { box-sizing: border-box; font-family: -apple-system, "Segoe UI", Roboto, Arial, sans-serif; }
    [hidden] { display: none !important; }
    .fab { position: fixed; right: 20px; bottom: 20px; z-index: 2147483646; background: #0f766e; color: #fff;
      border: none; border-radius: 22px; padding: 10px 16px; font-size: 13px; font-weight: 600; cursor: pointer;
      box-shadow: 0 4px 14px rgba(0,0,0,.25); }
    .fab:hover { background: #115e59; }
    .panel { position: fixed; right: 20px; bottom: 70px; z-index: 2147483647; width: 460px; max-width: calc(100vw - 40px);
      max-height: 80vh; display: flex; flex-direction: column; background: #fff; color: #1f2937; border-radius: 10px;
      box-shadow: 0 10px 30px rgba(0,0,0,.3); font-size: 13px; overflow: hidden; }
    .head { display: flex; align-items: center; justify-content: space-between; padding: 10px 14px; background: #0f766e; color: #fff; }
    .head b { font-size: 14px; }
    .head button { background: transparent; border: none; color: #fff; font-size: 20px; cursor: pointer; line-height: 1; }
    .body { padding: 12px 14px; overflow: auto; display: flex; flex-direction: column; gap: 10px; }
    .hint { color: #6b7280; font-size: 12px; line-height: 1.4; }
    textarea { width: 100%; height: 80px; resize: vertical; font-family: Consolas, monospace; font-size: 12px;
      border: 1px solid #d1d5db; border-radius: 6px; padding: 6px; }
    textarea::placeholder { color: #9ca3af; font-style: italic; }
    .row { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
    button.btn { border: 1px solid #d1d5db; background: #f9fafb; color: #111827; border-radius: 6px; padding: 6px 12px;
      font-size: 13px; cursor: pointer; }
    button.btn:disabled { opacity: .5; cursor: default; }
    button.primary { background: #16a34a; border-color: #16a34a; color: #fff; }
    button.danger { background: #dc2626; border-color: #dc2626; color: #fff; }
    label.opt { display: flex; gap: 6px; align-items: center; font-size: 12px; color: #374151; }
    .summary { font-weight: 600; white-space: pre-line; }
    .problems { color: #b91c1c; font-size: 12px; white-space: pre-line; max-height: 80px; overflow: auto; }
    .tablewrap { max-height: 34vh; overflow: auto; border: 1px solid #e5e7eb; border-radius: 6px; }
    table { width: 100%; border-collapse: collapse; font-size: 12px; }
    th, td { text-align: left; padding: 4px 6px; border-bottom: 1px solid #f3f4f6; vertical-align: top; }
    th { position: sticky; top: 0; background: #f9fafb; }
    .s-pending, .s-ready { color: #6b7280; } .s-checking, .s-working { color: #2563eb; }
    .s-done { color: #16a34a; } .s-failed, .s-notfound { color: #dc2626; }
    .progress { height: 6px; background: #e5e7eb; border-radius: 3px; overflow: hidden; }
    .progress > div { height: 100%; width: 0; background: #16a34a; transition: width .2s; }
    .credit { font-size: 11px; color: #9ca3af; text-align: right; }
    .credit a { color: #0f766e; }
  `;

  const host = document.createElement('div');
  host.id = 'acodax-excel-fill';
  const root = host.attachShadow({ mode: 'open' });
  root.innerHTML = `
    <style>${STYLE}</style>
    <button class="fab">Excel Fill</button>
    <div class="panel" hidden>
      <div class="head"><b>Fill items from Excel</b><button class="close" title="Close">×</button></div>
      <div class="body">
        <div class="hint">Column A = barcode, column B = quantity (empty quantity = 1).
          Fill the supplier and the other header fields first, then load the sheet.</div>
        <div class="row"><input type="file" class="file" accept=".xlsx,.xls,.csv"></div>
        <div class="hint">Or select the two columns in Excel, copy, and paste here:</div>
        <textarea class="paste" placeholder="Paste here. Example:&#10;1234567890123&#9;2&#10;9876543210987&#9;6"></textarea>
        <div class="row">
          <button class="btn load-paste">Use pasted rows</button>
          <label class="opt"><input type="checkbox" class="merge" checked> Add up repeated barcodes</label>
        </div>
        <div class="problems"></div>
        <div class="summary"></div>
        <div class="progress" hidden><div></div></div>
        <div class="tablewrap" hidden>
          <table><thead><tr><th>#</th><th>Barcode</th><th>Qty</th><th>Item</th><th>Status</th></tr></thead><tbody></tbody></table>
        </div>
        <div class="row">
          <button class="btn check" disabled>Check barcodes</button>
          <button class="btn primary start" disabled>Start filling</button>
          <button class="btn pause" disabled>Pause</button>
          <button class="btn danger stop" disabled>Stop</button>
          <button class="btn copy-failed" hidden>Copy failed rows</button>
        </div>
        <div class="credit">Developed by <a href="https://github.com/muhammed770/" target="_blank" rel="noopener">muhammedv</a></div>
      </div>
    </div>
  `;

  const $ = (selector) => root.querySelector(selector);
  const ui = {
    fab: $('.fab'),
    panel: $('.panel'),
    close: $('.close'),
    file: $('.file'),
    paste: $('.paste'),
    loadPaste: $('.load-paste'),
    merge: $('.merge'),
    problems: $('.problems'),
    summary: $('.summary'),
    progress: $('.progress'),
    progressBar: $('.progress > div'),
    tableWrap: $('.tablewrap'),
    tbody: $('tbody'),
    check: $('.check'),
    start: $('.start'),
    pause: $('.pause'),
    stop: $('.stop'),
    copyFailed: $('.copy-failed'),
  };

  // status: pending -> checking -> ready | notfound; ready -> working -> done | failed (failed is retried on Start)
  const state = { items: [], running: false, paused: false, stopRequested: false };

  const STATUS_LABELS = {
    pending: 'Waiting',
    checking: 'Checking…',
    ready: 'Found',
    notfound: 'Not found',
    working: 'Filling…',
    done: 'Added',
    failed: 'Failed',
  };

  function count(...statuses) {
    return state.items.filter((item) => statuses.includes(item.status)).length;
  }

  function setItems(result, emptyMessage) {
    state.items = result.items.map((item) => ({ ...item, status: 'pending', message: '', resolved: null }));
    ui.problems.textContent = result.problems.map((problem) => problem.message).join('\n');
    renderTable();
    const totalQty = Math.round(state.items.reduce((sum, item) => sum + item.qty, 0) * 10000) / 10000;
    let summary = state.items.length ? `${state.items.length} items, total qty ${totalQty}` : emptyMessage;
    if (/purchase-invoice/.test(window.location.pathname) && state.items.length > PI_LINE_LIMIT) {
      summary += `\nWarning: a purchase invoice takes at most ${PI_LINE_LIMIT} lines. Split the sheet.`;
    }
    ui.summary.textContent = summary;
    ui.copyFailed.hidden = true;
    ui.start.textContent = 'Start filling';
    setRunningUi(false);
    updateProgress();
  }

  function renderTable() {
    ui.tableWrap.hidden = state.items.length === 0;
    ui.tbody.innerHTML = '';
    state.items.forEach((item, index) => {
      const tr = document.createElement('tr');
      [String(index + 1), item.barcode, String(item.qty), '', ''].forEach((text) => {
        const td = document.createElement('td');
        td.textContent = text;
        tr.appendChild(td);
      });
      item.row = tr;
      ui.tbody.appendChild(tr);
      updateRow(item);
    });
  }

  function updateRow(item) {
    if (!item.row) {
      return;
    }
    const itemCell = item.row.children[3];
    const statusCell = item.row.children[4];
    itemCell.textContent = item.resolved ? item.resolved.stock_id : '';
    itemCell.title = item.resolved ? item.resolved.item_name : '';
    statusCell.textContent = item.message ? `${STATUS_LABELS[item.status]}: ${item.message}` : STATUS_LABELS[item.status];
    statusCell.className = `s-${item.status}`;
    if (item.status === 'checking' || item.status === 'working') {
      item.row.scrollIntoView({ block: 'nearest' });
    }
  }

  function setStatus(item, status, message = '') {
    item.status = status;
    item.message = message;
    updateRow(item);
    updateProgress();
  }

  function updateProgress() {
    ui.progress.hidden = state.items.length === 0;
    const finished = count('done', 'notfound');
    ui.progressBar.style.width = state.items.length ? `${(finished / state.items.length) * 100}%` : '0';
  }

  function setRunningUi(running) {
    state.running = running;
    const hasWork = state.items.some((item) => !['done', 'notfound'].includes(item.status));
    ui.check.disabled = running || count('pending') === 0;
    ui.start.disabled = running || !hasWork;
    ui.pause.disabled = !running;
    ui.stop.disabled = !running;
    ui.file.disabled = running;
    ui.loadPaste.disabled = running;
    ui.pause.textContent = state.paused ? 'Resume' : 'Pause';
  }

  async function waitWhilePaused() {
    while (state.paused && !state.stopRequested) {
      await new Promise((resolve) => setTimeout(resolve, 200));
    }
  }

  async function resolvePending() {
    const queue = state.items.filter((item) => item.status === 'pending');
    let fatalMessage = null;
    async function worker() {
      while (!state.stopRequested && !fatalMessage) {
        await waitWhilePaused();
        // Take the next item only after the await: the other workers share this queue.
        const item = queue.shift();
        if (!item) {
          return;
        }
        setStatus(item, 'checking');
        const result = await AXF.filler.resolveBarcode(item.barcode);
        if (result.ok) {
          item.resolved = result.item;
          setStatus(item, 'ready', result.note || result.item.item_name);
        } else if (result.fatal) {
          fatalMessage = result.message;
          setStatus(item, 'pending');
        } else {
          setStatus(item, 'notfound', result.message);
        }
      }
    }
    await Promise.all(Array.from({ length: RESOLVE_CONCURRENCY }, worker));
    return fatalMessage;
  }

  async function fillReady() {
    for (const item of state.items) {
      if (!['ready', 'failed'].includes(item.status)) {
        continue;
      }
      await waitWhilePaused();
      if (state.stopRequested) {
        return null;
      }
      setStatus(item, 'working');
      const result = await AXF.filler.fillItem(item, (message) => setStatus(item, 'working', message));
      if (result.ok) {
        setStatus(item, 'done');
      } else {
        setStatus(item, 'failed', result.message);
        if (result.fatal) {
          return result.message;
        }
      }
    }
    return null;
  }

  async function run({ fill }) {
    if (fill && !AXF.filler.findEntryRow()) {
      ui.summary.textContent = 'The item entry row is not on screen. Pick the supplier first, then try again.';
      return;
    }
    state.stopRequested = false;
    state.paused = false;
    setRunningUi(true);
    ui.copyFailed.hidden = true;

    let stopMessage = await resolvePending();
    if (!stopMessage && fill && !state.stopRequested) {
      stopMessage = await fillReady();
    }

    setRunningUi(false);
    const lines = [];
    if (stopMessage) {
      lines.push(`Stopped: ${stopMessage}`);
    } else if (state.stopRequested) {
      lines.push('Stopped by you.');
    }
    if (fill) {
      lines.push(`Added ${count('done')} of ${state.items.length}.`);
    } else {
      lines.push(`Found ${count('ready')} of ${state.items.length}.`);
    }
    if (count('notfound')) {
      lines.push(`${count('notfound')} barcodes not found — fix them in the sheet.`);
    }
    if (count('failed')) {
      lines.push(`${count('failed')} lines failed — press "Retry failed" to try them again.`);
    }
    if (fill && count('done')) {
      lines.push('Check the lines in the form, then press Save yourself.');
    }
    ui.summary.textContent = lines.join('\n');
    ui.copyFailed.hidden = count('notfound', 'failed') === 0;
    ui.start.textContent = count('failed') ? 'Retry failed' : 'Start filling';
  }

  function startRun(options) {
    run(options).catch((error) => {
      setRunningUi(false);
      ui.summary.textContent = `Stopped: ${error.message}`;
    });
  }

  ui.fab.addEventListener('click', () => {
    ui.panel.hidden = !ui.panel.hidden;
  });
  ui.close.addEventListener('click', () => {
    ui.panel.hidden = true;
  });

  ui.file.addEventListener('change', async () => {
    const file = ui.file.files && ui.file.files[0];
    if (!file) {
      return;
    }
    try {
      setItems(
        await AXF.parse.parseFile(file, { mergeDuplicates: ui.merge.checked }),
        'No barcodes found in the first sheet. Put barcodes in column A and quantities in column B.'
      );
    } catch (error) {
      ui.problems.textContent = `Could not read the file: ${error.message}`;
    }
    ui.file.value = '';
  });

  ui.loadPaste.addEventListener('click', () => {
    setItems(
      AXF.parse.parsePastedText(ui.paste.value, { mergeDuplicates: ui.merge.checked }),
      ui.paste.value.trim()
        ? 'No barcodes found in the pasted text. Column A must be the barcode, column B the quantity.'
        : 'The paste box is empty. In Excel, select the barcode and quantity cells, copy (Ctrl+C), then paste here (Ctrl+V).'
    );
  });

  ui.check.addEventListener('click', () => startRun({ fill: false }));
  ui.start.addEventListener('click', () => startRun({ fill: true }));

  ui.pause.addEventListener('click', () => {
    state.paused = !state.paused;
    ui.pause.textContent = state.paused ? 'Resume' : 'Pause';
  });

  ui.stop.addEventListener('click', () => {
    state.stopRequested = true;
    state.paused = false;
  });

  ui.copyFailed.addEventListener('click', async () => {
    const text = state.items
      .filter((item) => item.status === 'notfound' || item.status === 'failed')
      .map((item) => `${item.barcode}\t${item.qty}\t${item.message}`)
      .join('\n');
    try {
      await navigator.clipboard.writeText(text);
      ui.copyFailed.textContent = 'Copied — paste into Excel';
    } catch (error) {
      ui.paste.value = text;
      ui.copyFailed.textContent = 'Put in the paste box';
    }
    setTimeout(() => { ui.copyFailed.textContent = 'Copy failed rows'; }, 2500);
  });

  // The ERP is a single-page app: show the button only while an item entry row exists on the page.
  function syncVisibility() {
    const available = state.running || AXF.filler.hasEntryRow();
    ui.fab.hidden = !available;
    if (!available) {
      ui.panel.hidden = true;
    }
  }

  // The toolbar popup asks this page to open the panel.
  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (!message || message.type !== 'axf-open-panel') {
      return;
    }
    const available = state.running || AXF.filler.hasEntryRow();
    if (available) {
      ui.fab.hidden = false;
      ui.panel.hidden = false;
    }
    sendResponse({ available });
  });

  document.documentElement.appendChild(host);
  syncVisibility();
  setInterval(syncVisibility, 1000);
})();
