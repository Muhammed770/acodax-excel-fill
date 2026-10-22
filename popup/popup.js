const status = document.getElementById('status');

document.getElementById('open').addEventListener('click', async () => {
  status.textContent = '';
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  try {
    const reply = await chrome.tabs.sendMessage(tab.id, { type: 'axf-open-panel' });
    if (reply && reply.available) {
      window.close();
      return;
    }
    status.textContent = 'No item rows on this page. Open Add Purchase Invoice or Add Stock Transfer and pick the supplier first.';
  } catch (error) {
    status.textContent = 'This is not an Acodax page. Open Acodax first (reload the page if you just installed the extension).';
  }
});

document.getElementById('sample').addEventListener('click', () => {
  // Barcode cells are text so Excel keeps leading zeros.
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Barcode', 'Quantity'],
    [{ t: 's', v: '1234567890123', z: '@' }, 2],
    [{ t: 's', v: '9876543210987', z: '@' }, 3],
  ]);
  sheet['!cols'] = [{ wch: 20 }, { wch: 10 }];
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, 'Items');
  XLSX.writeFile(book, 'acodax-excel-fill-sample.xlsx');
});
