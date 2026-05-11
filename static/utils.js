const STORAGE_KEYS = {
  TERMINALS: 'atm_terminals',
  CASH_AMOUNTS: 'atm_cash_amounts',
  DRIVER_SELECTION: 'atm_drivers_today',
  TERMINAL_ASSIGNMENTS: 'atm_terminal_assignments',
};

function saveToSession(key, data) {
  sessionStorage.setItem(key, JSON.stringify(data));
}

function loadFromSession(key) {
  try {
    const raw = sessionStorage.getItem(key);
    return raw ? JSON.parse(raw) : null;
  } catch { return null; }
}

function guardStep(requiredKey, redirectTo = '/') {
  if (!loadFromSession(requiredKey)) {
    window.location.href = redirectTo;
    return false;
  }
  return true;
}

function computeDefaultCashToAdd(suggested) {
  const val = Math.max(0, suggested ?? 0);
  return Math.ceil(val / 500) * 500;
}

function urgencyClass(days) {
  if (days === null || days === undefined || days <= 0) return 'urgency-red';
  if (days <= 2) return 'urgency-orange';
  if (days <= 4) return 'urgency-yellow';
  return 'urgency-green';
}

function toTitleCase(str) {
  if (!str) return '—';
  return str.trim().split(' ')
    .map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join(' ');
}

function formatCurrency(n) {
  if (n === null || n === undefined || n === '') return '—';
  return '$' + Number(n).toLocaleString('en-US');
}

function attachSortHeaders(tableEl) {
  let sortCol = -1, sortDir = 1;
  const headers = Array.from(tableEl.querySelectorAll('thead th'));

  function applySort() {
    if (sortCol < 0) return;
    const tbody = tableEl.querySelector('tbody');
    const rows = Array.from(tbody.querySelectorAll('tr'));
    rows.sort((a, b) => {
      const ac = a.cells[sortCol], bc = b.cells[sortCol];
      if (!ac || !bc) return 0;
      const aText = ac.textContent.trim().replace(/[$,]/g, '');
      const bText = bc.textContent.trim().replace(/[$,]/g, '');
      if (aText === 'OVERDUE') return -sortDir;
      if (bText === 'OVERDUE') return sortDir;
      if (aText === '—' || aText === '') return sortDir;
      if (bText === '—' || bText === '') return -sortDir;
      const aN = parseFloat(aText), bN = parseFloat(bText);
      if (!isNaN(aN) && !isNaN(bN)) return (aN - bN) * sortDir;
      return aText.localeCompare(bText) * sortDir;
    });
    rows.forEach(r => tbody.appendChild(r));
  }

  headers.forEach((th, idx) => {
    if ('noSort' in th.dataset) return;
    th.classList.add('th-sortable');
    th.addEventListener('click', function () {
      if (sortCol === idx) {
        sortDir *= -1;
      } else {
        sortCol = idx;
        sortDir = 1;
      }
      headers.forEach(h => h.classList.remove('th-sort-asc', 'th-sort-desc'));
      this.classList.add(sortDir === 1 ? 'th-sort-asc' : 'th-sort-desc');
      applySort();
    });
  });

  return applySort;
}

function triggerBase64Download(b64, filename, mimeType) {
  const bytes = atob(b64);
  const arr = new Uint8Array(bytes.length);
  for (let i = 0; i < bytes.length; i++) arr[i] = bytes.charCodeAt(i);
  const blob = new Blob([arr], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
