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
