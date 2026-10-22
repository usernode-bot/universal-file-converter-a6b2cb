// The pages around the converter: Riwayat (history kept on this device),
// Profil and Edit File. Only the file's names, formats and sizes are kept;
// never its contents.
const HISTORY_KEY = 'conversion-history';
const HISTORY_MAX = 50;

export function loadHistory() {
  try {
    const list = JSON.parse(localStorage.getItem(HISTORY_KEY) || '[]');
    return Array.isArray(list) ? list : [];
  } catch {
    return [];
  }
}

export function addHistory(entry) {
  const list = [entry, ...loadHistory()].slice(0, HISTORY_MAX);
  try { localStorage.setItem(HISTORY_KEY, JSON.stringify(list)); } catch { /* storage refused */ }
}

export function clearHistory() {
  try { localStorage.removeItem(HISTORY_KEY); } catch { /* storage refused */ }
}

function row(tile, title, meta, result) {
  const li = document.createElement('li');
  li.className = 'list-row';
  const t = document.createElement('span');
  t.className = 'doc-tile';
  if (result) t.dataset.result = '';
  t.textContent = tile;
  const box = document.createElement('div');
  box.className = 'min-w-0 flex-1';
  const p1 = document.createElement('p');
  p1.className = 'truncate text-body font-medium';
  p1.textContent = title;
  const p2 = document.createElement('p');
  p2.className = 'text-small text-muted';
  p2.textContent = meta;
  box.append(p1, p2);
  li.append(t, box);
  return li;
}

const WHEN = new Intl.DateTimeFormat('id-ID', { dateStyle: 'medium', timeStyle: 'short' });

export function renderHistory() {
  const list = loadHistory();
  const ul = document.getElementById('history-list');
  ul.innerHTML = '';
  for (const e of list) {
    ul.appendChild(row(e.ext, e.name, `Dari ${e.from} · ${e.size} · ${WHEN.format(new Date(e.at))}`, true));
  }
  ul.hidden = !list.length;
  document.getElementById('history-empty').hidden = !!list.length;
  document.getElementById('history-clear').hidden = !list.length;
}

let profileToken = 0;

export async function renderProfile() {
  const mine = ++profileToken;
  const $ = (id) => document.getElementById(id);
  $('profile-loading').hidden = false;
  $('profile-error').hidden = true;
  $('profile-list').hidden = true;
  $('profile-guest').hidden = true;
  try {
    const token = new URLSearchParams(location.search).get('token');
    const res = await fetch('/api/me', token ? { headers: { 'x-usernode-token': token } } : {});
    if (!res.ok) throw new Error('status ' + res.status);
    const data = await res.json();
    if (mine !== profileToken) return;
    $('profile-loading').hidden = true;
    if (!data.user) {
      $('profile-guest').hidden = false;
      return;
    }
    const ul = $('profile-list');
    ul.innerHTML = '';
    const item = (label, value) => {
      const li = document.createElement('li');
      li.className = 'list-row justify-between';
      const a = document.createElement('span');
      a.className = 'text-body text-muted';
      a.textContent = label;
      const b = document.createElement('span');
      b.className = 'text-body font-medium';
      b.textContent = value;
      li.append(a, b);
      ul.appendChild(li);
    };
    item('Nama pengguna', data.user.username);
    item('Dompet', data.user.walletLinked ? 'Tertaut' : 'Belum tertaut');
    item('Konversi di perangkat ini', String(loadHistory().length));
    ul.hidden = false;
  } catch {
    if (mine !== profileToken) return;
    $('profile-loading').hidden = true;
    $('profile-error').hidden = false;
  }
}
