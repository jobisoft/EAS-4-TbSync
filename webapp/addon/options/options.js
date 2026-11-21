/**
 * Options page: setup hint while the helper app is missing, and while it is
 * connected the status, the address and token of the HTTP API, the Outlook
 * tabs and the latest requests. It only talks to background.js.
 */

const $ = id => document.getElementById(id);


let token = '';
let tokenVisible = false;

// ── OS specific parts ──────────────────────────────────────────────────────────

const { os } = await browser.runtime.getPlatformInfo();
for (const el of document.querySelectorAll('[data-os]')) {
  el.hidden = os === 'win' ? el.dataset.os !== 'win' : el.dataset.os !== 'unix';
}

// ── Helpers ────────────────────────────────────────────────────────────────────

function setBanner(kind, title, detail = '') {
  $('banner').className = `banner banner-${kind}`;
  $('banner-spinner').hidden = kind !== 'checking';
  $('banner-title').textContent = title;
  $('banner-detail').textContent = detail;
  $('retry-btn').hidden = kind !== 'error';
}

function formatUptime(seconds) {
  const h = Math.floor(seconds / 3600);
  const m = Math.floor((seconds % 3600) / 60);
  return [h && `${h}h`, (h || m) && `${m}m`, `${seconds % 60}s`].filter(Boolean).join(' ');
}

function el(tag, className, text) {
  const e = document.createElement(tag);
  if (className) e.className = className;
  if (text !== undefined) e.textContent = text;
  return e;
}

function showToken() {
  $('token').textContent = tokenVisible ? token : '*'.repeat(token.length);
  $('show-token-btn').textContent = tokenVisible ? 'Hide' : 'Show';
}

async function copy(text, button) {
  const label = button.textContent;
  try {
    await navigator.clipboard.writeText(text);
    button.textContent = 'Copied';
  } catch {
    button.textContent = 'Failed';
  }
  setTimeout(() => { button.textContent = label; }, 1500);
}

// ── Rendering ──────────────────────────────────────────────────────────────────

function renderInfo(reply) {
  const info = reply.info;
  token = info.adminToken;
  tokenVisible = false;
  showToken();
  $('addon-version').textContent = reply.addonVersion;
  $('host-version').textContent = `${info.hostVersion} (Python ${info.pythonVersion}), helper build: ${info.build}`;
  $('uptime').textContent = formatUptime(info.uptime);
  $('config-path').textContent = info.configPath;
  $('api-url').textContent = `http://127.0.0.1:${info.port}`;
  if (info.serverRunning) {
    setBanner('ok', 'The helper app is connected.', `The API listens on 127.0.0.1:${info.port}.`);
  } else {
    setBanner('error', 'The helper app runs, but its server does not.', info.serverError || '');
  }
  const box = $('recent');
  box.replaceChildren();
  if (!info.recent.length) {
    box.append(el('div', 'empty-state', 'No requests yet.'));
  } else {
    const table = el('table');
    for (const r of [...info.recent].reverse()) {
      const tr = el('tr');
      tr.append(el('td', '', new Date(r.time * 1000).toLocaleTimeString()), el('td', '', r.endpoint));
      const state = el('td');
      state.append(el('span', `badge badge-${r.ok ? 'ok' : 'error'}`, r.ok ? 'ok' : 'error'));
      tr.append(state, el('td', 'small', r.error || `${r.ms} ms`));
      table.append(tr);
    }
    box.append(table);
  }
}

async function renderTabs() {
  const reply = await browser.runtime.sendMessage({ cmd: 'outlookStatus' });
  const list = $('tab-list');
  list.replaceChildren();
  for (const tab of reply?.tabs || []) {
    const row = el('div', 'tab');
    row.append(el('span', 'tab-title', tab.title || tab.url));
    row.append(el('span', `badge badge-${tab.contentScript ? 'ok' : 'error'}`,
      tab.contentScript ? 'ready' : 'no content script'));
    list.append(row);
  }
}

async function refresh() {
  setBanner('checking', 'Checking the helper app ...');
  const reply = await browser.runtime.sendMessage({ cmd: 'reconnect' });
  const connected = !!reply?.connected;
  $('setup-section').hidden = connected;
  $('main').hidden = !connected;
  if (!connected) {
    setBanner('error', 'The helper app is not connected.', reply?.error || '');
    return;
  }
  renderInfo(reply);
  await renderTabs();
}

// ── Buttons ────────────────────────────────────────────────────────────────────

$('retry-btn').addEventListener('click', refresh);
$('open-outlook-btn').addEventListener('click', async () => {
  await browser.runtime.sendMessage({ cmd: 'openOutlook' });
  setTimeout(renderTabs, 1500);
});
$('copy-url-btn').addEventListener('click', e => copy($('api-url').textContent, e.target));
$('copy-token-btn').addEventListener('click', e => copy(token, e.target));
$('show-token-btn').addEventListener('click', () => {
  tokenVisible = !tokenVisible;
  showToken();
});
$('new-token-btn').addEventListener('click', async () => {
  if (!confirm('Create a new token? Scripts with the old token stop working.')) return;
  const reply = await browser.runtime.sendMessage({ cmd: 'newToken' });
  if (reply?.connected) renderInfo(reply);
});

await refresh();
