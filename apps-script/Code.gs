/**
 * INSULAR — ponte entre o app e esta planilha Google.
 *
 * PASSO A PASSO
 * 1) Troque o valor de TOKEN abaixo por uma senha sua (a mesma vai no app).
 * 2) Salve (Ctrl+S). No seletor de funções escolha "setup" e clique em Executar;
 *    autorize o acesso (Avançado > Acessar Insular). Isso cria as abas "Registros"
 *    e "Ajustes", com os dois dias da folha de papel já lançados.
 * 3) Implantar > Nova implantação > tipo "App da web":
 *       Executar como: "Eu"      |      Quem pode acessar: "Qualquer pessoa"
 *    Clique em Implantar e copie a URL (termina em /exec). Ela vai no app.
 * 4) Sempre que alterar este código, faça Implantar > Gerenciar implantações >
 *    (lápis) > Versão: "Nova versão" > Implantar. Sem isso a URL continua
 *    rodando o código antigo.
 */

const TOKEN = 'troque-esta-senha';

const SHEET_DAYS = 'Registros';
const SHEET_SETTINGS = 'Ajustes';

// Ordem das colunas da aba Registros (não altere a ordem; o app depende dela).
const KEYS = [
  'date',
  'jejum_hora', 'jejum', 'ins_manha_ui', 'ins_manha_hora', 'cafe_hora', 'apos_cafe_hora', 'apos_cafe',
  'pre_almoco_hora', 'pre_almoco', 'almoco_hora', 'apos_almoco_hora', 'apos_almoco',
  'pre_jantar_hora', 'pre_jantar', 'ins_jantar_ui', 'ins_jantar_hora', 'jantar_hora', 'apos_jantar_hora', 'apos_jantar',
  'dormir_hora', 'dormir', 'madrugada_hora', 'madrugada',
  'obs', 'atualizado_em'
];
const HEADERS = [
  'Data',
  'Jejum – hora', 'Jejum (mg/dL)', 'Insulina manhã (UI)', 'Insulina manhã – hora', 'Café – hora', '1h após café – hora', '1h após café (mg/dL)',
  'Antes almoço – hora', 'Antes almoço (mg/dL)', 'Almoço – hora', '1h após almoço – hora', '1h após almoço (mg/dL)',
  'Antes jantar – hora', 'Antes jantar (mg/dL)', 'Insulina jantar (UI)', 'Insulina jantar – hora', 'Jantar – hora', '1h após jantar – hora', '1h após jantar (mg/dL)',
  'Dormir – hora', 'Dormir (mg/dL)', 'Madrugada – hora', 'Madrugada (mg/dL)',
  'Observações', 'Atualizado em (ms)'
];
const NUMERIC = ['jejum', 'ins_manha_ui', 'apos_cafe', 'pre_almoco', 'apos_almoco', 'pre_jantar', 'ins_jantar_ui', 'apos_jantar', 'dormir', 'madrugada', 'atualizado_em'];

// Metas usadas só para colorir a planilha (o app tem as suas em Ajustes).
const META_JEJUM = 90, META_PRE = 100, META_POS = 140, HIPO = 70;

/* ---------------- configuração inicial (rodar uma vez) ---------------- */
function setup() {
  const sh = daysSheet();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setWrap(true);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);
  formatColumns(sh);
  applyColors(sh);
  sh.setColumnWidths(1, HEADERS.length, 92);
  sh.setColumnWidth(1, 100);
  sh.setColumnWidth(KEYS.indexOf('obs') + 1, 220);

  const st = settingsSheet();
  st.getRange(1, 1, 1, 2).setValues([['Chave', 'Valor']]).setFontWeight('bold');
  st.setFrozenRows(1);

  // Dias transcritos da folha de papel (apague as linhas na planilha se não quiser).
  upsertDay({ date: '2026-11-01', jejum_hora: '08:12', jejum: 97, apos_cafe: 142, pre_almoco_hora: '13:33', pre_almoco: 84, apos_almoco: 118, pre_jantar_hora: '22:14', pre_jantar: 87, apos_jantar: 95, obs: 'Transcrito da folha de papel.', atualizado_em: 1793745600000 });
  upsertDay({ date: '2026-11-02', jejum_hora: '08:15', jejum: 80, apos_cafe: 96, pre_almoco_hora: '13:34', pre_almoco: 67, apos_almoco: 124, pre_jantar_hora: '21:21', pre_jantar: 111, apos_jantar: 126, obs: 'Transcrito da folha de papel.', atualizado_em: 1793745600000 });
  writeSettings({ nome: 'Daniele', insulina: 'NPH', doseManha: 22, horaInsManha: '07:00', doseJantar: 22, metaJejum: META_JEJUM, metaPre: META_PRE, metaPos: META_POS, hipo: HIPO, mostrarNoite: true, updatedAt: 1793745600000 });
}

function formatColumns(sh) {
  const rows = Math.max(sh.getMaxRows(), 2);
  KEYS.forEach((k, i) => {
    const col = i + 1;
    const rg = sh.getRange(1, col, rows, 1);
    if (NUMERIC.indexOf(k) >= 0) rg.setNumberFormat('0'); else rg.setNumberFormat('@');
  });
}

function applyColors(sh) {
  const rules = [];
  const last = 1000;
  const col = k => KEYS.indexOf(k) + 1;
  const rng = k => sh.getRange(2, col(k), last - 1, 1);
  const high = (keys, lim) => SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(lim).setBackground('#FCE3DF').setFontColor('#B3261E')
    .setRanges(keys.map(rng)).build();
  const low = (keys) => SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(HIPO).setBackground('#E2E9FB').setFontColor('#2F4FB8')
    .setRanges(keys.map(rng)).build();
  const ok = (keys, lim) => SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(HIPO, lim - 1).setBackground('#DEF4E8').setFontColor('#1B7A50')
    .setRanges(keys.map(rng)).build();
  rules.push(low(['jejum', 'pre_almoco', 'pre_jantar', 'apos_cafe', 'apos_almoco', 'apos_jantar', 'dormir', 'madrugada']));
  rules.push(high(['jejum'], META_JEJUM), ok(['jejum'], META_JEJUM));
  rules.push(high(['pre_almoco', 'pre_jantar'], META_PRE), ok(['pre_almoco', 'pre_jantar'], META_PRE));
  rules.push(high(['apos_cafe', 'apos_almoco', 'apos_jantar'], META_POS), ok(['apos_cafe', 'apos_almoco', 'apos_jantar'], META_POS));
  sh.setConditionalFormatRules(rules);
}

/* ---------------- entrada do app ---------------- */
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  return respond(handle({ token: p.token, action: 'all' }));
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return respond({ ok: false, error: 'json' }); }
  return respond(handle(body));
}

function handle(req) {
  if (!req || String(req.token) !== TOKEN) return { ok: false, error: 'token' };
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const action = req.action || 'all';
    if (action === 'sync') {
      (req.days || []).forEach(upsertDay);
      (req.deletes || []).forEach(d => deleteDay(d.date, d.deletedAt));
      if (req.settings) writeSettings(req.settings);
    } else if (action !== 'all') {
      return { ok: false, error: 'action' };
    }
    return { ok: true, days: readDays(), settings: readSettings(), serverTime: Date.now() };
  } finally {
    lock.releaseLock();
  }
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- aba Registros ---------------- */
function sheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function daysSheet() {
  const sh = sheet(SHEET_DAYS);
  if (sh.getLastRow() < 1) { sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold'); formatColumns(sh); }
  return sh;
}
function findRow(sh, date) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const col = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < col.length; i++) if (String(col[i][0]).trim() === date) return i + 2;
  return -1;
}
function rowFrom(day) {
  return KEYS.map(k => {
    const v = day[k];
    if (v === undefined || v === null || v === '') return '';
    if (NUMERIC.indexOf(k) >= 0) { const n = Number(v); return isNaN(n) ? '' : n; }
    return String(v);
  });
}
function formatRow(sh, row) {
  KEYS.forEach((k, i) => sh.getRange(row, i + 1).setNumberFormat(NUMERIC.indexOf(k) >= 0 ? '0' : '@'));
}
function upsertDay(day) {
  if (!day || !/^\d{4}-\d\d-\d\d$/.test(String(day.date))) return;
  const sh = daysSheet();
  const date = String(day.date);
  const incoming = Number(day.atualizado_em) || 0;
  const row = findRow(sh, date);
  if (row > 0) {
    const current = Number(sh.getRange(row, KEYS.indexOf('atualizado_em') + 1).getDisplayValue()) || 0;
    if (incoming < current) return; // a planilha já tem uma versão mais nova
    formatRow(sh, row);
    sh.getRange(row, 1, 1, KEYS.length).setValues([rowFrom(day)]);
  } else {
    const r = sh.getLastRow() + 1;
    if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
    formatRow(sh, r);
    sh.getRange(r, 1, 1, KEYS.length).setValues([rowFrom(day)]);
  }
}
function deleteDay(date, deletedAt) {
  const sh = daysSheet();
  const row = findRow(sh, String(date));
  if (row < 0) return;
  const current = Number(sh.getRange(row, KEYS.indexOf('atualizado_em') + 1).getDisplayValue()) || 0;
  if ((Number(deletedAt) || 0) < current) return; // foi editado depois de apagado em outro aparelho
  sh.deleteRow(row);
}
function readDays() {
  const sh = daysSheet();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, KEYS.length).getDisplayValues();
  const out = [];
  vals.forEach(r => {
    const o = {};
    KEYS.forEach((k, i) => { o[k] = r[i]; });
    if (/^\d{4}-\d\d-\d\d$/.test(String(o.date).trim())) { o.date = String(o.date).trim(); out.push(o); }
  });
  out.sort((a, b) => (a.date < b.date ? -1 : 1));
  return out;
}

/* ---------------- aba Ajustes ---------------- */
function settingsSheet() {
  const sh = sheet(SHEET_SETTINGS);
  if (sh.getLastRow() < 1) sh.getRange(1, 1, 1, 2).setValues([['Chave', 'Valor']]).setFontWeight('bold');
  return sh;
}
function readSettings() {
  const sh = settingsSheet();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const vals = sh.getRange(2, 1, last - 1, 2).getDisplayValues();
  const o = {};
  vals.forEach(r => { if (!r[0]) return; try { o[r[0]] = JSON.parse(r[1]); } catch (e) { o[r[0]] = r[1]; } });
  return Object.keys(o).length ? o : null;
}
function writeSettings(s) {
  if (!s || typeof s !== 'object') return;
  const cur = readSettings();
  if (cur && (Number(cur.updatedAt) || 0) > (Number(s.updatedAt) || 0)) return;
  const sh = settingsSheet();
  const rows = Object.keys(s).map(k => [k, JSON.stringify(s[k])]);
  if (sh.getMaxRows() > 1) sh.getRange(2, 1, sh.getMaxRows() - 1, 2).clearContent();
  const rg = sh.getRange(2, 1, rows.length, 2);
  rg.setNumberFormat('@');
  rg.setValues(rows);
}
