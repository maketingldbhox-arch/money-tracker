/**
 * ຈົດລາຍຮັບ-ລາຍຈ່າຍ — Google Apps Script back-end
 * ຖານຂໍ້ມູນ: Google Sheets (ສ້າງແຜ່ນງານໃຫ້ເອງອັດຕະໂນມັດ)
 * AI: Gemini API (ຕັ້ງ GEMINI_API_KEY ໃນ Project Settings › Script Properties)
 */

const SHEET_TX = 'Transactions';
const SHEET_SET = 'Settings';
const SHEET_RULES = 'Rules';
const TX_HEAD = ['id', 'date', 'time', 'type', 'category', 'amount', 'currency', 'rate', 'amount_lak', 'note', 'ref', 'source', 'counterparty', 'created_at'];
const SET_DEFAULT = { rate_THB: 650, rate_USD: 21800, rate_CNY: 3000, owner_name: '' };
const CURRENCIES = ['LAK', 'THB', 'USD', 'CNY'];

const EXP_CATS = {
  coffee: 'ກາເຟ', breakfast: 'ອາຫານເຊົ້າ', lunch: 'ອາຫານສວາຍ', dinner: 'ອາຫານຄ່ຳ', snack: 'ເຂົ້າໜົມ - ເຄື່ອງດື່ມ',
  market: 'ອອກຕະຫຼາດ', fuel: 'ນ້ຳມັນລົດ', border: 'ຄ່າຜ່ານແດນ ຕມ', shipping: 'ຄ່າຂົນສົ່ງ - ພັດສະດຸ', hair: 'ຕັດຜົມ',
  cosmetic: 'ເຄື່ອງສຳອາງ', health: 'ສຸຂະພາບ', online: 'Shopping Online', topup: 'ຕື່ມເງິນໂທລະສັບ', general: 'ໃຊ້ຈ່າຍທົ່ວໄປ', other: 'ອື່ນໆ'
};
const INC_CATS = {
  salary: 'ເງິນເດືອນ', side: 'ລາຍຮັບເສີມ', bonus: 'ໂບນັດ', interest: 'ດອກເບ້ຍ', gift: 'ຂອງຂວັນ',
  pass: 'ເງິນຜ່ານມື (ບໍ່ນັບເປັນລາຍຮັບ)', other_inc: 'ລາຍຮັບອື່ນໆ'
};

/* ---------- Web app ---------- */
/* ---------- JSON API for the GitHub Pages front-end ----------
 * POST body (text/plain): {"token":"...","fn":"getInit","args":[...]}
 * Set APP_TOKEN in Project Settings › Script Properties.
 */
const API_FUNCS = { getInit: getInit, addTx: addTx, addTxBatch: addTxBatch, deleteTx: deleteTx, saveSettings: saveSettings, saveRule: saveRule, readSlip: readSlip, aiReport: aiReport };

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_({ ok: true, app: 'money-tracker-api' });
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const token = PropertiesService.getScriptProperties().getProperty('APP_TOKEN');
    if (!token || req.token !== token) return json_({ ok: false, error: 'ລະຫັດເຂົ້າໃຊ້ບໍ່ຖືກຕ້ອງ' });
    const fn = API_FUNCS[req.fn];
    if (!fn) return json_({ ok: false, error: 'Unknown function: ' + req.fn });
    return json_({ ok: true, data: fn.apply(null, req.args || []) });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

/* ---------- Sheet helpers ---------- */
const DEFAULT_SHEET_ID = '1raaJf_koTVPNnmwxK1blFaN7X__Xa5W24vRIVIi8hi0'; // Google Sheet "ຈົດລາຍຮັບ-ລາຍຈ່າຍ (Data)"

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID') || DEFAULT_SHEET_ID;
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name, head) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#021e3a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange('A:Z').setNumberFormat('@');
  }
  return sh;
}

function txSheet_() { return sheet_(SHEET_TX, TX_HEAD); }

function readRows_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  return sh.getRange(2, 1, last - 1, head.length).getDisplayValues().map(function (r) {
    const o = {};
    head.forEach(function (h, i) { o[h] = r[i]; });
    return o;
  });
}

function getSettings_() {
  const sh = sheet_(SHEET_SET, ['key', 'value']);
  const out = Object.assign({}, SET_DEFAULT);
  readRows_(sh).forEach(function (r) { if (r.key) out[r.key] = r.value; });
  ['rate_THB', 'rate_USD', 'rate_CNY'].forEach(function (k) { out[k] = Number(out[k]) || SET_DEFAULT[k]; });
  return out;
}

function rateOf_(cur, settings) {
  if (cur === 'LAK') return 1;
  return Number(settings['rate_' + cur]) || SET_DEFAULT['rate_' + cur] || 1;
}

function cleanTx_(t, settings) {
  const type = t.type === 'inc' ? 'inc' : 'exp';
  const cats = type === 'inc' ? INC_CATS : EXP_CATS;
  const category = cats[t.category] ? t.category : (type === 'inc' ? 'other_inc' : 'other');
  const currency = CURRENCIES.indexOf(t.currency) >= 0 ? t.currency : 'LAK';
  const amount = Math.abs(Number(t.amount) || 0);
  if (!amount) throw new Error('ຈຳນວນເງິນຕ້ອງຫຼາຍກວ່າ 0');
  const rate = rateOf_(currency, settings);
  const date = /^\d{4}-\d{2}-\d{2}$/.test(t.date || '') ? t.date : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const time = /^\d{2}:\d{2}$/.test(t.time || '') ? t.time : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');
  return {
    id: Utilities.getUuid(), date: date, time: time, type: type, category: category,
    amount: String(amount), currency: currency, rate: String(rate), amount_lak: String(Math.round(amount * rate)),
    note: String(t.note || '').slice(0, 300), ref: String(t.ref || '').slice(0, 80), source: t.source === 'slip' ? 'slip' : 'manual',
    counterparty: String(t.counterparty || '').slice(0, 120), created_at: new Date().toISOString()
  };
}

function toRow_(o) { return TX_HEAD.map(function (h) { return o[h] === undefined ? '' : o[h]; }); }

/* ---------- API (called by google.script.run) ---------- */
function getInit() {
  const props = PropertiesService.getScriptProperties();
  return {
    tx: readRows_(txSheet_()),
    settings: getSettings_(),
    rules: readRows_(sheet_(SHEET_RULES, ['match', 'type', 'category'])),
    aiKey: !!props.getProperty('GEMINI_API_KEY'),
    sheetUrl: ss_().getUrl(),
    demo: false
  };
}

function addTx(t) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const o = cleanTx_(t, getSettings_());
    txSheet_().appendRow(toRow_(o));
    return o;
  } finally { lock.releaseLock(); }
}

function addTxBatch(list) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const settings = getSettings_();
    const sh = txSheet_();
    const refs = {};
    readRows_(sh).forEach(function (r) { if (r.ref) refs[r.ref] = true; });
    const saved = [];
    let skipped = 0;
    (list || []).forEach(function (t) {
      if (t.ref && refs[t.ref]) { skipped++; return; }
      const o = cleanTx_(t, settings);
      if (o.ref) refs[o.ref] = true;
      saved.push(o);
    });
    if (saved.length) sh.getRange(sh.getLastRow() + 1, 1, saved.length, TX_HEAD.length).setValues(saved.map(toRow_));
    return { saved: saved, skipped: skipped };
  } finally { lock.releaseLock(); }
}

function deleteTx(id) {
  const sh = txSheet_();
  const last = sh.getLastRow();
  if (last < 2) return false;
  const ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (ids[i][0] === id) { sh.deleteRow(i + 2); return true; }
  }
  return false;
}

function saveSettings(obj) {
  const sh = sheet_(SHEET_SET, ['key', 'value']);
  const cur = getSettings_();
  Object.keys(SET_DEFAULT).forEach(function (k) { if (obj && obj[k] !== undefined) cur[k] = obj[k]; });
  const rows = Object.keys(SET_DEFAULT).map(function (k) { return [k, String(cur[k])]; });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  sh.getRange(2, 1, rows.length, 2).setValues(rows);
  return getSettings_();
}

function saveRule(rule) {
  if (!rule || !rule.match) return readRows_(sheet_(SHEET_RULES, ['match', 'type', 'category']));
  const sh = sheet_(SHEET_RULES, ['match', 'type', 'category']);
  const rows = readRows_(sh);
  const key = String(rule.match).trim().toUpperCase();
  const idx = rows.findIndex(function (r) { return String(r.match).toUpperCase() === key; });
  const row = [key, rule.type === 'inc' ? 'inc' : 'exp', String(rule.category || '')];
  if (idx >= 0) sh.getRange(idx + 2, 1, 1, 3).setValues([row]);
  else sh.appendRow(row);
  return readRows_(sh);
}

/* ---------- Gemini ---------- */
/* Fast models first. GEMINI_MODEL (Script Property) overrides; the last model that worked is remembered. */
const GEMINI_FALLBACK = ['gemini-3.5-flash-lite', 'gemini-flash-lite-latest', 'gemini-3.1-flash-lite', 'gemini-3.5-flash', 'gemini-flash-latest'];

function thinkingFor_(model) {
  if (/2\.5/.test(model)) return { thinkingBudget: 0 };          // Gemini 2.5: turn thinking off
  if (/flash|lite/.test(model)) return { thinkingLevel: 'minimal' }; // newer models: least thinking
  return null;
}

function callModel_(key, model, parts, asJson, withThinking) {
  const gen = asJson ? { responseMimeType: 'application/json', temperature: 0.1 } : { temperature: 0.4 };
  const th = withThinking ? thinkingFor_(model) : null;
  if (th) gen.thinkingConfig = th;
  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent', {
    method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key },
    payload: JSON.stringify({ contents: [{ role: 'user', parts: parts }], generationConfig: gen }), muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  let data = {};
  try { data = JSON.parse(res.getContentText()); } catch (e) { }
  if (code === 200) {
    const c = data.candidates && data.candidates[0];
    return { ok: true, text: c && c.content && c.content.parts ? c.content.parts.map(function (p) { return p.text || ''; }).join('') : '' };
  }
  return { ok: false, code: code, error: (data.error && data.error.message) || ('Gemini error ' + code) };
}

function gemini_(parts, asJson) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('ຍັງບໍ່ໄດ້ຕັ້ງ GEMINI_API_KEY ໃນ Script Properties');
  const cache = CacheService.getScriptCache();
  const remembered = cache.get('gemini_ok_model');
  const models = [props.getProperty('GEMINI_MODEL'), remembered].concat(GEMINI_FALLBACK)
    .filter(function (m, i, a) { return !!m && a.indexOf(m) === i; });
  let lastErr = '';
  for (let i = 0; i < models.length; i++) {
    let r = callModel_(key, models[i], parts, asJson, true);
    if (!r.ok && r.code === 400 && /think/i.test(r.error)) r = callModel_(key, models[i], parts, asJson, false);
    if (r.ok) { if (remembered !== models[i]) cache.put('gemini_ok_model', models[i], 21600); return r.text; }
    lastErr = r.error;
    if ([400, 404, 429, 500, 503].indexOf(r.code) < 0) break; // next model when this one is unavailable or busy
  }
  throw new Error(lastErr);
}

/* Run from the editor to see which Gemini models this key can use and how fast each answers. */
function geminiBenchmark() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const list = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', { headers: { 'x-goog-api-key': key }, muteHttpExceptions: true });
  const names = ((JSON.parse(list.getContentText()).models) || [])
    .filter(function (m) { return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0; })
    .map(function (m) { return m.name.replace('models/', ''); })
    .filter(function (n) { return /flash/.test(n); });
  Logger.log('FLASH MODELS: ' + names.join(', '));
  GEMINI_FALLBACK.forEach(function (m) {
    const t = Date.now();
    let r = callModel_(key, m, [{ text: 'Reply with the word OK' }], false, true);
    if (!r.ok && r.code === 400 && /think/i.test(r.error)) r = callModel_(key, m, [{ text: 'Reply with the word OK' }], false, false);
    Logger.log(m + ': ' + (r.ok ? (Date.now() - t) + ' ms' : 'ERR ' + r.code + ' ' + r.error));
  });
}

function parseJson_(text) {
  const t = String(text || '').replace(/^```(?:json)?/i, '').replace(/```$/, '').trim();
  return JSON.parse(t);
}

function readSlip(b64, mime, ownerName) {
  const expList = Object.keys(EXP_CATS).map(function (k) { return k + ' = ' + EXP_CATS[k]; }).join('; ');
  const incList = Object.keys(INC_CATS).map(function (k) { return k + ' = ' + INC_CATS[k]; }).join('; ');
  const prompt = [
    'You read one screenshot of a Lao/Thai bank transfer slip or card payment notice. The app user owns the account.',
    ownerName ? ('The account owner name is: ' + ownerName + '.') : '',
    'Return ONLY a JSON object with these keys:',
    'amount (number, no commas), currency (LAK, THB, USD or CNY; ₭/kip/ກີບ=LAK, ฿/บาท/ບາດ=THB, $=USD, ¥/元/RMB=CNY),',
    'amount_color ("red", "green" or "other": the color of the main amount text),',
    'direction ("out" if the owner paid/sent money, "in" if the owner received money, "unknown"),',
    'date ("YYYY-MM-DD"), time ("HH:mm"), counterparty (the merchant or other person, NOT the owner),',
    'memo (description/purpose text exactly as written), ref (reference, transaction or approval number),',
    'suggested_type ("exp" or "inc": red amount means exp, green means inc; if color is other, use direction),',
    'suggested_category (exactly one id. For exp choose from: ' + expList + '. For inc choose from: ' + incList + '),',
    'reason_lo (one short sentence in Lao: what this payment looks like, then ask the user which category it is).',
    'Use "" for anything not visible and 0 if no amount.'
  ].join('\n');
  const text = gemini_([{ text: prompt }, { inline_data: { mime_type: mime || 'image/jpeg', data: b64 } }], true);
  const o = parseJson_(text);
  o.amount = Math.abs(Number(String(o.amount || 0).replace(/,/g, ''))) || 0;
  if (CURRENCIES.indexOf(o.currency) < 0) o.currency = 'LAK';
  o.suggested_type = o.suggested_type === 'inc' ? 'inc' : 'exp';
  const cats = o.suggested_type === 'inc' ? INC_CATS : EXP_CATS;
  if (!cats[o.suggested_category]) o.suggested_category = o.suggested_type === 'inc' ? 'other_inc' : 'other';
  return o;
}

function aiReport(summary, question) {
  const prompt = [
    'ເຈົ້າເປັນຜູ້ຊ່ວຍການເງິນສ່ວນຕົວ ຕອບເປັນພາສາລາວທີ່ອ່ານງ່າຍ ສັ້ນ ກະທັດຮັດ.',
    'ໃຊ້ສະເພາະຕົວເລກທີ່ມີໃນຂໍ້ມູນ ຫ້າມແຕ່ງຕົວເລກຂຶ້ນເອງ. ຈຳນວນເງິນເປັນກີບ (₭). ລາຍການ "ເງິນຜ່ານມື" ບໍ່ນັບເປັນລາຍຮັບ.',
    'ຂໍ້ມູນ (JSON): ' + JSON.stringify(summary),
    question
      ? ('ຄຳຖາມຂອງຜູ້ໃຊ້: ' + String(question).slice(0, 500) + '\nຕອບຄຳຖາມນີ້ໂດຍກົງ ບໍ່ເກີນ 6 ແຖວ.')
      : 'ຂຽນລາຍງານປະຈຳເດືອນ ມີ 4 ຫົວຂໍ້: ສະຫຼຸບ, ບັນຫາທີ່ພົບ (ໝວດທີ່ໃຊ້ຫຼາຍ ຫຼື ເພີ່ມຂຶ້ນຈາກເດືອນກ່ອນ), ສິ່ງທີ່ເຮັດໄດ້ດີ, ຄຳແນະນຳ 3 ຂໍ້ທີ່ເຮັດໄດ້ຈິງ. ໃຊ້ **ຕົວໜາ** ສຳລັບຫົວຂໍ້ ແລະ "- " ສຳລັບລາຍການ.'
  ].join('\n');
  return gemini_([{ text: prompt }], false);
}

/* ---------- ໃຊ້ຄັ້ງທຳອິດ: ກົດ Run ຟັງຊັນນີ້ເພື່ອສ້າງແຜ່ນງານ ---------- */
function setup() {
  txSheet_();
  sheet_(SHEET_RULES, ['match', 'type', 'category']);
  saveSettings({});
  return 'OK';
}
