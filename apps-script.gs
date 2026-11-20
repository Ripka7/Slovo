/**
 * Слово за словом — сховище в Google-таблиці.
 * Аркуш "_data": одна строка на документ (meta, sessions, ch_0, ch_1 …), JSON порізаний на шматки по 40 000 символів.
 * Аркуші "Слова" і "Прогрес": читабельна копія, перебудовується після кожної зміни.
 * Історія змін — у самій таблиці: Файл → Історія версій.
 */
const CHUNK = 40000;

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function readDocs_() {
  const sh = sheet_('_data');
  const docs = {};
  if (sh.getLastRow() === 0) return docs;
  sh.getDataRange().getValues().forEach(r => {
    if (!r[0]) return;
    docs[r[0]] = JSON.parse(r.slice(2).join(''));
  });
  return docs;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  try { return out_({ ok: true, docs: readDocs_() }); }
  catch (e) { return out_({ ok: false, error: String(e) }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);
    const id = String(req.id || '');
    if (!/^(meta|sessions|ch_\d+)$/.test(id)) return out_({ ok: false, error: 'bad id' });
    const sh = sheet_('_data');
    const ids = sh.getLastRow() ? sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => r[0]) : [];
    let row = ids.indexOf(id) + 1;
    if (req.body === null) {            // видалення документа
      if (row) sh.deleteRow(row);
    } else {
      const json = JSON.stringify(req.body);
      const parts = json.match(new RegExp('[\\s\\S]{1,' + CHUNK + '}', 'g')) || [''];
      if (sh.getMaxColumns() < parts.length + 2) sh.insertColumnsAfter(sh.getMaxColumns(), parts.length + 2 - sh.getMaxColumns());
      if (!row) row = sh.getLastRow() + 1;
      else sh.getRange(row, 1, 1, sh.getMaxColumns()).clearContent();
      sh.getRange(row, 1, 1, parts.length + 2).setValues([[id, new Date()].concat(parts)]);
    }
    if (id !== 'sessions') rebuildView_();
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function fmt_(ts) {
  return ts ? Utilities.formatDate(new Date(ts), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
}

function rebuildView_() {
  const docs = readDocs_();
  const setName = {};
  ((docs.meta && docs.meta.sets) || []).forEach(s => { setName[s.id] = s.name; });
  const cards = [];
  Object.keys(docs).filter(k => k.indexOf('ch_') === 0).forEach(k => (docs[k].cards || []).forEach(c => cards.push(c)));
  cards.sort((a, b) => (a.created || 0) - (b.created || 0));
  const status = c => c.st === 'new' ? 'Нове' : (c.st === 'review' && c.ivl >= 21 ? 'Вивчено' : 'Вчу');
  const rows = cards.map(c => [c.en, c.uk, c.pos || '', c.ipa || '', c.ex || '', c.exUk || '', setName[c.set] || '', status(c), c.ivl || 0, c.st === 'new' ? '' : fmt_(c.due)]);
  const w = sheet_('Слова');
  w.clearContents();
  w.getRange(1, 1, 1, 10).setValues([['English', 'Українською', 'Частина мови', 'Транскрипція', 'Приклад', 'Переклад прикладу', 'Набір', 'Статус', 'Інтервал, дн.', 'Наступне повторення']]).setFontWeight('bold');
  if (rows.length) w.getRange(2, 1, rows.length, 10).setValues(rows);
  w.setFrozenRows(1);

  const daily = (docs.meta && docs.meta.daily) || {};
  const prog = Object.keys(daily).sort().reverse().map(k => { const d = daily[k]; return [k, d.new || 0, d.rev || 0, d.tot || 0, d.ok || 0, Math.round((d.ms || 0) / 60000)]; });
  const p = sheet_('Прогрес');
  p.clearContents();
  p.getRange(1, 1, 1, 6).setValues([['День', 'Нових', 'Повторень', 'Відповідей', 'Правильних', 'Хвилин']]).setFontWeight('bold');
  if (prog.length) p.getRange(2, 1, prog.length, 6).setValues(prog);
  p.setFrozenRows(1);
}
