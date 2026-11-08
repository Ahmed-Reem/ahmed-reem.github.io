/**
 * Ahmed & Reem — Wedding invitation tracker
 * Paste this into: Google Sheet > Extensions > Apps Script
 * Then run setup() once, and deploy as a Web App (Execute as: Me, Access: Anyone).
 */

const HEADERS = ['Guest Name', 'Phone (optional, e.g. 2010xxxxxxxx)', 'Seats', 'Guest ID', 'Invite Link',
  'Send on WhatsApp', 'Status', 'Opens', 'First Opened', 'Last Opened', 'Viewed Location',
  'Added to Calendar', 'RSVP', 'Attending', 'Message', 'Replied At'];
const C = { name: 1, phone: 2, seats: 3, id: 4, link: 5, wa: 6, status: 7, opens: 8, first: 9, last: 10,
  map: 11, cal: 12, rsvp: 13, count: 14, msg: 15, replied: 16 };

function onOpen() {
  SpreadsheetApp.getUi().createMenu('💍 Wedding')
    .addItem('Generate invite links', 'generateLinks')
    .addItem('Set up sheets (first time only)', 'setup')
    .addToUi();
}

/* ---------- one-time setup ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();

  const g = ss.getSheetByName('Guests') || ss.insertSheet('Guests');
  g.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#6e1f2c').setFontColor('#ffffff').setWrap(true);
  g.setFrozenRows(1);
  g.setColumnWidth(C.name, 180); g.setColumnWidth(C.link, 260); g.setColumnWidth(C.msg, 260);
  g.getRange('I2:J').setNumberFormat('dd MMM, hh:mm am/pm');
  g.getRange('K2:L').setNumberFormat('dd MMM, hh:mm am/pm');
  g.getRange('P2:P').setNumberFormat('dd MMM, hh:mm am/pm');
  const st = g.getRange('G2:G1000');
  g.setConditionalFormatRules([
    rule_(st, 'Attending', '#d9ead3', '#274e13'),
    rule_(st, 'Declined', '#f4cccc', '#990000'),
    rule_(st, 'Opened, no reply', '#fff2cc', '#7f6000'),
    rule_(st, 'Not opened', '#eeeeee', '#666666')
  ]);

  const s = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
  s.getRange('A1:B2').setValues([
    ['Invitation website URL', 'https://YOUR-SITE.netlify.app'],
    ['WhatsApp message (use {name} and {link})',
      'Dear {name} 🤍\n\nWith joyful hearts, Ahmed & Reem invite you to celebrate their wedding on Sunday, 29 November 2026 at Mountain Rose Hotel.\n\nYour personal invitation 💌\n{link}']
  ]);
  s.getRange('A1:A2').setFontWeight('bold'); s.setColumnWidth(1, 260); s.setColumnWidth(2, 520);
  s.getRange('B2').setWrap(true);

  const log = ss.getSheetByName('Log') || ss.insertSheet('Log');
  log.getRange(1, 1, 1, 6).setValues([['Time', 'Guest ID', 'Name', 'Event', 'Details', 'Device']])
    .setFontWeight('bold').setBackground('#4a1520').setFontColor('#ffffff');
  log.setFrozenRows(1);
  log.getRange('A2:A').setNumberFormat('dd MMM yyyy, hh:mm:ss am/pm');

  const sm = ss.getSheetByName('Summary') || ss.insertSheet('Summary', 0);
  sm.getRange('A1').setValue('Ahmed & Reem — RSVP Dashboard').setFontSize(18).setFontWeight('bold').setFontColor('#6e1f2c');
  sm.getRange('A3:B12').setValues([
    ['Invited (invitations)', '=COUNTA(Guests!A2:A)'],
    ['Opened the link', '=COUNTIF(Guests!H2:H,">0")'],
    ['Not opened yet', '=B3-B4'],
    ['Open rate', '=IFERROR(B4/B3,0)'],
    ['Accepted', '=COUNTIF(Guests!M2:M,"Attending")'],
    ['Declined', '=COUNTIF(Guests!M2:M,"Not attending")'],
    ['No reply yet', '=B3-B7-B8'],
    ['Expected people (total seats)', '=SUM(Guests!N2:N)'],
    ['Viewed location', '=COUNTA(Guests!K2:K)'],
    ['Added to calendar', '=COUNTA(Guests!L2:L)']
  ]);
  sm.getRange('B6').setNumberFormat('0%');
  sm.getRange('A3:A12').setFontWeight('bold');
  sm.getRange('B3:B12').setFontSize(14).setHorizontalAlignment('center');
  sm.getRange('B10').setFontColor('#6e1f2c').setFontWeight('bold');
  sm.setColumnWidth(1, 260); sm.setColumnWidth(2, 120);

  const def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);

  SpreadsheetApp.getActive().toast('Done! Put your website URL in Settings, add guests, then 💍 Wedding > Generate invite links.');
}

function rule_(range, text, bg, fg) {
  return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text)
    .setBackground(bg).setFontColor(fg).setRanges([range]).build();
}

/* ---------- create a personal link for every guest ---------- */
function generateLinks() {
  const ss = SpreadsheetApp.getActive();
  const g = ss.getSheetByName('Guests'), s = ss.getSheetByName('Settings');
  let base = String(s.getRange('B1').getValue()).trim().replace(/\/+$/, '');
  if (!base || base.indexOf('YOUR-SITE') > -1) {
    SpreadsheetApp.getUi().alert('First put your invitation website URL in Settings!B1');
    return;
  }
  const template = String(s.getRange('B2').getValue());
  const last = g.getLastRow();
  if (last < 2) return;

  const rng = g.getRange(2, 1, last - 1, HEADERS.length);
  const vals = rng.getValues();
  const used = {};
  vals.forEach(r => { if (r[C.id - 1]) used[r[C.id - 1]] = true; });

  vals.forEach((r, i) => {
    const row = i + 2;
    const name = String(r[C.name - 1]).trim();
    if (!name) return;
    let seats = parseInt(r[C.seats - 1], 10);
    if (!seats || seats < 1) { seats = 2; g.getRange(row, C.seats).setValue(seats); }
    let id = r[C.id - 1];
    if (!id) {
      do { id = slug_(name) + '-' + Math.random().toString(36).slice(2, 6); } while (used[id]);
      used[id] = true;
      g.getRange(row, C.id).setValue(id);
    }
    const link = base + '/?g=' + encodeURIComponent(id) + '&n=' + encodeURIComponent(name) + '&s=' + seats;
    g.getRange(row, C.link).setValue(link);

    const first = name.split(/\s+/)[0];
    const text = template.replace(/\{name\}/g, first).replace(/\{link\}/g, link);
    const phone = String(r[C.phone - 1]).replace(/\D/g, '');
    const wa = 'https://wa.me/' + (phone || '') + '?text=' + encodeURIComponent(text);
    g.getRange(row, C.wa).setRichTextValue(
      SpreadsheetApp.newRichTextValue().setText('Send 💬').setLinkUrl(wa).build());

    if (!r[C.status - 1]) g.getRange(row, C.status).setValue('Not opened');
  });
  ss.toast('Links ready ✨ Tap "Send 💬" next to each guest to send on WhatsApp.');
}

function slug_(s) {
  const out = s.toLowerCase().normalize('NFKD').replace(/[^\w؀-ۿ]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 24);
  return out || 'guest';
}

/* ---------- receives events from the invitation ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const p = JSON.parse(e.postData.contents || '{}');
    const id = String(p.g || '').slice(0, 80);
    const ev = String(p.event || '').slice(0, 20);
    if (!id || ['open', 'envelope', 'map', 'calendar', 'rsvp'].indexOf(ev) < 0) return out_('ignored');

    const ss = SpreadsheetApp.getActive();
    const now = new Date();
    ss.getSheetByName('Log').appendRow([now, id, safe_(p.n), ev, safe_(JSON.stringify(p.data || {})), safe_(p.ua)]);

    const g = ss.getSheetByName('Guests');
    let row = findRow_(g, id);
    // General link: a guest who replies adds themselves as a new row
    if (!row && p.pub === true && ev === 'rsvp' && /^pub-[a-z0-9]{4,12}$/.test(id)) {
      const nm = safe_(String((p.data || {}).name || '').trim().slice(0, 60));
      if (!nm) return out_('no name');
      g.appendRow([nm, '', 6, id, 'General link', '', '', 1, now, now]);
      row = g.getLastRow();
    }
    if (!row) return out_('unknown guest');
    if (p.pub === true && ev === 'rsvp' && (p.data || {}).name) {
      g.getRange(row, C.name).setValue(safe_(String(p.data.name).trim().slice(0, 60)));
    }

    const r = g.getRange(row, 1, 1, HEADERS.length).getValues()[0];
    if (ev === 'open') {
      g.getRange(row, C.opens).setValue((Number(r[C.opens - 1]) || 0) + 1);
      if (!r[C.first - 1]) g.getRange(row, C.first).setValue(now);
      g.getRange(row, C.last).setValue(now);
    } else if (ev === 'map') {
      if (!r[C.map - 1]) g.getRange(row, C.map).setValue(now);
    } else if (ev === 'calendar') {
      if (!r[C.cal - 1]) g.getRange(row, C.cal).setValue(now);
    } else if (ev === 'rsvp') {
      const d = p.data || {};
      const yes = d.attending === true;
      const seats = parseInt(r[C.seats - 1], 10) || 12;
      g.getRange(row, C.rsvp).setValue(yes ? 'Attending' : 'Not attending');
      g.getRange(row, C.count).setValue(yes ? Math.min(Math.max(parseInt(d.count, 10) || 1, 1), seats) : 0);
      g.getRange(row, C.msg).setValue(safe_(d.message));
      g.getRange(row, C.replied).setValue(now);
    }
    updateStatus_(g, row);
    return out_('ok');
  } catch (err) {
    return out_('error');
  } finally {
    lock.releaseLock();
  }
}

/* ---------- lets the invitation remember a guest's earlier reply ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'status' && p.g) {
    const g = SpreadsheetApp.getActive().getSheetByName('Guests');
    const row = findRow_(g, String(p.g).slice(0, 80));
    if (row) {
      const r = g.getRange(row, 1, 1, HEADERS.length).getValues()[0];
      if (r[C.rsvp - 1]) {
        return json_({ replied: true, attending: r[C.rsvp - 1] === 'Attending', count: r[C.count - 1], message: r[C.msg - 1] });
      }
    }
    return json_({ replied: false });
  }
  return out_('Wedding tracker is running 💍');
}

function updateStatus_(g, row) {
  const r = g.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  let st = 'Not opened';
  if (r[C.rsvp - 1] === 'Attending') st = 'Attending';
  else if (r[C.rsvp - 1] === 'Not attending') st = 'Declined';
  else if (Number(r[C.opens - 1]) > 0) st = 'Opened, no reply';
  g.getRange(row, C.status).setValue(st);
}

function findRow_(g, id) {
  const last = g.getLastRow();
  if (last < 2) return 0;
  const ids = g.getRange(2, C.id, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return 0;
}

function safe_(v) {
  let s = String(v == null ? '' : v).slice(0, 500);
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // stop formula injection
  return s;
}
function out_(t) { return ContentService.createTextOutput(t); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
