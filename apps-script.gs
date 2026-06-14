/**
 * RENOVATION BOARD — Google Sheets Sync Script  v1.3.2
 * =============================================
 * HOW TO INSTALL:
 *  1. Extensions → Apps Script → delete existing code → paste this → Save
 *  2. Deploy → New deployment → Type: Web app, Execute as: Me, Access: Anyone → Deploy
 *  3. Copy URL → Dashboard ⚙ Settings → Google Sheets URL → Connect Sheets
 *
 * COLUMN LAYOUT (auto-created on first run if missing):
 *   A  (formatting — ignored)   B  Name   C  Qty   D  Unit (€)
 *   E  Material (formula)       F  Labor  G  Total (formula)
 *   H  Payment Status  ← dropdown: Not Paid / Deposit Paid / Fully Paid
 *   I  Delivery Status ← dropdown: Not Ordered / Ordered / Received / Returned
 *   J  Actual spent    K  (blank)   L  Link   M  Notes
 *
 * PHASE DETECTION: only rows whose name (col B) starts with "Phase" or "Fase".
 */

// ── CONFIGURATION ─────────────────────────────────────────────
const SPREADSHEET_ID = '1670GVxMtkEsH67B0kQgGMAnCMOdEABqMOPtqSMS7nW8';
const SHEET_NAME     = 'Home Renovation Budget Template';

const COL = {
  NAME:     1,   // B
  QTY:      2,   // C
  UNIT:     3,   // D
  MAT:      4,   // E — formula, never overwritten
  LABOR:    5,   // F
  TOTAL:    6,   // G — formula, never overwritten
  PAYMENT:  7,   // H — payment status
  DELIVERY: 8,   // I — delivery status
  ACTUAL:   9,   // J — actual spent
  LINK:    11,   // L — link / URL
  NOTES:   12,   // M — notes
};

const PAY_DEFAULT = 'Not Paid';
const DEL_DEFAULT = 'Not Ordered';
// ──────────────────────────────────────────────────────────────

function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

// ── READ ──────────────────────────────────────────────────────
function doGet(e) {
  try {
    const sheet = getSheet();

    // Auto-setup: insert Payment/Delivery columns if not yet present
    const colHHeader = String(sheet.getRange(1, 8).getValue()).toLowerCase();
    if (!colHHeader.includes('payment')) {
      setupNewColumns(sheet);
    }

    const rows   = sheet.getDataRange().getValues();
    const budget = parseSheetToBudget(rows);
    return respond({ ok: true, budget, rows: rows.length, sheet: sheet.getName() });
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

function parseSheetToBudget(rows) {
  const phases       = [];
  let   currentPhase = null;

  rows.forEach((row, i) => {
    const name = String(row[COL.NAME] || '').trim();
    if (!name) return;
    if (name.toLowerCase().startsWith('subtotal')) return;

    if (isPhaseHeader(name)) {
      const lc = name.toLowerCase();
      let type = undefined, color = '#0052cc';

      if (lc.includes('subsidi') || lc.includes('subsidy')) {
        type = 'subsidy'; color = '#00875a';
      } else if (lc.includes('labour') || lc.includes('labor') || lc.includes('arbeid')) {
        type = 'labour'; color = '#0052cc';
      } else {
        if      (lc.includes('phase 11') || lc.includes('fase 11')) color = '#00875a';
        else if (lc.includes('phase 10') || lc.includes('fase 10')) color = '#6554c0';
        else if (lc.includes('phase 0')  || lc.includes('fase 0'))  color = '#0052cc';
        else if (lc.includes('phase 1')  || lc.includes('fase 1'))  color = '#de350b';
        else if (lc.includes('phase 2')  || lc.includes('fase 2'))  color = '#ff8b00';
        else if (lc.includes('phase 3')  || lc.includes('fase 3'))  color = '#00b8d9';
        else if (lc.includes('phase 4')  || lc.includes('fase 4'))  color = '#de350b';
        else if (lc.includes('phase 5')  || lc.includes('fase 5'))  color = '#6554c0';
        else if (lc.includes('phase 6')  || lc.includes('fase 6'))  color = '#ff8b00';
        else if (lc.includes('phase 7')  || lc.includes('fase 7'))  color = '#00875a';
        else if (lc.includes('phase 8')  || lc.includes('fase 8'))  color = '#00b8d9';
        else if (lc.includes('phase 9')  || lc.includes('fase 9'))  color = '#00875a';
      }

      const existing = phases.find(p => p.label === name);
      if (existing) { currentPhase = existing; }
      else { currentPhase = { id: 'ph_gs_' + i, label: name, color, type, items: [] }; phases.push(currentPhase); }

    } else if (currentPhase) {
      const actual   = row[COL.ACTUAL];
      const payment  = String(row[COL.PAYMENT]  || '').trim() || PAY_DEFAULT;
      const delivery = String(row[COL.DELIVERY] || '').trim() || DEL_DEFAULT;
      currentPhase.items.push({
        id: 'i_gs_' + i, name,
        qty:    numOrZero(row[COL.QTY]),
        unit:   numOrZero(row[COL.UNIT]),
        labor:  numOrZero(row[COL.LABOR]),
        actual: (actual !== '' && actual !== null && actual !== undefined) ? Number(actual) : null,
        payment, delivery,
        link:   String(row[COL.LINK]  || ''),
        notes:  String(row[COL.NOTES] || ''),
      });
    }
  });

  return phases.filter(p => p.items.length > 0);
}

// ── WRITE ─────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (!body.budget) throw new Error('No budget data in request body');
    const sheet = getSheet();
    writeBudgetToSheet(sheet, body.budget);
    return respond({ ok: true, updated: new Date().toISOString() });
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

function writeBudgetToSheet(sheet, budget) {
  const rows = sheet.getDataRange().getValues();
  const nameToRow = {}, phaseLastRow = {};
  let currentPhaseLabel = null;

  rows.forEach((row, i) => {
    const name = String(row[COL.NAME] || '').trim();
    if (!name) return;
    if (name.toLowerCase().startsWith('subtotal')) return;
    if (isPhaseHeader(name)) { currentPhaseLabel = name; }
    else if (currentPhaseLabel) { nameToRow[name] = i; phaseLastRow[currentPhaseLabel] = i; }
  });

  budget.forEach(phase => {
    (phase.items || []).forEach(item => {
      let rowIdx = nameToRow[item.name];

      if (rowIdx === undefined) {
        const afterIdx = phaseLastRow[phase.label];
        if (afterIdx === undefined) return;
        const afterRow1 = afterIdx + 1;
        sheet.insertRowAfter(afterRow1);
        const newRow1 = afterRow1 + 1;
        sheet.getRange(newRow1, COL.NAME     + 1).setValue(item.name);
        sheet.getRange(newRow1, COL.QTY      + 1).setValue(item.qty      || 0);
        sheet.getRange(newRow1, COL.UNIT     + 1).setValue(item.unit     || 0);
        sheet.getRange(newRow1, COL.LABOR    + 1).setValue(item.labor    || 0);
        sheet.getRange(newRow1, COL.PAYMENT  + 1).setValue(item.payment  || PAY_DEFAULT);
        sheet.getRange(newRow1, COL.DELIVERY + 1).setValue(item.delivery || DEL_DEFAULT);
        if (item.link  !== undefined) sheet.getRange(newRow1, COL.LINK  + 1).setValue(item.link  || '');
        if (item.notes !== undefined) sheet.getRange(newRow1, COL.NOTES + 1).setValue(item.notes || '');
        phaseLastRow[phase.label] = afterIdx + 1;
        nameToRow[item.name]      = afterIdx + 1;
        return;
      }

      const r = rowIdx + 1;
      sheet.getRange(r, COL.QTY      + 1).setValue(item.qty      || 0);
      sheet.getRange(r, COL.UNIT     + 1).setValue(item.unit     || 0);
      sheet.getRange(r, COL.LABOR    + 1).setValue(item.labor    || 0);
      sheet.getRange(r, COL.PAYMENT  + 1).setValue(item.payment  || PAY_DEFAULT);
      sheet.getRange(r, COL.DELIVERY + 1).setValue(item.delivery || DEL_DEFAULT);
      if (item.actual !== null && item.actual !== undefined) sheet.getRange(r, COL.ACTUAL + 1).setValue(item.actual);
      if (item.link   !== undefined) sheet.getRange(r, COL.LINK  + 1).setValue(item.link  || '');
      if (item.notes  !== undefined) sheet.getRange(r, COL.NOTES + 1).setValue(item.notes || '');
    });
  });

  SpreadsheetApp.flush();
}

// ── AUTO-SETUP ────────────────────────────────────────────────
/**
 * Called automatically by doGet on first run if columns H/I are missing.
 * Inserts Payment Status (H) and Delivery Status (I), adds dropdowns,
 * fills defaults, and applies conditional colour formatting.
 */
function setupNewColumns(sheet) {
  if (!sheet) sheet = getSheet();
  const lastRow  = sheet.getLastRow();

  // Insert 2 columns after G (col 7, 1-indexed)
  sheet.insertColumnAfter(7);  // H = Payment Status
  sheet.insertColumnAfter(8);  // I = Delivery Status

  // Headers
  sheet.getRange(1, 8).setValue('Payment Status').setFontWeight('bold');
  sheet.getRange(1, 9).setValue('Delivery Status').setFontWeight('bold');

  const dataRows = Math.max(lastRow - 1, 1);
  const payRange = sheet.getRange(2, 8, dataRows, 1);
  const delRange = sheet.getRange(2, 9, dataRows, 1);

  // Dropdown validation
  payRange.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Not Paid', 'Deposit Paid', 'Fully Paid'], true)
      .setAllowInvalid(false).build()
  );
  delRange.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Not Ordered', 'Ordered', 'Received', 'Returned'], true)
      .setAllowInvalid(false).build()
  );

  // Fill defaults
  const payVals = payRange.getValues();
  for (let i = 0; i < payVals.length; i++) if (!payVals[i][0]) payVals[i][0] = PAY_DEFAULT;
  payRange.setValues(payVals);

  const delVals = delRange.getValues();
  for (let i = 0; i < delVals.length; i++) if (!delVals[i][0]) delVals[i][0] = DEL_DEFAULT;
  delRange.setValues(delVals);

  // Conditional formatting — Payment Status
  const cfPay = [
    { text: 'Not Paid',     bg: '#FFEBE6', fg: '#DE350B' },
    { text: 'Deposit Paid', bg: '#FFF0B3', fg: '#FF8B00' },
    { text: 'Fully Paid',   bg: '#E3FCEF', fg: '#006644' },
  ].map(r => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(r.text).setBackground(r.bg).setFontColor(r.fg).setRanges([payRange]).build());

  // Conditional formatting — Delivery Status
  const cfDel = [
    { text: 'Not Ordered', bg: '#EBECF0', fg: '#5E6C84' },
    { text: 'Ordered',     bg: '#DEEBFF', fg: '#0052CC' },
    { text: 'Received',    bg: '#E3FCEF', fg: '#006644' },
    { text: 'Returned',    bg: '#FFF0B3', fg: '#FF8B00' },
  ].map(r => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(r.text).setBackground(r.bg).setFontColor(r.fg).setRanges([delRange]).build());

  sheet.setConditionalFormatRules([...sheet.getConditionalFormatRules(), ...cfPay, ...cfDel]);
  SpreadsheetApp.flush();
}

// ── HELPERS ───────────────────────────────────────────────────
function isPhaseHeader(name) { return /^(phase|fase)\s+/i.test(name); }

function numOrZero(v) {
  const n = Number(v);
  return (v !== '' && v !== null && v !== undefined && !isNaN(n)) ? n : 0;
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
