/**
 * RENOVATION BOARD — Google Sheets Sync Script
 * =============================================
 * HOW TO INSTALL (3 steps):
 *
 *  1. Go to https://script.google.com → New project
 *  2. Delete any existing code, paste this entire file → Save (Ctrl+S)
 *  3. Deploy → New deployment → Type: Web app
 *       Execute as:  Me
 *       Who has access:  Anyone
 *     → Deploy → Copy the Web App URL
 *  4. In the Renovation Board → ⚙ Settings → Google Sheets URL → paste URL → Connect Sheets
 *
 * COLUMN LAYOUT (confirmed):
 *   A  Item name / phase header
 *   B  (narrow hidden column — ignored)
 *   C  Quantity
 *   D  Cost per item (€)
 *   E  Material total  ← formula =IF(D*C=0,"",D*C), NEVER overwritten
 *   F  Labor (€)
 *   G  Total estimate  ← formula =IF(SUM(E:F)=0,"",SUM(E:F)), NEVER overwritten
 *   H  Actual spent (€)
 *   I  (blank)
 *   J  Link / URL
 */

// ── CONFIGURATION ─────────────────────────────────────────────
const SPREADSHEET_ID = '1670GVxMtkEsH67B0kQgGMAnCMOdEABqMOPtqSMS7nW8';
const SHEET_NAME     = 'Home Renovation Budget Template';

const COL = {
  NAME:   0,   // A  — item name / phase header
  QTY:    2,   // C  — quantity
  UNIT:   3,   // D  — cost per item
  MAT:    4,   // E  — material total (formula — never overwritten)
  LABOR:  5,   // F  — labor cost
  TOTAL:  6,   // G  — total estimate (formula — never overwritten)
  ACTUAL: 7,   // H  — actual spent
  LINK:   9,   // J  — link / URL
};
// ──────────────────────────────────────────────────────────────


function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

// ── READ ──────────────────────────────────────────────────────
function doGet(e) {
  try {
    const sheet  = getSheet();
    const rows   = sheet.getDataRange().getValues();
    const budget = parseSheetToBudget(rows);
    return respond({ ok: true, budget, rows: rows.length, sheet: sheet.getName() });
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

/**
 * Row classification:
 *   • blank name             → skip
 *   • name starts "Subtotal" → skip (section subtotal row)
 *   • no data in any of {qty, unit, labor, actual, link} → phase header
 *   • otherwise              → item row
 */
function parseSheetToBudget(rows) {
  const phases       = [];
  let   currentPhase = null;

  rows.forEach((row, i) => {
    const name = String(row[COL.NAME] || '').trim();
    if (!name) return;                                    // blank row

    // Skip subtotal / section-total rows
    if (name.toLowerCase().startsWith('subtotal')) return;

    const hasData = rowHasData(row);

    if (!hasData) {
      // ── Phase / section header ──────────────────────────────
      const lc = name.toLowerCase();
      let type  = undefined;
      let color = '#0052cc';

      if      (lc.includes('subsidi') || lc.includes('subsidy'))                { type = 'subsidy'; color = '#00875a'; }
      else if (lc.includes('labour')  || lc.includes('labor') || lc.includes('arbeid')) { type = 'labour'; color = '#0052cc'; }
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
      else if (lc.includes('phase 10') || lc.includes('fase 10')) color = '#6554c0';
      else if (lc.includes('phase 11') || lc.includes('fase 11')) color = '#00875a';

      currentPhase = { id: 'ph_gs_' + i, label: name, color, type, items: [] };
      phases.push(currentPhase);

    } else if (currentPhase) {
      // ── Item row ────────────────────────────────────────────
      const actual = row[COL.ACTUAL];
      currentPhase.items.push({
        id:     'i_gs_' + i,
        name,
        qty:    numOrZero(row[COL.QTY]),
        unit:   numOrZero(row[COL.UNIT]),
        labor:  numOrZero(row[COL.LABOR]),
        actual: (actual !== '' && actual !== null && actual !== undefined)
                ? Number(actual) : null,
        link:   String(row[COL.LINK] || ''),
      });
    }
  });

  // Drop phases that have no items (e.g. blank template sections)
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

  // Build name → row-index map for item rows only
  const nameToRow = {};
  rows.forEach((row, i) => {
    const name = String(row[COL.NAME] || '').trim();
    if (name && !name.toLowerCase().startsWith('subtotal') && rowHasData(row)) {
      nameToRow[name] = i;
    }
  });

  // Apply updates — only editable columns (C, D, F, H, J)
  // Columns E (Material) and G (Total) are formulas — never touched
  budget.forEach(phase => {
    (phase.items || []).forEach(item => {
      const rowIdx = nameToRow[item.name];
      if (rowIdx === undefined) return;
      const r = rowIdx + 1;   // 1-indexed

      sheet.getRange(r, COL.QTY   + 1).setValue(item.qty   || 0);
      sheet.getRange(r, COL.UNIT  + 1).setValue(item.unit  || 0);
      sheet.getRange(r, COL.LABOR + 1).setValue(item.labor || 0);

      if (item.actual !== null && item.actual !== undefined) {
        sheet.getRange(r, COL.ACTUAL + 1).setValue(item.actual);
      }
      if (item.link !== undefined) {
        sheet.getRange(r, COL.LINK + 1).setValue(item.link || '');
      }
    });
  });

  SpreadsheetApp.flush();
}

// ── HELPERS ───────────────────────────────────────────────────
function rowHasData(row) {
  return [COL.QTY, COL.UNIT, COL.LABOR, COL.ACTUAL, COL.LINK].some(c => {
    const v = row[c];
    return v !== '' && v !== null && v !== undefined && String(v).trim() !== '';
  });
}

function numOrZero(v) {
  const n = Number(v);
  return (v !== '' && v !== null && v !== undefined && !isNaN(n)) ? n : 0;
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
