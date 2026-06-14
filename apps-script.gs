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
 *   A  (narrow formatting column — ignored, always blank in data rows)
 *   B  Item name / phase header
 *   C  Quantity
 *   D  Cost per item (€)
 *   E  Material total  ← formula =IF(D*C=0,"",D*C), NEVER overwritten
 *   F  Labor (€)
 *   G  Total estimate  ← formula =IF(SUM(E:F)=0,"",SUM(E:F)), NEVER overwritten
 *   H  Actual spent (€)
 *   I  (blank)
 *   J  Link / URL
 *   K  Notes
 */

// ── CONFIGURATION ─────────────────────────────────────────────
const SPREADSHEET_ID = '1670GVxMtkEsH67B0kQgGMAnCMOdEABqMOPtqSMS7nW8';
const SHEET_NAME     = 'Home Renovation Budget Template';

const COL = {
  NAME:   1,   // B  — item name / phase header (column A is a narrow formatting column)
  QTY:    2,   // C  — quantity
  UNIT:   3,   // D  — cost per item
  MAT:    4,   // E  — material total (formula — never overwritten)
  LABOR:  5,   // F  — labor cost
  TOTAL:  6,   // G  — total estimate (formula — never overwritten)
  ACTUAL: 7,   // H  — actual spent
  LINK:   9,   // J  — link / URL
  NOTES: 10,   // K  — notes
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
    const range  = sheet.getDataRange();
    const rows   = range.getValues();
    const bgs    = range.getBackgrounds();
    const budget = parseSheetToBudget(rows, bgs);
    return respond({ ok: true, budget, rows: rows.length, sheet: sheet.getName() });
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

/**
 * Row classification — uses BACKGROUND COLOR as the primary signal:
 *   • blank name             → skip
 *   • name starts "Subtotal" → skip (section subtotal row)
 *   • dark/colored background → phase / section header
 *   • light/white background  → item row
 *
 * This is more reliable than checking for empty data columns, because some
 * phase headers in the template have 0-values in data columns, and some
 * items (e.g. "Self Made" work) have all-zero costs.
 */
function parseSheetToBudget(rows, bgs) {
  const phases       = [];
  let   currentPhase = null;

  rows.forEach((row, i) => {
    const name = String(row[COL.NAME] || '').trim();
    if (!name) return;                                    // blank row

    // Skip subtotal / section-total rows
    if (name.toLowerCase().startsWith('subtotal')) return;

    const bg       = bgs && bgs[i] ? bgs[i][COL.NAME] : null;
    const isHeader = isPhaseHeader(name, bg);

    if (isHeader) {
      // ── Phase / section header ──────────────────────────────
      const lc = name.toLowerCase();
      let type  = undefined;
      let color = '#0052cc';

      if      (lc.includes('subsidi') || lc.includes('subsidy'))                            { type = 'subsidy'; color = '#00875a'; }
      else if (lc.includes('labour')  || lc.includes('labor') || lc.includes('arbeid'))     { type = 'labour';  color = '#0052cc'; }
      // Check multi-digit phases BEFORE single-digit (e.g. "phase 1" is a substring of "phase 10")
      else if (lc.includes('phase 11') || lc.includes('fase 11'))  color = '#00875a';
      else if (lc.includes('phase 10') || lc.includes('fase 10'))  color = '#6554c0';
      else if (lc.includes('phase 0')  || lc.includes('fase 0'))   color = '#0052cc';
      else if (lc.includes('phase 1')  || lc.includes('fase 1'))   color = '#de350b';
      else if (lc.includes('phase 2')  || lc.includes('fase 2'))   color = '#ff8b00';
      else if (lc.includes('phase 3')  || lc.includes('fase 3'))   color = '#00b8d9';
      else if (lc.includes('phase 4')  || lc.includes('fase 4'))   color = '#de350b';
      else if (lc.includes('phase 5')  || lc.includes('fase 5'))   color = '#6554c0';
      else if (lc.includes('phase 6')  || lc.includes('fase 6'))   color = '#ff8b00';
      else if (lc.includes('phase 7')  || lc.includes('fase 7'))   color = '#00875a';
      else if (lc.includes('phase 8')  || lc.includes('fase 8'))   color = '#00b8d9';
      else if (lc.includes('phase 9')  || lc.includes('fase 9'))   color = '#00875a';

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
        link:   String(row[COL.LINK]  || ''),
        notes:  String(row[COL.NOTES] || ''),
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
  const range = sheet.getDataRange();
  const rows  = range.getValues();
  const bgs   = range.getBackgrounds();

  // Build name→rowIndex map (items only) and phase→lastItemRow map
  const nameToRow    = {};
  const phaseLastRow = {};   // phase label → 0-based index of its last item row
  let   currentPhaseLabel = null;

  rows.forEach((row, i) => {
    const name = String(row[COL.NAME] || '').trim();
    if (!name) return;
    if (name.toLowerCase().startsWith('subtotal')) return;

    const bg = bgs[i] ? bgs[i][COL.NAME] : null;
    if (isPhaseHeader(name, bg)) {
      currentPhaseLabel = name;   // phase header
    } else {
      nameToRow[name] = i;
      if (currentPhaseLabel) phaseLastRow[currentPhaseLabel] = i;
    }
  });

  // Apply updates; INSERT new rows after the last item of their phase
  budget.forEach(phase => {
    (phase.items || []).forEach(item => {
      let rowIdx = nameToRow[item.name];

      if (rowIdx === undefined) {
        // ── NEW ITEM: insert a row after the phase's last known item ──
        const afterIdx = phaseLastRow[phase.label];
        if (afterIdx === undefined) return;   // phase not in sheet — skip

        const afterRow1 = afterIdx + 1;       // 1-indexed sheet row
        sheet.insertRowAfter(afterRow1);
        const newRow1 = afterRow1 + 1;        // 1-indexed position of new row

        sheet.getRange(newRow1, COL.NAME  + 1).setValue(item.name);
        sheet.getRange(newRow1, COL.QTY   + 1).setValue(item.qty   || 0);
        sheet.getRange(newRow1, COL.UNIT  + 1).setValue(item.unit  || 0);
        sheet.getRange(newRow1, COL.LABOR + 1).setValue(item.labor || 0);
        if (item.link  !== undefined) sheet.getRange(newRow1, COL.LINK  + 1).setValue(item.link  || '');
        if (item.notes !== undefined) sheet.getRange(newRow1, COL.NOTES + 1).setValue(item.notes || '');

        // Advance tracker so next new item in same phase goes after this one
        const newIdx0 = afterIdx + 1;
        phaseLastRow[phase.label] = newIdx0;
        nameToRow[item.name]      = newIdx0;
        return;
      }

      // ── EXISTING ITEM: update editable columns ────────────────────
      const r = rowIdx + 1;   // 1-indexed
      sheet.getRange(r, COL.QTY   + 1).setValue(item.qty   || 0);
      sheet.getRange(r, COL.UNIT  + 1).setValue(item.unit  || 0);
      sheet.getRange(r, COL.LABOR + 1).setValue(item.labor || 0);

      if (item.actual !== null && item.actual !== undefined) {
        sheet.getRange(r, COL.ACTUAL + 1).setValue(item.actual);
      }
      if (item.link  !== undefined) sheet.getRange(r, COL.LINK  + 1).setValue(item.link  || '');
      if (item.notes !== undefined) sheet.getRange(r, COL.NOTES + 1).setValue(item.notes || '');
    });
  });

  SpreadsheetApp.flush();
}

// ── HELPERS ───────────────────────────────────────────────────

/**
 * Returns true if a row is a phase/section header.
 * Uses two signals — either is sufficient:
 *   1. Cell background is dark/colored (primary signal for most phases)
 *   2. Name matches "Phase N" / "Fase N" pattern (fallback for phases whose
 *      background may be lighter than expected, e.g. Phase 9, Phase 11)
 */
function isPhaseHeader(name, bg) {
  if (isHeaderBackground(bg)) return true;
  // Fallback: name literally starts with "phase" or "fase" followed by a number
  return /^(phase|fase)\s+\d+/i.test(name);
}

/**
 * Returns true if the cell background is a dark/colored (non-white, non-light) color.
 * Phase header rows in this sheet always have a dark background.
 * Item rows are white or very light grey.
 */
function isHeaderBackground(bg) {
  if (!bg || bg === '#ffffff' || bg === '#000000') return false;
  const hex = bg.replace('#', '');
  if (hex.length !== 6) return false;
  const r = parseInt(hex.substr(0, 2), 16);
  const g = parseInt(hex.substr(2, 2), 16);
  const b = parseInt(hex.substr(4, 2), 16);
  // If all channels > 200 → very light (white/light grey) → NOT a header
  return !(r > 200 && g > 200 && b > 200);
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
