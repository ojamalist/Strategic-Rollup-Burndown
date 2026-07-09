// ═══════════════════════════════════════════════════════════════════════════
//  Client Burndown Dashboard — Code.gs
//  Pure Google Sheets implementation. No web app required.
//
//  DATA MODEL:
//   • BurndownData — single editable sheet; all issues live here
//   • Tenant tabs  — live FILTER views per tenant (read-only, always in sync)
//   • Summary      — live dashboard (COUNTIF formulas on BurndownData)
//
//  BURNDOWN DATA COLUMNS (A–N):
//   A Row ID (auto)  B Tenant          C Date Added (auto)  D Added By
//   E Customer Owner F ST Owner        G Status             H Priority Level
//   I Topic          J Action Item     K Notes/Next Steps   L Due Date
//   M Date Resolved  N Support Case #
// ═══════════════════════════════════════════════════════════════════════════

const DATA_SHEET_NAME = 'BurndownData';

const DATA_HEADERS = [
  'Row ID', 'Tenant', 'Date Added', 'Added By', 'Customer Owner', 'ST Owner',
  'Status', 'Priority Level', 'Topic', 'Action Item Description',
  'Notes / Next Steps', 'Due Date', 'Date Resolved', 'Support Case #'
];

// Default statuses — used to seed the Statuses config sheet on first run
// AND whenever the sheet is found empty (e.g. on existing spreadsheets after upgrade).
const DEFAULT_STATUSES = [
  'Open', 'In Progress', 'Fixed', 'Closed', 'On Hold', 'Ticket Created'
];

// Priority levels — shared between the BurndownData dropdown and the Add Item dialog.
const PRIORITY_LEVELS = ['P1 - Critical', 'P2 - High', 'P3 - Medium', 'P4 - Low'];

// ─── Menu ────────────────────────────────────────────────────────────────────

function onOpen() {
  const ui   = SpreadsheetApp.getUi();
  const help = ui.createMenu('❓ Help')
    .addItem('About this dashboard',        'helpAbout')
    .addItem('How to add issues',           'helpAddIssues')
    .addItem('How to add tenants & topics', 'helpTenants')
    .addItem('About the Summary tab',       'helpSummary')
    .addItem('About tenant tabs',           'helpTenantTabs')
    .addItem('Troubleshooting',             'helpTroubleshoot');

  ui.createMenu('🔥 Burndown')
    .addItem('🚀 Set Up Dashboard (run first)',   'onboardClient')
    .addSeparator()
    .addItem('📝 Add Burndown Item',               'showAddIssueDialog')
    .addSeparator()
    .addItem('➕ Add Tenant',                     'addTenantFromMenu')
    .addItem('➕ Add Topic',                      'addTopicFromMenu')
    .addItem('➕ Add Status',                     'addStatusFromMenu')
    .addSeparator()
    .addItem('🔄 Rebuild Summary Formulas',       'rebuildSummaryFromMenu')
    .addItem('📊 Rebuild Dashboard Charts',       'rebuildDashboardFromMenu')
    .addItem('🎨 Refresh Dropdown Validation',    'refreshDataValidationFromMenu')
    .addSeparator()
    .addSubMenu(help)
    .addToUi();
}

// ─── Onboarding wizard ────────────────────────────────────────────────────────

function onboardClient() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const nameResult = ui.prompt(
    '🚀 Burndown Setup (1 of 2)',
    'What is the client or project name?',
    ui.ButtonSet.OK_CANCEL
  );
  if (nameResult.getSelectedButton() !== ui.Button.OK) return;
  const clientName = nameResult.getResponseText().trim() || 'Client Burndown';

  const tenantResult = ui.prompt(
    '🚀 Burndown Setup (2 of 2)',
    'Enter tenant names, separated by commas:\n' +
    'e.g.  AVR, C&C, Elite, Tiger Team\n\n' +
    '(You can add more later via Burndown → Add Tenant)',
    ui.ButtonSet.OK_CANCEL
  );
  if (tenantResult.getSelectedButton() !== ui.Button.OK) return;
  const tenants = tenantResult.getResponseText()
    .split(',').map(t => t.trim()).filter(Boolean);

  ss.rename(clientName + ' — Burndown');
  initializeSheets_(ss);
  tenants.forEach(t => addToConfigList('Tenants', t));
  buildDashboard_();

  ui.alert(
    '✅ Setup Complete!',
    'Sheets created:\n' +
    '• Summary  (live formula dashboard)\n' +
    '• BurndownData  (add & edit all issues here)\n' +
    '• Dashboard  (visual charts — refresh with the checkbox)\n' +
    '• Tenants / Topics / Statuses  (config lists)\n' +
    (tenants.length ? '• ' + tenants.join(', ') + '  (filtered views per tenant)\n' : '') +
    '\nOpen BurndownData to start adding issues, or use 🔥 Burndown → 📝 Add Burndown Item.\n' +
    'Dropdowns are ready on Status, Priority, Tenant, and Topic.',
    ui.ButtonSet.OK
  );

  ss.setActiveSheet(ss.getSheetByName(DATA_SHEET_NAME));
}

// ─── Sheet initialization ─────────────────────────────────────────────────────

function initializeSheets_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  // Step 1: create all data/config sheets FIRST so formula references can resolve
  getOrCreateDataSheet_();   // BurndownData must exist before Summary references it
  getListSheet_('Tenants', []);
  getListSheet_('Statuses', DEFAULT_STATUSES);
  getListSheet_('Topics', [
    'Estimating/SBET', 'Reporting', 'Accounting', 'Data', 'Bug/Issue',
    'App/Mobile', 'Desktop/Web', 'Dispatching/Scheduling', 'Settings/Config',
    'Call Center', 'Telecoms', 'Fintech', '3rd Party Integrations',
    'Mobile (Cordova)', 'Mobile (FMA)', 'Feature Request', 'Other'
  ]);
  SpreadsheetApp.flush();    // commit sheets so BurndownData/Topics exist before Summary formulas

  // Step 2: build Summary now that BurndownData + Topics already exist
  buildSummarySheet_(ss);
  SpreadsheetApp.flush();

  // Step 3: formatting and validation
  applyDataValidation_();
  applyConditionalFormatting_();
  applyTableFormat_();
  deleteDefaultSheet_(ss);   // remove the blank Sheet1 Google creates by default

  // Step 4: enforce tab order — Summary first, BurndownData second
  const summary = ss.getSheetByName('Summary');
  if (summary) ss.moveActiveSheet && (ss.setActiveSheet(summary), ss.moveActiveSheet(1));
  const bd = ss.getSheetByName(DATA_SHEET_NAME);
  if (bd) ss.moveActiveSheet && (ss.setActiveSheet(bd), ss.moveActiveSheet(2));
}

/**
 * Deletes the default "Sheet1" tab that Google Sheets creates automatically
 * when a new spreadsheet is opened. Safe — only runs if other sheets exist.
 */
function deleteDefaultSheet_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1) ss.deleteSheet(sheet1);
}

// ─── BurndownData sheet ───────────────────────────────────────────────────────

function getOrCreateDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(DATA_SHEET_NAME, 1); // always second tab, after Summary
  sheet.appendRow(DATA_HEADERS);
  sheet.setFrozenRows(1);
  sheet.getRange('1:1')
    .setBackground('#1e293b').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(10);

  const widths = [90, 110, 90, 100, 130, 130, 110, 120, 150, 280, 280, 90, 90, 120];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.setRowHeight(1, 30);
  sheet.setTabColor('#16a34a');
  return sheet;
}

/**
 * Applies table-like formatting to BurndownData:
 * alternating row banding + basic filter (dropdown arrows on every header cell).
 */
function applyTableFormat_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) return;

  // Remove existing bandings to avoid duplicates
  sheet.getBandings().forEach(b => b.remove());

  // Alternating row banding across a generous row range
  const bandRange = sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 2000), DATA_HEADERS.length);
  bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor('#1e293b')
    .setFirstRowColor('#ffffff')
    .setSecondRowColor('#f1f5f9');

  // Basic filter — adds the dropdown/sort arrows to every header cell
  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), DATA_HEADERS.length).createFilter();
}

// ─── Auto-fill Row ID + Date Added; Dashboard refresh checkbox ───────────────

function onEdit(e) {
  const sheet = e.range.getSheet();
  const name  = sheet.getName();

  // Dashboard refresh checkbox — row 3, col B (the checkbox cell)
  if (name === 'Dashboard') {
    if (e.range.getRow() === 3 && e.range.getColumn() === 2 && e.value === 'TRUE') {
      buildDashboard_();
    }
    return;
  }

  if (name !== DATA_SHEET_NAME) return;

  const startRow = e.range.getRow();
  const numRows  = e.range.getNumRows();
  const tz       = Session.getScriptTimeZone();
  let   email    = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (_) {}

  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    if (row <= 1) continue; // skip header row

    // Read cols B–N to check for any data in this row
    const rowVals = sheet.getRange(row, 2, 1, 13).getValues()[0];
    const hasData = rowVals.some(v => v !== '' && v !== null);
    if (!hasData) continue;

    // Col A — Row ID
    const rowIdCell = sheet.getRange(row, 1);
    if (!rowIdCell.getValue()) rowIdCell.setValue(getNextIssueId_(sheet));

    // Col C — Date Added
    const dateCell = sheet.getRange(row, 3);
    if (!dateCell.getValue())
      dateCell.setValue(Utilities.formatDate(new Date(), tz, 'MM/dd/yyyy'));

    // Col D — Added By
    const addedByCell = sheet.getRange(row, 4);
    if (!addedByCell.getValue() && email) addedByCell.setValue(email);
  }
}

/**
 * Returns the next sequential issue ID (ISS-001, ISS-002, …).
 * Reads existing IDs from col A, finds the max number, and increments.
 * Pads to 3 digits; goes beyond 3 if needed (ISS-1000, etc.).
 */
function getNextIssueId_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 'ISS-001';
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(r => String(r[0]).trim())
    .filter(id => /^ISS-\d+$/.test(id))
    .map(id => parseInt(id.replace('ISS-', ''), 10));
  const next = ids.length > 0 ? Math.max(...ids) + 1 : 1;
  return 'ISS-' + String(next).padStart(3, '0');
}

// ─── Data validation (dropdowns on BurndownData) ──────────────────────────────

function applyDataValidation_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) return;

  const ROWS = 2000;

  // Status (col G = 7) — references Statuses sheet (seeds defaults if sheet is empty)
  const statusSheet = getListSheet_('Statuses', DEFAULT_STATUSES);
  const statusSheetTab = ss.getSheetByName('Statuses');
  if (statusSheetTab && statusSheetTab.getLastRow() > 1) {
    sheet.getRange(2, 7, ROWS).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(statusSheetTab.getRange(2, 1, statusSheetTab.getLastRow() - 1, 1), true)
        .setAllowInvalid(false).setHelpText('Select a status').build()
    );
  }

  // Priority (col H = 8)
  sheet.getRange(2, 8, ROWS).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(PRIORITY_LEVELS, true)
      .setAllowInvalid(false).setHelpText('Select a priority').build()
  );

  // Tenant (col B = 2) — references Tenants sheet
  const tenantSheet = ss.getSheetByName('Tenants');
  if (tenantSheet && tenantSheet.getLastRow() > 1) {
    sheet.getRange(2, 2, ROWS).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(tenantSheet.getRange(2, 1, tenantSheet.getLastRow() - 1, 1), true)
        .setAllowInvalid(false).setHelpText('Select a tenant').build()
    );
  }

  // Topic (col I = 9) — references Topics sheet
  const topicSheet = ss.getSheetByName('Topics');
  if (topicSheet && topicSheet.getLastRow() > 1) {
    sheet.getRange(2, 9, ROWS).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(topicSheet.getRange(2, 1, topicSheet.getLastRow() - 1, 1), true)
        .setAllowInvalid(false).setHelpText('Select a topic').build()
    );
  }
}

// ─── Conditional formatting (Status, Priority, Tenant, Topic colors) ─────────

// Color palettes assigned round-robin to dynamic values (tenants, topics)
const CHIP_PALETTE = [
  { bg: '#dbeafe', fg: '#1e40af' }, // blue
  { bg: '#dcfce7', fg: '#14532d' }, // green
  { bg: '#fce7f3', fg: '#9d174d' }, // pink
  { bg: '#ede9fe', fg: '#5b21b6' }, // purple
  { bg: '#fed7aa', fg: '#9a3412' }, // orange
  { bg: '#fef9c3', fg: '#854d0e' }, // yellow
  { bg: '#ccfbf1', fg: '#134e4a' }, // teal
  { bg: '#fef2f2', fg: '#991b1b' }, // red
  { bg: '#f0fdf4', fg: '#166534' }, // light green
  { bg: '#f5f3ff', fg: '#4c1d95' }, // lavender
];

function applyConditionalFormatting_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) return;

  const ROWS = 2000;

  // ── Status (col G = 7) — curated colors for the 6 built-ins; CHIP_PALETTE for customs ──
  const STATUS_FIXED_COLORS = {
    'Open':           { bg: '#fef9c3', fg: '#713f12' },
    'In Progress':    { bg: '#dbeafe', fg: '#1e3a8a' },
    'Fixed':          { bg: '#dcfce7', fg: '#14532d' },
    'Closed':         { bg: '#f1f5f9', fg: '#475569' },
    'On Hold':        { bg: '#fce7f3', fg: '#831843' },
    'Ticket Created': { bg: '#ede9fe', fg: '#4c1d95' },
  };
  // getListSheet_ with DEFAULT_STATUSES seeds the sheet if it was created empty
  const allStatuses  = getListSheet_('Statuses', DEFAULT_STATUSES);
  const statusColors = allStatuses.map((value, i) =>
    STATUS_FIXED_COLORS[value]
      ? { value, ...STATUS_FIXED_COLORS[value] }
      : { value, ...CHIP_PALETTE[(i + 6) % CHIP_PALETTE.length] }
  );

  // ── Priority (col H = 8) ──
  const priorityColors = [
    { value: 'P1 - Critical', bg: '#fecaca', fg: '#7f1d1d' },
    { value: 'P2 - High',     bg: '#fed7aa', fg: '#7c2d12' },
    { value: 'P3 - Medium',   bg: '#fef9c3', fg: '#713f12' },
    { value: 'P4 - Low',      bg: '#dcfce7', fg: '#14532d' },
  ];

  // ── Tenants (col B = 2) ── dynamic, assigned by index
  const tenants = getListSheet_('Tenants', []);
  // ── Topics (col I = 9) ── dynamic, assigned by index
  const topics  = getListSheet_('Topics',  []);

  // Drop all existing rules on cols 2, 7, 8, 9 (Tenant, Status, Priority, Topic)
  const DYNAMIC_COLS = new Set([2, 7, 8, 9]);
  const kept = sheet.getConditionalFormatRules().filter(rule =>
    !rule.getRanges().some(r => DYNAMIC_COLS.has(r.getColumn()) && r.getRow() > 1)
  );

  const newRules = [];

  // Status rules
  const statusRange = sheet.getRange(2, 7, ROWS, 1);
  statusColors.forEach(({ value, bg, fg }) => {
    newRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(value).setBackground(bg).setFontColor(fg).setBold(true)
      .setRanges([statusRange]).build());
  });

  // Priority rules
  const priorityRange = sheet.getRange(2, 8, ROWS, 1);
  priorityColors.forEach(({ value, bg, fg }) => {
    newRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(value).setBackground(bg).setFontColor(fg).setBold(true)
      .setRanges([priorityRange]).build());
  });

  // Tenant rules (round-robin palette)
  const tenantRange = sheet.getRange(2, 2, ROWS, 1);
  tenants.forEach((tenant, i) => {
    const { bg, fg } = CHIP_PALETTE[i % CHIP_PALETTE.length];
    newRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(tenant).setBackground(bg).setFontColor(fg).setBold(true)
      .setRanges([tenantRange]).build());
  });

  // Topic rules (round-robin palette, offset by 3 so colors differ from tenants)
  const topicRange = sheet.getRange(2, 9, ROWS, 1);
  topics.forEach((topic, i) => {
    const { bg, fg } = CHIP_PALETTE[(i + 3) % CHIP_PALETTE.length];
    newRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(topic).setBackground(bg).setFontColor(fg)
      .setRanges([topicRange]).build());
  });

  sheet.setConditionalFormatRules([...kept, ...newRules]);
}

// ─── Tenant view tabs (FILTER-based, always in sync) ─────────────────────────

/**
 * Creates a per-tenant tab showing a live FILTER view of that tenant's rows
 * from BurndownData. The tenant name is stored in hidden cell Z1 so the
 * formula uses a cell reference rather than a hardcoded string — this means
 * names with special characters (C&C, R&R, etc.) never cause formula errors.
 */
function createTenantSheet_(tenantName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(tenantName)) return; // already exists

  // Insert just before the Tenants config tab so order is:
  // Summary → BurndownData → [tenant tabs] → Tenants → Topics
  const allSheets    = ss.getSheets();
  const tenantsSheet = ss.getSheetByName('Tenants');
  const insertIdx    = tenantsSheet
    ? allSheets.findIndex(s => s.getName() === 'Tenants')
    : allSheets.length;

  const sheet = ss.insertSheet(tenantName, insertIdx);
  sheet.setTabColor('#2563eb');

  // Row 1: Title bar (also readable label for the tab)
  sheet.getRange(1, 1, 1, DATA_HEADERS.length).merge()
    .setValue(tenantName + '  —  filtered view  |  To add or edit issues, go to BurndownData')
    .setBackground('#2563eb').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 32);

  // Row 2: Column headers (matching BurndownData)
  sheet.getRange(2, 1, 1, DATA_HEADERS.length).setValues([DATA_HEADERS])
    .setBackground('#1e293b').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(2);

  // Z1: hidden cell holding the tenant name as plain text.
  // The FILTER formula references Z1 instead of hardcoding the name.
  sheet.getRange('Z1').setValue(tenantName);
  sheet.hideColumns(26); // hide col Z

  // Row 3: FILTER formula — pulls all matching rows, auto-expands
  sheet.getRange('A3').setFormula(
    '=IFERROR(FILTER(' + DATA_SHEET_NAME + '!A2:N,' +
    DATA_SHEET_NAME + '!B2:B=Z1),"")'
  );

  // Column widths (match BurndownData)
  const widths = [90, 110, 90, 100, 130, 130, 110, 120, 150, 280, 280, 90, 90, 120];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Protect with warning so accidental typing doesn't break the formula
  sheet.protect()
    .setDescription('This is a read-only view. Add and edit issues in BurndownData.')
    .setWarningOnly(true);
}

function deleteTenantSheet_(tenantName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tenantName);
  if (sheet) ss.deleteSheet(sheet);
}

// ─── Visual Dashboard (charts) ────────────────────────────────────────────────

/**
 * Builds (or rebuilds) the Dashboard tab with three embedded charts:
 *   1. Status distribution — donut chart (% of all issues by status)
 *   2. Status breakdown by tenant — 100% stacked horizontal bar
 *   3. Issues by topic — horizontal bar sorted by volume
 *
 * Chart data is written to hidden columns (P onward) on the Dashboard sheet
 * and charts are created from those ranges. The refresh checkbox in B3
 * triggers this function via onEdit.
 */
function buildDashboard_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create or locate the Dashboard tab — insert after BurndownData
  let dash = ss.getSheetByName('Dashboard');
  if (!dash) {
    const allSheets = ss.getSheets();
    const bdIdx     = allSheets.findIndex(s => s.getName() === DATA_SHEET_NAME);
    dash = ss.insertSheet('Dashboard', bdIdx >= 0 ? bdIdx + 1 : 2);
  }

  // Clear everything
  dash.getCharts().forEach(c => dash.removeChart(c));
  dash.clearContents();
  dash.clearFormats();
  dash.setTabColor('#7c3aed');

  const DARK = '#1e293b', WHITE = '#ffffff';

  // ── Row 1: Title ──
  dash.setRowHeight(1, 44);
  const title = ss.getName().replace(' — Burndown', '') || 'Client';
  dash.getRange('A1:N1').merge()
    .setValue(title.toUpperCase() + ' — VISUAL DASHBOARD')
    .setBackground(DARK).setFontColor(WHITE)
    .setFontWeight('bold').setFontSize(15)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── Row 2: thin spacer ──
  dash.setRowHeight(2, 8);
  dash.getRange('A2:N2').setBackground('#e2e8f0');

  // ── Row 3: Refresh checkbox bar ──
  dash.setRowHeight(3, 34);
  dash.getRange('A3:N3').setBackground('#f8fafc');
  dash.setColumnWidth(1, 20);  // small left margin
  dash.setColumnWidth(2, 34);  // checkbox col
  dash.setColumnWidth(3, 200); // label col
  dash.getRange('B3').insertCheckboxes()
    .setBackground('#f8fafc').setHorizontalAlignment('center').setVerticalAlignment('middle');
  dash.getRange('C3')
    .setValue('Refresh Charts')
    .setFontWeight('bold').setFontSize(11).setFontColor('#1e293b')
    .setBackground('#f8fafc').setVerticalAlignment('middle');
  dash.getRange('D3')
    .setValue('← check this box to update charts with latest data')
    .setFontColor('#94a3b8').setFontSize(9).setBackground('#f8fafc').setVerticalAlignment('middle');
  dash.getRange('D3:N3').setBackground('#f8fafc');

  // ── Row 4: thin spacer ──
  dash.setRowHeight(4, 8);

  // ══ Compute data from BurndownData ══════════════════════════════════════════
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  const lastRow   = dataSheet ? dataSheet.getLastRow() : 1;

  const allStatuses  = getListSheet_('Statuses', DEFAULT_STATUSES);
  const STATUS_ORDER = allStatuses.length ? allStatuses : DEFAULT_STATUSES;
  const statusCounts       = {};
  const tenantStatusCounts = {}; // { tenant: { status: count } }
  const topicCounts        = {};
  let   total              = 0;

  if (dataSheet && lastRow > 1) {
    const data = dataSheet.getRange(2, 1, lastRow - 1, 14).getValues();
    data.forEach(row => {
      if (!row[9]) return; // skip rows with no Action Item (col J = index 9)
      const tenant = String(row[1]).trim();  // col B
      const status = String(row[6]).trim();  // col G
      const topic  = String(row[8]).trim();  // col I
      total++;
      statusCounts[status] = (statusCounts[status] || 0) + 1;
      if (tenant) {
        if (!tenantStatusCounts[tenant]) tenantStatusCounts[tenant] = {};
        tenantStatusCounts[tenant][status] = (tenantStatusCounts[tenant][status] || 0) + 1;
      }
      if (topic) topicCounts[topic] = (topicCounts[topic] || 0) + 1;
    });
  }

  // No data yet — show placeholder
  if (total === 0) {
    dash.getRange('B6:M7').merge()
      .setValue('No issue data yet. Add issues to BurndownData, then click the Refresh checkbox above.')
      .setHorizontalAlignment('center').setFontColor('#94a3b8').setFontSize(12)
      .setVerticalAlignment('middle');
    dash.setRowHeight(6, 60);
    return;
  }

  // ══ Write chart data to hidden columns (P = col 16 onward) ═════════════════
  const DC = 16; // data start column (P)

  // Clear stale data
  const clearRows = Math.max(dash.getLastRow(), 80);
  dash.getRange(1, DC, clearRows, 20).clearContent();

  // --- Status data: P1:Q{n} ---
  const statusRows = [['Status','% of Issues']];
  STATUS_ORDER.forEach(s => {
    const c = statusCounts[s] || 0;
    if (c > 0) statusRows.push([s, Math.round((c / total) * 1000) / 10]);
  });
  // catch any statuses not in the predefined list
  Object.keys(statusCounts).forEach(s => {
    if (!STATUS_ORDER.includes(s) && s && statusCounts[s] > 0)
      statusRows.push([s, Math.round((statusCounts[s] / total) * 1000) / 10]);
  });
  dash.getRange(1, DC, statusRows.length, 2).setValues(statusRows);

  // --- Tenant data: starts at row 14 ---
  const TENANT_ROW = 14;
  const tenants = Object.keys(tenantStatusCounts).sort();
  const tenantHeader = ['Tenant', ...STATUS_ORDER];
  const tenantRows = [tenantHeader];
  tenants.forEach(t => {
    const s = tenantStatusCounts[t];
    tenantRows.push([t, ...STATUS_ORDER.map(st => s[st] || 0)]);
  });
  dash.getRange(TENANT_ROW, DC, tenantRows.length, tenantHeader.length).setValues(tenantRows);

  // --- Topic data: starts at row 30 ---
  const TOPIC_ROW = 30;
  const sortedTopics = Object.entries(topicCounts)
    .filter(([t]) => t)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 15); // top 15
  const topicTotal = sortedTopics.reduce((s, [,c]) => s + c, 0);
  const topicRows = [['Topic','% of Issues']];
  sortedTopics.forEach(([topic, count]) => {
    topicRows.push([topic, Math.round((count / topicTotal) * 1000) / 10]);
  });
  dash.getRange(TOPIC_ROW, DC, topicRows.length, 2).setValues(topicRows);

  // Hide the data columns
  dash.hideColumns(DC, 20);

  // ══ Create charts ═══════════════════════════════════════════════════════════
  const STATUS_COLORS = ['#fbbf24','#60a5fa','#34d399','#94a3b8','#f472b6','#a78bfa','#fb923c'];

  // Chart 1 — Status Donut (top-left)
  dash.insertChart(
    dash.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(dash.getRange(1, DC, statusRows.length, 2))
      .setPosition(5, 2, 10, 0)
      .setOption('title', 'Issues by Status')
      .setOption('titleTextStyle', { fontSize: 13, bold: true, color: '#1e293b' })
      .setOption('pieHole', 0.45)
      .setOption('width', 420).setOption('height', 320)
      .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
      .setOption('colors', STATUS_COLORS)
      .setOption('backgroundColor', '#ffffff')
      .build()
  );

  // Chart 2 — 100% Stacked Bar by Tenant (top-right)
  if (tenants.length > 0) {
    dash.insertChart(
      dash.newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(dash.getRange(TENANT_ROW, DC, tenantRows.length, tenantHeader.length))
        .setPosition(5, 9, 10, 0)
        .setOption('title', 'Status Breakdown by Tenant')
        .setOption('titleTextStyle', { fontSize: 13, bold: true, color: '#1e293b' })
        .setOption('isStacked', 'percent')
        .setOption('width', 460).setOption('height', 320)
        .setOption('legend', { position: 'bottom', textStyle: { fontSize: 9 } })
        .setOption('colors', STATUS_COLORS)
        .setOption('backgroundColor', '#ffffff')
        .setOption('hAxis', { format: '#%', textStyle: { fontSize: 10 } })
        .setOption('vAxis', { textStyle: { fontSize: 10 } })
        .build()
    );
  }

  // Chart 3 — Issues by Topic horizontal bar (bottom, full width)
  if (sortedTopics.length > 0) {
    dash.insertChart(
      dash.newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(dash.getRange(TOPIC_ROW, DC, topicRows.length, 2))
        .setPosition(26, 2, 10, 0)
        .setOption('title', 'Issues by Topic')
        .setOption('titleTextStyle', { fontSize: 13, bold: true, color: '#1e293b' })
        .setOption('width', 880).setOption('height', 380)
        .setOption('legend', { position: 'none' })
        .setOption('colors', ['#6366f1'])
        .setOption('backgroundColor', '#ffffff')
        .setOption('hAxis', {
          title: '% of Issues', titleTextStyle: { color: '#64748b', fontSize: 10 },
          textStyle: { fontSize: 10 }
        })
        .setOption('vAxis', { textStyle: { fontSize: 10 } })
        .build()
    );
  }

  // Protect with warning
  dash.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  dash.protect()
    .setDescription('Dashboard is auto-generated. Use the Refresh checkbox or Burndown → Rebuild Dashboard Charts.')
    .setWarningOnly(true);
}

// ─── Add Burndown Item (dialog form) ──────────────────────────────────────────

/**
 * Opens a modal dialog form for adding a new BurndownData row without
 * typing directly into the sheet. Row ID, Date Added, and Added By are
 * still filled in automatically, exactly as they would be on a manual edit.
 */
function showAddIssueDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(DATA_SHEET_NAME)) {
    SpreadsheetApp.getUi().alert(
      'Run "🚀 Set Up Dashboard (run first)" before adding issues.'
    );
    return;
  }

  const formData = getAddIssueFormData_();
  const template = HtmlService.createTemplateFromFile('AddIssueDialog');
  template.tenants   = formData.tenants;
  template.statuses  = formData.statuses;
  template.topics    = formData.topics;
  template.priorities = formData.priorities;
  template.email     = formData.email;

  const html = template.evaluate().setWidth(480).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '📝 Add Burndown Item');
}

/**
 * Supplies the dropdown options (and current user email) to AddIssueDialog.html.
 */
function getAddIssueFormData_() {
  let email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (_) {}
  return {
    tenants:    getListSheet_('Tenants', []),
    statuses:   getListSheet_('Statuses', DEFAULT_STATUSES),
    topics:     getListSheet_('Topics', []),
    priorities: PRIORITY_LEVELS,
    email:      email
  };
}

/**
 * Appends a new row to BurndownData from the Add Burndown Item dialog.
 * Called via google.script.run from AddIssueDialog.html.
 */
function submitNewIssue(form) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
    if (!sheet) return { success: false, error: 'BurndownData sheet not found. Run Set Up Dashboard first.' };

    const tenant     = String((form && form.tenant) || '').trim();
    const topic      = String((form && form.topic) || '').trim();
    const actionItem = String((form && form.actionItem) || '').trim();
    if (!tenant) return { success: false, error: 'Tenant is required.' };
    if (!topic) return { success: false, error: 'Topic is required.' };
    if (!actionItem) return { success: false, error: 'Action Item Description is required.' };

    const tz = Session.getScriptTimeZone();
    let email = '';
    try { email = Session.getActiveUser().getEmail() || ''; } catch (_) {}

    const rowId     = getNextIssueId_(sheet);
    const dateAdded = Utilities.formatDate(new Date(), tz, 'MM/dd/yyyy');

    const row = [
      rowId,
      tenant,
      dateAdded,
      email,
      String((form && form.customerOwner) || '').trim(),
      String((form && form.stOwner) || '').trim(),
      String((form && form.status) || '').trim(),
      String((form && form.priority) || '').trim(),
      topic,
      actionItem,
      String((form && form.notes) || '').trim(),
      formatDateInput_(form && form.dueDate),
      '', // Date Resolved — left blank on creation
      String((form && form.supportCase) || '').trim()
    ];

    sheet.appendRow(row);
    return { success: true, rowId: rowId };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Converts an HTML date input value ("yyyy-mm-dd") to the same "MM/dd/yyyy"
 * text format used by the Date Added column. Returns '' for empty input.
 */
function formatDateInput_(value) {
  value = String(value || '').trim();
  if (!value) return '';
  const parts = value.split('-');
  if (parts.length !== 3) return value;
  const d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  if (isNaN(d.getTime())) return value;
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd/yyyy');
}

// ─── Menu actions ─────────────────────────────────────────────────────────────

function addTenantFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const r  = ui.prompt('➕ Add Tenant', 'Enter the tenant name:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const name = r.getResponseText().trim();
  if (!name) return;
  const res = addToConfigList('Tenants', name);
  if (res.success) {
    ui.alert('✅ Tenant "' + name + '" added.\n\nA filtered view tab has been created for this tenant.');
  } else {
    ui.alert('❌ ' + res.error);
  }
}

function addTopicFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const r  = ui.prompt('➕ Add Topic', 'Enter the topic name:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const name = r.getResponseText().trim();
  if (!name) return;
  const res = addToConfigList('Topics', name);
  ui.alert(res.success ? '✅ Topic "' + name + '" added.' : '❌ ' + res.error);
}

function addStatusFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const r  = ui.prompt('➕ Add Status', 'Enter the new status name:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const name = r.getResponseText().trim();
  if (!name) return;
  const res = addToConfigList('Statuses', name);
  ui.alert(res.success ? '✅ Status "' + name + '" added.' : '❌ ' + res.error);
}

function rebuildSummaryFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheetByName('Summary');
  if (existing) ss.deleteSheet(existing);
  buildSummarySheet_(ss);
  ui.alert('✅ Summary rebuilt from current Tenants and Topics lists.');
  ss.setActiveSheet(ss.getSheetByName('Summary'));
}

function rebuildDashboardFromMenu() {
  buildDashboard_();
  SpreadsheetApp.getUi().alert('✅ Dashboard charts rebuilt.');
  const dash = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  if (dash) SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(dash);
}

function refreshDataValidationFromMenu() {
  applyDataValidation_();
  applyConditionalFormatting_();
  SpreadsheetApp.getUi().alert('✅ Dropdown validation and colors refreshed.');
}

// ─── Config list management ───────────────────────────────────────────────────

function addToConfigList(sheetName, value) {
  try {
    value = String(value).trim();
    if (!value) return { success: false, error: 'Value cannot be empty' };

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sheet "' + sheetName + '" not found' };

    const lastRow  = Math.max(sheet.getLastRow(), 1);
    const existing = sheet.getRange(2, 1, lastRow, 1).getValues()
      .map(r => String(r[0]).trim().toLowerCase());
    if (existing.includes(value.toLowerCase()))
      return { success: false, error: '"' + value + '" already exists' };

    sheet.appendRow([value]);
    addRowToSummary_(value, sheetName === 'Tenants');

    if (sheetName === 'Tenants') {
      createTenantSheet_(value);
      applyDataValidation_();          // keep Tenant dropdown in sync
    } else {
      applyDataValidation_();          // keep Topic / Status dropdown in sync
    }
    applyConditionalFormatting_();     // assign color to the new entry

    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function removeFromConfigList(sheetName, value) {
  try {
    value = String(value).trim();
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sheet "' + sheetName + '" not found' };

    const lastRow = sheet.getLastRow();
    const data    = sheet.getRange(2, 1, Math.max(lastRow - 1, 1), 1).getValues();
    let found = false;
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0]).trim() === value) {
        sheet.deleteRow(i + 2);
        found = true;
        break;
      }
    }
    if (!found) return { success: false, error: '"' + value + '" not found' };

    removeRowFromSummary_(value, sheetName === 'Tenants');
    if (sheetName === 'Tenants') deleteTenantSheet_(value);
    applyDataValidation_();

    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// ─── Summary sheet ────────────────────────────────────────────────────────────

function buildSummarySheet_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  const tenants = getListSheet_('Tenants', []);
  const topics  = getListSheet_('Topics', []);

  let sheet = ss.getSheetByName('Summary');
  if (!sheet) {
    sheet = ss.insertSheet('Summary', 0);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }
  sheet.setTabColor('#1e3a5f');

  const DARK = '#1e293b', HEADER = '#dbeafe', WHITE = '#ffffff',
        ALT  = '#f8fafc', BORDER = '#cbd5e1';

  sheet.setColumnWidth(1, 220);
  [2,3,4,5,6].forEach(c => sheet.setColumnWidth(c, 105));
  sheet.setRowHeight(1, 44);

  const title = ss.getName().replace(' — Burndown', '') || 'Client';
  sheet.getRange('A1:F1').merge()
    .setValue(title.toUpperCase() + ' — BURNDOWN DASHBOARD')
    .setBackground(DARK).setFontColor(WHITE)
    .setFontWeight('bold').setFontSize(15)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.getRange('A2').setValue('Last updated:').setFontWeight('bold').setFontColor('#64748b');
  sheet.getRange('B2').setFormula('=NOW()').setNumberFormat('mmm d, yyyy h:mm AM/PM').setFontColor('#64748b');

  // Overall Summary
  const dn = DATA_SHEET_NAME;
  sheet.getRange('A4:C4').merge().setValue('OVERALL SUMMARY')
    .setBackground(DARK).setFontColor(WHITE).setFontWeight('bold');
  sheet.getRange('A5:B5').setValues([['Metric','Count']])
    .setBackground(HEADER).setFontWeight('bold');
  sheet.getRange('B5').setHorizontalAlignment('center');

  const overallRows = [
    ['Total Issues',           `=COUNTA(${dn}!J:J)-1`],
    ['Open',                   `=COUNTIF(${dn}!G:G,"Open")`],
    ['In Progress',            `=COUNTIF(${dn}!G:G,"In Progress")`],
    ['Fixed / Closed',         `=COUNTIFS(${dn}!G:G,"Fixed")+COUNTIFS(${dn}!G:G,"Closed")`],
    ['On Hold',                `=COUNTIF(${dn}!G:G,"On Hold")`],
    ['Ticket Created',         `=COUNTIF(${dn}!G:G,"Ticket Created")`],
    ['Unresolved (no status)', `=COUNTA(${dn}!J2:J10000)-COUNTA(${dn}!G2:G10000)`],
  ];
  overallRows.forEach(([label, formula], i) => {
    const r = i + 6;
    sheet.getRange(r, 1).setValue(label);
    sheet.getRange(r, 2).setFormula(formula).setHorizontalAlignment('center').setNumberFormat('0');
    sheet.getRange(r, 1, 1, 2).setBackground(i % 2 === 0 ? WHITE : ALT);
  });
  sheet.getRange('A5:B12').setBorder(true,true,true,true,true,true, BORDER, SpreadsheetApp.BorderStyle.SOLID);

  // BY TENANT
  const tenantHeaderRow = 14;
  sheet.getRange(tenantHeaderRow, 1, 1, 6).merge()
    .setValue('BY TENANT').setBackground(DARK).setFontColor(WHITE).setFontWeight('bold');
  sheet.getRange(tenantHeaderRow + 1, 1, 1, 6)
    .setValues([['Tenant','Total','Open','In Progress','Fixed/Closed','On Hold']])
    .setBackground(HEADER).setFontWeight('bold');
  sheet.getRange(tenantHeaderRow + 1, 2, 1, 5).setHorizontalAlignment('center');

  const tenantDataStart = tenantHeaderRow + 2;
  tenants.forEach((t, i) => setTenantSummaryRow_(sheet, tenantDataStart + i, t, i));
  if (tenants.length) {
    sheet.getRange(tenantHeaderRow + 1, 1, tenants.length + 1, 6)
      .setBorder(true,true,true,true,true,true, BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }

  // BY TOPIC
  const topicHeaderRow = tenantDataStart + tenants.length + 2;
  sheet.getRange(topicHeaderRow, 1, 1, 6).merge()
    .setValue('BY TOPIC').setBackground(DARK).setFontColor(WHITE).setFontWeight('bold');
  sheet.getRange(topicHeaderRow + 1, 1, 1, 6)
    .setValues([['Topic','Total','Open','In Progress','Fixed/Closed','On Hold']])
    .setBackground(HEADER).setFontWeight('bold');
  sheet.getRange(topicHeaderRow + 1, 2, 1, 5).setHorizontalAlignment('center');

  const topicDataStart = topicHeaderRow + 2;
  topics.forEach((t, i) => setTopicSummaryRow_(sheet, topicDataStart + i, t, i));
  if (topics.length) {
    sheet.getRange(topicHeaderRow + 1, 1, topics.length + 1, 6)
      .setBorder(true,true,true,true,true,true, BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }

  sheet.setFrozenRows(1);

  // Protect with warning
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  sheet.protect()
    .setDescription('Summary is auto-calculated. Use Burndown → Rebuild Summary Formulas to refresh.')
    .setWarningOnly(true);

  return sheet;
}

function setTenantSummaryRow_(sheet, r, tenant, idx) {
  const dn = DATA_SHEET_NAME;
  sheet.getRange(r, 1).setValue(tenant);
  sheet.getRange(r, 2).setFormula(`=COUNTIF(${dn}!$B:$B,$A${r})`);
  sheet.getRange(r, 3).setFormula(`=COUNTIFS(${dn}!$B:$B,$A${r},${dn}!$G:$G,"Open")`);
  sheet.getRange(r, 4).setFormula(`=COUNTIFS(${dn}!$B:$B,$A${r},${dn}!$G:$G,"In Progress")`);
  sheet.getRange(r, 5).setFormula(
    `=COUNTIFS(${dn}!$B:$B,$A${r},${dn}!$G:$G,"Fixed")+COUNTIFS(${dn}!$B:$B,$A${r},${dn}!$G:$G,"Closed")`
  );
  sheet.getRange(r, 6).setFormula(`=COUNTIFS(${dn}!$B:$B,$A${r},${dn}!$G:$G,"On Hold")`);
  sheet.getRange(r, 1, 1, 6).setBackground(idx % 2 === 0 ? '#ffffff' : '#f8fafc');
  sheet.getRange(r, 2, 1, 5).setHorizontalAlignment('center').setNumberFormat('0');
}

function setTopicSummaryRow_(sheet, r, topic, idx) {
  const dn = DATA_SHEET_NAME;
  sheet.getRange(r, 1).setValue(topic);
  sheet.getRange(r, 2).setFormula(`=COUNTIF(${dn}!$I:$I,$A${r})`);
  sheet.getRange(r, 3).setFormula(`=COUNTIFS(${dn}!$I:$I,$A${r},${dn}!$G:$G,"Open")`);
  sheet.getRange(r, 4).setFormula(`=COUNTIFS(${dn}!$I:$I,$A${r},${dn}!$G:$G,"In Progress")`);
  sheet.getRange(r, 5).setFormula(
    `=COUNTIFS(${dn}!$I:$I,$A${r},${dn}!$G:$G,"Fixed")+COUNTIFS(${dn}!$I:$I,$A${r},${dn}!$G:$G,"Closed")`
  );
  sheet.getRange(r, 6).setFormula(`=COUNTIFS(${dn}!$I:$I,$A${r},${dn}!$G:$G,"On Hold")`);
  sheet.getRange(r, 1, 1, 6).setBackground(idx % 2 === 0 ? '#ffffff' : '#f8fafc');
  sheet.getRange(r, 2, 1, 5).setHorizontalAlignment('center').setNumberFormat('0');
}

// ─── Summary row helpers ──────────────────────────────────────────────────────

function addRowToSummary_(value, isTenant) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName('Summary');
  if (!summary) return;

  const allVals = summary.getRange(1, 1, summary.getLastRow(), 1).getValues();

  if (isTenant) {
    let topicHeaderRow = -1;
    for (let i = 0; i < allVals.length; i++) {
      if (String(allVals[i][0]).trim() === 'BY TOPIC') { topicHeaderRow = i + 1; break; }
    }
    if (topicHeaderRow === -1) return;
    const insertAt = topicHeaderRow - 1;
    summary.insertRowBefore(insertAt);
    const idx = getListSheet_('Tenants', []).length - 1;
    setTenantSummaryRow_(summary, insertAt, value, idx % 2);
  } else {
    let topicHeaderRow = -1;
    for (let i = 0; i < allVals.length; i++) {
      if (String(allVals[i][0]).trim() === 'BY TOPIC') { topicHeaderRow = i + 1; break; }
    }
    if (topicHeaderRow === -1) return;
    const refreshed = summary.getRange(1, 1, summary.getLastRow(), 1).getValues();
    let lastTopicRow = topicHeaderRow + 1;
    for (let i = topicHeaderRow + 1; i < refreshed.length; i++) {
      if (String(refreshed[i][0]).trim() !== '') lastTopicRow = i + 1;
      else break;
    }
    const idx = getListSheet_('Topics', []).length - 1;
    setTopicSummaryRow_(summary, lastTopicRow + 1, value, idx % 2);
  }
}

function removeRowFromSummary_(value, isTenant) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName('Summary');
  if (!summary) return;

  const sectionMarker = isTenant ? 'BY TENANT' : 'BY TOPIC';
  const allVals       = summary.getRange(1, 1, summary.getLastRow(), 1).getValues();

  let sectionStart = -1;
  for (let i = 0; i < allVals.length; i++) {
    if (String(allVals[i][0]).trim() === sectionMarker) { sectionStart = i + 1; break; }
  }
  if (sectionStart === -1) return;

  for (let i = sectionStart + 2; i < allVals.length; i++) {
    if (String(allVals[i][0]).trim() === value) {
      summary.deleteRow(i + 1);
      return;
    }
  }
}

// ─── Config sheet helper ──────────────────────────────────────────────────────

function getListSheet_(name, defaults) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  let wroteDefaults = false;

  if (!sheet) {
    // Brand-new sheet: create it, write header, always seed defaults
    sheet = ss.insertSheet(name);
    sheet.getRange('A1').setValue(name)
      .setBackground('#1e293b').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setColumnWidth(1, 200);
    if (defaults.length > 0) {
      // Batch write — faster and avoids lazy-flush issues with getLastRow()
      sheet.getRange(2, 1, defaults.length, 1).setValues(defaults.map(v => [v]));
      wroteDefaults = true;
    }
  } else if (sheet.getLastRow() <= 1 && defaults.length > 0) {
    // Existing but empty sheet (e.g. user deleted all rows, or upgrade scenario)
    sheet.getRange(2, 1, defaults.length, 1).setValues(defaults.map(v => [v]));
    wroteDefaults = true;
  }

  // Flush so getLastRow() reflects the writes we just made
  if (wroteDefaults) SpreadsheetApp.flush();

  const lastRow = Math.max(sheet.getLastRow(), 1);
  return sheet.getRange(2, 1, lastRow, 1).getValues()
    .map(r => String(r[0]).trim()).filter(Boolean);
}

// ─── Help dialogs ─────────────────────────────────────────────────────────────

function helpAbout() {
  SpreadsheetApp.getUi().alert(
    'About This Dashboard',
    'This is a client burndown tracker built in Google Sheets.\n\n' +
    'SHEET STRUCTURE:\n' +
    '• BurndownData  — where all issues are added and edited\n' +
    '• [Tenant tabs]  — live filtered views per tenant (read-only, always in sync)\n' +
    '• Summary  — live dashboard with counts by status, tenant, and topic\n' +
    '• Tenants / Topics / Statuses  — config lists that drive the dropdowns\n\n' +
    'Use the 🔥 Burndown menu for all setup and configuration.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function helpAddIssues() {
  SpreadsheetApp.getUi().alert(
    'How to Add Issues',
    'Fastest way: 🔥 Burndown → 📝 Add Burndown Item\n' +
    'Opens a form so you can fill in an issue without touching the sheet directly —\n' +
    'Row ID, Date Added, and Added By are still filled in automatically.\n\n' +
    'You can also open BurndownData and type directly into the next empty row:\n\n' +
    '• Row ID — fills in automatically when you start typing\n' +
    '• Tenant — dropdown (values come from the Tenants config tab)\n' +
    '• Date Added — fills in automatically\n' +
    '• Status — dropdown (values come from the Statuses config tab)\n' +
    '• Priority — dropdown: P1 Critical → P4 Low\n' +
    '• Topic — dropdown (values come from the Topics config tab)\n' +
    '• Action Item Description — what the issue is\n' +
    '• Notes / Next Steps — running notes and follow-ups\n' +
    '• Due Date / Date Resolved / Support Case # — optional fields\n\n' +
    'TIP: The tenant tabs and Summary update instantly as you add rows — no refresh needed.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function helpTenants() {
  SpreadsheetApp.getUi().alert(
    'How to Add Tenants, Topics & Statuses',
    '➕ Add Tenant  (Burndown menu)\n' +
    'Prompts for a name, then:\n' +
    '  - Creates a new filtered view tab for that tenant\n' +
    '  - Adds it to the Tenants config tab\n' +
    '  - Adds a row to the Summary dashboard\n' +
    '  - Updates the Tenant dropdown in BurndownData\n\n' +
    '➕ Add Topic  (Burndown menu)\n' +
    'Same flow — adds to Topics tab, updates Summary, refreshes the Topic dropdown.\n\n' +
    '➕ Add Status  (Burndown menu)\n' +
    'Adds a new status to the Statuses config tab and refreshes the Status dropdown.\n\n' +
    'You can also type directly into the Tenants, Topics, or Statuses config tabs, then use\n' +
    '🎨 Refresh Dropdown Validation to sync the dropdowns.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function helpSummary() {
  SpreadsheetApp.getUi().alert(
    'About the Summary Tab',
    'The Summary tab is a live dashboard — all numbers update automatically as you edit BurndownData.\n\n' +
    'SECTIONS:\n' +
    '• Overall Summary — total issue counts by status across all tenants\n' +
    '• By Tenant — breakdown per tenant (Total, Open, In Progress, Fixed/Closed, On Hold)\n' +
    '• By Topic — same breakdown by issue category\n\n' +
    'The Summary is protected with a warning to prevent accidental edits.\n' +
    'To regenerate it from scratch:\n' +
    '🔥 Burndown → 🔄 Rebuild Summary Formulas',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function helpTenantTabs() {
  SpreadsheetApp.getUi().alert(
    'About Tenant Tabs',
    'Each tenant gets its own tab (e.g. "AVR", "C&C") showing a live filtered view\n' +
    'of only their issues from BurndownData.\n\n' +
    '• Updates instantly — no refresh needed\n' +
    '• READ-ONLY — to add or edit issues, go to BurndownData\n' +
    '• Safe to share with a client team — they only see their own issues\n\n' +
    'The tab is created automatically when you add a tenant via ➕ Add Tenant.\n\n' +
    'The filtered view uses a cell reference (not a hardcoded name) in its formula,\n' +
    'so tenant names with special characters like & always work correctly.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function helpTroubleshoot() {
  SpreadsheetApp.getUi().alert(
    'Troubleshooting',
    '🔥 Burndown menu not appearing?\n' +
    '→ Reload the page.\n\n' +
    'Tenant / Topic / Status dropdown not showing a new value?\n' +
    '→ Use 🎨 Refresh Dropdown Validation.\n\n' +
    'Summary numbers look wrong?\n' +
    '→ Use 🔄 Rebuild Summary Formulas.\n\n' +
    'Tenant tab not created after adding a tenant?\n' +
    '→ Remove and re-add the tenant via ➕ Add Tenant.\n\n' +
    'Row ID or Date Added not auto-filling?\n' +
    '→ Make sure you are typing in BurndownData, not a tenant tab.\n' +
    '  Tenant tabs are read-only filtered views.\n\n' +
    'Formula errors on a tenant name with & (e.g. C&C)?\n' +
    '→ Run 🔄 Rebuild Summary Formulas. The script uses cell references\n' +
    '  so special characters are always handled safely.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
