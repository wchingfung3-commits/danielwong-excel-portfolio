/**
 * Google Sheets Order & Reporting Automation
 * Self-initiated portfolio project by Daniel Wong.
 *
 * The script never sends email automatically. It can only create a Gmail draft
 * after the user supplies a recipient in the Config sheet and runs the command.
 */

const SHEETS = Object.freeze({
  RAW: 'Raw_Orders',
  CLEAN: 'Clean_Orders',
  DASHBOARD: 'Dashboard',
  CONFIG: 'Config',
  LOG: 'Automation_Log',
});

const CLEAN_HEADERS = Object.freeze([
  'Order ID', 'Order Date', 'Customer', 'Email', 'Region', 'Product',
  'Quantity', 'Unit Price', 'Revenue', 'Status', 'QA Status', 'QA Notes',
]);

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Operations Automation')
    .addItem('1. Set up demo workbook', 'setupDemoWorkbook')
    .addSeparator()
    .addItem('2. Process and validate orders', 'processOrders')
    .addItem('3. Refresh dashboard', 'refreshDashboard')
    .addItem('4. Create weekly Gmail draft', 'createWeeklySummaryDraft')
    .addToUi();
}

function setupDemoWorkbook() {
  const ss = SpreadsheetApp.getActive();
  const raw = getOrCreateSheet_(ss, SHEETS.RAW);
  const config = getOrCreateSheet_(ss, SHEETS.CONFIG);

  raw.clear();
  raw.getRange(1, 1, 1, 9).setValues([[
    'Order ID', 'Order Date', 'Customer', 'Email', 'Region',
    'Product', 'Quantity', 'Unit Price', 'Status',
  ]]);
  raw.getRange(2, 1, 8, 9).setValues([
    ['ORD-1001', new Date('2026-09-01'), '  Acme studio ', 'SALES@ACME.COM', 'usa', 'Analytics Setup', 2, 450, 'paid'],
    ['ORD-1002', new Date('2026-09-02'), 'Northstar Retail', 'ops@northstar.com', 'Australia', 'Monthly Report', 1, 780, 'completed'],
    ['ORD-1003', new Date('2026-09-03'), 'bright labs', 'hello@brightlabs.com', 'APAC', 'Data Cleanup', 3, 220, 'paid'],
    ['ORD-1004', new Date('2026-09-04'), 'Harbour Works', 'finance@harbourworks.com', 'UK', 'Dashboard', 1, 920, 'pending'],
    ['ORD-1005', new Date('2026-09-05'), 'Cedar Group', 'bad-email', 'Europe', 'Monthly Report', 1, 780, 'completed'],
    ['ORD-1006', new Date('2026-09-06'), 'Blue Peak Co', 'admin@bluepeak.co', 'Aus', 'Data Cleanup', 0, 220, 'paid'],
    ['ORD-1002', new Date('2026-09-02'), 'Northstar Retail', 'ops@northstar.com', 'Australia', 'Monthly Report', 1, 780, 'completed'],
    ['ORD-1007', new Date('2026-09-07'), 'Keystone Ltd', 'contact@keystone.example', 'North America', 'Automation', 1, 1250, 'refunded'],
  ]);
  raw.setFrozenRows(1);
  raw.autoResizeColumns(1, 9);

  config.clear();
  config.getRange('A1:B4').setValues([
    ['Setting', 'Value'],
    ['Weekly summary recipient', ''],
    ['Report title', 'Weekly Order Operations Summary'],
    ['Currency', 'USD'],
  ]);
  config.setFrozenRows(1);
  config.autoResizeColumns(1, 2);

  processOrders();
  SpreadsheetApp.getUi().alert('Demo workbook created. Review the Clean_Orders, Dashboard and Automation_Log sheets.');
}

function processOrders() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActive();
    const raw = ss.getSheetByName(SHEETS.RAW);
    if (!raw) throw new Error(`Missing required sheet: ${SHEETS.RAW}`);

    const values = raw.getDataRange().getValues();
    if (values.length < 2) throw new Error('Raw_Orders does not contain any order rows.');

    validateHeaders_(values[0]);
    const seenIds = new Set();
    const output = values.slice(1).filter(row => row.some(value => value !== '')).map(row => {
      const orderId = cleanText_(row[0]).toUpperCase();
      const orderDate = normalizeDate_(row[1]);
      const customer = toTitleCase_(cleanText_(row[2]));
      const email = cleanText_(row[3]).toLowerCase();
      const region = normalizeRegion_(row[4]);
      const product = cleanText_(row[5]);
      const quantity = Number(row[6]);
      const unitPrice = Number(row[7]);
      const status = normalizeStatus_(row[8]);
      const notes = [];

      if (!orderId) notes.push('Missing order ID');
      if (orderId && seenIds.has(orderId)) notes.push('Duplicate order ID');
      if (orderId) seenIds.add(orderId);
      if (!orderDate) notes.push('Invalid order date');
      if (!customer) notes.push('Missing customer');
      if (!isValidEmail_(email)) notes.push('Invalid email');
      if (!region) notes.push('Unrecognized region');
      if (!Number.isFinite(quantity) || quantity <= 0) notes.push('Quantity must be greater than zero');
      if (!Number.isFinite(unitPrice) || unitPrice <= 0) notes.push('Unit price must be greater than zero');
      if (status === 'Excluded') notes.push('Excluded transaction status');
      if (status === 'Pending') notes.push('Pending transaction');

      const ready = notes.length === 0;
      const revenue = ready ? quantity * unitPrice : 0;
      return [
        orderId, orderDate || '', customer, email, region, product,
        quantity || '', unitPrice || '', revenue, status,
        ready ? 'Ready' : 'Review', notes.join('; '),
      ];
    });

    const clean = getOrCreateSheet_(ss, SHEETS.CLEAN);
    clean.clear();
    clean.getRange(1, 1, 1, CLEAN_HEADERS.length).setValues([CLEAN_HEADERS]);
    if (output.length) clean.getRange(2, 1, output.length, CLEAN_HEADERS.length).setValues(output);
    clean.setFrozenRows(1);
    clean.getRange('B:B').setNumberFormat('yyyy-mm-dd');
    clean.getRange('H:I').setNumberFormat('$#,##0.00');
    clean.autoResizeColumns(1, CLEAN_HEADERS.length);
    applyQaFormatting_(clean, output.length);

    refreshDashboard_();
    logRun_('Process orders', 'Success', `${output.length} rows processed`);
  } catch (error) {
    logRun_('Process orders', 'Failed', error.message);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function refreshDashboard() {
  refreshDashboard_();
  logRun_('Refresh dashboard', 'Success', 'Dashboard refreshed');
  SpreadsheetApp.getUi().alert('Dashboard refreshed.');
}

function refreshDashboard_() {
  const ss = SpreadsheetApp.getActive();
  const clean = ss.getSheetByName(SHEETS.CLEAN);
  if (!clean || clean.getLastRow() < 2) throw new Error('Run Process and validate orders first.');

  const rows = clean.getRange(2, 1, clean.getLastRow() - 1, CLEAN_HEADERS.length).getValues();
  const readyRows = rows.filter(row => row[10] === 'Ready');
  const reviewRows = rows.filter(row => row[10] === 'Review');
  const revenue = readyRows.reduce((sum, row) => sum + Number(row[8] || 0), 0);
  const regionTotals = readyRows.reduce((acc, row) => {
    acc[row[4]] = (acc[row[4]] || 0) + Number(row[8] || 0);
    return acc;
  }, {});

  const dashboard = getOrCreateSheet_(ss, SHEETS.DASHBOARD);
  dashboard.clear();
  dashboard.getCharts().forEach(chart => dashboard.removeChart(chart));
  dashboard.getRange('A1:F1').merge().setValue('Order Operations Dashboard');
  dashboard.getRange('A3:B6').setValues([
    ['Metric', 'Value'],
    ['Ready orders', readyRows.length],
    ['Rows requiring review', reviewRows.length],
    ['Validated revenue', revenue],
  ]);
  dashboard.getRange('B6').setNumberFormat('$#,##0.00');

  const regionRows = Object.entries(regionTotals).sort((a, b) => b[1] - a[1]);
  dashboard.getRange('D3:E3').setValues([['Region', 'Revenue']]);
  if (regionRows.length) {
    dashboard.getRange(4, 4, regionRows.length, 2).setValues(regionRows);
    dashboard.getRange(4, 5, regionRows.length, 1).setNumberFormat('$#,##0.00');
    const chart = dashboard.newChart()
      .asColumnChart()
      .addRange(dashboard.getRange(3, 4, regionRows.length + 1, 2))
      .setPosition(8, 1, 0, 0)
      .setOption('title', 'Validated Revenue by Region')
      .setOption('legend', { position: 'none' })
      .build();
    dashboard.insertChart(chart);
  }

  dashboard.getRange('A1:F1').setBackground('#0b253d').setFontColor('#ffffff').setFontWeight('bold').setFontSize(16);
  dashboard.getRange('A3:B3').setBackground('#dff6ed').setFontWeight('bold');
  dashboard.getRange('D3:E3').setBackground('#dff6ed').setFontWeight('bold');
  dashboard.autoResizeColumns(1, 6);
}

function createWeeklySummaryDraft() {
  const ss = SpreadsheetApp.getActive();
  const config = ss.getSheetByName(SHEETS.CONFIG);
  const dashboard = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!config || !dashboard) throw new Error('Run Set up demo workbook first.');

  const recipient = cleanText_(config.getRange('B2').getDisplayValue());
  if (!isValidEmail_(recipient)) {
    SpreadsheetApp.getUi().alert('Enter a valid recipient in Config!B2. No draft was created.');
    return;
  }

  const title = cleanText_(config.getRange('B3').getDisplayValue()) || 'Weekly Order Operations Summary';
  const readyOrders = dashboard.getRange('B4').getDisplayValue();
  const reviewRows = dashboard.getRange('B5').getDisplayValue();
  const revenue = dashboard.getRange('B6').getDisplayValue();
  const body = [
    'Hello,',
    '',
    'The latest order data has been processed and validated.',
    '',
    `Ready orders: ${readyOrders}`,
    `Rows requiring review: ${reviewRows}`,
    `Validated revenue: ${revenue}`,
    '',
    'Please review the Dashboard and Clean_Orders sheets for details.',
    '',
    'This message was created as a draft and has not been sent automatically.',
  ].join('\n');

  GmailApp.createDraft(recipient, title, body);
  logRun_('Create Gmail draft', 'Success', `Draft created for ${recipient}`);
  SpreadsheetApp.getUi().alert('Gmail draft created. Review it before sending.');
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function validateHeaders_(headers) {
  const required = ['Order ID', 'Order Date', 'Customer', 'Email', 'Region', 'Product', 'Quantity', 'Unit Price', 'Status'];
  required.forEach((header, index) => {
    if (cleanText_(headers[index]) !== header) {
      throw new Error(`Expected column ${index + 1} to be "${header}".`);
    }
  });
}

function cleanText_(value) {
  return String(value == null ? '' : value).trim().replace(/\s+/g, ' ');
}

function toTitleCase_(value) {
  return value.toLowerCase().replace(/\b\w/g, char => char.toUpperCase());
}

function normalizeDate_(value) {
  const date = value instanceof Date ? value : new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function normalizeRegion_(value) {
  const key = cleanText_(value).toLowerCase();
  const map = {
    'us': 'North America', 'usa': 'North America', 'north america': 'North America',
    'au': 'Australia', 'aus': 'Australia', 'australia': 'Australia',
    'uk': 'Europe', 'eu': 'Europe', 'europe': 'Europe',
    'apac': 'Asia-Pacific', 'asia': 'Asia-Pacific', 'asia-pacific': 'Asia-Pacific',
  };
  return map[key] || '';
}

function normalizeStatus_(value) {
  const key = cleanText_(value).toLowerCase();
  if (['paid', 'completed', 'complete'].includes(key)) return 'Completed';
  if (['pending', 'on hold', 'on-hold'].includes(key)) return 'Pending';
  if (['refunded', 'cancelled', 'canceled'].includes(key)) return 'Excluded';
  return 'Review';
}

function isValidEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleanText_(value));
}

function applyQaFormatting_(sheet, rowCount) {
  if (!rowCount) return;
  const qaRange = sheet.getRange(2, 11, rowCount, 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Ready').setBackground('#dff6ed').setFontColor('#087451').setRanges([qaRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Review').setBackground('#fff1d8').setFontColor('#8a5a00').setRanges([qaRange]).build(),
  ];
  sheet.setConditionalFormatRules(rules);
}

function logRun_(action, status, details) {
  const ss = SpreadsheetApp.getActive();
  const log = getOrCreateSheet_(ss, SHEETS.LOG);
  if (log.getLastRow() === 0) log.appendRow(['Timestamp', 'User', 'Action', 'Status', 'Details']);
  log.appendRow([new Date(), Session.getActiveUser().getEmail() || 'Active user', action, status, details]);
  log.setFrozenRows(1);
}
