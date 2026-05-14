// ============================================================
// Orders.gs — Order sync logic
// ============================================================

const LAST_ORDER_SYNC_KEY = 'LAST_ORDER_SYNC';

function setupOrdersTab() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ORDERS_TAB);
  if (!sheet) sheet = ss.insertSheet(ORDERS_TAB);

  sheet.clearContents();

  const headerRange = sheet.getRange(1, 1, 1, ORDERS_HEADERS.length);
  headerRange.setValues([ORDERS_HEADERS]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground(HEADER_COLOUR);
  sheet.setFrozenRows(1);

  COL_WIDTHS.forEach((w, i) => {
    sheet.setColumnWidth(i + 1, w);
    if (MANUAL_COLS.includes(i + 1)) {
      sheet.getRange(1, i + 1).setBackground('#FFE082'); // darker yellow for manual headers
    }
  });

  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  const protection = sheet.protect().setDescription('Synced BrickLink data — do not edit');
  const manualRanges = MANUAL_COLS.map(c => sheet.getRange(2, c, sheet.getMaxRows() - 1, 1));
  protection.setUnprotectedRanges(manualRanges);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) protection.setDomainEdit(false);
}

function syncOrders() {
  const ui    = SpreadsheetApp.getUi();
  const props = PropertiesService.getUserProperties();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(ORDERS_TAB);
  if (!sheet) {
    setupOrdersTab();
    sheet = ss.getSheetByName(ORDERS_TAB);
  }

  logStatus('Syncing orders…');
  resetApiCallLog();

  try {
    const data = bricklinkRequest('orders', 'GET', { direction: 'in' });

    if (!data.meta || data.meta.code !== 200) {
      ui.alert(`⚠️ API error:\n\n${JSON.stringify(data.meta)}`);
      logStatus(`Order sync failed: ${new Date().toLocaleString()}`);
      return;
    }

    const orders = data.data || [];

    const lastRow = sheet.getLastRow();
    const existingIds = new Set();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 1).getValues()
        .forEach(r => { if (r[0]) existingIds.add(String(r[0])); });
    }

    const newOrders = orders.filter(o => !existingIds.has(String(o.order_id)));

    if (newOrders.length === 0) {
      ui.alert('✅ Orders up to date — nothing new to import.');
      logStatus(`Orders synced: ${new Date().toLocaleString()} — no new orders`);
      flushApiCallLog('syncOrders');
      return;
    }

    newOrders.sort((a, b) => new Date(b.date_ordered) - new Date(a.date_ordered));

    const rows = newOrders.map(o => [
      o.order_id,
      o.date_ordered ? new Date(o.date_ordered).toLocaleDateString('en-GB') : '',
      o.buyer_name   || '',
      o.total_count  || 0,
      o.unique_count || 0,
      o.cost && o.cost.shipping    != null ? parseFloat(o.cost.shipping)    : '',
      '',  // Shipping Actual — manual entry
      o.cost && o.cost.grand_total != null ? parseFloat(o.cost.grand_total) : '',
      o.cost  ? o.cost.currency_code : '',
      o.payment ? o.payment.method   : '',
      o.status  || '',
      '',  // Refund — manual entry
      ''   // Notes — manual entry
    ]);

    const insertRow = Math.max(sheet.getLastRow(), 1) + 1;
    sheet.getRange(insertRow, 1, rows.length, ORDERS_HEADERS.length).setValues(rows);

    MANUAL_COLS.forEach(c => {
      sheet.getRange(insertRow, c, rows.length, 1).setBackground(MANUAL_COLOUR);
    });

    props.setProperty(LAST_ORDER_SYNC_KEY, new Date().toISOString());
    flushApiCallLog('syncOrders');

    ui.alert(`✅ Synced ${newOrders.length} new order(s).`);
    logStatus(`Orders synced: ${new Date().toLocaleString()} — ${newOrders.length} new`);

  } catch (e) {
    ui.alert(`❌ Sync failed:\n\n${e.message}`);
    logStatus(`Order sync failed: ${new Date().toLocaleString()}`);
  }
}
