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

    // Fetch full detail for each new order — the list endpoint returns summaries only
    // and does not include cost.shipping; GET /orders/{id} returns the full cost object.
    const rows = newOrders.map(o => {
      const detail     = bricklinkRequest(`orders/${o.order_id}`, 'GET');
      const d          = (detail.meta && detail.meta.code === 200) ? detail.data : o;
      const shipping   = d.cost && d.cost.shipping != null ? parseFloat(d.cost.shipping) : '';
      debugLog('syncOrders', `Order ${o.order_id}: cost.shipping=${shipping !== '' ? shipping : 'not returned'}`);

      return [
        d.order_id,
        d.date_ordered ? new Date(d.date_ordered).toLocaleDateString('en-GB') : '',
        d.buyer_name   || '',
        d.total_count  || 0,
        d.unique_count || 0,
        shipping,
        '',  // Shipping Actual — manual entry
        d.cost && d.cost.grand_total != null ? parseFloat(d.cost.grand_total) : '',
        d.cost    ? d.cost.currency_code : '',
        d.payment ? d.payment.method     : '',
        d.status  || '',
        '',  // Refund — manual entry
        ''   // Notes — manual entry
      ];
    });

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
    debugLog('syncOrders', `Error: ${e.message}`);
    ui.alert(`❌ Sync failed:\n\n${e.message}`);
    logStatus(`Order sync failed: ${new Date().toLocaleString()}`);
  }
}
