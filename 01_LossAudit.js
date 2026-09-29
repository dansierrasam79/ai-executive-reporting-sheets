/**
 * ============================================================================
 * PROJECT 1: PROFIT MARGIN ANOMALY & LOSS AUDIT WATCHDOG
 * 
 * Description: Scans ~15,000 sales transactions in memory, detects unprofitable
 *              orders, extracts them to an "audit_loss_log" tab, and formats
 *              the output for financial review.
 * ============================================================================
 */

const AUDIT_CONFIG = {
  SOURCE_SHEET: 'sales',
  AUDIT_SHEET: 'audit_loss_log',
  // Flag any order where profit is negative or margin is below this rate (0.00 = negative only)
  MARGIN_THRESHOLD: 0.00,
  HEADER_COLOR: '#b71c1c',      // Dark crimson red
  HEADER_FONT_COLOR: '#ffffff'
};

/**
 * Main execution function: Audits sales transactions and generates the audit sheet.
 */
function runProfitAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(AUDIT_CONFIG.SOURCE_SHEET);

  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(`Source sheet "${AUDIT_CONFIG.SOURCE_SHEET}" was not found.`);
    return;
  }

  // 1. Batch read entire dataset into memory (sub-second performance for 15k rows)
  const data = sourceSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('No transaction records found to audit.');
    return;
  }

  const headers = data[0];

  // Dynamically map column positions
  const idxDate = headers.indexOf('Order Date');
  const idxSales = headers.indexOf('Sales');
  const idxCost = headers.indexOf('Cost of Sales');
  const idxProfit = headers.indexOf('Profit');
  const idxChannel = headers.indexOf('Channel');
  const idxProduct = headers.indexOf('Product Name');
  const idxCategory = headers.indexOf('Product Category');
  const idxRegion = headers.indexOf('Region');
  const idxCountry = headers.indexOf('Country');

  if (idxProfit === -1 || idxSales === -1) {
    SpreadsheetApp.getUi().alert('Error: Missing required columns ("Profit" or "Sales").');
    return;
  }

  // 2. Scan rows for loss-making orders
  const flaggedRows = [];
  let totalLoss = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const sales = Number(row[idxSales]) || 0;
    const profit = Number(row[idxProfit]) || 0;
    const margin = sales > 0 ? (profit / sales) : 0;

    if (profit < 0 || margin < AUDIT_CONFIG.MARGIN_THRESHOLD) {
      let orderDate = row[idxDate];
      if (orderDate instanceof Date) {
        orderDate = Utilities.formatDate(orderDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }

      flaggedRows.push([
        orderDate,
        row[idxProduct] || 'N/A',
        row[idxCategory] || 'N/A',
        row[idxChannel] || 'N/A',
        row[idxRegion] || 'N/A',
        row[idxCountry] || 'N/A',
        row[idxCost] || 0,
        sales,
        profit,
        margin
      ]);

      totalLoss += profit;
    }
  }

  // 3. Prepare or reset the Audit Sheet
  let auditSheet = ss.getSheetByName(AUDIT_CONFIG.AUDIT_SHEET);
  if (!auditSheet) {
    auditSheet = ss.insertSheet(AUDIT_CONFIG.AUDIT_SHEET);
  } else {
    auditSheet.clear();
  }

  // 4. Define and write headers
  const auditHeaders = [
    'Order Date',
    'Product Name',
    'Category',
    'Channel',
    'Region',
    'Country',
    'Cost of Sales',
    'Sales',
    'Net Loss',
    'Margin %'
  ];

  auditSheet.getRange(1, 1, 1, auditHeaders.length).setValues([auditHeaders]);

  // 5. Write and format data
  if (flaggedRows.length > 0) {
    // Sort ascending by profit (largest dollar losses at the top)
    flaggedRows.sort((a, b) => a[8] - b[8]);

    auditSheet.getRange(2, 1, flaggedRows.length, auditHeaders.length).setValues(flaggedRows);

    // Apply currency format to Cost, Sales, Net Loss (Cols 7, 8, 9)
    auditSheet.getRange(2, 7, flaggedRows.length, 3).setNumberFormat('$#,##0.00');
    // Apply percentage format to Margin % (Col 10)
    auditSheet.getRange(2, 10, flaggedRows.length, 1).setNumberFormat('0.00%');
  }

  // 6. Professional visual hygiene
  const headerRange = auditSheet.getRange(1, 1, 1, auditHeaders.length);
  headerRange
    .setBackground(AUDIT_CONFIG.HEADER_COLOR)
    .setFontColor(AUDIT_CONFIG.HEADER_FONT_COLOR)
    .setFontWeight('bold');

  auditSheet.setFrozenRows(1);
  auditSheet.autoResizeColumns(1, auditHeaders.length);

  // Switch view to the audit tab
  ss.setActiveSheet(auditSheet);

  // 7. Executive Alert Dialog
  const formattedLoss = Utilities.formatString('$%,.2f', Math.abs(totalLoss));
  SpreadsheetApp.getUi().alert(
    'Audit Complete',
    `Identified ${flaggedRows.length} loss-making transactions totaling ${formattedLoss} in negative margins.\n\nResults have been extracted to "${AUDIT_CONFIG.AUDIT_SHEET}".`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}