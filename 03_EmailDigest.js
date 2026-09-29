/**
 * ============================================================================
 * PROJECT 3: AUTOMATED SCHEDULED EMAIL DIGEST & PDF EXPORTER
 * 
 * Description: Computes high-level sales and margin KPIs across ~15k records,
 *              creates an executive summary tab, exports it as a PDF attachment,
 *              and emails a styled HTML digest to designated stakeholders.
 * ============================================================================
 */

const DIGEST_CONFIG = {
  SOURCE_SHEET: 'sales',
  TEMP_SUMMARY_SHEET: 'exec_summary_digest',
  // Change to your desired recipient or leave dynamic:
  RECIPIENT_EMAIL: Session.getActiveUser().getEmail() || 'nutan@example.com',
  EMAIL_SUBJECT_PREFIX: '📈 Executive Sales & Profit Digest',
  BRAND_COLOR: '#1a73e8',
  SUCCESS_COLOR: '#137333',
  ACCENT_COLOR: '#202124'
};

/**
 * Main execution function: Computes KPIs, creates PDF, and sends executive email.
 */
function sendExecutiveDigest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(DIGEST_CONFIG.SOURCE_SHEET);

  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(`Source sheet "${DIGEST_CONFIG.SOURCE_SHEET}" not found.`);
    return;
  }

  // 1. In-Memory Calculation of Core KPIs
  const data = sourceSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('No data found for executive digest.');
    return;
  }

  const headers = data[0];
  const idxSales = headers.indexOf('Sales');
  const idxCost = headers.indexOf('Cost of Sales');
  const idxProfit = headers.indexOf('Profit');
  const idxProduct = headers.indexOf('Product Name');
  const idxRegion = headers.indexOf('Region');
  const idxChannel = headers.indexOf('Channel');

  let totalSales = 0;
  let totalCost = 0;
  let totalProfit = 0;

  const productSalesMap = {};
  const regionMap = {};
  const channelMap = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const s = Number(row[idxSales]) || 0;
    const c = Number(row[idxCost]) || 0;
    const p = Number(row[idxProfit]) || 0;
    const prod = row[idxProduct] || 'Unknown';
    const reg = row[idxRegion] || 'Unknown';
    const chan = row[idxChannel] || 'Unknown';

    totalSales += s;
    totalCost += c;
    totalProfit += p;

    // Track product sales
    productSalesMap[prod] = (productSalesMap[prod] || 0) + s;

    // Track region metrics
    if (!regionMap[reg]) regionMap[reg] = { sales: 0, profit: 0 };
    regionMap[reg].sales += s;
    regionMap[reg].profit += p;

    // Track channel metrics
    if (!channelMap[chan]) channelMap[chan] = { sales: 0, profit: 0 };
    channelMap[chan].sales += s;
    channelMap[chan].profit += p;
  }

  const marginPct = totalSales > 0 ? (totalProfit / totalSales) * 100 : 0;

  // Rank Top 5 Products
  const topProducts = Object.keys(productSalesMap)
    .map(name => ({ name: name, sales: productSalesMap[name] }))
    .sort((a, b) => b.sales - a.sales)
    .slice(0, 5);

  // 2. Build or Overwrite the Executive Summary Tab
  let summarySheet = ss.getSheetByName(DIGEST_CONFIG.TEMP_SUMMARY_SHEET);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(DIGEST_CONFIG.TEMP_SUMMARY_SHEET);
  } else {
    summarySheet.clear();
  }

  // Populate Summary Tab Content
  summarySheet.getRange('A1').setValue('Executive Performance Digest');
  summarySheet.getRange('A1:D1').merge()
    .setBackground(DIGEST_CONFIG.BRAND_COLOR)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center');

  summarySheet.getRange('A3:B6').setValues([
    ['Total Revenue', totalSales],
    ['Cost of Goods Sold', totalCost],
    ['Net Profit', totalProfit],
    ['Profit Margin %', marginPct / 100]
  ]);
  summarySheet.getRange('A3:A6').setFontWeight('bold');
  summarySheet.getRange('B3:B5').setNumberFormat('$#,##0.00');
  summarySheet.getRange('B6').setNumberFormat('0.00%');

  // Add Region Breakdown table to sheet
  summarySheet.getRange('A8:C8').setValues([['Region', 'Revenue', 'Profit']]);
  summarySheet.getRange('A8:C8').setBackground('#f1f3f4').setFontWeight('bold');
  const regionRows = Object.keys(regionMap).map(r => [r, regionMap[r].sales, regionMap[r].profit]);
  summarySheet.getRange(9, 1, regionRows.length, 3).setValues(regionRows);
  summarySheet.getRange(9, 2, regionRows.length, 2).setNumberFormat('$#,##0.00');

  // Add Top Products table to sheet
  const startProdRow = 10 + regionRows.length;
  summarySheet.getRange(startProdRow, 1, 1, 2).setValues([['Top 5 Products', 'Revenue']]);
  summarySheet.getRange(startProdRow, 1, 1, 2).setBackground('#f1f3f4').setFontWeight('bold');
  const prodRows = topProducts.map(p => [p.name, p.sales]);
  summarySheet.getRange(startProdRow + 1, 1, prodRows.length, 2).setValues(prodRows);
  summarySheet.getRange(startProdRow + 1, 2, prodRows.length, 1).setNumberFormat('$#,##0.00');

  summarySheet.autoResizeColumns(1, 4);
  SpreadsheetApp.flush();

  // 3. Export Summary Tab as PDF
  const pdfBlob = digest_createPdfBlob(ss, summarySheet);

  // 4. Generate HTML Email Body
  const emailHtml = digest_buildHtmlEmail({
    totalSales,
    totalCost,
    totalProfit,
    marginPct,
    regionMap,
    channelMap,
    topProducts
  });

  // 5. Send Email via GmailApp
  const recipient = DIGEST_CONFIG.RECIPIENT_EMAIL;
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const subject = `${DIGEST_CONFIG.EMAIL_SUBJECT_PREFIX} - ${today}`;

  GmailApp.sendEmail(recipient, subject, 'Please view this email in an HTML compatible reader.', {
    htmlBody: emailHtml,
    attachments: [pdfBlob],
    name: 'Sales Ops Automation'
  });

  SpreadsheetApp.getUi().alert(
    'Email Digest Sent',
    `Executive sales report & PDF export successfully delivered to ${recipient}.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Exports a specific sheet to a PDF Blob using UrlFetchApp and Google Sheets export URL.
 */
function digest_createPdfBlob(spreadsheet, sheet) {
  const ssId = spreadsheet.getId();
  const sheetId = sheet.getSheetId();

  // Sheets PDF export URL parameters
  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?` +
    `format=pdf&size=letter&portrait=true&fitw=true&gridlines=false` +
    `&printtitle=false&sheetnames=false&fzr=false&gid=${sheetId}`;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': `Bearer ${token}` },
    muteHttpExceptions: true
  });

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return response.getBlob().setName(`Executive_Sales_Digest_${today}.pdf`);
}

/**
 * Builds a styled, modern responsive HTML email body with metric cards.
 */
function digest_buildHtmlEmail(metrics) {
  const formatCurrency = (val) => '$' + Number(val).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const formatPct = (val) => Number(val).toFixed(2) + '%';

  // Region breakdown HTML rows
  let regionRowsHtml = '';
  for (const reg in metrics.regionMap) {
    regionRowsHtml += `
      <tr>
        <td style="padding: 8px 12px; border-bottom: 1px solid #eee;">${reg}</td>
        <td style="padding: 8px 12px; border-bottom: 1px solid #eee; text-align: right;">${formatCurrency(metrics.regionMap[reg].sales)}</td>
        <td style="padding: 8px 12px; border-bottom: 1px solid #eee; text-align: right; color: #137333;">${formatCurrency(metrics.regionMap[reg].profit)}</td>
      </tr>`;
  }

  // Top products HTML rows
  let productRowsHtml = '';
  metrics.topProducts.forEach((p, idx) => {
    productRowsHtml += `
      <tr>
        <td style="padding: 8px 12px; border-bottom: 1px solid #eee;">${idx + 1}. ${p.name}</td>
        <td style="padding: 8px 12px; border-bottom: 1px solid #eee; text-align: right; font-weight: bold;">${formatCurrency(p.sales)}</td>
      </tr>`;
  });

  return `
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; max-width: 650px; margin: auto; padding: 20px; color: #333; background-color: #fafbfc;">
      <div style="background-color: #ffffff; padding: 24px; border-radius: 8px; border: 1px solid #e1e4e8;">
        
        <h2 style="color: #1a73e8; margin-top: 0; margin-bottom: 6px;">📊 Executive Sales & Margin Digest</h2>
        <p style="color: #586069; font-size: 13px; margin-top: 0; margin-bottom: 20px;">Automated summary compiled across global transactions (~15k orders).</p>
        
        <!-- KPI CARDS -->
        <div style="display: flex; gap: 12px; margin-bottom: 24px;">
          <div style="flex: 1; background-color: #f6f8fa; padding: 14px; border-radius: 6px; border: 1px solid #e1e4e8; text-align: center;">
            <div style="font-size: 11px; text-transform: uppercase; color: #586069; font-weight: bold;">Gross Revenue</div>
            <div style="font-size: 18px; font-weight: bold; color: #24292e; margin-top: 4px;">${formatCurrency(metrics.totalSales)}</div>
          </div>
          <div style="flex: 1; background-color: #f6f8fa; padding: 14px; border-radius: 6px; border: 1px solid #e1e4e8; text-align: center;">
            <div style="font-size: 11px; text-transform: uppercase; color: #586069; font-weight: bold;">Net Profit</div>
            <div style="font-size: 18px; font-weight: bold; color: #137333; margin-top: 4px;">${formatCurrency(metrics.totalProfit)}</div>
          </div>
          <div style="flex: 1; background-color: #f6f8fa; padding: 14px; border-radius: 6px; border: 1px solid #e1e4e8; text-align: center;">
            <div style="font-size: 11px; text-transform: uppercase; color: #586069; font-weight: bold;">Profit Margin</div>
            <div style="font-size: 18px; font-weight: bold; color: #1a73e8; margin-top: 4px;">${formatPct(metrics.marginPct)}</div>
          </div>
        </div>

        <!-- REGION TABLE -->
        <h3 style="color: #24292e; font-size: 15px; margin-bottom: 8px; border-bottom: 2px solid #eaecef; padding-bottom: 6px;">Regional Performance</h3>
        <table style="width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 24px;">
          <thead>
            <tr style="background-color: #f6f8fa; text-align: left;">
              <th style="padding: 8px 12px; border-bottom: 1px solid #ddd;">Region</th>
              <th style="padding: 8px 12px; border-bottom: 1px solid #ddd; text-align: right;">Sales</th>
              <th style="padding: 8px 12px; border-bottom: 1px solid #ddd; text-align: right;">Profit</th>
            </tr>
          </thead>
          <tbody>
            ${regionRowsHtml}
          </tbody>
        </table>

        <!-- TOP PRODUCTS TABLE -->
        <h3 style="color: #24292e; font-size: 15px; margin-bottom: 8px; border-bottom: 2px solid #eaecef; padding-bottom: 6px;">Top 5 Revenue Drivers</h3>
        <table style="width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 20px;">
          <thead>
            <tr style="background-color: #f6f8fa; text-align: left;">
              <th style="padding: 8px 12px; border-bottom: 1px solid #ddd;">Product</th>
              <th style="padding: 8px 12px; border-bottom: 1px solid #ddd; text-align: right;">Revenue</th>
            </tr>
          </thead>
          <tbody>
            ${productRowsHtml}
          </tbody>
        </table>

        <div style="font-size: 11px; color: #6a737d; margin-top: 24px; padding-top: 12px; border-top: 1px solid #eaecef; text-align: center;">
          📎 <em>A clean single-page PDF snapshot is attached to this email.</em><br>
          Generated automatically via <strong>Sales Ops Suite (Google Apps Script)</strong>.
        </div>
      </div>
    </div>
  `;
}

/**
 * Creates an automatic time-driven trigger to run this digest every Monday at 8:00 AM.
 */
function digest_scheduleWeeklyTrigger() {
  // Delete existing triggers for this function to prevent duplicate emails
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendExecutiveDigest') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create new Monday 8:00 AM trigger
  ScriptApp.newTrigger('sendExecutiveDigest')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  SpreadsheetApp.getUi().alert(
    'Schedule Activated',
    'Executive Digest has been scheduled to run automatically every Monday at 8:00 AM.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}