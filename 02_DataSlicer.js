/**
 * ============================================================================
 * PROJECT 2: CUSTOM UI SIDEBAR & REGIONAL DATA SLICER (BACKEND)
 * ============================================================================
 */

const SLICER_CONFIG = {
  SOURCE_SHEET: 'sales',
  TARGET_SHEET: 'filtered_segment',
  HEADER_COLOR: '#1a73e8', // Material Blue
  HEADER_FONT_COLOR: '#ffffff'
};

/**
 * Opens the interactive HTML sidebar inside Google Sheets.
 */
function openDataSlicerSidebar() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('02_Sidebar')
    .setTitle('Regional Data Slicer')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

/**
 * Fetches distinct filter choices dynamically from the data sheet.
 * Called automatically by client-side JS when the sidebar opens.
 */
function slicer_getFilterOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SLICER_CONFIG.SOURCE_SHEET);
  if (!sheet) throw new Error(`Source sheet "${SLICER_CONFIG.SOURCE_SHEET}" not found.`);

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { years: [], regions: [], channels: [] };

  const headers = data[0];
  const idxYear = headers.indexOf('Year');
  const idxRegion = headers.indexOf('Region');
  const idxChannel = headers.indexOf('Channel');

  const years = new Set();
  const regions = new Set();
  const channels = new Set();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (idxYear !== -1 && row[idxYear]) years.add(String(row[idxYear]));
    if (idxRegion !== -1 && row[idxRegion]) regions.add(String(row[idxRegion]));
    if (idxChannel !== -1 && row[idxChannel]) channels.add(String(row[idxChannel]));
  }

  return {
    years: Array.from(years).sort(),
    regions: Array.from(regions).sort(),
    channels: Array.from(channels).sort()
  };
}

/**
 * Filters the dataset based on selected criteria, exports matching rows
 * to the target sheet, and returns high-level summary KPIs to the sidebar.
 */
function slicer_applyFilterAndExtract(criteria) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SLICER_CONFIG.SOURCE_SHEET);
  if (!sourceSheet) throw new Error('Source sheet not found.');

  const data = sourceSheet.getDataRange().getValues();
  const headers = data[0];

  const idxYear = headers.indexOf('Year');
  const idxRegion = headers.indexOf('Region');
  const idxChannel = headers.indexOf('Channel');
  const idxSales = headers.indexOf('Sales');
  const idxProfit = headers.indexOf('Profit');
  const idxDate = headers.indexOf('Order Date');

  const matchingRows = [];
  let totalSales = 0;
  let totalProfit = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const matchYear = !criteria.year || String(row[idxYear]) === criteria.year;
    const matchRegion = !criteria.region || String(row[idxRegion]) === criteria.region;
    const matchChannel = !criteria.channel || String(row[idxChannel]) === criteria.channel;

    if (matchYear && matchRegion && matchChannel) {
      const rowClone = [...row];
      if (idxDate !== -1 && rowClone[idxDate] instanceof Date) {
        rowClone[idxDate] = Utilities.formatDate(rowClone[idxDate], Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      matchingRows.push(rowClone);
      totalSales += Number(row[idxSales]) || 0;
      totalProfit += Number(row[idxProfit]) || 0;
    }
  }

  // Create or refresh target sheet
  let targetSheet = ss.getSheetByName(SLICER_CONFIG.TARGET_SHEET);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(SLICER_CONFIG.TARGET_SHEET);
  } else {
    targetSheet.clear();
  }

  // Write headers
  targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (matchingRows.length > 0) {
    targetSheet.getRange(2, 1, matchingRows.length, headers.length).setValues(matchingRows);

    // Apply currency formatting to Sales (col E) and Profit (col F)
    if (idxSales !== -1) {
      targetSheet.getRange(2, idxSales + 1, matchingRows.length, 1).setNumberFormat('$#,##0.00');
    }
    if (idxProfit !== -1) {
      targetSheet.getRange(2, idxProfit + 1, matchingRows.length, 1).setNumberFormat('$#,##0.00');
    }
  }

  // Style header row
  const headerRange = targetSheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground(SLICER_CONFIG.HEADER_COLOR)
             .setFontColor(SLICER_CONFIG.HEADER_FONT_COLOR)
             .setFontWeight('bold');
  targetSheet.setFrozenRows(1);
  targetSheet.autoResizeColumns(1, Math.min(headers.length, 15));

  ss.setActiveSheet(targetSheet);

  const marginPct = totalSales > 0 ? (totalProfit / totalSales) * 100 : 0;

  return {
    rowCount: matchingRows.length,
    totalSales: totalSales,
    totalProfit: totalProfit,
    marginPct: marginPct
  };
}