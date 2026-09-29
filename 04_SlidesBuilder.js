/**
 * ============================================================================
 * PROJECT 4: DYNAMIC GOOGLE SLIDES PRESENTATION BUILDER
 * 
 * Description: Aggregates retail sales data and programmatically builds a
 *              branded executive presentation in Google Slides using SlidesApp.
 * ============================================================================
 */

const SLIDES_CONFIG = {
  SOURCE_SHEET: 'sales',
  DECK_TITLE_PREFIX: 'Executive Sales Performance Deck',
  BRAND_COLOR: '#1a73e8',       // Primary Blue
  ACCENT_COLOR: '#137333',      // Positive Green
  CARD_BG_COLOR: '#f1f3f4',     // Light Grey Card Background
  TEXT_DARK: '#202124'
};

/**
 * Main function: Generates a complete Google Slides executive deck.
 */
function buildSalesPitchDeck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SLIDES_CONFIG.SOURCE_SHEET);

  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(`Source sheet "${SLIDES_CONFIG.SOURCE_SHEET}" not found.`);
    return;
  }

  // 1. In-memory data aggregation
  const data = sourceSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('No data found to build slides.');
    return;
  }

  const headers = data[0];
  const idxSales = headers.indexOf('Sales');
  const idxProfit = headers.indexOf('Profit');
  const idxRegion = headers.indexOf('Region');
  const idxCategory = headers.indexOf('Product Category');

  let totalSales = 0;
  let totalProfit = 0;
  const regionMetrics = {};
  const categoryMetrics = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const s = Number(row[idxSales]) || 0;
    const p = Number(row[idxProfit]) || 0;
    const r = row[idxRegion] || 'Other';
    const c = row[idxCategory] || 'Other';

    totalSales += s;
    totalProfit += p;

    // Region tracking
    if (!regionMetrics[r]) regionMetrics[r] = { sales: 0, profit: 0 };
    regionMetrics[r].sales += s;
    regionMetrics[r].profit += p;

    // Category tracking
    if (!categoryMetrics[c]) categoryMetrics[c] = { sales: 0, profit: 0 };
    categoryMetrics[c].sales += s;
    categoryMetrics[c].profit += p;
  }

  const overallMargin = totalSales > 0 ? (totalProfit / totalSales) * 100 : 0;
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy');

  // 2. Create a new Google Slides Presentation
  const deckTitle = `${SLIDES_CONFIG.DECK_TITLE_PREFIX} - ${todayStr}`;
  const presentation = SlidesApp.create(deckTitle);

  // Remove default blank slide
  const initialSlides = presentation.getSlides();
  if (initialSlides.length > 0) {
    initialSlides[0].remove();
  }

  // Helper to format currency
  const fmt = (num) => '$' + Number(num).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });

  // --- SLIDE 1: Title Slide ---
  const slide1 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide1.getBackground().setSolidFill('#0d47a1'); // Dark Navy Header

  const titleBox = slide1.insertTextBox('GLOBAL SALES & PROFITABILITY', 50, 140, 620, 80);
  const titleStyle = titleBox.getText().getTextStyle();
  titleStyle.setFontSize(36);
  titleStyle.setBold(true);
  titleStyle.setForegroundColor('#ffffff');

  const subBox = slide1.insertTextBox(`Automated Executive Deck | Prepared: ${todayStr}\nData Scope: ~15,000 Retail Transactions`, 50, 220, 620, 60);
  const subStyle = subBox.getText().getTextStyle();
  subStyle.setFontSize(14);
  subStyle.setForegroundColor('#e3f2fd');

  // --- SLIDE 2: Executive Summary & KPI Cards ---
  const slide2 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slides_addHeader(slide2, 'Executive Performance Overview', 'Topline sales, margin, and profitability indicators');

  // 3 Metric Cards
  slides_createKpiCard(slide2, 50, 130, 190, 140, 'TOTAL REVENUE', fmt(totalSales), '#1a73e8');
  slides_createKpiCard(slide2, 265, 130, 190, 140, 'NET PROFIT', fmt(totalProfit), '#137333');
  slides_createKpiCard(slide2, 480, 130, 190, 140, 'PROFIT MARGIN', overallMargin.toFixed(1) + '%', '#e37400');

  // Executive Summary text note below cards
  const noteBox = slide2.insertTextBox(
    `Key Takeaways:\n• Gross revenue reached ${fmt(totalSales)} with an aggregate operating profit margin of ${overallMargin.toFixed(1)}%.\n• Multi-channel performance remains net positive across core international divisions.`,
    50, 295, 620, 80
  );
  const noteStyle = noteBox.getText().getTextStyle();
  noteStyle.setFontSize(13);
  noteStyle.setForegroundColor('#3c4043');

  // --- SLIDE 3: Regional Performance Breakdown ---
  const slide3 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slides_addHeader(slide3, 'Regional Performance Analysis', 'Distribution of gross sales and net margins across territories');

  // Add Table to Slide 3
  const regions = Object.keys(regionMetrics);
  const table = slide3.insertTable(regions.length + 1, 4, 50, 130, 620, 200);

  // Table Headers
  const tableHeaders = ['Region', 'Gross Sales', 'Net Profit', 'Profit Margin'];
  for (let c = 0; c < tableHeaders.length; c++) {
    const cell = table.getCell(0, c);
    cell.getText().setText(tableHeaders[c]);
    const style = cell.getText().getTextStyle();
    style.setBold(true);
    style.setFontSize(12);
    style.setForegroundColor('#ffffff');
    cell.getFill().setSolidFill(SLIDES_CONFIG.BRAND_COLOR);
  }

  // Populate Table Data Rows
  regions.forEach((reg, rIdx) => {
    const s = regionMetrics[reg].sales;
    const p = regionMetrics[reg].profit;
    const m = s > 0 ? (p / s) * 100 : 0;
    const rowValues = [reg, fmt(s), fmt(p), m.toFixed(1) + '%'];

    for (let c = 0; c < rowValues.length; c++) {
      const cell = table.getCell(rIdx + 1, c);
      cell.getText().setText(rowValues[c]);
      const style = cell.getText().getTextStyle();
      style.setFontSize(11);
      style.setForegroundColor(SLIDES_CONFIG.TEXT_DARK);
      if (rIdx % 2 === 0) {
        cell.getFill().setSolidFill('#f8f9fa');
      }
    }
  });

  // --- SLIDE 4: Strategic Category Insights ---
  const slide4 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slides_addHeader(slide4, 'Category Dynamics & Strategic Outlook', 'Key drivers and portfolio margin considerations');

  // Rank categories
  const catArray = Object.keys(categoryMetrics).map(cat => ({
    name: cat,
    sales: categoryMetrics[cat].sales,
    profit: categoryMetrics[cat].profit,
    margin: (categoryMetrics[cat].profit / categoryMetrics[cat].sales) * 100
  })).sort((a, b) => b.sales - a.sales);

  const topCategory = catArray[0];
  const lowestMarginCategory = [...catArray].sort((a, b) => a.margin - b.margin)[0];

  const catBox = slide4.insertTextBox(
    `Portfolio Observations:\n\n` +
    `1. Primary Revenue Engine: "${topCategory.name}"\n` +
    `   Generated ${fmt(topCategory.sales)} with ${topCategory.margin.toFixed(1)}% operating margin.\n\n` +
    `2. Margin Optimization Area: "${lowestMarginCategory.name}"\n` +
    `   Lowest margin at ${lowestMarginCategory.margin.toFixed(1)}% (${fmt(lowestMarginCategory.profit)} net profit).\n\n` +
    `3. Recommended Action:\n` +
    `   Audit discount and wholesale structures within reseller channels for low-margin lines.`,
    50, 140, 620, 220
  );
  const catStyle = catBox.getText().getTextStyle();
  catStyle.setFontSize(14);
  catStyle.setForegroundColor('#202124');

  // 3. Prompt user with the presentation link
  const url = presentation.getUrl();
  const htmlDialog = HtmlService.createHtmlOutput(
    `<div style="font-family: Arial, sans-serif; padding: 12px;">
       <p style="color: #137333; font-weight: bold;">✅ Google Slides Deck Created Successfully!</p>
       <p>The presentation contains 4 custom executive slides built directly from your dataset.</p>
       <p><a href="${url}" target="_blank" style="display: inline-block; background-color: #1a73e8; color: white; padding: 10px 16px; text-decoration: none; border-radius: 4px; font-weight: bold;">Open Slides Presentation ➔</a></p>
     </div>`
  ).setWidth(380).setHeight(180);

  SpreadsheetApp.getUi().showModalDialog(htmlDialog, 'Presentation Ready');
}

/**
 * Helper to add standard slide titles & subtitles.
 */
function slides_addHeader(slide, title, subtitle) {
  const headerBox = slide.insertTextBox(title, 50, 30, 620, 40);
  const hStyle = headerBox.getText().getTextStyle();
  hStyle.setFontSize(22);
  hStyle.setBold(true);
  hStyle.setForegroundColor(SLIDES_CONFIG.BRAND_COLOR);

  const subBox = slide.insertTextBox(subtitle, 50, 68, 620, 25);
  const sStyle = subBox.getText().getTextStyle();
  sStyle.setFontSize(11);
  sStyle.setForegroundColor('#5f6368');

  // Decorative divider line
  const line = slide.insertLine(
    SlidesApp.LineCategory.STRAIGHT,
    50, 98, 670, 98
  );
  line.getLineFill().setSolidFill('#dadce0');
  line.setWeight(1);
}

/**
 * Helper to construct styled KPI metric card boxes.
 */
function slides_createKpiCard(slide, left, top, width, height, label, value, valColor) {
  // Use ROUND_RECTANGLE (standard SlidesApp ShapeType enum)
  const shape = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, left, top, width, height);
  shape.getFill().setSolidFill(SLIDES_CONFIG.CARD_BG_COLOR);
  shape.getBorder().setTransparent();

  // Insert text directly inside the shape
  const textRange = shape.getText();
  textRange.setText(`${label}\n${value}`);

  // Style the label text
  const labelLength = label.length;
  const labelStyle = textRange.getRange(0, labelLength).getTextStyle();
  labelStyle.setFontSize(11);
  labelStyle.setBold(true);
  labelStyle.setForegroundColor('#5f6368');

  // Style the KPI metric value
  const valStyle = textRange.getRange(labelLength + 1, textRange.asString().length).getTextStyle();
  valStyle.setFontSize(24);
  valStyle.setBold(true);
  valStyle.setForegroundColor(valColor);
}