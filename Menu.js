/**
 * Unified custom menu for the entire Sales Automation Suite.
 * Creates a custom toolbar menu when the spreadsheet is opened.
 **/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Sales Ops Suite')
    .addItem('1. Run Margin & Loss Audit', 'runProfitAudit')
    .addItem('2. Open Regional Data Slicer', 'openDataSlicerSidebar')
    .addItem('3. Send Executive Email Digest & PDF', 'sendExecutiveDigest')
    .addItem('4. Generate Google Slides Deck', 'buildSalesPitchDeck')
    .addItem('5. Generate AI Executive Commentary', 'generateAiCommentary')
    .addItem('⏰ Schedule Weekly Email (Monday 8 AM)', 'digest_scheduleWeeklyTrigger')
    .addToUi();

  // Dedicated AI Executive Suite Menu:
  ui.createMenu("🤖 AI Suite")
    .addItem("Generate Executive Commentary", "generateAIExecutiveCommentary")
    .addItem("Send Executive Email Briefing", "sendExecutiveBriefingEmail")
    .addItem("Run Forensic Loss Audit", "runLossLeakageAudit")
    .addSeparator()
    .addItem("💬 Open AI Data Chatbot", "openAIChatSidebar")
    .addToUi();
}

