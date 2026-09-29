/**
 * PART 2: Scans proj1_audit_loss_log and has Gemini run a root-cause forensic audit.
 */
function runLossLeakageAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lossSheet = ss.getSheetByName("proj1_audit_loss_log");

  if (!lossSheet) {
    SpreadsheetApp.getUi().alert("Sheet 'proj1_audit_loss_log' not found!");
    return;
  }

  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) {
    SpreadsheetApp.getUi().alert("API Key not found in Script Properties!");
    return;
  }

  ss.toast("Aggregating loss records...", "Audit In Progress", 5);

  // Read data: Headers in Row 1, data starts Row 2
  // Columns: Order Date(0), Product Name(1), Category(2), Channel(3), Region(4), Country(5), Cost of Sales(6), Sales(7), Net Loss(8), Margin %(9)
  const lastRow = lossSheet.getLastRow();
  const data = lossSheet.getRange(2, 1, lastRow - 1, 10).getValues();

  let totalLoss = 0;
  const categoryLoss = {};
  const channelLoss = {};
  const countryLoss = {};

  for (let i = 0; i < data.length; i++) {
    const category = data[i][2];
    const channel  = data[i][3];
    const country  = data[i][5];
    const lossVal  = Math.abs(parseFloat(data[i][8]) || 0);

    totalLoss += lossVal;
    categoryLoss[category] = (categoryLoss[category] || 0) + lossVal;
    channelLoss[channel]   = (channelLoss[channel] || 0) + lossVal;
    countryLoss[country]   = (countryLoss[country] || 0) + lossVal;
  }

  // Sort and pick top 3 contributors
  const topCategories = Object.entries(categoryLoss).sort((a,b) => b[1] - a[1]).slice(0, 3);
  const topChannels   = Object.entries(channelLoss).sort((a,b) => b[1] - a[1]).slice(0, 3);
  const topCountries  = Object.entries(countryLoss).sort((a,b) => b[1] - a[1]).slice(0, 3);

  const prompt = `You are an elite Operations Auditor and Financial Controller.
Analyze this summary of unprofitable transactions (${data.length} total loss transactions amounting to $${totalLoss.toLocaleString()} in net loss):

Top Loss Categories:
${topCategories.map(([k, v]) => `- ${k}: $${Math.round(v).toLocaleString()}`).join("\n")}

Top Loss Channels:
${topChannels.map(([k, v]) => `- ${k}: $${Math.round(v).toLocaleString()}`).join("\n")}

Top Loss Countries:
${topCountries.map(([k, v]) => `- ${k}: $${Math.round(v).toLocaleString()}`).join("\n")}

Please provide:
1. Root-Cause Analysis (Why are these specific categories/channels hemorrhaging margin?)
2. Financial Risk Assessment
3. Actionable Remediation Plan (3 immediate operational steps to plug these leakages)`;

  ss.toast("Gemini is diagnosing root causes...", "AI Analysis", 5);

  const url = "https://generativelanguage.googleapis.com/v1beta/interactions";
  const payload = {
    model: "gemini-3.8-flash",
    input: prompt
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "x-goog-api-key": apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const res = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(res.getContentText());

    let auditReport = "";
    if (json.output_text) {
      auditReport = json.output_text;
    } else if (json.steps && json.steps.length > 0) {
      for (const step of json.steps) {
        if (step.content && Array.isArray(step.content)) {
          for (const item of step.content) {
            if (item.text) auditReport += item.text + "\n";
          }
        }
      }
    }

    // Place the report into proj1_audit_loss_log (Column L3)
    lossSheet.getRange("L1").setValue("AI Forensic Audit Report").setFontWeight("bold").setFontSize(12);
    const reportCell = lossSheet.getRange("L3");
    reportCell.setValue(auditReport.trim());
    reportCell.setWrap(true);
    reportCell.setVerticalAlignment("top");
    lossSheet.setColumnWidth(12, 500);

    ss.toast("Audit complete! Report written to Column L.", "Success", 5);
  } catch (err) {
    SpreadsheetApp.getUi().alert("Loss Audit error: " + err.message);
  }
}