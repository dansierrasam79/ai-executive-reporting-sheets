/**
 * PART 3: Opens the AI Chat Sidebar and processes user questions.
 */
function openAIChatSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("06_AIChatModal")
    .setTitle("Gemini Data Analyst")
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

function askGeminiAnalyst(userQuestion) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) throw new Error("API Key not found in Script Properties!");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const digestSheet = ss.getSheetByName("proj3_exec_summary_digest");

  const revenue = digestSheet ? digestSheet.getRange("B3").getDisplayValue() : "N/A";
  const profit  = digestSheet ? digestSheet.getRange("B5").getDisplayValue() : "N/A";
  const margin  = digestSheet ? digestSheet.getRange("B6").getDisplayValue() : "N/A";

  const prompt = `You are an AI financial analyst embedded in the company's Google Sheets system.
Current Performance Context:
- Revenue: ${revenue}
- Net Profit: ${profit}
- Profit Margin: ${margin}

User Question: "${userQuestion}"

Provide a concise, direct, professional answer (2-3 sentences max) based on this business context.`;

  const url = "https://generativelanguage.googleapis.com/v1beta/interactions";
  const payload = {
    model: "gemini-3.8-flash",
    input: prompt
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: { "x-goog-api-key": apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const json = JSON.parse(response.getContentText());
  let reply = "";
  if (json.output_text) {
    reply = json.output_text;
  } else if (json.steps && json.steps.length > 0) {
    for (const step of json.steps) {
      if (step.content && Array.isArray(step.content)) {
        for (const item of step.content) {
          if (item.text) reply += item.text + "\n";
        }
      }
    }
  }
  return reply.trim() || "No response generated.";
}