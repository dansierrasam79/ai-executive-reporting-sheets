function generateAIExecutiveCommentary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("proj3_exec_summary_digest");
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Error: Tab 'proj3_exec_summary_digest' not found!");
    return;
  }

  // 1. Fetch saved key from Script Properties
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");

  if (!apiKey) {
    SpreadsheetApp.getUi().alert("API Key not found in Script Properties.");
    return;
  }

  // 2. Read performance digest data from the sheet
  const totalRevenue = sheet.getRange("B3").getDisplayValue();
  const cogs         = sheet.getRange("B4").getDisplayValue();
  const netProfit    = sheet.getRange("B5").getDisplayValue();
  const margin       = sheet.getRange("B6").getDisplayValue();
  
  // Regional breakdown (A8:C11)
  const regionRows = sheet.getRange("A8:C11").getDisplayValues();
  let regionalText = "";
  for (let i = 1; i < regionRows.length; i++) {
    regionalText += `- ${regionRows[i][0]}: Revenue ${regionRows[i][1]} | Profit ${regionRows[i][2]}\n`;
  }

  // Top 5 Products (A13:B18)
  const prodRows = sheet.getRange("A13:B18").getDisplayValues();
  let productText = "";
  for (let i = 1; i < prodRows.length; i++) {
    productText += `- ${prodRows[i][0]}: ${prodRows[i][1]}\n`;
  }

  // 3. Formulate the prompt
  const prompt = `You are a Chief Financial Officer / Executive Financial Analyst.
Analyze the following business performance metrics and write a clear, 3-section executive commentary:

Key Financials:
- Total Revenue: ${totalRevenue}
- Cost of Goods Sold: ${cogs}
- Net Profit: ${netProfit}
- Profit Margin: ${margin}

Regional Breakdown:
${regionalText}

Top 5 Revenue Generators:
${productText}

Please structure your response into:
1. Executive Summary & Top-Line Performance
2. Regional Dynamics & Profitability
3. Product Concentration & Strategic Recommendations

Keep it professional, high-impact, and formatted with clean bullet points.`;

  // 4. Call Gemini 3.8 Flash Interactions API
  const url = "https://generativelanguage.googleapis.com/v1beta/interactions";

  const payload = {
    model: "gemini-3.8-flash",
    input: prompt
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-goog-api-key": apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    ss.toast("Analyzing data and generating commentary...", "AI Analysis", 5);
    
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (statusCode !== 200) {
      Logger.log("API Error: " + responseText);
      SpreadsheetApp.getUi().alert(`Gemini API Error (HTTP ${statusCode}):\n${responseText}`);
      return;
    }

    const json = JSON.parse(responseText);
    
    // Extract commentary from Interactions API response structure
    let commentary = "";
    if (json.output_text) {
      commentary = json.output_text;
    } else if (json.steps && json.steps.length > 0) {
      for (const step of json.steps) {
        if (step.content && Array.isArray(step.content)) {
          for (const item of step.content) {
            if (item.text) {
              commentary += item.text + "\n";
            }
          }
        }
      }
    } else if (json.candidates && json.candidates[0]?.content?.parts?.[0]?.text) {
      commentary = json.candidates[0].content.parts[0].text;
    }

    if (!commentary) {
      SpreadsheetApp.getUi().alert("Received empty response from API: " + responseText);
      return;
    }

    // 5. Output to Column E
    sheet.getRange("E1").setValue("AI Executive Commentary")
      .setFontWeight("bold")
      .setFontSize(12);

    const cell = sheet.getRange("E3");
    cell.setValue(commentary.trim());
    cell.setWrap(true);
    cell.setVerticalAlignment("top");
    
    sheet.setColumnWidth(5, 550);
    ss.toast("Executive commentary generated in E3!", "Success", 5);

  } catch (err) {
    Logger.log("Execution error: " + err);
    SpreadsheetApp.getUi().alert("Execution error: " + err.message);
  }
}