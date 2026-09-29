/**
 * PART 1: Sends the Executive Summary & Gemini Commentary via an HTML email briefing.
 */
function sendExecutiveBriefingEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("proj3_exec_summary_digest");

  if (!sheet) {
    SpreadsheetApp.getUi().alert("Sheet 'proj3_exec_summary_digest' not found!");
    return;
  }

  // 1. Gather Data
  const totalRevenue = sheet.getRange("B3").getDisplayValue();
  const netProfit    = sheet.getRange("B5").getDisplayValue();
  const margin       = sheet.getRange("B6").getDisplayValue();
  const commentary   = sheet.getRange("E3").getDisplayValue();

  if (!commentary || commentary.trim() === "") {
    SpreadsheetApp.getUi().alert("Please generate the AI Commentary in E3 before sending the email!");
    return;
  }

  // 2. Ask user for recipient email (or defaults to the user's current email)
  const recipientEmail = Session.getActiveUser().getEmail();

  // Convert commentary markdown line breaks to HTML
  const formattedCommentary = commentary
    .replace(/\n\n/g, "<br><br>")
    .replace(/\n/g, "<br>")
    .replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");

  // 3. Assemble clean HTML template
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 650px; margin: auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px;">
      <h2 style="color: #1a73e8; margin-bottom: 5px;">📊 Executive Performance Briefing</h2>
      <p style="color: #5f6368; font-size: 13px; margin-top: 0;">Automated Digest & AI Analysis</p>
      <hr style="border: none; border-top: 1px solid #eee; margin: 15px 0;">
      
      <!-- Key Metric Cards -->
      <table style="width: 100%; text-align: center; margin-bottom: 20px; border-collapse: collapse;">
        <tr>
          <td style="background: #f8f9fa; padding: 12px; border-radius: 6px; width: 32%;">
            <span style="font-size: 12px; color: #5f6368;">Total Revenue</span><br>
            <strong style="font-size: 18px; color: #202124;">${totalRevenue}</strong>
          </td>
          <td style="width: 2%;"></td>
          <td style="background: #f8f9fa; padding: 12px; border-radius: 6px; width: 32%;">
            <span style="font-size: 12px; color: #5f6368;">Net Profit</span><br>
            <strong style="font-size: 18px; color: #1e8e3e;">${netProfit}</strong>
          </td>
          <td style="width: 2%;"></td>
          <td style="background: #f8f9fa; padding: 12px; border-radius: 6px; width: 32%;">
            <span style="font-size: 12px; color: #5f6368;">Profit Margin</span><br>
            <strong style="font-size: 18px; color: #1a73e8;">${margin}</strong>
          </td>
        </tr>
      </table>

      <!-- AI Commentary Section -->
      <h3 style="color: #202124; margin-bottom: 8px;">🤖 AI Strategic Commentary</h3>
      <div style="background-color: #f1f3f4; padding: 15px; border-left: 4px solid #1a73e8; font-size: 14px; line-height: 1.6; color: #3c4043; border-radius: 4px;">
        ${formattedCommentary}
      </div>

      <hr style="border: none; border-top: 1px solid #eee; margin: 20px 0;">
      <p style="font-size: 11px; color: #80868b; text-align: center;">Generated automatically via Google Apps Script & Gemini API</p>
    </div>
  `;

  // 4. Dispatch Email
  GmailApp.sendEmail(recipientEmail, "Executive Financial Briefing & AI Analysis", "", {
    htmlBody: htmlBody
  });

  SpreadsheetApp.getUi().alert(`Briefing successfully sent to ${recipientEmail}!`);
}
